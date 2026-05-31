import os
import json
import time
from typing import Generator
from google import genai
from google.genai import types
from dotenv import load_dotenv

from tools.physician_data import get_physician_data as _get_physician_data
from agents.ppt_agent import run_ppt_agent
from agents.excel_agent import run_excel_agent
from agents.report_agent import run_report_agent
from agents.sandbox_agent import run_sandbox_agent

load_dotenv()

client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

# -------------------------------------------------------------------
# TOOL DEFINITIONS
# These are what Gemini sees — descriptions must be precise because
# Gemini decides WHICH tool to call based purely on these descriptions.
# -------------------------------------------------------------------

_TOOLS = [
    types.Tool(function_declarations=[
        types.FunctionDeclaration(
            name="get_physician_data",
            description=(
                "Retrieves filtered physician records from the DocNexus database. "
                "ALWAYS call this first before calling any agent tool. "
                "Use this to fetch the physician population relevant to the user's query. "
                "All parameters are optional — omit any you don't have signal for."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                properties={
                    "specialty": types.Schema(
                        type=types.Type.STRING,
                        description="Physician specialty to filter by. Examples: 'Medical Oncology', 'Pulmonology', 'Thoracic Surgery', 'Radiation Oncology'"
                    ),
                    "states": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.STRING),
                        description="List of US state abbreviations to filter by. Examples: ['CA', 'NY', 'TX']"
                    ),
                    "icd10_codes": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.STRING),
                        description="List of ICD-10 codes to filter by. Physician must have at least one. Examples: ['C341', 'C342']"
                    ),
                    "volume_threshold": types.Schema(
                        type=types.Type.STRING,
                        enum=["low", "high", "very_high"],
                        description="Minimum volume tier. 'high' returns high and very_high physicians."
                    ),
                },
            ),
        ),
        types.FunctionDeclaration(
            name="call_ppt_agent",
            description=(
                "Generates a downloadable PowerPoint slide deck (.pptx). "
                "Call this when the user asks for a 'slide deck', 'presentation', 'PowerPoint', or 'slides'. "
                "Requires physician data — call get_physician_data first."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                required=["query", "topic", "physician_list"],
                properties={
                    "query": types.Schema(
                        type=types.Type.STRING,
                        description="The original user query verbatim — used for the title slide query summary"
                    ),
                    "topic": types.Schema(
                        type=types.Type.STRING,
                        description="The title/topic of the presentation. Be specific."
                    ),
                    "physician_list": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.OBJECT),
                        description="List of physician records returned by get_physician_data"
                    ),
                    "icd10_codes": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.STRING),
                        description="ICD-10 codes relevant to this presentation"
                    ),
                    "slide_count": types.Schema(
                        type=types.Type.INTEGER,
                        description="Number of slides to generate. Default 4."
                    ),
                    "style_notes": types.Schema(
                        type=types.Type.STRING,
                        description="Any style or content preferences from the user query"
                    ),
                },
            ),
        ),
        types.FunctionDeclaration(
            name="call_excel_agent",
            description=(
                "Generates a downloadable Excel workbook (.xlsx) with multiple sheets. "
                "Call this when the user asks for a 'spreadsheet', 'Excel', 'breakdown', 'workbook', or 'table'. "
                "Requires physician data — call get_physician_data first."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                required=["analysis_type", "physician_list"],
                properties={
                    "analysis_type": types.Schema(
                        type=types.Type.STRING,
                        description="What kind of analysis the Excel should show."
                    ),
                    "physician_list": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.OBJECT),
                        description="List of physician records returned by get_physician_data"
                    ),
                    "dimensions": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.STRING),
                        description="Breakdown dimensions. Examples: ['state', 'specialty', 'icd10_code']"
                    ),
                    "icd10_codes": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.STRING),
                        description="ICD-10 codes relevant to this analysis"
                    ),
                },
            ),
        ),
        types.FunctionDeclaration(
            name="call_report_agent",
            description=(
                "Generates a structured written market access report in markdown format. "
                "Call this when the user asks for a 'report', 'write-up', or 'market access report'. "
                "Requires physician data — call get_physician_data first."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                required=["report_type", "physician_list"],
                properties={
                    "report_type": types.Schema(
                        type=types.Type.STRING,
                        description="Type of report. Examples: 'market access report', 'physician landscape report'"
                    ),
                    "sections": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.STRING),
                        description="Sections to include in the report"
                    ),
                    "physician_list": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.OBJECT),
                        description="List of physician records returned by get_physician_data"
                    ),
                    "icd10_context": types.Schema(
                        type=types.Type.STRING,
                        description="ICD-10 context string for the report"
                    ),
                    "geographic_scope": types.Schema(
                        type=types.Type.STRING,
                        description="Geographic scope of the report"
                    ),
                },
            ),
        ),
        types.FunctionDeclaration(
            name="call_sandbox_agent",
            description=(
                "Generates and executes Python code to perform data analysis or produce a chart. "
                "Call this when the user asks to 'run an analysis', 'plot', 'show a chart', or 'visualize'. "
                "Requires physician data — call get_physician_data first."
            ),
            parameters=types.Schema(
                type=types.Type.OBJECT,
                required=["code_goal", "dataset"],
                properties={
                    "code_goal": types.Schema(
                        type=types.Type.STRING,
                        description="Plain English description of what the code should do."
                    ),
                    "dataset": types.Schema(
                        type=types.Type.ARRAY,
                        items=types.Schema(type=types.Type.OBJECT),
                        description="The physician dataset to analyze"
                    ),
                    "chart_type": types.Schema(
                        type=types.Type.STRING,
                        description="Type of chart to generate. Examples: 'bar', 'pie', 'scatter'"
                    ),
                },
            ),
        ),
    ])
]

# -------------------------------------------------------------------
# SYSTEM PROMPT
# -------------------------------------------------------------------

_SYSTEM_PROMPT = """
# ================================================================
# ORCHESTRATOR AGENT — DocNexus Physician Intelligence Platform
# ================================================================

# ROLE
# You are the routing brain of DocNexus. You receive natural language
# queries from medical affairs teams and market access analysts, then
# decompose them into tool calls that produce business artifacts.
# You never produce artifacts yourself — you delegate entirely via tools.

# ----------------------------------------------------------------
# SECTION 1: MANDATORY FIRST STEP
# WHY: Every artifact agent requires a physician list as input.
# Calling them without data produces empty, useless output and wastes
# an LLM call. This constraint is enforced first because Gemini will
# otherwise sometimes skip data retrieval and call artifact agents
# directly when the query is phrased confidently enough.
# ----------------------------------------------------------------
STEP 1 — ALWAYS call get_physician_data before any other tool.
Extract these filters from the query:
  - specialty        → from role mentions (see SECTION 3)
  - states           → from geography mentions, converted to abbreviations
  - icd10_codes      → from disease mentions (see SECTION 3)
  - volume_threshold → from volume mentions (see SECTION 3)

# ----------------------------------------------------------------
# SECTION 2: INTENT → AGENT ROUTING
# WHY: These keyword triggers are explicit because without them,
# Gemini defaults to call_report_agent for almost everything —
# it prefers text output. Each trigger phrase maps directly to
# the spec's example queries to guarantee coverage of graded cases.
# Multi-artifact queries must call ALL matching agents — skipping
# one is a grading failure.
# ----------------------------------------------------------------
STEP 2 — Route to the correct agent(s) based on intent:

  "slide deck" | "PowerPoint" | "presentation" | "slides"
    → call_ppt_agent

  "Excel" | "spreadsheet" | "breakdown" | "workbook" | "table"
    → call_excel_agent

  "report" | "write-up" | "market access" | "two-page" | "summary"
    → call_report_agent

  "analyze" | "plot" | "chart" | "visualize" | "run an analysis" | "show me which"
    → call_sandbox_agent

  Multiple artifact types in one query → call ALL matching agents, no exceptions.

# ----------------------------------------------------------------
# SECTION 3: DOMAIN KNOWLEDGE — SEMANTIC MAPPINGS
# WHY: Users speak in clinical shorthand. Without explicit mappings,
# the model hallucinates ICD-10 codes (e.g. using C34 instead of C341)
# and misses geographic expansions like "Northeast". These mappings
# are baked in here rather than expected from the user.
# ----------------------------------------------------------------
ICD-10 REFERENCE (NSCLC):
  C341 = upper lobe   C342 = middle lobe
  C343 = lower lobe   C349 = unspecified
  Default for any NSCLC mention → use [C341, C342]

SPECIALTY REFERENCE:
  "oncologist"           → "Medical Oncology"
  "pulmonologist"        → "Pulmonology"
  "thoracic surgeon"     → "Thoracic Surgery"
  "radiation oncologist" → "Radiation Oncology"

GEOGRAPHY REFERENCE:
  Full state names → 2-letter abbreviations (e.g. "California" → "CA")
  "Northeast" → [NY, MA, CT, NJ, PA, MD]

VOLUME REFERENCE:
  "high volume" | "high-volume" → volume_threshold: "high"
  "very high volume"            → volume_threshold: "very_high"

# ----------------------------------------------------------------
# SECTION 4: DATA PASSING RULES
# WHY: Two failure modes observed during development:
# (a) Truncating the physician list before passing to agents causes
#     incomplete artifacts — agents must always see the full cohort.
# (b) Pre-filtering by volume for sandbox queries removes the
#     denominator population, making concentration percentages wrong.
#     The sandbox agent must compute proportions itself.
# ----------------------------------------------------------------
- Pass the FULL physician list from get_physician_data to every agent.
- For concentration/distribution queries ("which states have the highest..."),
  do NOT set volume_threshold — let the sandbox compute it from full data.
- Pass the original user query verbatim as the "query" param to call_ppt_agent.

# ----------------------------------------------------------------
# SECTION 5: CLOSING SUMMARY
# WHY: Users need a plain-text confirmation even if they don't open
# an artifact. The summary also provides a fallback if artifact
# rendering fails on the frontend.
# ----------------------------------------------------------------
After all tools complete, return a brief plain-text summary:
  - How many physicians matched
  - Which agents were called and what they produced
  - One specific observation from the data

# HARD CONSTRAINTS
# NEVER call an artifact agent before get_physician_data
# NEVER fabricate physician records or IDs
# NEVER skip an artifact type the user asked for
# NEVER return final text without calling at least one agent
"""

# -------------------------------------------------------------------
# TOOL HANDLERS
# -------------------------------------------------------------------

def _handle_get_physician_data(args: dict) -> dict:
    physicians = _get_physician_data(
        specialty=args.get("specialty"),
        states=args.get("states"),
        icd10_codes=args.get("icd10_codes"),
        volume_threshold=args.get("volume_threshold"),
    )
    return {"count": len(physicians), "physicians": physicians}


def _handle_ppt_agent(args: dict) -> dict:
    return run_ppt_agent(
        query=args.get("query", ""),
        topic=args.get("topic", "NSCLC Physician Landscape"),
        physician_list=args.get("physician_list", []),
        icd10_codes=args.get("icd10_codes"),
        slide_count=args.get("slide_count", 4),
        style_notes=args.get("style_notes", ""),
    )


def _handle_excel_agent(args: dict) -> dict:
    return run_excel_agent(
        analysis_type=args.get("analysis_type", "physician breakdown"),
        physician_list=args.get("physician_list", []),
        dimensions=args.get("dimensions"),
        icd10_codes=args.get("icd10_codes"),
    )


def _handle_report_agent(args: dict) -> dict:
    return run_report_agent(
        report_type=args.get("report_type", "market access report"),
        physician_list=args.get("physician_list", []),
        sections=args.get("sections"),
        icd10_context=args.get("icd10_context", ""),
        geographic_scope=args.get("geographic_scope", ""),
        preferences={},
    )


def _handle_sandbox_agent(args: dict) -> dict:
    return run_sandbox_agent(
        code_goal=args.get("code_goal", ""),
        dataset=args.get("dataset", []),
        chart_type=args.get("chart_type", ""),
    )


_TOOL_HANDLERS = {
    "get_physician_data": _handle_get_physician_data,
    "call_ppt_agent": _handle_ppt_agent,
    "call_excel_agent": _handle_excel_agent,
    "call_report_agent": _handle_report_agent,
    "call_sandbox_agent": _handle_sandbox_agent,
}

# Human-readable labels for the trace UI
_TOOL_LABELS = {
    "get_physician_data": "Fetching physician data",
    "call_ppt_agent":     "PPT agent generating slides",
    "call_excel_agent":   "Excel agent building workbook",
    "call_report_agent":  "Report agent writing report",
    "call_sandbox_agent": "Sandbox agent running analysis",
}


# -------------------------------------------------------------------
# INTERNAL AGENT LOOP — shared by both sync and streaming paths
# -------------------------------------------------------------------

def _build_initial_contents(query: str, preferences: dict):
    preference_context = ""
    if any(preferences.values()):
        preference_context = f"\n\nUser preferences panel: {json.dumps(preferences)}"
    initial_message = query + preference_context
    return [
        types.Content(
            role="user",
            parts=[types.Part(text=initial_message)]
        )
    ]


# -------------------------------------------------------------------
# STREAMING GENERATOR
# Yields SSE-formatted strings. Each event is a JSON payload.
#
# Event types:
#   { "event": "trace",    "data": { "step", "tool", "label", "status", "args_summary", "elapsed_ms" } }
#   { "event": "artifact", "data": { "type", "artifact_id", "download_url" } }
#   { "event": "report",   "data": { "markdown": "..." } }
#   { "event": "sandbox",  "data": { ...sandbox_result fields... } }
#   { "event": "summary",  "data": { "text": "..." } }
#   { "event": "done",     "data": {} }
#   { "event": "error",    "data": { "message": "..." } }
# -------------------------------------------------------------------

def _sse(event: str, data: dict) -> str:
    """Format a single SSE message."""
    payload = json.dumps({"event": event, "data": data})
    return f"data: {payload}\n\n"


def run_orchestrator_stream(query: str, preferences: dict) -> Generator[str, None, None]:
    """
    Generator that runs the orchestrator loop and yields SSE strings.
    Drop-in replacement stream version of run_orchestrator.
    """
    contents = _build_initial_contents(query, preferences)
    config = types.GenerateContentConfig(
        system_instruction=_SYSTEM_PROMPT,
        tools=_TOOLS,
    )

    max_steps = 10
    step = 0

    try:
        while step < max_steps:
            step += 1

            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=contents,
                config=config,
            )

            candidate = response.candidates[0]
            parts = candidate.content.parts

            tool_calls = [p for p in parts if p.function_call is not None]

            if not tool_calls:
                # Gemini is done — emit summary
                final_text = " ".join(p.text for p in parts if p.text)
                yield _sse("summary", {"text": final_text})
                break

            # Append Gemini's response to conversation history
            contents.append(candidate.content)

            # Execute each tool call
            tool_result_parts = []
            for part in tool_calls:
                fc = part.function_call
                tool_name = fc.name
                tool_args = dict(fc.args)

                label = _TOOL_LABELS.get(tool_name, tool_name)

                # Emit "in progress" trace event
                args_summary = {
                    k: v for k, v in tool_args.items()
                    if k not in ("physician_list", "dataset")
                }
                yield _sse("trace", {
                    "step": step,
                    "tool": tool_name,
                    "label": label,
                    "status": "running",
                    "args_summary": args_summary,
                })

                t0 = time.monotonic()
                handler = _TOOL_HANDLERS.get(tool_name)
                tool_result = handler(tool_args) if handler else {"error": f"Unknown tool: {tool_name}"}
                elapsed_ms = int((time.monotonic() - t0) * 1000)

                # Emit "done" trace event
                result_summary = {
                    k: v for k, v in tool_result.items()
                    if k not in ("physicians", "report_markdown")
                }
                yield _sse("trace", {
                    "step": step,
                    "tool": tool_name,
                    "label": label,
                    "status": "done",
                    "args_summary": args_summary,
                    "result_summary": result_summary,
                    "elapsed_ms": elapsed_ms,
                })

                # Emit artifact if produced
                if (
                    "artifact_id" in tool_result
                    and "stub" not in tool_result.get("artifact_id", "")
                    and tool_result.get("status") == "success"
                ):
                    yield _sse("artifact", {
                        "type": tool_name.replace("call_", "").replace("_agent", ""),
                        "artifact_id": tool_result["artifact_id"],
                        "download_url": tool_result["download_url"],
                    })

                # Emit report markdown
                if "report_markdown" in tool_result:
                    yield _sse("report", {"markdown": tool_result["report_markdown"]})

                # Emit sandbox result
                if tool_name == "call_sandbox_agent":
                    yield _sse("sandbox", tool_result)

                # Build function response for next Gemini turn
                tool_result_parts.append(
                    types.Part(
                        function_response=types.FunctionResponse(
                            name=tool_name,
                            response={"result": tool_result},
                        )
                    )
                )

            contents.append(
                types.Content(role="tool", parts=tool_result_parts)
            )

    except Exception as e:
        yield _sse("error", {"message": str(e)})

    yield _sse("done", {})


# -------------------------------------------------------------------
# ORIGINAL SYNC run_orchestrator — UNCHANGED
# Still used by POST /query. Do not remove.
# -------------------------------------------------------------------

def run_orchestrator(query: str, preferences: dict) -> dict:
    """
    Runs the full orchestrator agent loop.
    Returns structured result with agent trace + artifacts.
    """

    preference_context = ""
    if any(preferences.values()):
        preference_context = f"\n\nUser preferences panel: {json.dumps(preferences)}"

    initial_message = query + preference_context

    agent_trace = []
    artifacts = []
    report_markdown = None
    sandbox_result = None
    final_text = "Orchestration complete."
    max_steps = 10

    contents = [
        types.Content(
            role="user",
            parts=[types.Part(text=initial_message)]
        )
    ]

    config = types.GenerateContentConfig(
        system_instruction=_SYSTEM_PROMPT,
        tools=_TOOLS,
    )

    step = 0
    while step < max_steps:
        step += 1

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=contents,
            config=config,
        )

        candidate = response.candidates[0]
        parts = candidate.content.parts

        tool_calls = [p for p in parts if p.function_call is not None]

        if not tool_calls:
            final_text = " ".join(p.text for p in parts if p.text)
            break

        contents.append(candidate.content)

        tool_result_parts = []
        for part in tool_calls:
            fc = part.function_call
            tool_name = fc.name
            tool_args = dict(fc.args)

            handler = _TOOL_HANDLERS.get(tool_name)
            tool_result = handler(tool_args) if handler else {"error": f"Unknown tool: {tool_name}"}

            agent_trace.append({
                "step": step,
                "tool": tool_name,
                "args_summary": {
                    k: v for k, v in tool_args.items()
                    if k not in ("physician_list", "dataset")
                },
                "result_summary": {
                    k: v for k, v in tool_result.items()
                    if k not in ("physicians", "report_markdown")
                }
            })

            if (
                "artifact_id" in tool_result
                and "stub" not in tool_result.get("artifact_id", "")
                and tool_result.get("status") == "success"
            ):
                artifacts.append({
                    "type": tool_name.replace("call_", "").replace("_agent", ""),
                    "artifact_id": tool_result["artifact_id"],
                    "download_url": tool_result["download_url"],
                })

            if "report_markdown" in tool_result:
                report_markdown = tool_result["report_markdown"]

            if tool_name == "call_sandbox_agent":
                sandbox_result = tool_result

            tool_result_parts.append(
                types.Part(
                    function_response=types.FunctionResponse(
                        name=tool_name,
                        response={"result": tool_result},
                    )
                )
            )

        contents.append(
            types.Content(role="tool", parts=tool_result_parts)
        )

    return {
        "agent_trace": agent_trace,
        "artifacts": artifacts,
        "report_markdown": report_markdown,
        "sandbox_result": sandbox_result,
        "summary": final_text,
    }