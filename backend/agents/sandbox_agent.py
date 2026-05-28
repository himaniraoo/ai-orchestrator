import os
import json
import base64
from google import genai
from google.genai import types
from dotenv import load_dotenv
from e2b_code_interpreter import Sandbox

load_dotenv()

client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

# -------------------------------------------------------------------
# SYSTEM PROMPT — code generation only
# LLM generates Python code as a plain string, nothing else.
# We execute it separately in E2B.
# -------------------------------------------------------------------

_CODE_GEN_SYSTEM_PROMPT = """
You are a Python code generation agent for DocNexus, a pharmaceutical intelligence platform.

## YOUR JOB
Generate clean, executable Python code to analyze physician data and produce the requested output.

## DATASET FORMAT
The dataset is available as a CSV file at the path: /home/user/dataset.csv
It has these columns:
id, npi, firstName, lastName, specialty, affiliation, city, state,
totalNSCLCClaims, volumeTier, boardCertified

## STRICT RULES
1. Return ONLY the Python code — no explanation, no markdown, no backticks
2. Always import what you use — pandas, matplotlib, etc are available
3. If producing a chart, save it to /home/user/chart.png using plt.savefig()
4. Always call plt.tight_layout() before saving
5. Print a brief text summary of findings to stdout
6. Do not use plt.show() — only plt.savefig()
7. Use matplotlib with a clean non-interactive backend: import matplotlib; matplotlib.use('Agg')
8. Handle edge cases — empty dataframe, missing columns, zero division
9. Make charts look professional — use clear labels, title, and colors
10. The code must be self-contained and run without any arguments
"""


# -------------------------------------------------------------------
# LLM CALL — generate Python code for the analysis goal
# -------------------------------------------------------------------

def _generate_code(code_goal: str, dataset_summary: str, error_context: str = "") -> str:
    """
    Generates Python code for the given analysis goal.
    If error_context is provided, it means we're self-correcting after a failure.
    """

    # Self-correction prompt is different — we tell it what went wrong
    if error_context:
        prompt = f"""
The following code failed with this error:
{error_context}

Fix the code so it runs correctly. The goal is still:
{code_goal}

Dataset info:
{dataset_summary}

Return only the fixed Python code.
"""
    else:
        prompt = f"""
Analysis goal: {code_goal}

Dataset info:
{dataset_summary}

Generate Python code to accomplish this goal.
"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=_CODE_GEN_SYSTEM_PROMPT,
        ),
    )

    code = response.text.strip()

    # Strip markdown code fences if model wraps in ```python ... ```
    if code.startswith("```"):
        lines = code.split("\n")
        code = "\n".join(lines[1:-1])

    return code.strip()


# -------------------------------------------------------------------
# E2B EXECUTION
# -------------------------------------------------------------------

def _execute_in_sandbox(code: str, dataset_csv: str) -> dict:
    from e2b_code_interpreter import Sandbox

    with Sandbox() as sandbox:   # <-- remove api_key=... entirely
        sandbox.files.write("/home/user/dataset.csv", dataset_csv)
        execution = sandbox.run_code(code)

        stdout = "\n".join(execution.logs.stdout) if execution.logs.stdout else ""
        stderr = "\n".join(execution.logs.stderr) if execution.logs.stderr else ""

        if execution.error:
            return {
                "success": False,
                "stdout": stdout,
                "stderr": execution.error.value or stderr,
                "chart_base64": None,
            }

        chart_base64 = None
        try:
            chart_bytes = sandbox.files.read("/home/user/chart.png", format="bytes")
            chart_base64 = base64.b64encode(chart_bytes).decode("utf-8")
        except Exception:
            pass

        return {
            "success": True,
            "stdout": stdout,
            "stderr": stderr,
            "chart_base64": chart_base64,
        }


# -------------------------------------------------------------------
# DATASET HELPERS
# -------------------------------------------------------------------

def _physicians_to_csv(physicians: list[dict]) -> str:
    """Convert physician list to CSV string for sandbox upload."""
    if not physicians:
        return ""

    headers = [
        "id", "npi", "firstName", "lastName", "specialty",
        "affiliation", "city", "state", "totalNSCLCClaims",
        "volumeTier", "boardCertified"
    ]

    lines = [",".join(headers)]
    for p in physicians:
        row = [
            str(p.get("id", "")),
            str(p.get("npi", "")),
            str(p.get("firstName", "")),
            str(p.get("lastName", "")),
            str(p.get("specialty", "")),
            f'"{p.get("affiliation", "")}"',  # quote affiliation — may contain commas
            str(p.get("city", "")),
            str(p.get("state", "")),
            str(p.get("totalNSCLCClaims", 0)),
            str(p.get("volumeTier", "")),
            str(p.get("boardCertified", False)),
        ]
        lines.append(",".join(row))

    return "\n".join(lines)


def _build_dataset_summary(physicians: list[dict]) -> str:
    """Build a compact description of the dataset for the code gen prompt."""
    if not physicians:
        return "Empty dataset"

    states = sorted(set(p["state"] for p in physicians))
    specialties = sorted(set(p["specialty"] for p in physicians))
    volume_tiers = sorted(set(p["volumeTier"] for p in physicians))

    return (
        f"Total physicians: {len(physicians)}\n"
        f"States: {', '.join(states)}\n"
        f"Specialties: {', '.join(specialties)}\n"
        f"Volume tiers: {', '.join(volume_tiers)}\n"
        f"Columns: id, npi, firstName, lastName, specialty, affiliation, "
        f"city, state, totalNSCLCClaims, volumeTier, boardCertified\n"
        f"\n"
        f"IMPORTANT:\n"
        f"- volumeTier values are lowercase\n"
        f"- 'high' and 'very_high' should both be treated as high-volume physicians\n"
    )


# -------------------------------------------------------------------
# MAIN ENTRY POINT
# -------------------------------------------------------------------

def run_sandbox_agent(
    code_goal: str,
    dataset: list[dict],
    chart_type: str = "",
) -> dict:
    """
    Generates Python code for the analysis goal and executes it in E2B.
    Self-corrects once on failure before giving up.
    Returns code, stdout, chart (base64), and execution status.
    """

    if not dataset:
        return {"error": "No dataset provided to Sandbox agent"}

    dataset_csv = _physicians_to_csv(dataset)
    dataset_summary = _build_dataset_summary(dataset)

    # Enrich the goal with chart type hint if provided
    enriched_goal = code_goal
    if chart_type:
        enriched_goal += f" Use a {chart_type} chart."

    # --- Attempt 1 — generate and execute code ---
    code = _generate_code(enriched_goal, dataset_summary)

    result = _execute_in_sandbox(code, dataset_csv)

    # --- Self-correction — if attempt 1 failed, try once more ---
    if not result["success"]:
        print(f"[Sandbox Agent] Attempt 1 failed: {result['stderr']}")
        print("[Sandbox Agent] Self-correcting — regenerating code...")

        corrected_code = _generate_code(
            code_goal=enriched_goal,
            dataset_summary=dataset_summary,
            error_context=result["stderr"],
        )

        result = _execute_in_sandbox(corrected_code, dataset_csv)

        if not result["success"]:
            # Both attempts failed — return what we have
            print(f"[Sandbox Agent] Attempt 2 also failed: {result['stderr']}")
            return {
                "status": "failed",
                "code": corrected_code,
                "output": result["stdout"],
                "error": result["stderr"],
                "chart_base64": None,
                "attempts": 2,
            }

        # Self-correction succeeded
        return {
            "status": "success_after_correction",
            "code": corrected_code,
            "output": result["stdout"],
            "chart_base64": result["chart_base64"],
            "attempts": 2,
        }

    # Attempt 1 succeeded
    return {
        "status": "success",
        "code": code,
        "output": result["stdout"],
        "chart_base64": result["chart_base64"],
        "attempts": 1,
    }