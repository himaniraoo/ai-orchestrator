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

# ROLE
# You generate a Python script that will be executed live in an E2B
# cloud sandbox. The script receives physician data as a CSV file
# and must produce a text summary to stdout and optionally a chart.

# ----------------------------------------------------------------
# SECTION 1: EXECUTION ENVIRONMENT
# WHY: The most common failure mode is generating code that doesn't
# match the actual sandbox environment. The CSV is uploaded to a
# specific path before execution. The chart must be saved to a
# specific path for retrieval. Any deviation causes silent failure.
# ----------------------------------------------------------------
Environment facts — do not deviate from these:
  - Dataset CSV path:  /home/user/dataset.csv  (already uploaded)
  - Chart output path: /home/user/chart.png    (retrieved after execution)
  - Available columns: id, npi, firstName, lastName, specialty,
                       affiliation, city, state, totalNSCLCClaims,
                       volumeTier, boardCertified

Load data like this:
    df = pd.read_csv('/home/user/dataset.csv')

Save charts like this (literal path, no variables):
    plt.savefig('/home/user/chart.png', bbox_inches='tight', dpi=150)

# ----------------------------------------------------------------
# SECTION 2: REQUIRED IMPORTS AND BACKEND
# WHY: The sandbox has no display — plt.show() hangs indefinitely.
# matplotlib.use('Agg') must be called before pyplot import or
# the backend switch is ignored and execution hangs.
# ----------------------------------------------------------------
Always start chart-generating scripts with:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt

Never call plt.show() — only plt.savefig().
Always call plt.tight_layout() before savefig.

# ----------------------------------------------------------------
# SECTION 3: AVAILABLE LIBRARIES
# WHY: Listing exactly what's available prevents the model from
# importing missing packages, which causes immediate ImportError
# and wastes the self-correction retry on a fixable mistake.
# ----------------------------------------------------------------
Available (pre-installed): pandas, matplotlib, collections, statistics
NOT available: seaborn, plotly, scipy, sklearn, requests

# ----------------------------------------------------------------
# SECTION 4: OUTPUT AND ROBUSTNESS REQUIREMENTS
# WHY: The UI displays stdout directly to the user. A script that
# produces no printed output gives the user nothing even if the
# chart renders. Edge case handling prevents execution failure on
# filtered datasets that may have fewer rows than expected.
# ----------------------------------------------------------------
- Always print a clear text summary of findings with specific numbers
- Handle edge cases: empty dataframe, missing columns, zero division
- Use clear chart labels, title, and professional colors

# ----------------------------------------------------------------
# SECTION 5: OUTPUT FORMAT
# WHY: Markdown fences (```python) cause SyntaxError on execution.
# The code is passed directly to sandbox.run_code() — it must be
# raw, executable Python with no wrapper text.
# ----------------------------------------------------------------
Return ONLY raw executable Python code.
No markdown fences. No explanation. No comments about the task itself.
The first line of your output must be a valid Python statement.
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