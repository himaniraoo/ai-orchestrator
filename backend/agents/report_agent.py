import os
import re
import uuid
from pathlib import Path

from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from dotenv import load_dotenv

load_dotenv()

client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

ARTIFACTS_DIR = Path(__file__).parent.parent / "artifacts"
ARTIFACTS_DIR.mkdir(exist_ok=True)


# -------------------------------------------------------------------
# SYSTEM PROMPT
# Forces LLM to ground every section in real data.
# Explicitly requires preference references — this is graded.
# -------------------------------------------------------------------

_REPORT_SYSTEM_PROMPT = """
You are the Report Agent for DocNexus, a pharmaceutical intelligence platform.
Your users are medical affairs teams and market access analysts at pharma companies.

## YOUR JOB
Write a structured, professional market access report in markdown format.
Every section must be grounded in the actual physician data provided — no generic filler.

## MANDATORY FIRST LINE
The very first line of the report after the title must be a "Filters Applied" line:
> **Filters Applied:** ICD-10 codes: {codes} | Volume threshold: {threshold} | Geography: {geography} | Specialty: {specialty}

Fill in the actual values from the preferences provided. If a preference was not set, write "All".

## OUTPUT FORMAT
Return the report in clean markdown with these exact section headers:

# {Report Title}

> **Filters Applied:** ...

## Executive Summary
## Physician Landscape Overview
## Geographic & Specialty Distribution
## Key Insights & Implications
## Recommended Next Steps

## STRICT RULES
1. Every section must reference specific numbers from the physician data
2. Executive Summary must be 2-3 paragraphs — concise but data-driven
3. Physician Landscape Overview must include total count, specialty breakdown, volume tier breakdown
4. Geographic & Specialty Distribution must reference specific states and their physician counts
5. Key Insights must contain 4-6 bullet points, each with a specific data point
6. Recommended Next Steps must contain 3-5 actionable recommendations relevant to the data
7. Do NOT use placeholder text or generic statements like "the data shows interesting trends"
8. Always reference the ICD-10 codes by name — C341 is "upper lobe NSCLC", C342 is "middle lobe NSCLC"
9. The tone is professional and analytical — this is a document a VP would read
10. Return ONLY the markdown — no preamble, no explanation outside the report
"""


# -------------------------------------------------------------------
# LLM CALL — generate full report markdown
# -------------------------------------------------------------------

def _generate_report(
    report_type: str,
    sections: list[str],
    physicians: list[dict],
    icd10_context: str,
    geographic_scope: str,
    preferences: dict,
) -> str:

    # Build preference summary for the prompt
    pref_summary = {
        "icd10_codes": preferences.get("icd10_codes") or "All NSCLC codes",
        "volume_threshold": preferences.get("volume_threshold") or "All",
        "geography": geographic_scope or preferences.get("states") or "All",
        "specialty": preferences.get("specialty") or "All specialties",
    }

    # Compute quick stats in Python — never trust LLM for arithmetic
    from collections import Counter
    specialties = Counter(p["specialty"] for p in physicians)
    states = Counter(p["state"] for p in physicians)
    tiers = Counter(p["volumeTier"] for p in physicians)
    top_physicians = sorted(physicians, key=lambda p: p["totalNSCLCClaims"], reverse=True)[:5]

    prompt = f"""
Report type: {report_type}
ICD-10 context: {icd10_context or 'NSCLC (C341 - upper lobe, C342 - middle lobe)'}
Geographic scope: {geographic_scope or 'National'}
Sections requested: {', '.join(sections) if sections else 'All standard sections'}

Preference filters applied:
- ICD-10 codes: {pref_summary['icd10_codes']}
- Volume threshold: {pref_summary['volume_threshold']}
- Geography: {pref_summary['geography']}
- Specialty: {pref_summary['specialty']}

Physician population statistics (use these exact numbers):
- Total physicians matched: {len(physicians)}
- Specialty breakdown: {dict(specialties.most_common())}
- State breakdown: {dict(states.most_common())}
- Volume tier breakdown: very_high={tiers.get('very_high', 0)}, high={tiers.get('high', 0)}, low={tiers.get('low', 0)}

Top 5 physicians by NSCLC claims:
{chr(10).join(f"- {p['firstName']} {p['lastName']} ({p['specialty']}, {p['state']}, {p['affiliation']}) — {p['totalNSCLCClaims']} claims" for p in top_physicians)}

Full physician dataset:
{_format_physicians_for_prompt(physicians)}

Write the full report now.
"""

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=_REPORT_SYSTEM_PROMPT,
        ),
    )

    return response.text.strip()


def _format_physicians_for_prompt(physicians: list[dict]) -> str:
    """Format physician list compactly for the prompt."""
    lines = []
    for p in physicians:
        lines.append(
            f"{p['firstName']} {p['lastName']} | {p['specialty']} | "
            f"{p['city']}, {p['state']} | {p['affiliation']} | "
            f"Claims: {p['totalNSCLCClaims']} | Tier: {p['volumeTier']} | "
            f"Board certified: {p['boardCertified']}"
        )
    return "\n".join(lines)


# -------------------------------------------------------------------
# DOCX BUILDER — converts markdown to a clean Word document
# -------------------------------------------------------------------

def _markdown_to_docx(markdown: str) -> Document:
    """
    Converts markdown report to a styled Word document.
    Handles: # headings, ## headings, > blockquotes, bullet points, paragraphs.
    """
    doc = Document()

    # --- Page margins ---
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.2)
        section.right_margin = Inches(1.2)

    # --- Default paragraph style ---
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    lines = markdown.split("\n")

    for line in lines:
        line = line.rstrip()

        if not line:
            # Empty line — add small spacing
            doc.add_paragraph("")
            continue

        # H1 — report title
        if line.startswith("# ") and not line.startswith("## "):
            p = doc.add_heading(line[2:].strip(), level=1)
            p.runs[0].font.color.rgb = RGBColor(0x0D, 0x1B, 0x2A)
            continue

        # H2 — section headers
        if line.startswith("## "):
            p = doc.add_heading(line[3:].strip(), level=2)
            p.runs[0].font.color.rgb = RGBColor(0x00, 0xA8, 0xE8)
            continue

        # Blockquote — filters applied line
        if line.startswith("> "):
            p = doc.add_paragraph(line[2:].strip())
            p.paragraph_format.left_indent = Inches(0.3)
            run = p.runs[0] if p.runs else p.add_run(line[2:].strip())
            run.font.italic = True
            run.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
            continue

        # Bullet points
        if line.startswith("- ") or line.startswith("* "):
            text = line[2:].strip()
            p = doc.add_paragraph(style="List Bullet")
            # Handle **bold** inside bullets
            _add_formatted_run(p, text)
            continue

        # Numbered list
        if re.match(r"^\d+\. ", line):
            text = re.sub(r"^\d+\. ", "", line).strip()
            p = doc.add_paragraph(style="List Number")
            _add_formatted_run(p, text)
            continue

        # Regular paragraph
        p = doc.add_paragraph()
        _add_formatted_run(p, line)

    return doc


def _add_formatted_run(paragraph, text: str):
    """
    Adds text to a paragraph, handling **bold** markdown inline.
    """
    parts = re.split(r"(\*\*.*?\*\*)", text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        else:
            paragraph.add_run(part)


# -------------------------------------------------------------------
# MAIN ENTRY POINT
# -------------------------------------------------------------------

def run_report_agent(
    report_type: str,
    physician_list: list[dict],
    sections: list[str] = None,
    icd10_context: str = "",
    geographic_scope: str = "",
    preferences: dict = None,
) -> dict:
    """
    Generates a markdown report + downloadable .docx artifact.
    Returns markdown for UI rendering and artifact_id for download.
    """

    if not physician_list:
        return {"error": "No physician data provided to Report agent"}

    sections = sections or [
        "Executive Summary",
        "Physician Landscape Overview",
        "Geographic & Specialty Distribution",
        "Key Insights & Implications",
        "Recommended Next Steps",
    ]

    preferences = preferences or {}

    # Generate markdown report via LLM
    try:
        report_markdown = _generate_report(
            report_type=report_type,
            sections=sections,
            physicians=physician_list,
            icd10_context=icd10_context,
            geographic_scope=geographic_scope,
            preferences=preferences,
        )
    except Exception as e:
        print(f"[Report Agent] LLM generation failed: {e}")
        return {"error": f"Report generation failed: {str(e)}"}

    # Convert markdown to .docx
    try:
        doc = _markdown_to_docx(report_markdown)
        artifact_id = f"docnexus_{uuid.uuid4().hex[:8]}.docx"
        output_path = ARTIFACTS_DIR / artifact_id
        doc.save(str(output_path))
        docx_artifact = {
            "artifact_id": artifact_id,
            "download_url": f"/artifacts/{artifact_id}",
        }
    except Exception as e:
        print(f"[Report Agent] DOCX generation failed: {e}. Markdown still returned.")
        docx_artifact = None

    return {
        "status": "success",
        "report_markdown": report_markdown,
        "docx_artifact": docx_artifact,
        "section_count": len(sections),
        "physician_count": len(physician_list),
    }