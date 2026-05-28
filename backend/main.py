import os
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from dotenv import load_dotenv

from tools.physician_data import get_physician_data
from agents.orchestrator import run_orchestrator

load_dotenv()

app = FastAPI(title="DocNexus API", version="1.0.0")

# -------------------------------------------------------------------
# CORS — React dev server runs on 5173
# -------------------------------------------------------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

ARTIFACTS_DIR = Path(__file__).parent / "artifacts"
ARTIFACTS_DIR.mkdir(exist_ok=True)


# -------------------------------------------------------------------
# Request / Response models
# -------------------------------------------------------------------
class Preferences(BaseModel):
    specialty: Optional[str] = None
    states: Optional[list[str]] = None
    icd10_codes: Optional[list[str]] = None
    volume_threshold: Optional[str] = None


class QueryRequest(BaseModel):
    query: str
    preferences: Preferences = Preferences()


# -------------------------------------------------------------------
# Routes
# -------------------------------------------------------------------
@app.get("/health")
def health_check():
    return {"status": "ok"}


@app.get("/physicians")
def list_physicians(
    specialty: Optional[str] = Query(None),
    states: Optional[list[str]] = Query(None),
    icd10_codes: Optional[list[str]] = Query(None),
    volume_threshold: Optional[str] = Query(None),
):
    """
    Returns filtered physician list.
    Example: /physicians?states=CA&states=NY&volume_threshold=high
    """
    results = get_physician_data(
        specialty=specialty,
        states=states,
        icd10_codes=icd10_codes,
        volume_threshold=volume_threshold,
    )
    return {"count": len(results), "physicians": results}


@app.get("/artifacts/{artifact_id}")
def download_artifact(artifact_id: str):
    """
    Serves a generated file (pptx / xlsx / docx) for download.
    artifact_id is the filename without path.
    """
    # Sanitize — prevent path traversal
    safe_name = Path(artifact_id).name
    file_path = ARTIFACTS_DIR / safe_name

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Artifact not found")

    return FileResponse(
        path=file_path,
        filename=safe_name,
        media_type="application/octet-stream",
    )


@app.post("/query")
async def run_query(request: QueryRequest):
    """
    Main endpoint — receives natural language query + preferences.
    Runs the full orchestrator agent loop and returns results.
    """
    result = run_orchestrator(
        query=request.query,
        preferences=request.preferences.model_dump(),
    )
    return result