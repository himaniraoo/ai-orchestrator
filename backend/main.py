import os
import uuid
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware # allows browser frontend to call backend
from fastapi.responses import FileResponse # lets API download files
from pydantic import BaseModel # validates the incoming json
from dotenv import load_dotenv

from tools.physician_data import get_physician_data # imports the filteirng engine

load_dotenv() # activates the .env

app = FastAPI(title="DocNexus API", version="1.0.0")

# -------------------------------------------------------------------
# CORS — React dev server runs on 5173
# -------------------------------------------------------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_credentials=True,
    allow_methods=["*"], # get , post , put etc.
    allow_headers=["*"],
)

ARTIFACTS_DIR = Path(__file__).parent / "artifacts"
ARTIFACTS_DIR.mkdir(exist_ok=True)


# -------------------------------------------------------------------
# Request / Response models
# -------------------------------------------------------------------

class Preferences(BaseModel): # defines the expected JSON Schema
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
def health_check():  # used to verify if the server is alive 
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


@app.get("/artifacts/{artifact_id}")  # artifcat download routine
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
    Orchestrator will be wired in here in the next step.
    For now returns a stub so we can confirm the endpoint works.
    """
    return {
        "status": "stub", # will be updated lateron
        "message": "Orchestrator not wired yet",
        "received_query": request.query,
        "received_preferences": request.preferences.model_dump(),
    }