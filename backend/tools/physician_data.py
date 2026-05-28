import json #read JSON file
from pathlib import Path #builds file path safely
from typing import Optional #optinally gives typing hints

# Load once at module import time
_DATA_PATH = Path(__file__).parent.parent / "data" / "physicians.json"

with open(_DATA_PATH, "r") as f:
    _PHYSICIANS: list[dict] = json.load(f) #make the JSON datastructure as python based

def get_physician_data( #default filtering engine - paramaters can be none / they are optional
    specialty: Optional[str] = None,
    states: Optional[list[str]] = None,
    icd10_codes: Optional[list[str]] = None,
    volume_threshold: Optional[str] = None,
) -> list[dict]:
    """
    Filter physicians from mock data based on provided criteria.
    All filters are optional — omitting one means 'no filter on that dimension'.
    Multiple filters are AND-ed together.
    icd10_codes filter = physician must have at least one of the codes (OR logic).
    """
    results = _PHYSICIANS #intially all datasets are included

    if specialty: #progressively narrow results
        specialty_lower = specialty.lower()
        results = [
            p for p in results
            if specialty_lower in p["specialty"].lower()
        ]

    if states:
        states_upper = [s.upper() for s in states]
        results = [
            p for p in results
            if p["state"].upper() in states_upper
        ]

    if icd10_codes:
        codes_upper = [c.upper() for c in icd10_codes]
        results = [
            p for p in results
            if any(code in p["icd10ClaimVolume"] for code in codes_upper)
        ]

    if volume_threshold:
        tier_rank = {"low": 1, "high": 2, "very_high": 3}
        min_rank = tier_rank.get(volume_threshold.lower(), 1)
        results = [
            p for p in results
            if tier_rank.get(p["volumeTier"], 0) >= min_rank
        ]

    return results # returns filtered physicians list