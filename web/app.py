"""
Web API for the CAAI flight logbook converter.

Upload a flight logbook (Excel, CSV, PDF, or LogTen Pro export) and
get back a filled Israeli CAAI tofes-shaot form.

Usage:
    python -m uvicorn web.app:app --reload
    # Open http://localhost:8000
"""

import os
import sys
import uuid
import shutil
import tempfile
from io import StringIO
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

# Ensure project root is on the path
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.universal_importer import create_standardized_logbook, detect_format
from src.logbook_creator import create_logbook
from src.distance_calculator import add_distances
from src.caai_columns import add_caai_columns
from src.caai_form_filler import fill_caai_form

app = FastAPI(title="CAAI Flight Logbook Converter")

# Template path
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "tofes-shaot-blank.xlsx"

# Store completed jobs for download (job_id -> file path)
_jobs = {}

# Serve static files
STATIC_DIR = Path(__file__).parent / "static"
if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    """Serve the main upload page."""
    html_path = Path(__file__).parent / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.post("/api/convert")
async def convert(
    file: UploadFile = File(...),
    pilot_name: str = Form(""),
):
    """Convert an uploaded logbook to a filled CAAI tofes-shaot form.

    Runs the full pipeline:
    1. Import logbook (auto-detect format and columns)
    2. Calculate distances
    3. Add CAAI classification columns
    4. Fill the tofes-shaot form

    Returns JSON with stats and a download URL.
    """
    # Validate file
    if not file.filename:
        raise HTTPException(400, "No file uploaded")

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ('.xlsx', '.xls', '.xlsm', '.csv', '.tsv', '.txt', '.pdf'):
        raise HTTPException(400, f"Unsupported file type: {ext}. Use Excel, CSV, TSV, or PDF.")

    # Create temp directory for this job
    job_id = str(uuid.uuid4())[:8]
    work_dir = tempfile.mkdtemp(prefix=f"caai_{job_id}_")

    try:
        # Save uploaded file
        upload_path = os.path.join(work_dir, file.filename)
        with open(upload_path, "wb") as f:
            content = await file.read()
            f.write(content)

        logbook_path = os.path.join(work_dir, "logbook.xlsx")
        caai_output = os.path.join(work_dir, "CAAI_Tofes_Shaot_Filled.xlsx")

        # Capture stdout
        old_stdout = sys.stdout
        sys.stdout = log_capture = StringIO()

        try:
            # Step 1: Import
            fmt = detect_format(upload_path)
            if fmt == 'logten':
                create_logbook(upload_path, logbook_path, pilot_name)
            else:
                create_standardized_logbook(
                    input_file=upload_path,
                    output_file=logbook_path,
                    pilot_name=pilot_name,
                    fmt=fmt,
                )

            # Step 2: Distances
            add_distances(logbook_path)

            # Step 3: CAAI columns
            add_caai_columns(logbook_path)

            # Step 4: Fill form
            result = fill_caai_form(logbook_path, str(TEMPLATE_PATH), caai_output)

        finally:
            sys.stdout = old_stdout

        log_output = log_capture.getvalue()

        # Count flights from logbook
        from openpyxl import load_workbook as _lwb
        _wb = _lwb(logbook_path, read_only=True)
        flight_count = _wb['Flight Log'].max_row - 1  # minus header
        _wb.close()

        # Extract key stats
        grand = result.get('grand', {})
        stats = {
            'flights': flight_count,
            'total_hours': round(grand.get('total', 0), 1),
            'pic': round(grand.get('pic', 0), 1),
            'sic': round(grand.get('sic', 0), 1),
            'student': round(grand.get('student', 0), 1),
            'safety_pilot': round(grand.get('safety_pilot_se', 0), 1),
            'night': round(grand.get('night', 0), 1),
            'xc': round(grand.get('xc_all_roles', 0), 1),
            'instrument': round(
                grand.get('actual_inst', 0) + grand.get('sim_inst_air', 0), 1
            ),
            'caai_grand_total': round(result.get('caai_grand_total', 0), 1),
            'form_total': round(grand.get('form_total', 0), 1),
            'format_detected': fmt,
        }

        # Store for download
        _jobs[job_id] = {
            'caai_file': caai_output,
            'logbook_file': logbook_path,
            'work_dir': work_dir,
        }

        return JSONResponse({
            'success': True,
            'job_id': job_id,
            'stats': stats,
            'download_url': f'/api/download/{job_id}',
            'logbook_url': f'/api/download/{job_id}/logbook',
        })

    except Exception as e:
        # Cleanup on error
        shutil.rmtree(work_dir, ignore_errors=True)
        return JSONResponse(
            status_code=500,
            content={'success': False, 'error': str(e)},
        )


@app.get("/api/download/{job_id}")
async def download_caai(job_id: str):
    """Download the filled CAAI tofes-shaot form."""
    job = _jobs.get(job_id)
    if not job or not os.path.exists(job['caai_file']):
        raise HTTPException(404, "File not found or expired")

    return FileResponse(
        job['caai_file'],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="CAAI_Tofes_Shaot_Filled.xlsx",
    )


@app.get("/api/download/{job_id}/logbook")
async def download_logbook(job_id: str):
    """Download the standardized logbook."""
    job = _jobs.get(job_id)
    if not job or not os.path.exists(job['logbook_file']):
        raise HTTPException(404, "File not found or expired")

    return FileResponse(
        job['logbook_file'],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Flight_Logbook.xlsx",
    )
