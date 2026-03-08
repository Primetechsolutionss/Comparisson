"""FastAPI server for Leveransplan comparison service."""
import os
import uuid
import shutil
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from comparison_engine import compare_and_report, DEFAULT_ALLOWLIST

app = FastAPI(title="Leveransplan Comparison API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = Path("/tmp/leveransplan_uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

REPORT_DIR = Path("/tmp/leveransplan_reports")
REPORT_DIR.mkdir(exist_ok=True)


@app.get("/api/health")
async def health():
    return {"status": "ok", "version": "1.0.0"}


@app.post("/api/compare")
async def compare(
    master_file: UploadFile = File(...),
    delivery_file: UploadFile = File(...),
    allowlist: str = Form(default=""),
):
    """Upload master + delivery files and run comparison."""
    job_id = str(uuid.uuid4())[:8]
    job_dir = UPLOAD_DIR / job_id
    job_dir.mkdir(exist_ok=True)
    
    try:
        # Save uploaded files
        master_path = job_dir / f"master_{master_file.filename}"
        delivery_path = job_dir / f"delivery_{delivery_file.filename}"
        
        with open(master_path, "wb") as f:
            shutil.copyfileobj(master_file.file, f)
        with open(delivery_path, "wb") as f:
            shutil.copyfileobj(delivery_file.file, f)
        
        # Parse allowlist
        if allowlist.strip():
            ext_list = {ext.strip().lower() for ext in allowlist.split(",") if ext.strip()}
            ext_set = {ext if ext.startswith('.') else f'.{ext}' for ext in ext_list}
        else:
            ext_set = DEFAULT_ALLOWLIST
        
        # Generate report filename
        delivery_stem = Path(delivery_file.filename).stem
        report_filename = f"{delivery_stem}_vs_Master_ComparisonReport.xlsx"
        report_path = REPORT_DIR / f"{job_id}_{report_filename}"
        
        # Run comparison
        result, error = compare_and_report(
            master_path=str(master_path),
            delivery_path=str(delivery_path),
            output_path=str(report_path),
            allowlist=ext_set,
        )
        
        if error:
            raise HTTPException(status_code=400, detail=error)
        
        return JSONResponse({
            "job_id": job_id,
            "report_filename": report_filename,
            "download_url": f"/api/download/{job_id}/{report_filename}",
            "summary": result["summary_text"],
            "stats": result["stats"],
        })
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # Cleanup uploads
        if job_dir.exists():
            shutil.rmtree(job_dir, ignore_errors=True)


@app.get("/api/download/{job_id}/{filename}")
async def download_report(job_id: str, filename: str):
    """Download the generated comparison report."""
    report_path = REPORT_DIR / f"{job_id}_{filename}"
    if not report_path.exists():
        raise HTTPException(status_code=404, detail="Report not found or expired.")
    return FileResponse(
        path=str(report_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/allowlist/default")
async def get_default_allowlist():
    """Return the default allowlist."""
    return {"allowlist": sorted(DEFAULT_ALLOWLIST)}
