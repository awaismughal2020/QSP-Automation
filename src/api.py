"""
FastAPI server for QSP Quarterly Report Automation.

Provides REST API endpoints to generate quarterly reports programmatically.
"""

from fastapi import FastAPI, HTTPException, BackgroundTasks, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field
from typing import Optional, List
from pathlib import Path
from datetime import datetime
import shutil
import tempfile
import uuid
import os
from loguru import logger

from .orchestrator import QuarterlyReportOrchestrator, QuarterlyReportConfig

# Initialize FastAPI app
app = FastAPI(
    title="QSP Quarterly Report API",
    description="API for generating QSP ESS B.V. quarterly reports",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

# Store job statuses
job_status = {}

# Base paths
BASE_DIR = Path(__file__).resolve().parent.parent
INPUTS_DIR = BASE_DIR / "inputs"
OUTPUTS_DIR = BASE_DIR / "outputs"
UPLOADS_DIR = BASE_DIR / "uploads"

# Ensure directories exist
INPUTS_DIR.mkdir(exist_ok=True)
OUTPUTS_DIR.mkdir(exist_ok=True)
UPLOADS_DIR.mkdir(exist_ok=True)


# Request/Response Models
class GenerateRequest(BaseModel):
    """Request model for generating quarterly report."""
    year: int = Field(..., description="Report year (e.g., 2025)", ge=2020, le=2030)
    quarter: int = Field(..., description="Quarter number (1-4)", ge=1, le=4)
    bdo_file: str = Field(..., description="Path to BDO financials Excel file")
    prev_ma_file: str = Field(..., description="Path to previous Management Accounts file")
    rent_roll_file: str = Field(..., description="Path to rent roll Excel file")
    sales_tracker_file: str = Field(..., description="Path to sales tracker Excel file")
    prev_compliance_file: str = Field(..., description="Path to previous Compliance Certificate file")
    word_template_file: str = Field(..., description="Path to Word template file")
    output_dir: Optional[str] = Field(None, description="Output directory (default: outputs/)")
    dry_run: bool = Field(False, description="Validate inputs without generating files")

    class Config:
        json_schema_extra = {
            "example": {
                "year": 2025,
                "quarter": 3,
                "bdo_file": "inputs/Cijfers_QSP_30-09-2025_d_d__14-10-2025.xlsx",
                "prev_ma_file": "inputs/Management Accounts Q2 2025 - Draft 1.xlsx",
                "rent_roll_file": "inputs/QSP_huurlijst_1-10-2025.xlsx",
                "sales_tracker_file": "inputs/Unit_Sales_tracker_Q3_updated.xlsx",
                "prev_compliance_file": "inputs/Compliance Certificate Berekening QSP - Q2 2025_updated.xlsx",
                "word_template_file": "inputs/Quarterly_QSP_-_Q3_2025_-_Draft.docx",
                "dry_run": False
            }
        }


class GenerateResponse(BaseModel):
    """Response model for report generation."""
    status: str
    job_id: str
    message: str
    output_files: Optional[List[str]] = None
    warnings: Optional[List[str]] = None
    errors: Optional[List[str]] = None
    execution_time_seconds: Optional[float] = None


class JobStatusResponse(BaseModel):
    """Response model for job status check."""
    job_id: str
    status: str  # pending, running, completed, failed
    progress: Optional[str] = None
    result: Optional[dict] = None


class HealthResponse(BaseModel):
    """Health check response."""
    status: str
    timestamp: str
    version: str


# API Endpoints

@app.get("/", response_model=HealthResponse)
async def root():
    """Root endpoint - health check."""
    return HealthResponse(
        status="healthy",
        timestamp=datetime.now().isoformat(),
        version="1.0.0"
    )


@app.get("/health", response_model=HealthResponse)
async def health_check():
    """Health check endpoint."""
    return HealthResponse(
        status="healthy",
        timestamp=datetime.now().isoformat(),
        version="1.0.0"
    )


@app.post("/api/v1/generate", response_model=GenerateResponse)
async def generate_report(request: GenerateRequest):
    """
    Generate a quarterly report.
    
    This endpoint runs the full quarterly report generation pipeline:
    1. Parse BDO quarterly financials
    2. Parse rent roll
    3. Parse sales tracker
    4. Build Management Accounts
    5. Build Compliance Certificate
    6. Calculate compliance metrics
    7. Update Word template
    8. Assemble PDF (if LibreOffice available)
    """
    start_time = datetime.now()
    job_id = str(uuid.uuid4())[:8]
    
    logger.info(f"[Job {job_id}] Starting report generation for Q{request.quarter} {request.year}")
    
    try:
        # Resolve file paths
        output_dir = Path(request.output_dir) if request.output_dir else OUTPUTS_DIR
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Validate input files exist
        input_files = {
            "bdo_file": request.bdo_file,
            "prev_ma_file": request.prev_ma_file,
            "rent_roll_file": request.rent_roll_file,
            "sales_tracker_file": request.sales_tracker_file,
            "prev_compliance_file": request.prev_compliance_file,
            "word_template_file": request.word_template_file,
        }
        
        missing_files = []
        for name, path in input_files.items():
            if not Path(path).exists():
                missing_files.append(f"{name}: {path}")
        
        if missing_files:
            raise HTTPException(
                status_code=400,
                detail=f"Missing input files: {', '.join(missing_files)}"
            )
        
        # Create orchestrator config
        config = QuarterlyReportConfig(
            year=request.year,
            quarter=request.quarter,
            bdo_file=Path(request.bdo_file),
            previous_management_accounts=Path(request.prev_ma_file),
            rent_roll_file=Path(request.rent_roll_file),
            sales_tracker_file=Path(request.sales_tracker_file),
            previous_compliance_calc=Path(request.prev_compliance_file),
            word_template=Path(request.word_template_file),
            output_dir=output_dir
        )
        
        # Run orchestrator
        orchestrator = QuarterlyReportOrchestrator(config)
        result = orchestrator.run(dry_run=request.dry_run)
        
        execution_time = (datetime.now() - start_time).total_seconds()
        
        # Build response
        if result.get('status') == 'success':
            return GenerateResponse(
                status="success",
                job_id=job_id,
                message=f"Quarterly report Q{request.quarter} {request.year} generated successfully",
                output_files=result.get('output_files', []),
                warnings=result.get('warnings', []),
                errors=[],
                execution_time_seconds=execution_time
            )
        else:
            return GenerateResponse(
                status="failed",
                job_id=job_id,
                message="Report generation failed",
                output_files=result.get('output_files', []),
                warnings=result.get('warnings', []),
                errors=result.get('errors', []),
                execution_time_seconds=execution_time
            )
            
    except HTTPException:
        raise
    except Exception as e:
        logger.exception(f"[Job {job_id}] Error generating report: {e}")
        raise HTTPException(
            status_code=500,
            detail=f"Report generation failed: {str(e)}"
        )


@app.post("/api/v1/generate/async", response_model=JobStatusResponse)
async def generate_report_async(request: GenerateRequest, background_tasks: BackgroundTasks):
    """
    Start asynchronous report generation.
    
    Returns immediately with a job ID that can be used to check status.
    """
    job_id = str(uuid.uuid4())[:8]
    job_status[job_id] = {"status": "pending", "progress": "Queued"}
    
    background_tasks.add_task(run_generation_task, job_id, request)
    
    return JobStatusResponse(
        job_id=job_id,
        status="pending",
        progress="Job queued for processing"
    )


@app.get("/api/v1/jobs/{job_id}", response_model=JobStatusResponse)
async def get_job_status(job_id: str):
    """Get the status of an async generation job."""
    if job_id not in job_status:
        raise HTTPException(status_code=404, detail=f"Job {job_id} not found")
    
    job = job_status[job_id]
    return JobStatusResponse(
        job_id=job_id,
        status=job.get("status", "unknown"),
        progress=job.get("progress"),
        result=job.get("result")
    )


@app.get("/api/v1/outputs")
async def list_outputs():
    """List all generated output files."""
    if not OUTPUTS_DIR.exists():
        return {"files": []}
    
    files = []
    for f in OUTPUTS_DIR.iterdir():
        if f.is_file():
            files.append({
                "name": f.name,
                "size_bytes": f.stat().st_size,
                "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat()
            })
    
    return {"files": sorted(files, key=lambda x: x["modified"], reverse=True)}


@app.get("/api/v1/outputs/{filename}")
async def download_output(filename: str):
    """Download a generated output file."""
    file_path = OUTPUTS_DIR / filename
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f"File not found: {filename}")
    
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/octet-stream"
    )


@app.post("/api/v1/upload")
async def upload_file(file: UploadFile = File(...)):
    """
    Upload an input file.
    
    Files are saved to the inputs/ directory.
    """
    try:
        file_path = INPUTS_DIR / file.filename
        
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        return {
            "status": "success",
            "filename": file.filename,
            "path": str(file_path),
            "size_bytes": file_path.stat().st_size
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Upload failed: {str(e)}")


@app.get("/api/v1/inputs")
async def list_inputs():
    """List all input files."""
    if not INPUTS_DIR.exists():
        return {"files": []}
    
    files = []
    for f in INPUTS_DIR.iterdir():
        if f.is_file():
            files.append({
                "name": f.name,
                "size_bytes": f.stat().st_size,
                "modified": datetime.fromtimestamp(f.stat().st_mtime).isoformat()
            })
    
    return {"files": sorted(files, key=lambda x: x["name"])}


@app.delete("/api/v1/outputs/{filename}")
async def delete_output(filename: str):
    """Delete an output file."""
    file_path = OUTPUTS_DIR / filename
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail=f"File not found: {filename}")
    
    file_path.unlink()
    return {"status": "success", "message": f"Deleted {filename}"}


# Background task for async generation
async def run_generation_task(job_id: str, request: GenerateRequest):
    """Background task to run report generation."""
    try:
        job_status[job_id]["status"] = "running"
        job_status[job_id]["progress"] = "Starting generation..."
        
        output_dir = Path(request.output_dir) if request.output_dir else OUTPUTS_DIR
        output_dir.mkdir(parents=True, exist_ok=True)
        
        config = QuarterlyReportConfig(
            year=request.year,
            quarter=request.quarter,
            bdo_file=Path(request.bdo_file),
            previous_management_accounts=Path(request.prev_ma_file),
            rent_roll_file=Path(request.rent_roll_file),
            sales_tracker_file=Path(request.sales_tracker_file),
            previous_compliance_calc=Path(request.prev_compliance_file),
            word_template=Path(request.word_template_file),
            output_dir=output_dir
        )
        
        orchestrator = QuarterlyReportOrchestrator(config)
        result = orchestrator.run(dry_run=request.dry_run)
        
        job_status[job_id]["status"] = "completed" if result.get('status') == 'success' else "failed"
        job_status[job_id]["progress"] = "Completed"
        job_status[job_id]["result"] = result
        
    except Exception as e:
        logger.exception(f"[Job {job_id}] Background task failed: {e}")
        job_status[job_id]["status"] = "failed"
        job_status[job_id]["progress"] = f"Failed: {str(e)}"
        job_status[job_id]["result"] = {"errors": [str(e)]}


# Run with: uvicorn src.api:app --reload --port 8000
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)

