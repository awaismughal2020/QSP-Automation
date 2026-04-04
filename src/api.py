"""
FastAPI server for QSP Quarterly Report Automation.

Provides REST API endpoints to generate quarterly reports programmatically.
"""

from fastapi import FastAPI, HTTPException, BackgroundTasks, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from typing import Optional, List
from pathlib import Path
from datetime import datetime
import shutil
import tempfile
import uuid
import os
from loguru import logger

from .orchestrator import (
    QuarterlyReportOrchestrator, 
    QuarterlyReportConfig,
    WordReportOrchestrator,
    WordReportConfig,
    FinalPDFOrchestrator,
    FinalPDFConfig
)

# Initialize FastAPI app
app = FastAPI(
    title="QSP Quarterly Report API",
    description="API for generating QSP ESS B.V. quarterly reports",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

# Add CORS middleware for n8n and other clients
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins for development
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
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
                "year": 2025,  # Dynamic: Use the report year
                "quarter": 3,  # Dynamic: 1-4 for Q1-Q4
                "bdo_file": "inputs/Cijfers_QSP_<date>.xlsx",  # BDO quarterly financials
                "prev_ma_file": "inputs/Management Accounts Q<prev_quarter> <year> - Draft 1.xlsx",
                "rent_roll_file": "inputs/QSP_huurlijst_<date>.xlsx",  # Rent roll
                "sales_tracker_file": "inputs/Unit_Sales_tracker_Q<quarter>_updated.xlsx",
                "prev_compliance_file": "inputs/Compliance Certificate Berekening QSP - Q<prev_quarter> <year>_updated.xlsx",
                "word_template_file": "inputs/Quarterly_QSP_-_Q<prev_quarter>_<year>_-_Draft.docx",
                "dry_run": False
            }
        }


class AIVerificationStatus(BaseModel):
    """AI verification step result."""
    status: str = "skipped"
    patches_applied: int = 0
    revalidation_passed: bool = False
    issues_found: int = 0
    notes: str = ""


class GenerateResponse(BaseModel):
    """Response model for report generation."""
    status: str
    job_id: str
    message: str
    output_files: Optional[List[str]] = None
    management_accounts_path: Optional[str] = None
    compliance_certificate_path: Optional[str] = None
    warnings: Optional[List[str]] = None
    errors: Optional[List[str]] = None
    execution_time_seconds: Optional[float] = None
    ai_verification: Optional[AIVerificationStatus] = None


class GenerateAccountsRequest(BaseModel):
    """Request model for Phase 1: Generate Management Accounts + Compliance Certificate."""
    year: int = Field(..., description="Report year (e.g., 2025)", ge=2020, le=2030)
    quarter: int = Field(..., description="Quarter number (1-4)", ge=1, le=4)
    bdo_file: str = Field(..., description="Path to BDO financials Excel file")
    prev_ma_file: str = Field(..., description="Path to previous Management Accounts file")
    rent_roll_file: str = Field(..., description="Path to rent roll Excel file")
    sales_tracker_file: str = Field(..., description="Path to sales tracker Excel file")
    prev_compliance_file: str = Field(..., description="Path to previous Compliance Certificate file")
    output_dir: Optional[str] = Field(None, description="Output directory (default: outputs/)")
    dry_run: bool = Field(False, description="Validate inputs without generating files")

    class Config:
        json_schema_extra = {
            "example": {
                "year": 2025,
                "quarter": 3,
                "bdo_file": "inputs/Cijfers_QSP_Q3.xlsx",
                "prev_ma_file": "inputs/Management Accounts Q2 2025 - Draft 1.xlsx",
                "rent_roll_file": "inputs/QSP_huurlijst_Q3.xlsx",
                "sales_tracker_file": "inputs/Unit_Sales_tracker_Q3.xlsx",
                "prev_compliance_file": "inputs/Compliance Certificate Berekening QSP - Q2 2025.xlsx",
                "dry_run": False
            }
        }


class GenerateWordReportRequest(BaseModel):
    """Request model for Phase 2: Generate Word Report."""
    year: int = Field(..., description="Report year (e.g., 2025)", ge=2020, le=2030)
    quarter: int = Field(..., description="Quarter number (1-4)", ge=1, le=4)
    management_accounts_file: str = Field(..., description="Path to Management Accounts file (from Phase 1)")
    compliance_certificate_file: str = Field(..., description="Path to Compliance Certificate file (edited by client)")
    rent_roll_file: str = Field(..., description="Path to rent roll Excel file")
    sales_tracker_file: str = Field(..., description="Path to sales tracker Excel file")
    word_template_file: str = Field(..., description="Path to Word template file")
    output_dir: Optional[str] = Field(None, description="Output directory (default: outputs/)")

    class Config:
        json_schema_extra = {
            "example": {
                "year": 2025,
                "quarter": 3,
                "management_accounts_file": "outputs/Management Accounts Q3 2025.xlsx",
                "compliance_certificate_file": "inputs/Compliance Certificate Berekening QSP - Q3 2025_edited.xlsx",
                "rent_roll_file": "inputs/QSP_huurlijst_Q3.xlsx",
                "sales_tracker_file": "inputs/Unit_Sales_tracker_Q3.xlsx",
                "word_template_file": "inputs/Quarterly_QSP_-_Q2_2025_-_Draft.docx"
            }
        }


class GenerateFinalPDFRequest(BaseModel):
    """Request model for Phase 3: Generate Final PDF."""
    year: int = Field(..., description="Report year (e.g., 2025)", ge=2020, le=2030)
    quarter: int = Field(..., description="Quarter number (1-4)", ge=1, le=4)
    word_report_file: str = Field(..., description="Path to final Word report (from Phase 2, potentially edited)")
    management_accounts_file: str = Field(..., description="Path to final Management Accounts file")
    compliance_certificate_file: str = Field(..., description="Path to final Compliance Certificate (client-approved)")
    rent_roll_file: str = Field(..., description="Path to rent roll Excel file")
    sales_tracker_file: str = Field(..., description="Path to sales tracker Excel file")
    output_dir: Optional[str] = Field(None, description="Output directory (default: outputs/)")

    class Config:
        json_schema_extra = {
            "example": {
                "year": 2025,
                "quarter": 3,
                "word_report_file": "outputs/Quarterly QSP - Q3 2025 - Draft.docx",
                "management_accounts_file": "outputs/Management Accounts Q3 2025.xlsx",
                "compliance_certificate_file": "inputs/Compliance Certificate Berekening QSP - Q3 2025_final.xlsx",
                "rent_roll_file": "inputs/QSP_huurlijst_Q3.xlsx",
                "sales_tracker_file": "inputs/Unit_Sales_tracker_Q3.xlsx"
            }
        }


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


# Helper function for path resolution
def resolve_input_path(path_str: str, base_dir: Path = None) -> Path:
    """
    Resolve input file path, handling both absolute and relative paths.
    
    Args:
        path_str: Path string (can be absolute like /app/inputs/... or relative)
        base_dir: Base directory for resolving relative paths (defaults to BASE_DIR)
    
    Returns:
        Resolved Path object
    """
    if base_dir is None:
        base_dir = BASE_DIR
    
    path = Path(path_str)
    # If path is absolute, use as-is
    if path.is_absolute():
        return path
    else:
        # Try relative to BASE_DIR first, then INPUTS_DIR
        resolved = base_dir / path_str
        if resolved.exists():
            return resolved
        return INPUTS_DIR / path_str


def _extract_ai_verification(result: dict) -> Optional[AIVerificationStatus]:
    """Extract AI verification status from orchestrator results."""
    ai_step = result.get('steps', {}).get('ai_verification')
    if ai_step:
        return AIVerificationStatus(
            status=ai_step.get('status', 'skipped'),
            patches_applied=ai_step.get('patches_applied', 0),
            revalidation_passed=ai_step.get('revalidation_passed', False),
            issues_found=ai_step.get('issues_found', 0),
            notes=ai_step.get('notes', ''),
        )
    return None


def _extract_named_paths(result: dict) -> dict:
    """Derive management_accounts_path and compliance_certificate_path from output_files."""
    ma_path = None
    cc_path = None
    for fp in result.get('output_files', []):
        lower = fp.lower()
        if 'management accounts' in lower and lower.endswith('.xlsx'):
            ma_path = fp
        elif 'compliance' in lower and lower.endswith('.xlsx'):
            cc_path = fp
    return {'management_accounts_path': ma_path, 'compliance_certificate_path': cc_path}


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
        
        # Validate input files exist and resolve paths
        # Handle both absolute paths (e.g., /app/inputs/...) and relative paths
        input_files = {
            "bdo_file": request.bdo_file,
            "prev_ma_file": request.prev_ma_file,
            "rent_roll_file": request.rent_roll_file,
            "sales_tracker_file": request.sales_tracker_file,
            "prev_compliance_file": request.prev_compliance_file,
            "word_template_file": request.word_template_file,
        }
        
        # Resolve file paths (handle both absolute and relative paths)
        resolved_files = {}
        missing_files = []
        
        for name, path_str in input_files.items():
            resolved_path = resolve_input_path(path_str)
            
            if not resolved_path.exists():
                missing_files.append(f"{name}: {path_str} (resolved: {resolved_path})")
            else:
                resolved_files[name] = resolved_path
                logger.info(f"[Job {job_id}] Resolved {name}: {path_str} -> {resolved_path}")
        
        if missing_files:
            raise HTTPException(
                status_code=400,
                detail=f"Missing input files: {', '.join(missing_files)}"
            )
        
        # Create orchestrator config with resolved paths
        config = QuarterlyReportConfig(
            year=request.year,
            quarter=request.quarter,
            bdo_file=resolved_files["bdo_file"],
            previous_management_accounts=resolved_files["prev_ma_file"],
            rent_roll_file=resolved_files["rent_roll_file"],
            sales_tracker_file=resolved_files["sales_tracker_file"],
            previous_compliance_calc=resolved_files["prev_compliance_file"],
            word_template=resolved_files["word_template_file"],
            output_dir=output_dir
        )
        
        # Run orchestrator
        orchestrator = QuarterlyReportOrchestrator(config)
        result = orchestrator.run(dry_run=request.dry_run)
        
        execution_time = (datetime.now() - start_time).total_seconds()
        
        # Build response
        ai_verification = _extract_ai_verification(result)
        named = _extract_named_paths(result)

        if result.get('status') == 'success':
            return GenerateResponse(
                status="success",
                job_id=job_id,
                message=f"Quarterly report Q{request.quarter} {request.year} generated successfully",
                output_files=result.get('output_files', []),
                management_accounts_path=named['management_accounts_path'],
                compliance_certificate_path=named['compliance_certificate_path'],
                warnings=result.get('warnings', []),
                errors=[],
                execution_time_seconds=execution_time,
                ai_verification=ai_verification,
            )
        else:
            return GenerateResponse(
                status="failed",
                job_id=job_id,
                message="Report generation failed",
                output_files=result.get('output_files', []),
                warnings=result.get('warnings', []),
                errors=result.get('errors', []),
                execution_time_seconds=execution_time,
                ai_verification=ai_verification,
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


# ============================================================================
# PHASE 1: Generate Management Accounts + Compliance Certificate
# ============================================================================

@app.post("/api/v1/generate-accounts", response_model=GenerateResponse)
async def generate_accounts(request: GenerateAccountsRequest):
    """
    Phase 1: Generate Management Accounts and Draft Compliance Certificate.
    
    This is the first phase of the split workflow:
    1. Parse BDO quarterly financials
    2. Parse rent roll
    3. Parse sales tracker  
    4. Build Management Accounts
    5. Build Compliance Certificate
    
    Outputs:
    - Management Accounts Excel
    - Compliance Certificate Excel (Draft)
    
    The client can download and review/edit the Compliance Certificate
    before proceeding to Phase 2.
    """
    start_time = datetime.now()
    job_id = str(uuid.uuid4())[:8]
    
    logger.info(f"[Job {job_id}] [Phase 1] Starting accounts generation for Q{request.quarter} {request.year}")
    
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
        }
        
        resolved_files = {}
        missing_files = []
        
        for name, path_str in input_files.items():
            resolved_path = resolve_input_path(path_str)
            if not resolved_path.exists():
                missing_files.append(f"{name}: {path_str} (resolved: {resolved_path})")
            else:
                resolved_files[name] = resolved_path
                logger.info(f"[Job {job_id}] Resolved {name}: {path_str} -> {resolved_path}")
        
        if missing_files:
            raise HTTPException(
                status_code=400,
                detail=f"Missing input files: {', '.join(missing_files)}"
            )
        
        # Create orchestrator config
        config = QuarterlyReportConfig(
            year=request.year,
            quarter=request.quarter,
            bdo_file=resolved_files["bdo_file"],
            previous_management_accounts=resolved_files["prev_ma_file"],
            rent_roll_file=resolved_files["rent_roll_file"],
            sales_tracker_file=resolved_files["sales_tracker_file"],
            previous_compliance_calc=resolved_files["prev_compliance_file"],
            word_template=Path("/dev/null"),  # Not used in Phase 1
            output_dir=output_dir
        )
        
        # Run Phase 1 only
        orchestrator = QuarterlyReportOrchestrator(config)
        result = orchestrator.run_accounts_only(dry_run=request.dry_run)
        
        execution_time = (datetime.now() - start_time).total_seconds()
        
        ai_verification = _extract_ai_verification(result)
        named = _extract_named_paths(result)

        if result.get('status') == 'success':
            return GenerateResponse(
                status="success",
                job_id=job_id,
                message=f"Phase 1 complete: Management Accounts and Compliance Certificate generated for Q{request.quarter} {request.year}",
                output_files=result.get('output_files', []),
                management_accounts_path=named['management_accounts_path'],
                compliance_certificate_path=named['compliance_certificate_path'],
                warnings=result.get('warnings', []),
                errors=[],
                execution_time_seconds=execution_time,
                ai_verification=ai_verification,
            )
        else:
            return GenerateResponse(
                status="failed",
                job_id=job_id,
                message="Phase 1 generation failed",
                output_files=result.get('output_files', []),
                warnings=result.get('warnings', []),
                errors=result.get('errors', []),
                execution_time_seconds=execution_time,
                ai_verification=ai_verification,
            )
            
    except HTTPException:
        raise
    except Exception as e:
        logger.exception(f"[Job {job_id}] Error in Phase 1: {e}")
        raise HTTPException(
            status_code=500,
            detail=f"Phase 1 generation failed: {str(e)}"
        )


# ============================================================================
# PHASE 2: Generate Word Report
# ============================================================================

@app.post("/api/v1/generate-word-report", response_model=GenerateResponse)
async def generate_word_report(request: GenerateWordReportRequest):
    """
    Phase 2: Generate Draft Word Report.
    
    Uses Management Accounts and Compliance Certificate from Phase 1
    (potentially edited by client) to generate the Word report.
    
    Inputs:
    - Management Accounts file (from Phase 1)
    - Compliance Certificate (edited/updated by client)
    - All source files (rent roll, sales tracker)
    - Word template
    
    Output:
    - Draft Word Report
    """
    start_time = datetime.now()
    job_id = str(uuid.uuid4())[:8]
    
    logger.info(f"[Job {job_id}] [Phase 2] Starting Word report generation for Q{request.quarter} {request.year}")
    
    try:
        # Resolve file paths
        output_dir = Path(request.output_dir) if request.output_dir else OUTPUTS_DIR
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Validate input files exist
        input_files = {
            "management_accounts_file": request.management_accounts_file,
            "compliance_certificate_file": request.compliance_certificate_file,
            "rent_roll_file": request.rent_roll_file,
            "sales_tracker_file": request.sales_tracker_file,
            "word_template_file": request.word_template_file,
        }
        
        resolved_files = {}
        missing_files = []
        
        for name, path_str in input_files.items():
            resolved_path = resolve_input_path(path_str)
            if not resolved_path.exists():
                missing_files.append(f"{name}: {path_str} (resolved: {resolved_path})")
            else:
                resolved_files[name] = resolved_path
                logger.info(f"[Job {job_id}] Resolved {name}: {path_str} -> {resolved_path}")
        
        if missing_files:
            raise HTTPException(
                status_code=400,
                detail=f"Missing input files: {', '.join(missing_files)}"
            )
        
        # Create Word report config
        config = WordReportConfig(
            year=request.year,
            quarter=request.quarter,
            management_accounts_file=resolved_files["management_accounts_file"],
            compliance_certificate_file=resolved_files["compliance_certificate_file"],
            rent_roll_file=resolved_files["rent_roll_file"],
            sales_tracker_file=resolved_files["sales_tracker_file"],
            word_template=resolved_files["word_template_file"],
            output_dir=output_dir
        )
        
        # Run Phase 2
        orchestrator = WordReportOrchestrator(config)
        result = orchestrator.run()
        
        execution_time = (datetime.now() - start_time).total_seconds()
        
        if result.get('status') == 'success':
            return GenerateResponse(
                status="success",
                job_id=job_id,
                message=f"Phase 2 complete: Word report generated for Q{request.quarter} {request.year}",
                output_files=result.get('output_files', []),
                warnings=result.get('warnings', []),
                errors=[],
                execution_time_seconds=execution_time
            )
        else:
            return GenerateResponse(
                status="failed",
                job_id=job_id,
                message="Phase 2 generation failed",
                output_files=result.get('output_files', []),
                warnings=result.get('warnings', []),
                errors=result.get('errors', []),
                execution_time_seconds=execution_time
            )
            
    except HTTPException:
        raise
    except Exception as e:
        logger.exception(f"[Job {job_id}] Error in Phase 2: {e}")
        raise HTTPException(
            status_code=500,
            detail=f"Phase 2 generation failed: {str(e)}"
        )


# ============================================================================
# PHASE 3: Generate Final PDF
# ============================================================================

@app.post("/api/v1/generate-final-pdf", response_model=GenerateResponse)
async def generate_final_pdf(request: GenerateFinalPDFRequest):
    """
    Phase 3: Generate Final PDF Report.
    
    Takes all finalized documents and assembles them into the final PDF.
    
    Inputs:
    - Final Compliance Certificate (client-approved)
    - Final Word Report (from Phase 2, potentially edited)
    - All source files (rent roll, sales tracker)
    
    Output:
    - Final PDF Report
    
    Note: Requires LibreOffice to be installed for PDF conversion.
    """
    start_time = datetime.now()
    job_id = str(uuid.uuid4())[:8]
    
    logger.info(f"[Job {job_id}] [Phase 3] Starting final PDF generation for Q{request.quarter} {request.year}")
    
    try:
        # Resolve file paths
        output_dir = Path(request.output_dir) if request.output_dir else OUTPUTS_DIR
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Validate input files exist
        input_files = {
            "word_report_file": request.word_report_file,
            "management_accounts_file": request.management_accounts_file,
            "compliance_certificate_file": request.compliance_certificate_file,
            "rent_roll_file": request.rent_roll_file,
            "sales_tracker_file": request.sales_tracker_file,
        }
        
        resolved_files = {}
        missing_files = []
        
        for name, path_str in input_files.items():
            resolved_path = resolve_input_path(path_str)
            if not resolved_path.exists():
                missing_files.append(f"{name}: {path_str} (resolved: {resolved_path})")
            else:
                resolved_files[name] = resolved_path
                logger.info(f"[Job {job_id}] Resolved {name}: {path_str} -> {resolved_path}")
        
        if missing_files:
            raise HTTPException(
                status_code=400,
                detail=f"Missing input files: {', '.join(missing_files)}"
            )
        
        # Create Final PDF config
        config = FinalPDFConfig(
            year=request.year,
            quarter=request.quarter,
            word_report_file=resolved_files["word_report_file"],
            management_accounts_file=resolved_files["management_accounts_file"],
            compliance_certificate_file=resolved_files["compliance_certificate_file"],
            rent_roll_file=resolved_files["rent_roll_file"],
            sales_tracker_file=resolved_files["sales_tracker_file"],
            output_dir=output_dir
        )
        
        # Run Phase 3
        orchestrator = FinalPDFOrchestrator(config)
        result = orchestrator.run()
        
        execution_time = (datetime.now() - start_time).total_seconds()
        
        if result.get('status') == 'success':
            return GenerateResponse(
                status="success",
                job_id=job_id,
                message=f"Phase 3 complete: Final PDF generated for Q{request.quarter} {request.year}",
                output_files=result.get('output_files', []),
                warnings=result.get('warnings', []),
                errors=[],
                execution_time_seconds=execution_time
            )
        else:
            return GenerateResponse(
                status="failed",
                job_id=job_id,
                message="Phase 3 generation failed",
                output_files=result.get('output_files', []),
                warnings=result.get('warnings', []),
                errors=result.get('errors', []),
                execution_time_seconds=execution_time
            )
            
    except HTTPException:
        raise
    except Exception as e:
        logger.exception(f"[Job {job_id}] Error in Phase 3: {e}")
        raise HTTPException(
            status_code=500,
            detail=f"Phase 3 generation failed: {str(e)}"
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
        # Sanitize filename
        safe_filename = file.filename.replace(" ", "_") if file.filename else f"upload_{uuid.uuid4()[:8]}"
        file_path = INPUTS_DIR / safe_filename
        
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        logger.info(f"Uploaded file: {safe_filename} ({file_path.stat().st_size} bytes)")
        
        return {
            "status": "success",
            "filename": safe_filename,
            "original_filename": file.filename,
            "path": str(file_path),
            "size_bytes": file_path.stat().st_size
        }
    except Exception as e:
        logger.error(f"Upload failed: {e}")
        raise HTTPException(status_code=500, detail=f"Upload failed: {str(e)}")


@app.post("/api/v1/upload/batch")
async def upload_files_batch(files: List[UploadFile] = File(...)):
    """
    Upload multiple input files at once.
    
    Files are saved to the inputs/ directory.
    """
    results = []
    errors = []
    
    for file in files:
        try:
            safe_filename = file.filename.replace(" ", "_") if file.filename else f"upload_{uuid.uuid4()[:8]}"
            file_path = INPUTS_DIR / safe_filename
            
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            
            results.append({
                "filename": safe_filename,
                "original_filename": file.filename,
                "path": str(file_path),
                "size_bytes": file_path.stat().st_size
            })
            logger.info(f"Uploaded file: {safe_filename} ({file_path.stat().st_size} bytes)")
        except Exception as e:
            errors.append({
                "filename": file.filename,
                "error": str(e)
            })
            logger.error(f"Failed to upload {file.filename}: {e}")
    
    return {
        "status": "success" if not errors else "partial",
        "uploaded": results,
        "errors": errors,
        "total_uploaded": len(results),
        "total_failed": len(errors)
    }


@app.get("/form", response_class=HTMLResponse)
async def upload_form():
    """
    Simple HTML form for uploading files and generating reports.
    
    This provides a web interface without requiring n8n.
    """
    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <title>QSP Report Generator</title>
        <style>
            * { box-sizing: border-box; }
            body { 
                font-family: 'Segoe UI', Arial, sans-serif; 
                max-width: 900px; 
                margin: 0 auto; 
                padding: 20px; 
                background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
                min-height: 100vh;
                color: #eee;
            }
            .container { 
                background: rgba(30, 41, 59, 0.95); 
                padding: 40px; 
                border-radius: 16px; 
                box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5);
                border: 1px solid rgba(255,255,255,0.1);
            }
            h1 { 
                color: #60a5fa; 
                text-align: center;
                margin-bottom: 10px;
                font-size: 2.2em;
            }
            .subtitle {
                text-align: center;
                color: #94a3b8;
                margin-bottom: 30px;
            }
            .form-group { 
                margin-bottom: 20px; 
            }
            label { 
                display: block; 
                margin-bottom: 8px; 
                font-weight: 500;
                color: #cbd5e1;
            }
            input[type="number"], select { 
                width: 100%; 
                padding: 12px 15px; 
                border: 2px solid #334155; 
                border-radius: 8px; 
                font-size: 16px;
                background: #1e293b;
                color: #fff;
                transition: border-color 0.3s;
            }
            input[type="number"]:focus, select:focus { 
                outline: none;
                border-color: #60a5fa;
            }
            input[type="file"] { 
                width: 100%; 
                padding: 12px; 
                border: 2px dashed #334155; 
                border-radius: 8px;
                background: #1e293b;
                color: #94a3b8;
                cursor: pointer;
                transition: all 0.3s;
            }
            input[type="file"]:hover { 
                border-color: #60a5fa;
                background: #263447;
            }
            .row { 
                display: flex; 
                gap: 20px; 
            }
            .row .form-group { 
                flex: 1; 
            }
            button { 
                width: 100%; 
                padding: 16px 30px; 
                background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
                color: white; 
                border: none; 
                border-radius: 10px; 
                font-size: 18px;
                font-weight: 600;
                cursor: pointer;
                transition: all 0.3s;
                margin-top: 20px;
            }
            button:hover { 
                transform: translateY(-2px);
                box-shadow: 0 10px 25px -5px rgba(59,130,246,0.5);
            }
            button:disabled { 
                background: #475569; 
                cursor: not-allowed;
                transform: none;
                box-shadow: none;
            }
            .result { 
                margin-top: 30px; 
                padding: 20px; 
                border-radius: 10px; 
                display: none;
            }
            .success { 
                background: rgba(34, 197, 94, 0.15); 
                border: 1px solid #22c55e;
            }
            .error { 
                background: rgba(239, 68, 68, 0.15); 
                border: 1px solid #ef4444;
            }
            .loader { 
                display: none; 
                text-align: center; 
                padding: 30px;
            }
            .loader::after {
                content: '';
                display: inline-block;
                width: 40px;
                height: 40px;
                border: 3px solid #334155;
                border-top-color: #60a5fa;
                border-radius: 50%;
                animation: spin 1s linear infinite;
            }
            @keyframes spin {
                to { transform: rotate(360deg); }
            }
            .download-link { 
                display: inline-block; 
                margin: 8px 8px 8px 0; 
                padding: 10px 20px; 
                background: #22c55e; 
                color: white; 
                text-decoration: none; 
                border-radius: 6px;
                font-weight: 500;
                transition: all 0.2s;
            }
            .download-link:hover { 
                background: #16a34a;
                transform: translateY(-1px);
            }
            .file-icon { margin-right: 8px; }
            .progress-text {
                text-align: center;
                color: #94a3b8;
                margin-top: 10px;
            }
            .section-header {
                font-size: 1.1em;
                color: #60a5fa;
                margin: 30px 0 15px;
                padding-bottom: 8px;
                border-bottom: 1px solid #334155;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>📊 QSP Report Generator</h1>
            <p class="subtitle">Upload your quarterly files to generate the complete report</p>
            
            <form id="reportForm" enctype="multipart/form-data">
                <div class="section-header">📅 Report Period</div>
                <div class="row">
                    <div class="form-group">
                        <label for="year">Year</label>
                        <input type="number" id="year" name="year" min="2020" max="2035" required>
                    </div>
                    <div class="form-group">
                        <label for="quarter">Quarter</label>
                        <select id="quarter" name="quarter" required>
                            <option value="1">Q1</option>
                            <option value="2">Q2</option>
                            <option value="3">Q3</option>
                            <option value="4">Q4</option>
                        </select>
                    </div>
                </div>
                <script>
                    // Set dynamic defaults based on current date
                    (function() {
                        const now = new Date();
                        const currentYear = now.getFullYear();
                        const currentMonth = now.getMonth() + 1; // 1-12
                        // Determine current quarter
                        const currentQuarter = Math.ceil(currentMonth / 3);
                        
                        document.getElementById('year').value = currentYear;
                        document.getElementById('quarter').value = currentQuarter;
                    })();
                </script>
                
                <div class="section-header">📁 Input Files</div>
                
                <div class="form-group">
                    <label for="bdo_file">BDO Quarterly Financials (.xlsx)</label>
                    <input type="file" id="bdo_file" name="bdo_file" accept=".xlsx" required>
                </div>
                
                <div class="form-group">
                    <label for="prev_ma_file">Previous Management Accounts (.xlsx)</label>
                    <input type="file" id="prev_ma_file" name="prev_ma_file" accept=".xlsx" required>
                </div>
                
                <div class="form-group">
                    <label for="rent_roll_file">Rent Roll (.xlsx)</label>
                    <input type="file" id="rent_roll_file" name="rent_roll_file" accept=".xlsx" required>
                </div>
                
                <div class="form-group">
                    <label for="sales_tracker_file">Unit Sales Tracker (.xlsx)</label>
                    <input type="file" id="sales_tracker_file" name="sales_tracker_file" accept=".xlsx" required>
                </div>
                
                <div class="form-group">
                    <label for="prev_compliance_file">Previous Compliance Certificate (.xlsx)</label>
                    <input type="file" id="prev_compliance_file" name="prev_compliance_file" accept=".xlsx" required>
                </div>
                
                <div class="form-group">
                    <label for="word_template_file">Word Report Template (.docx)</label>
                    <input type="file" id="word_template_file" name="word_template_file" accept=".docx" required>
                </div>
                
                <button type="submit" id="submitBtn">🚀 Generate Quarterly Report</button>
            </form>
            
            <div class="loader" id="loader">
                <p class="progress-text">Generating your report... This may take a few minutes.</p>
            </div>
            
            <div class="result" id="result"></div>
        </div>
        
        <script>
            document.getElementById('reportForm').addEventListener('submit', async (e) => {
                e.preventDefault();
                
                const submitBtn = document.getElementById('submitBtn');
                const loader = document.getElementById('loader');
                const result = document.getElementById('result');
                
                submitBtn.disabled = true;
                submitBtn.textContent = '⏳ Processing...';
                loader.style.display = 'block';
                result.style.display = 'none';
                
                try {
                    // Upload all files first
                    const fileInputs = ['bdo_file', 'prev_ma_file', 'rent_roll_file', 
                                        'sales_tracker_file', 'prev_compliance_file', 'word_template_file'];
                    const uploadedFiles = {};
                    
                    for (const inputId of fileInputs) {
                        const input = document.getElementById(inputId);
                        if (input.files[0]) {
                            const formData = new FormData();
                            formData.append('file', input.files[0]);
                            
                            const uploadRes = await fetch('/api/v1/upload', {
                                method: 'POST',
                                body: formData
                            });
                            const uploadData = await uploadRes.json();
                            
                            if (uploadData.status !== 'success') {
                                throw new Error(`Failed to upload ${input.files[0].name}`);
                            }
                            uploadedFiles[inputId] = uploadData.filename;
                        }
                    }
                    
                    // Generate report
                    const year = document.getElementById('year').value;
                    const quarter = document.getElementById('quarter').value;
                    
                    const generateRes = await fetch('/api/v1/generate', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            year: parseInt(year),
                            quarter: parseInt(quarter),
                            bdo_file: `inputs/${uploadedFiles.bdo_file}`,
                            prev_ma_file: `inputs/${uploadedFiles.prev_ma_file}`,
                            rent_roll_file: `inputs/${uploadedFiles.rent_roll_file}`,
                            sales_tracker_file: `inputs/${uploadedFiles.sales_tracker_file}`,
                            prev_compliance_file: `inputs/${uploadedFiles.prev_compliance_file}`,
                            word_template_file: `inputs/${uploadedFiles.word_template_file}`,
                            dry_run: false
                        })
                    });
                    
                    const data = await generateRes.json();
                    
                    if (data.status === 'success') {
                        let html = '<h3>✅ Report Generated Successfully!</h3>';
                        html += `<p><strong>Execution Time:</strong> ${data.execution_time_seconds?.toFixed(1)} seconds</p>`;
                        html += '<h4>📥 Download Generated Files:</h4>';
                        
                        if (data.output_files && data.output_files.length > 0) {
                            for (const file of data.output_files) {
                                const filename = file.split('/').pop();
                                const icon = filename.endsWith('.pdf') ? '📄' : 
                                            filename.endsWith('.docx') ? '📝' : '📊';
                                html += `<a href="/api/v1/outputs/${filename}" class="download-link" download>
                                    <span class="file-icon">${icon}</span>${filename}
                                </a>`;
                            }
                        }
                        
                        if (data.warnings && data.warnings.length > 0) {
                            html += '<h4>⚠️ Warnings:</h4><ul>';
                            data.warnings.forEach(w => html += `<li>${w}</li>`);
                            html += '</ul>';
                        }
                        
                        result.className = 'result success';
                        result.innerHTML = html;
                    } else {
                        let html = '<h3>❌ Report Generation Failed</h3>';
                        if (data.errors && data.errors.length > 0) {
                            html += '<ul>';
                            data.errors.forEach(e => html += `<li>${e}</li>`);
                            html += '</ul>';
                        }
                        if (data.detail) {
                            html += `<p>${data.detail}</p>`;
                        }
                        result.className = 'result error';
                        result.innerHTML = html;
                    }
                    
                } catch (err) {
                    result.className = 'result error';
                    result.innerHTML = `<h3>❌ Error</h3><p>${err.message}</p>`;
                }
                
                submitBtn.disabled = false;
                submitBtn.textContent = '🚀 Generate Quarterly Report';
                loader.style.display = 'none';
                result.style.display = 'block';
            });
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html)


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
        
        # Resolve input file paths (handle both absolute and relative paths)
        input_files = {
            "bdo_file": request.bdo_file,
            "prev_ma_file": request.prev_ma_file,
            "rent_roll_file": request.rent_roll_file,
            "sales_tracker_file": request.sales_tracker_file,
            "prev_compliance_file": request.prev_compliance_file,
            "word_template_file": request.word_template_file,
        }
        
        resolved_files = {}
        for name, path_str in input_files.items():
            resolved_files[name] = resolve_input_path(path_str)
            logger.debug(f"[Job {job_id}] Resolved {name}: {path_str} -> {resolved_files[name]}")
        
        config = QuarterlyReportConfig(
            year=request.year,
            quarter=request.quarter,
            bdo_file=resolved_files["bdo_file"],
            previous_management_accounts=resolved_files["prev_ma_file"],
            rent_roll_file=resolved_files["rent_roll_file"],
            sales_tracker_file=resolved_files["sales_tracker_file"],
            previous_compliance_calc=resolved_files["prev_compliance_file"],
            word_template=resolved_files["word_template_file"],
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

