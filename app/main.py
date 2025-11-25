from fastapi import FastAPI, File, UploadFile, Request, Form, Response, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from dotenv import load_dotenv
from starlette.middleware.base import BaseHTTPMiddleware
import os, sys
import shutil
import logging
from datetime import datetime
import time
from app.src.JBGLanguageImprover import JBGLanguageImprover
import uuid

KEEP_FILES_DAYS = 0
KEEP_FILES_HOURS = 1

def clean_old_files(directory, logger, max_age_days=KEEP_FILES_DAYS, max_age_hours=KEEP_FILES_HOURS):
    now = time.time()
    max_age_seconds = (max_age_days * 24  + max_age_hours) * 3600  

    num_old_files_deleted = 0
    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        if os.path.isfile(filepath):
            file_age = now - os.path.getmtime(filepath)
            if file_age > max_age_seconds:
                try:
                    os.remove(filepath)
                    logger.info(f"🧹 Deleted old file: {filename}")
                    num_old_files_deleted +=1
                except Exception as e:
                    logger.warning(f"⚠️ Could not delete {filename}: {e}")
    
    if num_old_files_deleted > 0:
        logger.info(f"🧼 Upload cleanup completed: {num_old_files_deleted} files deleted.")

def setup_run_logger(log_path):
    logger = logging.getLogger(f"run-{log_path}")
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    logger.handlers.clear()

    # File
    file_handler = logging.FileHandler(log_path, mode='w', encoding='utf-8')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    # Console (stdout)
    console_handler = logging.StreamHandler(sys.stdout)  # <== explicitly stdout
    console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

def load_prompt_parts(prompt_path):
    """
    Splits the prompt policy into editable and locked parts cleanly.
    Removes all comment markers like <!-- START_EDITABLE -->, <!-- START_LOCKED -->, etc.
    """
    with open(prompt_path, encoding="utf-8") as f:
        content = f.read()

    start_edit = content.find("<!-- START_EDITABLE -->")
    end_edit = content.find("<!-- END_EDITABLE -->")

    if start_edit == -1 or end_edit == -1:
        raise ValueError("Prompt policy missing START/END_EDITABLE markers!")

    # Extract parts
    editable_part = content[start_edit + len("<!-- START_EDITABLE -->"):end_edit].strip()
    locked_part_before = content[:start_edit].strip()
    locked_part_after = content[end_edit + len("<!-- END_EDITABLE -->"):].strip()

    # Remove any other markers inside locked parts
    for marker in ["<!-- START_LOCKED -->", "<!-- END_LOCKED -->",
                   "<!-- START_EDITABLE -->", "<!-- END_EDITABLE -->"]:
        locked_part_before = locked_part_before.replace(marker, "").strip()
        locked_part_after = locked_part_after.replace(marker, "").strip()

    return editable_part, locked_part_before, locked_part_after

def validate_prompt(editable_part: str):
    """
    Validates the user-edited prompt to ensure it does not accidentally include forbidden system markers.
    Raises ValueError if illegal markers are found.
    """
    forbidden_tags = [
        "<!-- START_EDITABLE -->",
        "<!-- END_EDITABLE -->",
        "<!-- START_LOCKED -->",
        "<!-- END_LOCKED -->",
    ]

    for tag in forbidden_tags:
        if tag in editable_part:
            raise ValueError(f"❌ Forbidden tag '{tag}' found in editable prompt area!")

    return True  # If OK


# We are using FastAPI for the backend
app = FastAPI()

# Directories
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)
RESULTS_DIR = "results"
os.makedirs(RESULTS_DIR, exist_ok=True)

# In-memory job status store for language improvement jobs
# job_id -> {"status": str, "done": bool, "error": Optional[str]}
jobs_lang = {}

# Mount static folders
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

class FrameOptionsMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response: Response = await call_next(request)
        # Remove 'x-frame-options' if it exists (case-insensitive)
        if "x-frame-options" in response.headers:
            del response.headers["x-frame-options"]
        # Allow embedding from any origin (use with caution)
        response.headers["Content-Security-Policy"] = f"frame-ancestors {FRAME_ANCESTORS}"    
        return response

# Allow embedding via iframes
load_dotenv()
FRAME_ANCESTORS = os.getenv("FRAME_ANCESTORS", "*")
if not FRAME_ANCESTORS or FRAME_ANCESTORS == "*":
    print("⚠️ Warning: Using default FRAME_ANCESTORS='*'. Set in .env (localhost) or Azure App Settings (deployed).")
app.add_middleware(FrameOptionsMiddleware)

@app.get("/config")
def get_config():
    title = os.getenv("APP_TITLE", "JBG Klarspråkning")
    print(f" Appens titel: {title}")
    return {"title": title}

@app.get("/get_editable_prompt/")
def get_editable_prompt():
    editable_part, _, _ = load_prompt_parts(os.path.join(BASE_DIR, "policy", "prompt_policy.md"))
    return {"editable_prompt": editable_part}

@app.post("/upload/")
async def upload_file(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    api_key: str = Form(...),
    model: str = Form(...),
    editable_prompt: str = Form(""),
    temperature: float = Form(0.7),
    include_motivations: bool = Form(True),
    docx_mode: str = Form("simple")
):
    # Generate job ID and paths
    job_id = str(uuid.uuid4())
    input_path = os.path.join(UPLOAD_DIR, job_id + "_" + file.filename)
    output_path = os.path.join(RESULTS_DIR, f"{job_id}.result.docx")
    log_path = os.path.join(LOG_DIR, f"{job_id}.log")

     # Register job with initial status
    jobs_lang[job_id] = {
        "status": "Fil uppladdad. Bearbetar dokument...",
        "done": False,
        "error": None,
    }

    # Save uploaded file
    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # Setup logger and cleanup old files
    logger = setup_run_logger(log_path)
    clean_old_files(UPLOAD_DIR, logger, max_age_days=0, max_age_hours=1)
    clean_old_files(LOG_DIR, logger, max_age_days=30, max_age_hours=0)
    clean_old_files(RESULTS_DIR, logger, max_age_days=0, max_age_hours=1)
    logger.info(f"📥 Received file: {file.filename}")
    logger.info(f"Saving to: {input_path}")
    logger.info(f"🌡️ Temperature: {temperature}")
    logger.info(f"💬 Include motivations: {include_motivations}")
    logger.info(f"📝 DOCX markup mode: {docx_mode}")

    # Extract prompt parts and validate
    editable_part = editable_prompt.strip()
    _, locked_before, locked_after = load_prompt_parts(os.path.join(BASE_DIR, "policy", "prompt_policy.md"))
    full_prompt = f"{locked_before}\n\n{editable_part}\n\n{locked_after}"
    validate_prompt(full_prompt)

        # Background processing
    def run_language_improvement():
        def progress_callback(message: str):
            job = jobs_lang.get(job_id)
            if job is not None:
                job["status"] = message

        try:
            improver = JBGLanguageImprover(
                input_path=input_path,
                api_key=api_key,
                model=model,
                prompt_policy=full_prompt,
                temperature=temperature,
                include_motivations=include_motivations,
                docx_mode=docx_mode,
                logger=logger,
                progress_callback=progress_callback,
            )
            result = improver.run()
            shutil.move(result, output_path)

            job = jobs_lang.get(job_id)
            if job is not None:
                job["done"] = True
                if not job.get("status"):
                    job["status"] = "Språkgranskning klar."
                job["error"] = None

        except Exception as e:
            logger.error(f"Language improvement error for job {job_id}: {e}")
            job = jobs_lang.setdefault(job_id, {})
            job["done"] = True
            job["error"] = str(e)
            job["status"] = "Ett fel uppstod vid språkgranskningen."

    background_tasks.add_task(run_language_improvement)

    return JSONResponse({
    "job_id": job_id,
    "original_filename": file.filename
})

@app.get("/status/{job_id}")
def check_status(job_id: str):
    output_path = os.path.join(RESULTS_DIR, f"{job_id}.result.docx")
    job = jobs_lang.get(job_id)

    # If we know about the job, use that as source of truth
    if job is not None:
        # Error recorded
        if job.get("error"):
            return JSONResponse(
                {
                    "done": True,
                    "status": job.get("status") or "Ett fel uppstod vid språkgranskningen.",
                    "error": job["error"],
                },
                status_code=500,
            )

        # Not done yet
        if not job.get("done"):
            return JSONResponse(
                {
                    "done": False,
                    "status": job.get("status") or "Bearbetar dokument...",
                },
                status_code=202,
            )

        # Done according to job state → we expect file to exist
        if not os.path.exists(output_path):
            return JSONResponse(
                {
                    "done": True,
                    "status": "Jobbet är markerat som klart men resultatfilen saknas.",
                    "error": "Result file missing",
                },
                status_code=500,
            )

        return {
            "done": True,
            "status": job.get("status") or "Språkgranskning klar.",
        }

    # If job is not in memory (e.g. another instance), fall back to file existence

    if os.path.exists(output_path):
        return {
            "done": True,
            "status": "Språkgranskning klar.",
        }

    # No job and no file: assume processing (or very early state)
    return JSONResponse(
        {
            "done": False,
            "status": "Bearbetar dokument...",
        },
        status_code=202,
    )


@app.get("/download/{job_id}")
def download_result(job_id: str):
    output_path = os.path.join(RESULTS_DIR, f"{job_id}.result.docx")
    if not os.path.exists(output_path):
        return {"error": "Result not found yet"}
    return FileResponse(path=output_path, filename=f"{job_id}.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

async def upload_file_old(file: UploadFile = File(...), api_key: str = Form(...), model: str = Form(...), \
    editable_prompt: str = Form(""), temperature: float = Form(0.7), include_motivations: bool = Form(True), \
    docx_mode: str = Form("simple")):
    
    # Create fresh timestamp for each request
    # Create timestamped log file for this run
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_path = os.path.join(LOG_DIR, f"run-{timestamp}.log")
    
    # Save uploaded file
    input_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(input_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
        
    # Log upload and saving path
    logger = setup_run_logger(log_path)
    clean_old_files(UPLOAD_DIR, logger, max_age_days=0, max_age_hours=1)
    clean_old_files(LOG_DIR, logger, max_age_days=7, max_age_hours=0)
    logger.info(f"📥 Received file: {file.filename}")
    logger.info(f"Saving to: {input_path}")
    logger.info(f"🌡️ Temperature: {temperature}")
    logger.info(f"💬 Include motivations: {include_motivations}")
    logger.info(f"📝 DOCX markup mode: {docx_mode}")
    
    # Load base prompt policy from file
    # Instead of opening whole prompt_policy.md
    editable_part = editable_prompt.strip()
    _, locked_before, locked_after = load_prompt_parts(os.path.join(BASE_DIR, "policy", "prompt_policy.md"))

    # Build full prompt
    full_prompt = f"{locked_before}\n\n{editable_part}\n\n{locked_after}"
    validate_prompt(full_prompt)

    # Run language improvement pipeline
    improver = JBGLanguageImprover(
        input_path=input_path,
        api_key=api_key,
        model=model,
        prompt_policy=full_prompt,
        temperature=temperature,
        include_motivations=include_motivations,
        docx_mode=docx_mode,
        logger=logger
    )
    output_path = improver.run()

    return FileResponse(
        path=output_path,
        filename=os.path.basename(output_path),
        media_type='application/octet-stream'
    )
    
