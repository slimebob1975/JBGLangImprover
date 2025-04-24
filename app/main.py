from fastapi import FastAPI, File, UploadFile, Request, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import os
import shutil
import logging
from datetime import datetime
import time
from app.src.JBGLanguageImprover import JBGLanguageImprover

KEEP_FILES_HOURS = 1

def clean_old_uploads(directory, logger, max_age_hours=KEEP_FILES_HOURS):
    now = time.time()
    max_age_seconds = max_age_hours * 3600

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
    logger.propagate = False  # 🔴 Prevent duplicate output from parent loggers

    # Clear any previous handlers for safety
    logger.handlers.clear()

    # File handler for this run
    file_handler = logging.FileHandler(log_path, mode='w', encoding='utf-8')
    file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

# We are using FastAPI for the backend
app = FastAPI()

# Directories
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_DIR = "logs"
os.makedirs(LOG_DIR, exist_ok=True)

# Mount static folders
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/upload/")
async def upload_file(file: UploadFile = File(...), api_key: str = Form(...), model: str = Form(...), \
    custom_prompt: str = Form(""), temperature: float = Form(0.7), include_motivations: bool = Form(True), \
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
    clean_old_uploads(UPLOAD_DIR, logger)
    logger.info(f"📥 Received file: {file.filename}")
    logging.info(f"Saving to: {input_path}")
    logger.info(f"🌡️ Temperature: {temperature}")
    logger.info(f"💬 Include motivations: {include_motivations}")
    logger.info(f"📝 DOCX markup mode: {docx_mode}")
    
    # Load base prompt policy from file
    with open(os.path.join(BASE_DIR, "policy", "prompt_policy.md"), encoding="utf-8") as f:
        base_prompt = f.read().strip()

    # Merge with custom prompt if provided
    full_prompt = base_prompt
    if custom_prompt:
        full_prompt += "\n\nSpecifika instruktioner:\n" + custom_prompt.strip()

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
    
