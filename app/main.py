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

# Mount static folders
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/get_editable_prompt/")
def get_editable_prompt():
    editable_part, _, _ = load_prompt_parts(os.path.join(BASE_DIR, "policy", "prompt_policy.md"))
    return {"editable_prompt": editable_part}

@app.post("/upload/")
async def upload_file(file: UploadFile = File(...), api_key: str = Form(...), model: str = Form(...), \
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
    logging.info(f"Saving to: {input_path}")
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
    
