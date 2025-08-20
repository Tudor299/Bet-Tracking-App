import subprocess
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pathlib import Path
import shutil

app = FastAPI()
BASE_DIR = Path(__file__).resolve().parent
app.mount("/static", StaticFiles(directory=BASE_DIR), name="static")

@app.get("/")
async def root():
    return FileResponse("index.html")

@app.post("/upload_images")
async def upload_images(files: list[UploadFile] = File(...)):
    saved_files = []
    for file in files:
        file_path = BASE_DIR / file.filename
        with open(file_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
        saved_files.append(file.filename)
    return {"saved": saved_files}

@app.get("/refresh_bets")
def extract_bet_info():
    result = subprocess.run(["python3", "extract_bet_info.py"], capture_output=True, text=True)
    return {"output": result.stdout}

@app.get("/refresh_teams")
def get_teams_stats():
    result = subprocess.run(["python3", "get_teams_stats.py"], capture_output=True, text=True)
    return {"output": result.stdout}