import subprocess
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/refresh_bets")
def extract_bet_info():
    result = subprocess.run(["python3", "extract_bet_info.py"], capture_output=True, text=True)
    return {"output": result.stdout}

@app.get("/refresh_teams")
def get_teams_stats():
    result = subprocess.run(["python3", "get_teams_stats.py"], capture_output=True, text=True)
    return {"output": result.stdout}