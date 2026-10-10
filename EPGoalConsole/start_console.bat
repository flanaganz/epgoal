@echo off
cd /d C:\EPGoal\EPGoalConsole
if not exist .venv\Scripts\python.exe (
  py -m venv .venv
  .venv\Scripts\python.exe -m pip install --upgrade pip
  .venv\Scripts\python.exe -m pip install -r requirements.txt
)
start "EPGoal Console" http://127.0.0.1:8088
.venv\Scripts\python.exe app.py
pause
