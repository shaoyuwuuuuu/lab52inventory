@echo off
cd /d "%~dp0"
echo Starting server...
start "" "http://localhost:8765/influencer_tracker.html"
node "%~dp0server.js"
pause
