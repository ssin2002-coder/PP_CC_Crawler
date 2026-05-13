@echo off
REM start.bat — doc_intelligence launcher
REM Runs: python -m doc_intelligence.run
REM       (the server exits automatically when the browser tab is closed)
cd /d "%~dp0"
python -m doc_intelligence.run
