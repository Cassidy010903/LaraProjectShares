@echo off
title Hammerfall Website Demo
cd /d "%~dp0demo"
echo.
echo  Hammerfall Website Demo is running at http://localhost:8090/
echo  Your browser opens in a moment. Keep this window open while you use the demo.
echo  Close this window to stop it.
echo.
start "" "http://localhost:8090/"
python -m http.server 8090 --bind 127.0.0.1
if errorlevel 1 py -m http.server 8090 --bind 127.0.0.1
if errorlevel 1 (
  echo.
  echo  Python was not found. Install Python from https://www.python.org/downloads/ and try again.
  pause
)
