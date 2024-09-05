@echo off
cd %~dp0
set cwd=%cd%

if exist "%cwd%\.venv\Scripts\python.exe" (
    %cwd%\.venv\Scripts\python.exe %cwd%\robots\main.py
) else if exist "%cwd%\venv\Scripts\python.exe" (
    %cwd%\venv\Scripts\python.exe %cwd%\robots\main.py
) else (
    echo No virtual environment found in .venv or venv folders.
    exit /b 1
)