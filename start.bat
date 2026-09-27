@echo off
title Marsh AI Corporate Pitch Generator & Compliance Platform
echo ======================================================================
echo           MARSH AI MULTI-AGENT PITCH & COMPLIANCE PLATFORM
echo ======================================================================
echo.

:: 1. Check Python installation
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python is not installed or not in PATH. Please install Python 3.10+.
    pause
    exit /b 1
)

:: 2. Activate virtual environment if present
if exist "venv\Scripts\activate.bat" (
    echo [INFO] Activating virtual environment 'venv'...
    call venv\Scripts\activate.bat
) else (
    echo [WARNING] 'venv' not found. Running with global python environment...
)

:: 3. Verify .env file existence
if not exist ".env" (
    echo [WARNING] '.env' file not found!
    echo Please make sure GROQ_API_KEY or OPENAI_API_KEY is defined in .env.
    echo.
)

:: 4. Start FastAPI server
echo [INFO] Launching Marsh AI application on http://127.0.0.1:8000 ...
echo [INFO] Press Ctrl+C in this terminal to stop the server.
echo.

python server.py

pause
