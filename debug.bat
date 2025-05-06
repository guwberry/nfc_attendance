@echo off
set VENV_DIR=venv

:: Check if the virtual environment exists
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo Error: Virtual environment directory '%VENV_DIR%' or activate script not found.
    echo Please ensure the virtual environment is created and set up.
    echo You can create it by running 'python -m venv venv' and installing requirements.
    pause
    exit /b 1
)

:: Open a new Command Prompt with the virtual environment activated
echo Opening a new Command Prompt with the virtual environment activated...
start cmd /k "%VENV_DIR%\Scripts\activate.bat && echo Virtual environment activated. && echo Run 'python app.py' to start the application. && echo To deactivate the virtual environment, run 'deactivate'."

:: Check if the command was successful
if %ERRORLEVEL% neq 0 (
    echo Error: Failed to open a new Command Prompt.
    pause
    exit /b 1
)

echo New Command Prompt opened successfully with the virtual environment activated.
pause