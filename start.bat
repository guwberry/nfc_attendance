@echo off
set VENV_DIR=venv
set HOST=192.168.156.189
set PORT=5000

:: Check if the virtual environment exists
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo Error: Virtual environment directory '%VENV_DIR%' or activate script not found.
    echo Please ensure the virtual environment is created and set up.
    echo You can create it by running 'install_requirements.bat'.
    pause
    exit /b 1
)

:: Activate the virtual environment
echo Activating virtual environment...
call %VENV_DIR%\Scripts\activate
if %ERRORLEVEL% neq 0 (
    echo Error: Failed to activate virtual environment.
    pause
    exit /b 1
)

:: Check if uvicorn is installed
echo Checking if uvicorn is installed...
uvicorn --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo uvicorn is not installed. Installing uvicorn...
    pip install uvicorn==0.23.2
    if %ERRORLEVEL% neq 0 (
        echo Error: Failed to install uvicorn.
        pause
        exit /b 1
    )
    echo uvicorn installed successfully.
) else (
    echo uvicorn is already installed.
)

:: Check if app.py exists
if not exist "app.py" (
    echo Error: app.py not found in the current directory.
    pause
    exit /b 1
)

:: Start the application with uvicorn
echo Starting the application with uvicorn on %HOST%:%PORT%...
uvicorn app:asgi_app --host %HOST% --port %PORT% --reload
if %ERRORLEVEL% neq 0 (
    echo Error: Failed to start the application with uvicorn.
    pause
    exit /b 1
)

:: Normally, this line won't be reached because uvicorn runs in the foreground.
:: If uvicorn exits unexpectedly, the script will pause to show the error.
pause