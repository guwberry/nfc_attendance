@echo off
:: Get local IPv4 address
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4 Address"') do (
    echo %%a | findstr /r "[0-9]" >nul
    if not errorlevel 1 (
        set "ip=%%a"
    )
)

:: Trim leading space
for /f "tokens=* delims= " %%b in ("%ip%") do set ip=%%b

:: Run in a new terminal window
start cmd /k "call venv\Scripts\activate && echo Using IP: %ip% && uvicorn app:asgi_app --host %ip% --port 5000 --reload"
