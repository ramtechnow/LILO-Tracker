@REM @echo off

@REM py "D:\New folder (2)\working python projects\LILO Tracker\LILO_Main.py"

@echo off
REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed. Please install Python and try again.
    pause
    exit /b
)

REM Ensure pip is up-to-date
echo Updating pip...
python -m pip install --upgrade pip

REM Install required Python modules
echo Installing required modules...
python -m pip install tkinter pandas openpyxl smtplib

REM Run the Python script
echo Starting the LILO Tracker...
py "D:\New folder (2)\working python projects\LILO Tracker\LILO_Main.py"
