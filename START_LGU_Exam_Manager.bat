@echo off
setlocal
cd /d "%~dp0"
title LGU Exam Manager - Desktop

if exist "%~dp0Windows App\LGU_Exam_Manager.exe" (
    start "" "%~dp0Windows App\LGU_Exam_Manager.exe"
    exit /b 0
)

rem Detect a launcher opened alone from inside a ZIP before creating environments.
if not exist "%~dp0app.py" (
    echo The application files are not beside this launcher.
    echo Extract the COMPLETE ZIP to a new folder, then run this launcher from that folder.
    pause
    exit /b 1
)

rem Reuse the existing virtual environment when possible.
if not exist ".venv\Scripts\python.exe" (
    python --version >nul 2>&1
    if errorlevel 1 (
        echo Python was not found. Install Python 3.10 or newer and tick Add to PATH.
        goto :fail
    )
    echo Creating the Python environment on this PC...
    python -m venv ".venv"
    if errorlevel 1 goto :fail
)

".venv\Scripts\python.exe" --version

rem Dependencies already installed do not need to be downloaded on each launch.
".venv\Scripts\python.exe" -c "import customtkinter, fpdf, openpyxl, pandas, docx, matplotlib, PIL, python_calamine" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    ".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
    if errorlevel 1 goto :fail
)

rem Catch syntax errors before trying to open the desktop interface.
echo Verifying application source files...
".venv\Scripts\python.exe" -m py_compile app.py reports.py importer.py dashboard.py academic_engine.py academic_frame.py academic_reports.py sample_data.py diagnose.py gazette_layout.py course_reports.py course_ui.py session_store.py mapping_ui.py pdf_fonts.py draft_template.py calculation_policy.py
if errorlevel 1 goto :fail

echo Opening LGU Exam Manager. Keep this window open to see errors.
".venv\Scripts\python.exe" app.py
if errorlevel 1 goto :fail
exit /b 0

:fail
echo.
echo The application could not start. Please copy the error message above.
pause
exit /b 1
