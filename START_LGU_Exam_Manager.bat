@echo off

setlocal

cd /d "%~dp0"

title LGU Exam Manager - Setup and Start


echo ============================================================
echo  LGU Exam Manager - Gazette v4.6.9
echo ============================================================
echo.
echo  This window sets everything up by itself, then opens
echo  the app. The first time it takes a few minutes.
echo  Please wait and do not close this window.
echo.


rem If a ready-made program file is present, use it.
if exist "%~dp0Windows App\LGU_Exam_Manager.exe" (
    start "" "%~dp0Windows App\LGU_Exam_Manager.exe"
    exit /b 0
)


rem This file must sit beside the app files. If it was opened
rem from inside the ZIP without extracting, nothing else is here.
if not exist "%~dp0app.py" (
    echo  The app files are not beside this file.
    echo  Right-click the ZIP, choose Extract All, then run
    echo  this file from the folder you extracted.
    echo.
    pause
    exit /b 1
)


rem ---- Step 1: find Python on this computer -------------------
call :find_python
if not defined PYCMD (
    echo  Python was not found on this computer. Installing it now...
    echo  This needs internet and takes a few minutes.
    echo.
    call :install_python
    call :find_python
)
if not defined PYCMD (
    echo.
    echo  Python could not be installed automatically.
    echo  Please install Python 3.10 or newer from python.org,
    echo  tick "Add Python to PATH" during install, then run
    echo  this file again.
    echo.
    pause
    exit /b 1
)
echo  Python found.
"%PYCMD%" --version


rem ---- Step 2: use one shared Python environment ---------------
rem The environment lives outside the version folder, so every
rem version reuses it. Packages are installed once, then every
rem later version opens without installing again.
set "SHARED_VENV=%USERPROFILE%\LGU-Exam-Manager\.venv"
set "VENV_PY=%SHARED_VENV%\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo.
    echo  Creating the shared app environment on this computer...
    echo  This happens once. Later versions reuse it.
    "%PYCMD%" -m venv "%SHARED_VENV%"
    if errorlevel 1 goto :fail
)
"%VENV_PY%" --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  The shared environment is broken. Creating it again...
    rmdir /s /q "%SHARED_VENV%"
    "%PYCMD%" -m venv "%SHARED_VENV%"
    if errorlevel 1 goto :fail
)


rem ---- Step 3: install the packages the app needs --------------
rem If they are already installed in the shared environment,
rem this check passes and nothing is downloaded.
"%VENV_PY%" -c "import customtkinter, fpdf, openpyxl, pandas, docx, matplotlib, PIL, python_calamine" >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Installing the packages the app needs...
    echo  This needs internet and takes a few minutes the first time.
    "%VENV_PY%" -m pip install --disable-pip-version-check -r requirements.txt
    if errorlevel 1 goto :fail
)


rem ---- Step 4: check the app files -----------------------------
echo.
echo  Checking the app files...
for %%F in (*.py) do (
    "%VENV_PY%" -m py_compile "%%F"
    if errorlevel 1 goto :fail
)


rem ---- Step 5: open the app ------------------------------------
echo.
echo  Opening LGU Exam Manager...
echo  Keep this window open while you use the app.
"%VENV_PY%" app.py
if errorlevel 1 goto :fail
exit /b 0


:find_python
set PYCMD=
py --version >nul 2>&1
if not errorlevel 1 (
    set PYCMD=py
    exit /b 0
)
python --version >nul 2>&1
if not errorlevel 1 (
    set PYCMD=python
    exit /b 0
)
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PYCMD=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    exit /b 0
)
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set "PYCMD=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    exit /b 0
)
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "PYCMD=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    exit /b 0
)
if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" (
    set "PYCMD=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    exit /b 0
)
if exist "%ProgramFiles%\Python312\python.exe" (
    set "PYCMD=%ProgramFiles%\Python312\python.exe"
    exit /b 0
)
if exist "%ProgramFiles%\Python313\python.exe" (
    set "PYCMD=%ProgramFiles%\Python313\python.exe"
    exit /b 0
)
exit /b 0


:install_python
rem Try the Windows package manager first.
where winget >nul 2>&1
if not errorlevel 1 (
    echo  Installing Python with the Windows package manager...
    winget install Python.Python.3.12 --silent --accept-source-agreements --accept-package-agreements
    if not errorlevel 1 exit /b 0
    echo  That did not work. Trying the direct download instead...
)
rem Fallback: download the official installer and run it silently.
echo  Downloading Python from python.org...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe' -OutFile '%TEMP%\python-installer.exe'"
if errorlevel 1 exit /b 1
if not exist "%TEMP%\python-installer.exe" exit /b 1
echo  Installing Python, please wait...
"%TEMP%\python-installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
exit /b 0


:fail
echo.
echo  The app could not start. Please copy the message above
echo  and send it to Abdur Rehman.
echo.
pause
exit /b 1
