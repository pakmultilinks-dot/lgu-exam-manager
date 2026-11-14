@echo off
setlocal
cd /d "%~dp0"
if exist "%~dp0Windows App\LGU_Exam_Manager.exe" (
    "%~dp0Windows App\LGU_Exam_Manager.exe" --self-check --self-check-output "%~dp0LGU_Self_Check.json"
) else (
    if not exist "%~dp0.venv\Scripts\python.exe" goto missing
    "%~dp0.venv\Scripts\python.exe" "%~dp0app.py" --self-check --self-check-output "%~dp0LGU_Self_Check.json"
)
if errorlevel 1 goto failed
if not exist "%~dp0LGU_Self_Check.json" goto failed
echo Self-check completed. Result:
type "%~dp0LGU_Self_Check.json"
pause
exit /b 0
:missing
echo Extract the complete application ZIP before running this check.
pause
exit /b 1
:failed
echo Self-check did not complete successfully. Any existing result may be from an earlier run.
echo Check application.log in your user folder under .lgu-exam-manager for this build's failure details.
pause
exit /b 1
