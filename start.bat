@echo off
setlocal
cd /d "%~dp0"
if errorlevel 1 goto folder_error

where py >nul 2>nul
if errorlevel 1 goto use_python
py -3.11 -c "import sys" >nul 2>nul
if errorlevel 1 goto use_python
py -3.11 -m src.main
goto finished

:use_python
where python >nul 2>nul
if errorlevel 1 goto python_error
python -m src.main

:finished
if errorlevel 1 (
    echo.
    echo Application failed to start or exited with an error.
    echo Please copy the error above or take a screenshot.
    pause
)
exit /b

:python_error
echo Python was not found. Install Python 3.11 and the project dependencies.
pause
exit /b 1

:folder_error
echo Cannot open the application folder.
pause
exit /b 1
