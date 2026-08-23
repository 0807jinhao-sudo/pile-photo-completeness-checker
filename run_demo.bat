@echo off
setlocal
cd /d "%~dp0"

python scripts\create_demo_data.py
if errorlevel 1 goto :error

python pile_photo_checker.py examples\demo_input --output examples\demo_output
if errorlevel 1 goto :error

echo.
echo Demo complete. Open examples\demo_output\pile_photo_completeness.xlsx
pause
exit /b 0

:error
echo.
echo Demo failed. Please install dependencies with: pip install -r requirements.txt
pause
exit /b 1
