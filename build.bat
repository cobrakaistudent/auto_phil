@echo off
REM ============================================================
REM  AutoPhil - Build standalone .exe
REM  Run this on ANY Windows machine with Python 3.6+ installed.
REM  The output .exe needs NOTHING else to run.
REM ============================================================

echo.
echo  Building AutoPhil...
echo  ===================================
echo.

pip install pyinstaller 2>nul
if errorlevel 1 (
    echo ERROR: pip install failed. Make sure Python is in PATH.
    pause
    exit /b 1
)

pyinstaller --onefile --windowed --name "AutoPhil" auto_phil.py

if errorlevel 1 (
    echo.
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo.
echo  =============================================
echo   BUILD COMPLETE
echo   Output: dist\AutoPhil.exe
echo  =============================================
echo.
echo  Copy AutoPhil.exe to any Windows machine.
echo  No Python or anything else needed.
echo.
pause
