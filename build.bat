@echo off
REM ============================================================
REM  AutoPhil 2 - Build standalone .exe
REM  Run this on ANY Windows machine with Python 3.6+ installed.
REM  The output .exe needs NOTHING else to run.
REM ============================================================

echo.
echo  Building AutoPhil 2...
echo  ===================================
echo.

pip install pyinstaller 2>nul
if errorlevel 1 (
    echo ERROR: pip install failed. Make sure Python is in PATH.
    pause
    exit /b 1
)

pyinstaller --onefile --windowed --name "AutoPhil2" auto_phil2.py

if errorlevel 1 (
    echo.
    echo ERROR: Build failed.
    pause
    exit /b 1
)

echo.
echo  =============================================
echo   BUILD COMPLETE
echo   Output: dist\AutoPhil2.exe
echo  =============================================
echo.
echo  Copy AutoPhil2.exe to any Windows machine.
echo  No Python or anything else needed.
echo.
echo  If Verizon endpoint security blocks the .exe,
echo  run the source directly: python auto_phil2.py
echo.
pause
