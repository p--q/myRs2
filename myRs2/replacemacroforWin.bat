@echo off
cd tools
echo.
echo Do not close the error dialog until finished.
echo.
"C:\Program Files\LibreOffice\program\python.exe" replaceEmbeddedScripts.py
cd ..
echo.
echo Finished.
echo.
pause