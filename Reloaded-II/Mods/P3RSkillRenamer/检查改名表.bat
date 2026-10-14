@echo off
REM ============================================================
REM  P3R skill renamer - table checker.
REM  Edit the .xlsx in this folder first, then run this file.
REM ============================================================
setlocal
pushd "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "check-rename.ps1"
popd
echo.
pause