:: Purpose:         Batch script to check if process is still on-going.
:: Parameters:      %1 : process name
:: Requirements:    none
:: Author:          prdedumo.acn
:: Version:         1.0.0 . Initial write
@echo off

:: loop check if process is still running
:while
tasklist /fi "imagename eq %~1" | find ":" > nul
if errorlevel 1 (
    goto :while
)

call %~dp0function\log_with_date.bat "   Done."

:: check complete