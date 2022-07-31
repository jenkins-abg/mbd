:: Purpose:         Batch script to find and copy .c for unittest fix.
:: Parameters:      none
:: Requirements:    1. winams environment
:: Author:          prdedumo.acn
:: Version:         1.0.0 . Initial write
@echo off
setlocal enabledelayedexpansion

:: delete grepped c-source path
If Exist %~dp0tmp\c_source_extract_path.txt del /F %~dp0tmp\c_source_extract_path.txt
:: delete all copied c-source files in the tool folder
del /Q %~dp0c-source\

:: create c_source_extract_path.txt
set c_source=%~dp0tmp\c_source_extract_path.txt

:: loop for different .c in target IDs
for /F %%s in ( %~dp0tmp\source_file.txt) do (
    :: loop for to find .c in winAMS env then store in c_source_extract_path.txt
    echo:%%s
    for %%a in (%WINAMS_PATH%) do (
        if exist "%%a\" (
            dir "%%a\AP\%%s" /b /s /a-d
        )
    )>>%c_source%
)

:: loop to copy found c files to c-source folder
for /F %%s in ( %c_source%) do (
    echo:"copying %%s ...."
    xcopy %%s %~dp0c-source\ /Y
)