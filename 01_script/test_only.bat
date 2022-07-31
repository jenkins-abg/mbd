:: Purpose:         Batch script containing all commands to fix and prepare unittest design.
::                  In case a parameter is missing, it will result to an error and create log file
:: Parameters:      %1 : initial path of ATG IDs
::                  %2 : path where target IDs will be copied
::                  %3 : winams environment
::                  %4 : project component
::                  %5 : project model version
::                  %6 : kanri filename
::                  %7 : keisoku file
:: Requirements:    1. input_param.vbs
::                  2. transfer_target_ids.vbs
::                  3. test_fix.vbs
:: Author:          prdedumo.acn
:: Version:         1.0.3 + feature: Prcocesses multiple folders in zip input.
::                  1.0.2 + feature: Changes argument to process zip input and extract it in local PC.
::                        + feature: Add zipping archive at the end of the process.
::                  1.0.1 + feature: Changes argument assignment due to keisoku path
::                  1.0.0 . Initial write
@echo off
setlocal enabledelayedexpansion

:::::::::::::::::::::::::::::::::
:: STAGE 1A: EXTRACT ZIP FILES ::
:::::::::::::::::::::::::::::::::
:: create first instance of log file
call %~dp0function\init_log.bat
call %~dp0function\log_with_date.bat "  stage_1A_extracting environments begin..."

:: set variables
set atg_path=%1
set job_path=C:\jenkins\workspace\DSS_TestSheetFix

:: delete existing folder
echo:deleting %job_path%\tmp ...
rmdir %job_path%\tmp /q /s 2>nul
echo:deleting %job_path%\winams ...
rmdir %job_path%\winams /q /s 2>nul

:: create and initialize tmp folder
echo:creating %job_path%\tmp ...
mkdir %job_path%\tmp
echo:creating %job_path%\winams ...
rem mkdir %job_path%\winams

:: extract IDs
"C:\Program Files\7-Zip\7z.exe" x -tzip %atg_path% -o%job_path%\tmp
:: extract winAMS
rem "C:\Program Files\7-Zip\7z.exe" x -tzip %3 -o%job_path%\winams

call %~dp0function\log_with_date.bat "   Done."
:: sleep for 5 seconds before proceeding
ping -n 5 127.0.0.1 > nul

set process_output=_TestSheetFixResult
set prv_process=_ATGResult

:: replace previous zip foldername
for /D %%F in ( %job_path%\tmp\* ) do (
    set foldername=%%~nxF
    echo.!foldername! | findstr /C:%prv_process% 1>nul
    if errorlevel 1 (
set output_foldername=!foldername!%process_output%
    ) else (
set output_foldername=!foldername:_ATGResult=%process_output%!
    )

rename %job_path%\tmp\!foldername! !output_foldername!
)

:::::::::::::::::::::::::::::::::::::::
:: STAGE 1B: VARIABLE PREP AND CHECK ::
:::::::::::::::::::::::::::::::::::::::
call %~dp0function\log_with_date.bat "  stage_1B_variable_prep_and_check begin..."

::  pass bat parameter values to initialize script-internal variables
set transfer_atg_path=%2
set compo=%4
set model_ver=%5

set test_editor=%~dp0..\02_testsheet_editor\testsheet_editor.xlsm
set static_editor=%~dp0..\02_testsheet_editor\テストシート修正マクロ.xlsm

:: loop to get number of folders in input ATG and set folder name
set /a id_folder_count=0
for /D %%a in ("%job_path%\tmp\*") do (
    set /a id_folder_count=id_folder_count+1
    set atg_foldername[!id_folder_count!]=%%~nxa
)

:: loop to get number of folders in winAMS and set folder name
set /a winams_folder_count=0
for /D %%w in (%job_path%\winams\*) do (
    set /a winams_folder_count=winams_folder_count+1
    set winAMS_foldername[!winams_folder_count!]=%%~nxw
)

set tmp_csv=%~dp0tmp\tmp.csv
set func=関数一覧表
set keisoku=計測適合一覧表
set WINAMS_PATH=%job_path%\winams\!winAMS_foldername[1]!
set code_path=%~dp0c-source

:: process data per ID folders
for /L  %%a in (1,1,!id_folder_count!) Do (
    :: set kanri file path
    set files_path=%job_path%\tmp\!atg_foldername[%%a]!
    set full_kanri=!files_path!\%6
    set keisoku_file=!files_path!\%7

rem     for %%f in (%job_path%\tmp\!atg_foldername[%%a]!\*) do (
rem         echo.%%f | findstr  /C:%func% 1>nul
rem         if errorlevel 1 (
rem             echo:searching...
rem         ) else (
rem             :: set func file path
rem             set full_kanri=%%f
rem             echo:%%f
rem         )

rem         echo.%%f | findstr  /C:%keisoku% 1>nul
rem         if errorlevel 1 (
rem             echo:searching...
rem         ) else (
rem             :: set func file path
rem             set keisoku_file=%%f
rem             echo:%%f
rem         )
rem     )
    call %~dp0function\log_with_date.bat "   Done."

rem::::::::::::::::::::::::::::::::::::
rem:: STAGE 2: CREATE PARAMETER FILE ::
rem::::::::::::::::::::::::::::::::::::
    call %~dp0function\log_with_date.bat "  stage_2_creating_parameter begin..."

    :: transfer param details from to csv
    echo:%atg_path%>%tmp_csv%
    echo:%transfer_atg_path%>>%tmp_csv%
    echo:%compo%>>%tmp_csv%
    echo:%model_ver%>>%tmp_csv%
    echo:!full_kanri!>>%tmp_csv%
    echo:!keisoku_file!>>%tmp_csv%
    echo:!WINAMS_PATH!>>%tmp_csv%
    echo:!files_path!>>%tmp_csv%

    call %~dp0function\log_with_date.bat "   Done."

rem::::::::::::::::::::::::::::::::::::::::
rem:: STAGE 3: POPULATE TESTSHEET EDITOR ::
rem::::::::::::::::::::::::::::::::::::::::
    call %~dp0function\log_with_date.bat "  stage_3_populate_testeditor begin..."
:: Call input_param.vbs
    cscript.exe %~dp0vbs\input_param.vbs %test_editor% %tmp_csv%
:: check if there is an error
    if !ERRORLEVEL! NEQ 0 (
        call %~dp0function\log_with_date.bat "  Error in stage_3_populate_testeditor, terminating build..."
        goto error_exit
    )
:: loop check if vbs is still running
    call %~dp0process_loop_check.bat "input_param.vbs"
    call %~dp0function\log_with_date.bat "   Done."

    if not exist !keisoku_file! (
        call %~dp0function\log_with_date.bat " File: !keisoku_file! "
        call %~dp0function\log_with_date.bat " Keisoku File is not existing, terminating build..."
        goto error_exit
    )
    if not exist !full_kanri! (
        call %~dp0function\log_with_date.bat " File: !full_kanri! "
        call %~dp0function\log_with_date.bat " 関数一覧表 File is not existing, terminating build..."
        goto error_exit
    )
    if not exist %test_editor% (
        call %~dp0function\log_with_date.bat " File: %test_editor% "
        call %~dp0function\log_with_date.bat " File is not existing, terminating build..."
        goto error_exit
    )

rem:::::::::::::::::::::::::::::
rem:: STAGE 4: FIX TEST SHEET ::
rem:::::::::::::::::::::::::::::
    call %~dp0function\log_with_date.bat "  stage_4_test_fix begin..."
:: call vbs to perform testheet fixing
    cscript.exe %~dp0vbs\test_fix.vbs !full_kanri! %test_editor% !keisoku_file! %~dp0log
    if !ERRORLEVEL! NEQ 0 (
        call %~dp0function\log_with_date.bat "  Error in stage_4_test_fix, terminating build..."
        goto error_exit
    )
:: loop check if vbs is still running
    call %~dp0process_loop_check.bat "test_fix.vbs"

:: check if there is winAMS environment
    if "%~3" NEQ "" (
    :: fix for static variables
        call %~dp0find_c.bat
    :: loop check if bat is still running
        call %~dp0process_loop_check.bat "find_c.bat"

    :: call vbs to perform static fixing
        cscript.exe %~dp0vbs\test_fix_static.vbs %static_editor% !full_kanri! !keisoku_file! %code_path% %job_path%\tmp\!atg_foldername[%%a]!
    :: check if there is an error
        if !ERRORLEVEL! NEQ 0 (
            call %~dp0function\log_with_date.bat "  Error in stage_5_static_fix, terminating build..."
            goto error_exit
        )
    :: loop check if vbs is still running
        call %~dp0process_loop_check.bat "test_fix_static.vbs"
    )
rem set LOGTIME=%CUR_DATE%%TIME%
rem set LOGTIME=!LOGTIME:-=!
rem set LOGTIME=!LOGTIME::=!
rem set LOGTIME=!LOGTIME:.=!

rem ::copy log.txt to server
rem ::xcopy /s /i %~dp0log %job_path%\tmp\!atg_foldername[%%a]!_UnitTestFixResult_log /Y

)

call %~dp0function\log_with_date.bat "   Done."

::zipped output
echo:zipping...

:: replace previous zip foldername
for %%F in ( %atg_path% ) do (
    set foldername=%%~nF
    echo.!foldername! | findstr /C:%prv_process% 1>nul
    if errorlevel 1 (
set output_foldername=!foldername!%process_output%
    ) else (
set output_foldername=!foldername:_ATGResult=%process_output%!
    )
)

::set output_foldername=%foldername:_ATGResult=!process_output!%

"C:\Program Files\7-Zip\7z.exe" a -tzip %transfer_atg_path%\%output_foldername%.zip %job_path%\tmp\* -sdel -y
::copy log.txt to server
xcopy /s /i %~dp0log %transfer_atg_path% /Y
exit /b 0

::copy log.txt to server
:error_exit
xcopy /s /i %~dp0log %transfer_atg_path% /Y
exit /b 1
