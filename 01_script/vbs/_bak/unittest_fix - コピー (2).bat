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
:: Version:         1.0.2 + feature: Changes argument to process zip input and extract it in local PC.
::			  + feature: Add zipping archive at the end of the process.
:: 		    1.0.1 + feature: Changes argument assignment due to keisoku path
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
::rmdir %job_path%\tmp /q /s 2>nul
::rmdir %job_path%\winams /q /s 2>nul

:: create and initialize tmp folder
::mkdir %job_path%\tmp
::mkdir %job_path%\winams

:: extract IDs
::"C:\Program Files\7-Zip\7z.exe" x -tzip %atg_path% -o%job_path%\tmp
:: extract winAMS
::"C:\Program Files\7-Zip\7z.exe" x -tzip %3 -o%job_path%\winams

call %~dp0function\log_with_date.bat "   Done."
:: sleep for 5 seconds before proceeding
ping -n 5 127.0.0.1 > nul

:::::::::::::::::::::::::::::::::::::::
:: STAGE 1B: VARIABLE PREP AND CHECK ::
:::::::::::::::::::::::::::::::::::::::
:: create first instance of log file
call %~dp0function\init_log.bat
call %~dp0function\log_with_date.bat "  stage_1B_variable_prep_and_check begin..."

::  pass bat parameter values to initialize script-internal variables
set transfer_atg_path=%2
set compo=%4
set model_ver=%5

set test_editor=%~dp0..\02_testsheet_editor\testsheet_editor.xlsm
set static_editor=%~dp0..\02_testsheet_editor\テストシート修正マクロ.xlsm

:: loop to get number of folders in input ATG and set folder name
set /a id_folder_count=0
for /D %%a in ("C:\jenkins\workspace\DSS_TestSheetFix\tmp\*") do (
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

    for %%f in (%job_path%\tmp\!atg_foldername[%%a]!\*) do (
        echo.%%f | findstr  /C:%func% 1>nul
	    if errorlevel 1 (
       	        echo:searching...
            ) else (
                :: set func file path
		set full_kanri=%%f
		echo:%%f
	    ) 

        echo.%%f | findstr  /C:%keisoku% 1>nul
	    if errorlevel 1 (
       	        echo:searching...
            ) else (
                :: set func file path
		set keisoku_file=%%f
		echo:%%f
	    ) 
    )
    call %~dp0function\log_with_date.bat "   Done."

::::::::::::::::::::::::::::::::::::
:: STAGE 2: CREATE PARAMETER FILE ::
::::::::::::::::::::::::::::::::::::
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

::::::::::::::::::::::::::::::::::::::::
:: STAGE 3: POPULATE TESTSHEET EDITOR ::
::::::::::::::::::::::::::::::::::::::::
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

if not exist %test_editor% (
    call %~dp0function\log_with_date.bat " File: %test_editor% "
    call %~dp0function\log_with_date.bat " File is not existing, terminating build..."
:error_exit
    exit /b 1
)

:::::::::::::::::::::::::::::
:: STAGE 5: FIX TEST SHEET ::
:::::::::::::::::::::::::::::
    call %~dp0function\log_with_date.bat "  stage_5_test_fix begin..."
:: call vbs to perform testheet fixing
    cscript.exe %~dp0vbs\test_fix.vbs !full_kanri! %test_editor% !keisoku_file!
    if !ERRORLEVEL! NEQ 0 (
        call %~dp0function\log_with_date.bat "  Error in stage_5_test_fix, terminating build..."
        goto error_exit
    )
:: loop check if vbs is still running
    call %~dp0process_loop_check.bat "test_fix.vbs"

:: check if there is winAMS environment
if exist %WINAMS_PATH% (
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
set LOGTIME=%CUR_DATE%%TIME%
set LOGTIME=!LOGTIME:-=!
set LOGTIME=!LOGTIME::=!
set LOGTIME=!LOGTIME:.=!

::copy log.txt to server
xcopy /s /i %~dp0log %job_path%\tmp\!atg_foldername[%%a]!_UnitTestFixResult_log /Y

)

::zipped output
echo:zipping...


:: replace previous zip filename
for %%F in ( %atg_path% ) do set filename=%%~nF
set process_output=_TestSheetFixResult
set output_filename=%filename:_ATGResult=!process_output!%

"C:\Program Files\7-Zip\7z.exe" a -tzip %transfer_atg_path%\%output_filename%.zip %job_path%\tmp\* -sdel -y




