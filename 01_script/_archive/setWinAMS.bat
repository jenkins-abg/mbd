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
:: Version:         1.0.1 + Changes argument assignment due to keisoku path
::                  1.0.0 . Initial write
@echo off
setlocal enabledelayedexpansion

set WINAMS_PATH="C:\jenkins_edc\4_unittest_fix\ver5_2\03_test\input\ACL_ATGResult.zip"
for %%F in ( %WINAMS_PATH% ) do set filename=%%~nF
echo:%filename%
set process_output=_TestFixResult
set output_filename=%filename:_ATGResult=!process_output!%
echo:%output_filename%