:: Purpose:         Creates the first instance of log file
:: Parameters:
:: Requirements:
:: Author:          prdedumo.acn
:: Version:         1.0.0 + Initial write
@echo off

:: Force path to some system utilites
set WMIC=%SystemRoot%\System32\wbem\wmic.exe
set FIND=%SystemRoot%\System32\find.exe

:: Get the date into ISO 8601 standard format (yyyy-mm-dd)
for /f %%a in ('^<NUL %WMIC% OS GET LocalDateTime ^| %FIND% "."') DO set DTS=%%a
set CUR_DATE=%DTS:~0,4%-%DTS:~4,2%-%DTS:~6,2%
del %~dp0..\log\*.txt
echo # ------------------------------->%~dp0..\log\log.txt
echo # MBD ATG TooL 1.0.0>>%~dp0..\log\log.txt
echo # ------------------------------->>%~dp0..\log\log.txt
echo # >>%~dp0..\log\log.txt
echo # ------------------------------->>%~dp0..\log\log.txt
echo # Process: Unit Test Fix>>%~dp0..\log\log.txt
echo # ------------------------------->>%~dp0..\log\log.txt
echo # >>%~dp0..\log\log.txt