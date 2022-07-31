:: Process #4 log function with embedded date and time group  (DTG)

:: When the whole argument (%1) string is wrapped in double quotes, it is sent as an argument
:: The tilde (%~1) syntax removes the double quotes around the argument.
@echo off
echo:%CUR_DATE% %TIME% %~1>>%~dp0..\log\log.txt
echo:%CUR_DATE% %TIME% %~1