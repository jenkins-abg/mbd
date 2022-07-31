'****************************************************************************************************************
' test_fix_static.vbs
'*  Explanation     : Fix the static variables in the testsheet                                                 *
'*  Author          : Ryan                                                                                      *
'*  Version         : 1.0.0    . Initial write                                                                  *
'*  Edited by       :                                                                                           *
'****************************************************************************************************************
Option Explicit

Dim oXLApp                         ' excel object
Dim oXLBook                        ' excel workbook object
Dim returnVal

On error Resume next
' inital value of script return
returnVal = 0

Set oXLApp = CreateObject("Excel.Application")                      ' new instance of excel

oXLApp.visible = False       ' excel can be seen by the user, set to false for visibility

' open testsheet_static_tool.xlsm to transfer data
set oXLBook = oXLApp.Workbooks.Open(WScript.Arguments.Item(0))   ' open ' open testsheet_static_tool.xlsm to transfer data.xlsm
Dim errorsen
errorsen = oXLApp.Run("'テストシート修正マクロ.xlsm'!Macro", _
        WScript.Arguments.Item(1), WScript.Arguments.Item(2), _
        WScript.Arguments.Item(3), WScript.Arguments.Item(4))    ' run macro

' catch error in the process
if Err.Number <> 0 then
    WScript.echo Err.Description
    Err.clear
    returnVal = 1
end if

' close the tool file
oXLBook.close False
oXLApp.quit

' relrease objects
Set oXLApp = Nothing
set oXLBook = Nothing

WScript.quit returnVal
' end of script