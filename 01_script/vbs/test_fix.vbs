'****************************************************************************************************************
' test_fix.vbs
'*  Explanation     : Modify Target IDs testshets in the working folder                                         *
'*  Author          : Ryan                                                                                      *
'*  Version         : 1.0.1    - Changes testsheet editor filename                                              *
'*                  : 1.0.0    . Initial write                                                                  *
'*  Edited by       :                                                                                           *
'****************************************************************************************************************
Option Explicit

Dim oXLApp                         ' excel object
Dim oXLBook_TE, oxlBook_Kanri, oxlBook_Keisoku                           ' excel workbook object
Dim returnVal
Dim objFS

On error Resume next
' inital value of script return
returnVal = 0

Set oXLApp = CreateObject("Excel.Application")                      ' new instance of excel
Set objFS = CreateObject("Scripting.FileSystemObject")

oXLApp.visible = False       ' excel can be seen by the user, set to false for visibility

' open kanri file
set oxlBook_Kanri = oXLApp.Workbooks.Open(WScript.Arguments.Item(0))
' open keisoku file
set oxlBook_Keisoku = oXLApp.Workbooks.Open(WScript.Arguments.Item(2))

' open testsheet editor to transfer data
set oXLBook_TE = oXLApp.Workbooks.Open(WScript.Arguments.Item(1))   ' open testsheet editor file
oXLApp.Application.run "'testsheet_editor.xlsm'!Main_Fix"       ' run macro

' catch error in the process
if Err.Number <> 0 then
    WScript.echo Err.Description
    Err.clear
    returnVal = 1
end if

' catch if error log is created in Main_Fix Macro
if objFS.FileExists(WScript.Arguments.Item(3) & "\testsheeteditor_errlog.txt") then
    WScript.echo "Error... Check testsheeteditor_errlog.txt"
    Err.clear
    returnVal = 2
end if

' close the tool and kanri files
oxlBook_Kanri.close false
oxlBook_Keisoku.close false
oXLBook_TE.close false
oXLApp.quit

' relrease objects
Set oXLApp = Nothing
set oxlBook_Kanri = Nothing
set oxlBook_Keisoku = Nothing
set oXLBook_TE = Nothing

WScript.quit returnVal
' end of script