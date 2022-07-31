'****************************************************************************************************************
' transfer_target_ids.vbs
'*  Explanation     : Transfer target IDS to working folder                                                     *
'*  Author          : Ryan                                                                                      *
'*  Version         : 1.0.1    - Changes testsheet editor filename                                              *
'*                  : 1.0.0    . Initial write                                                                  *
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

' open testsheet editor to transfer data
set oXLBook = oXLApp.Workbooks.Open(WScript.Arguments.Item(0))   ' open testsheet editor file
oXLApp.Application.run "'testsheet_editor.xlsm'!ATG_Folder_get"     ' run macro

' catch error in the process
if Err.Number <> 0 then
    WScript.echo Err.Description
    Err.clear
    returnVal = 1
end if

' close the tool file
oXLBook.close True
oXLApp.quit

' relrease objects
Set oXLApp = Nothing
set oXLBook = Nothing

WScript.quit returnVal
' end of script