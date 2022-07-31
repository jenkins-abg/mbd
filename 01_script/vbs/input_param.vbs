'****************************************************************************************************************
' input_param.vbs
'*  Explanation     : Transfer data from tmp.csv to script_generated worksheet                                  *
'*  Author          : Ryan                                                                                      *
'*  Version         : 1.0.0                                                                                     *
'*  Edited by       :                                                                                           *
'****************************************************************************************************************
Option Explicit

Dim oXLApp                              ' excel object
Dim oXLBook                             ' excel workbook object
Dim oXLSheet, oXLSheet_Main             ' excel worksheets object
Dim intEndRow, i, returnVal

On error resume next
' inital value of script return
returnVal = 0

' setup csv var to get end row
Set oXLApp = CreateObject("Excel.Application")                                      ' new instance of excel

oXLApp.visible = False       ' excel can be seen by the user, set to false for visibility
oXLApp.DisplayAlerts = False

set oXLBook = oXLApp.Workbooks.Open(WScript.Arguments.Item(1))                      ' open csv file
Set oXLSheet = oXLBook.Worksheets("tmp")                                            ' csv file worksheet name

' read data in tmp.csv
intEndRow = oXLSheet.UsedRange.Rows.Count       ' last data entry
' store data to var
Dim dataArray()
redim dataArray(intEndRow)
for i = 1 to intEndRow
    dataArray(i) = oXLSheet.range("A" & i)
next

' catch error in the process
if Err.Number <> 0 then
    WScript.Echo Err.Description
    Err.clear
    returnVal = 1
    WScript.quit returnVal
end if

' close csv file
oXLBook.close True

' check if excel instance is running and testsheet editor is open
Dim xl
dim wb_test
wb_test = 0

on error resume next
set xl = GetObject(, "Excel.Application")
if xl Is Nothing then
    ' do nothing
else
    dim obj
    for each obj in xl.Workbooks
        if obj.name = "testsheet_editor.xlsm" then
            wb_test = 1
        end if
    next
end if

' set workbook
if wb_test = 0 then
    ' open testsheet editor to transfer data
    set oXLBook = oXLApp.Workbooks.Open(WScript.Arguments.Item(0))      ' open testsheet editor file
else
    set oXLBook = xl.Workbooks("testsheet_editor.xlsm")
end if

' Set oXLSheet = oXLBook.Worksheets("script_generated")
set oXLSheet_Main = oXLBook.Worksheets("Main")

' write data to testeditor file
With oXLSheet_Main
    .Range("B2").Value = dataArray(1)
    .Range("B3").Value = dataArray(2)
    .Range("F2").Value = dataArray(3)
    .Range("F3").Value = dataArray(4)
    .Range("F4").Value = dataArray(5)
    .Range("F5").Value = dataArray(6)
    .Range("B4").Value = dataArray(7)
    .Range("B5").Value = dataArray(8)
End with

' catch error in the process
if Err.Number <> 0 then
    WScript.echo Err.Description
    Err.clear
    returnVal = 2
end if

' close the tool file
oXLBook.close True
oXLApp.quit

' relrease objects
set oXLSheet_Main = Nothing
Set oXLApp = Nothing
set oXLBook = Nothing
Set oXLSheet = Nothing

WScript.quit returnVal
' end of script