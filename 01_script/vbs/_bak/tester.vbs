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

oXLApp.visible = True       ' excel can be seen by the user, set to false for visibility

' open testsheet_static_tool.xlsm to transfer data
set oXLBook = oXLApp.Workbooks.Open("C:\jenkins_edc\4_unittest_fix\ver5\02_soft\02_testsheet_editor\testsheet_static_tool.xlsm")   ' open testsheet_static_tool.xlsm
Dim errorsen
errorsen = oXLApp.Run("'testsheet_static_tool.xlsm'!Macro", _
        "C:\jenkins_edc\4_unittest_fix\ver5\03_test\ACC_CTL-STD_21PF_TSS3-09-256D-01\関数一覧表_TSS3_256D_1A_ACCMBD_ACC_CTL.xlsm", _
        "C:\jenkins_edc\4_unittest_fix\ver5\03_test\ACC_CTL-STD_21PF_TSS3-09-256D-01\計測適合一覧表_TSS3_256D_1A_ACCMBD_ACC_CTL.xlsm", _
        "C:\jenkins_edc\4_unittest_fix\ver5\02_soft\01_script\c-source", _
        "C:\jenkins_edc\4_unittest_fix\ver5\03_test\output")    ' run macro

' catch error in the process
if Err.Number <> 0 then
    WScript.echo Err.Description
    Err.clear
    returnVal = 1
end if

' close the tool file
'oXLBook.close True
'oXLApp.quit

' relrease objects
'Set oXLApp = Nothing
'set oXLBook = Nothing

WScript.quit returnVal
' end of script