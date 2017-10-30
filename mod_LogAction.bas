Attribute VB_Name = "mod_LogAction"
Option Compare Database
Option Explicit

Sub fLogAction(sChangeType As String, sTableName As String, iRecID As Long, sFldName As String, _
               Optional sOldVal As Variant, Optional sNewVal As Variant)
'******************************************************************************************
'Description:   Sub fLogAction(sChangeType As String, sTableName As String, iRecID as Long, sFldName As String, _
'                              Optional sOldVal As Variant, Optional sNewVal As Variant)
'               Logging of actions in tables in table tbl_ChangeLog
'               Put in Form.BeforeUpdate to have old and new value
'Input:         Changetype Add, Delete, Update
'Uses/Assumes:  -
'Output:        -
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          14-09-2017     Initial version
'******************************************************************************************

    Dim db As DAO.Database
    Dim i As Integer
    Dim sOSUname As String
    
    Dim ssql As String

    On Error GoTo fLogAction_error

    Set db = CurrentDb()
    sOSUname = fOSUserName
    
    Select Case sChangeType
        Case "Add"
            ssql = "INSERT INTO tbl_ChangeLog (LogDateTime, LogUserName, LogActionType, LogTable, LogRecID, LogFieldName, LogUpdNewVal) VALUES " _
                 & "(#" & Now() & "#, '" & sOSUname & "', 'Add', '" & sTableName & "', " & iRecID & ", '" & sFldName & "', '" & CStr(sNewVal) & "')"
        Case "Delete"
            ssql = "INSERT INTO tbl_ChangeLog (LogDateTime, LogUserName, LogActionType, LogTable, LogRecID, LogFieldName, LogUpdOldVal) VALUES " _
                 & "(#" & Now() & "#, '" & sOSUname & "', 'Delete', '" & sTableName & "', " & iRecID & ",' " & sFldName & "', '" & CStr(sOldVal) & "')"
        Case "Update"
            ssql = "INSERT INTO tbl_ChangeLog (LogDateTime, LogUserName, LogActionType, LogTable, LogRecID, LogFieldName, LogUpdOldVal, LogUpdNewVal) VALUES " _
                 & "(#" & Now() & "#, '" & sOSUname & "', 'Update', '" & sTableName & "', " & iRecID & ", '" & sFldName & "', '" & CStr(sOldVal) & "', '" & CStr(sNewVal) & "')"
    End Select
    
    'Debug.Print ssql
    DoCmd.SetWarnings False
    db.Execute ssql
    DoCmd.SetWarnings True
         
fLogAction_exit:
    On Error Resume Next
    db.Close
    Set db = Nothing
    Exit Sub

fLogAction_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fLogAction_exit

End Sub


