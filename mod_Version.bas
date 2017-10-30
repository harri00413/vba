Attribute VB_Name = "mod_Version"
Option Compare Database
Option Explicit

Function fGetDBVersion() As String
'******************************************************************************************
'Description:   Function fGetDBVersion() As String
'Assumption     db version structure is Y.M.#
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          16-05-2017     Initial version
'******************************************************************************************

    Dim db As DAO.DATABASE
    Dim i As Integer
    On Error GoTo fGetDBVersion_error

    Set db = CurrentDb()
    
    fGetDBVersion = db.Containers("Databases").Documents("UserDefined").Properties("Version")

fGetDBVersion_exit:
    On Error Resume Next
    db.Close
    Set db = Nothing
    Exit Function

fGetDBVersion_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fGetDBVersion_exit

End Function

