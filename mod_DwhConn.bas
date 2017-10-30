Attribute VB_Name = "mod_DwhConn"
Option Compare Database

Function fIVW2Conn()
'******************************************************************************************
'Description:   Function fIVW2Conn()
'               Testen verbinding met Netezza IVW2 met ADO!!
'Input:         Connection id from uSYStbl_Conn
'
'Output:
'Example:
'Calls/Uses:
'Assumes:       -
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          29-08-2016     Initial version
'******************************************************************************************

    Dim oconn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim rs_c As dao.Recordset
    Dim i As Integer
    Dim strSQL As String
    Dim iConnid As Long
    'Dim strCon As String
    fIVW2Conn = False
    
    ' Set iConnid manual for testing purposes
    iConnid = 11
                         
    Set rs_c = CurrentDb.OpenRecordset("SELECT * FROM usystbl_Conn WHERE ConnID = " & iConnid)
    With rs_c
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            For i = 1 To .RecordCount
                oconn.Open "Driver=" & .Fields("ODBCStr_Driver") & ";" & _
                       "Servername=" & .Fields("ODBCStr_Server") & ";" & _
                       "port=" & .Fields("ODBCStr_Port") & ";" & _
                       "Database=" & .Fields("ODBCStr_Db") & ";" & _
                       "Username=" & .Fields("ODBCStr_Uid") & ";" & _
                       "Password=" & .Fields("ODBCStr_PW") & ";"
                       
                Set rs = oconn.Execute("Select * FROM _v_sys_columns")
                'Debug.Print rs.Fields(0)
                oconn.Close
                .MoveNext
            Next i
        End If
    End With
    fIVW2Conn = True
    Debug.Print ftconn

fIVW2Conn_exit:
    On Error Resume Next
    rs.Close
    oconn.Close
    rs_c.Close
    Set rs = Nothing
    Set oconn = Nothing
    Set rs_c = Nothing
    Exit Function

fIVW2Conn_error:
    If Err < 0 Then
        MsgBox "Geen verbinding met de Netezza database"
    End If
    MsgBox "fout: " & Err & ", " & Err.Description
    fIVW2Conn = False
    Resume fIVW2Conn_exit
End Function

