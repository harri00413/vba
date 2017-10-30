Attribute VB_Name = "mod_createSchema_ini"
Option Compare Database
Option Explicit

Sub fCreateSchemaIni(sRptName As String, sPath As String, sFileName As String, Optional sRecType)
'******************************************************************************************
'Description:   Sub fCreateSchemaIni(sRptName As String, sPath As String, sFileName as String, Optional sRecType)
'               Create schema.ini file to read an exported fixed length file back into Access
'               Loop through the recordtypes to create one recordtype table
'Input:         sRptName As String, Reportname for which this is the entry in the schema.ini
'               sPath As String, Path where the .ini file will be located.
'               sFileName as String
'               Optional sRecType, Recordtype for this table. There will be a table per recordtype.
'               Assumption is there is only one recordtype per run of this procedure.
'               Each time the proc is run, the schema.ini will be deleted / overwritten with
'               the details of the current table / recordtype.
'Output:        Schema.ini text file at location sPath
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          11-10-2017     Initial version
'******************************************************************************************

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim bMoveNext As Boolean
    Dim i As Integer, iNwPos As Integer
    Dim ssql As String
    Dim iFile As Integer
    Dim iFillLen As Integer
    On Error GoTo fCreateSchemaIni_error
    DoCmd.Hourglass True
    
    Set db = CurrentDb()
    ssql = "SELECT a.Column_name, a.Length, a.Position, a.[Record Type] FROM tbl_RCRpt_Attributes AS a " _
         & "WHERE a.DEL_REC_IND = 0 And a.Report = '" & sRptName & "' "
        If Not IsMissing(sRecType) Then
           ssql = ssql & "And a.[Record Type] = '" & sRecType & "' "
        End If
        ssql = ssql & "ORDER BY a.[Record Type], a.Position; "
    'Debug.Print ssql
    Set rs = db.OpenRecordset(ssql)
    iFile = FreeFile()
    
    ' Create & open schema.ini file
    Open fCompletePath(sPath, "schema.ini") For Output As #iFile
    ' Print the header information of the file
    Print #iFile, "[" & sFileName & "]"
    Print #iFile, "ColNameHeader=False"
    Print #iFile, "Format=FixedLength"
    ' Now cycle through the fields and add one row in the ini file for eacht field formatted like:
    ' Col1=CustomerNumber Text Width 10
    ' Col2=CustomerName Text Width 30
    With rs
        .MoveLast
        .MoveFirst
        iNwPos = 1          ' Init iNwPos for checking position length coherence
        SysCmd acSysCmdInitMeter, "Creating schema.ini file...", .RecordCount
        For i = 1 To .RecordCount
            ' When length plus position of the field are not coherent with
            ' position of next line, a field must be inserted with a name like "filler#". The next record
            ' position must be equal to the current record position plus length, else create field.
            If iNwPos = .Fields("Position") Then
                Print #iFile, "Col" & i; "=" & .Fields(0) & " Text Width " & CStr(.Fields("Length"))
                bMoveNext = True
            Else
                iFillLen = .Fields("Position") - iNwPos
                Print #iFile, "Col" & i; "=" & "Filler" & i & " Text Width " & iFillLen
                bMoveNext = False
            End If
            If Not .EOF Then
                If Not bMoveNext Then
                    iNwPos = .Fields("Position")
                Else
                    iNwPos = .Fields("Position") + .Fields("Length")
                    .MoveNext
                End If
            End If
            SysCmd acSysCmdUpdateMeter, i
        Next i
    End With
    
    Close iFile
    
fCreateSchemaIni_exit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    db.Close
    Set db = Nothing
    DoCmd.Hourglass False
    SysCmd acSysCmdRemoveMeter
    Exit Sub

fCreateSchemaIni_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fCreateSchemaIni_exit

End Sub

