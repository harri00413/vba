Attribute VB_Name = "mod_RefreshCadertables"
Option Compare Database
Option Explicit

Function fRefreshCDR() As Integer
'******************************************************************************************
'Description:   Sub fRefreshCDR() as Integer
'               Reload oracle Cader tables from ODBC Source in Backend database.
'               Delete and reload version...
'Input:         Zo min mogelijk
'Uses/Assumes:  Access tables in back-end, databasename extracted from object in front-end
'Output:        Number of tables refreshed
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          18-03-2017     Initial version
'******************************************************************************************

    Dim db As DAO.DATABASE
    Dim dbbe As DAO.DATABASE
    Dim rs As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim i As Integer
    Dim ssql, sconn As String

    On Error GoTo fRefreshCDR_error
    DoCmd.Hourglass True
    fRefreshCDR = 0
    
    ' Get back-end path and put in db var
    Set dbbe = DBEngine.OpenDatabase(fGetStrSysitem("Cader_BE_Database"))
        
    ' Prep qdf
    ' Remove Base Querydef and recreate new one.
    ' Query in ssql die in cader back-end checkt of query bestaat, zo ja verwijderen
    ssql = "SELECT Count(*) FROM msysObjects WHERE Name = 'qry_RefreshCDR'"
    Set rs = dbbe.OpenRecordset(ssql)
    With rs
        If .Fields(0) > 0 Then
            dbbe.QueryDefs.Delete "qry_RefreshCDR"
        End If
    End With
    rs.Close
    
    ' Base query for inserting data into back-end tables
    Set qdf = dbbe.CreateQueryDef("qry_RefreshCDR")
    qdf.ReturnsRecords = True
    
    Set db = CurrentDb()
    ssql = "SELECT usystbl_Refreshtbl.ConnectString, usystbl_Refreshtbl.NameBE, usystbl_Refreshtbl.NameODBC " _
         & "FROM usystbl_Refreshtbl; "
    
    ' Loop through table with tablenames that need updating and update one by one
    Set rs = db.OpenRecordset(ssql)
    If fGetStrSysitem("CdrUpdCompleet") = "Ja" Then
        With rs
            If .RecordCount > 0 Then
                sconn = "ODBC;" & .Fields("ConnectString")
                qdf.Connect = sconn
                .MoveLast
                .MoveFirst
                SysCmd acSysCmdInitMeter, "Updating CADER Tables...", .RecordCount
                For i = 1 To .RecordCount
                    Call fZaptbl(.Fields("NameBE"))
                    SysCmd acSysCmdUpdateMeter, i
                    sconn = "ODBC;" & .Fields("ConnectString")
                    qdf.Connect = sconn
                    ssql = "SELECT * FROM " & .Fields("NameODBC") & ";"
                    qdf.SQL = ssql
                    'Debug.Print ssql
                    dbbe.Execute ("INSERT INTO " & .Fields("NameBE") & " SELECT * FROM qry_RefreshCDR;")
                    .MoveNext
                Next i
            End If
        End With
    Else                    ' Update, not replace
        With rs
            If .RecordCount > 0 Then
                .MoveLast
                .MoveFirst
                Call LogCdrUpd("Start Cader update", "", 0)
                Call LogCdrUpd("CR", "", 0)
                SysCmd acSysCmdInitMeter, "Updating CADER Tables...", .RecordCount
                For i = 1 To .RecordCount
                    Call LogCdrUpd("Update table ", .Fields("NameBE"), 0)
                    Call fRefreshCDRTable(.Fields("NameBE"))        ' Refresh per table, add missing, delete or update records
                    SysCmd acSysCmdUpdateMeter, i
                    .MoveNext
                Next i
            End If
        End With
    End If
    fRefreshCDR = i - 1
    
    ' Update 'Ververst datum & tijd' in de systeem parameter en verversen indien form "Requirements en Cader Begrippen open staat.
    Call fstorDateSysitem("CDRVerverstDT", Now())
    If CurrentProject.AllForms("frm_CADER_vs_Requirements").IsLoaded Then
        Forms!frm_CADER_vs_Requirements!txtCDRVerverstDT = fGetDateSysitem("CDRVerverstDT")
    End If
    
fRefreshCDR_exit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    dbbe.Close
    Set dbbe = Nothing
    db.Close
    Set db = Nothing
    Call LogCdrUpd("End Cader update ", "", 0)
    SysCmd acSysCmdRemoveMeter
    DoCmd.Hourglass False
    Exit Function

fRefreshCDR_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fRefreshCDR_exit

End Function

Sub fRefreshCDRTable(sTableName As String)
'******************************************************************************************
'Description:   Sub fRefreshCDRTable(sTableName As String)
'               Refresh Table in Access with data from odbc source
'Input:         Access tablename that should have an identical counterparty in the odbc source
'Output:        Log is written to a log table
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          07-06-2017     Initial version
'******************************************************************************************

    Dim db As DAO.DATABASE
    Dim rs As DAO.Recordset
    Dim rsC As DAO.Recordset
    Dim qdf As DAO.QueryDef
    Dim i As Integer
    Dim ssql As String, sCrit As String
    Dim ssqlC As String
    Const sTblPrefix As String = "CADER."

    On Error GoTo fRefreshCDRTable_error
    
    Set db = CurrentDb()

    ' 1. Are there new records?
    ssql = "SELECT * FROM " & sTblPrefix & sTableName & ";"
    ' Criteria for check if helpqry exists and delete / recreate helpquery
    sCrit = "Name = 'qry_CdrTbl'"
    If DCount("[Name]", "[msysObjects]", sCrit) > 0 Then
        db.QueryDefs.Delete "qry_CdrTbl"
    End If
    
    Set qdf = db.CreateQueryDef("qry_CdrTbl")
    With qdf
        .ReturnsRecords = True
        .Connect = fConnString(1)               ' Connectionstring 1 = CADER_PRD
        .SQL = ssql
    End With
    
    ' Directly insert the records into the Access table (first report on changes with loop)
    ssql = "SELECT qry_CdrTbl.* FROM qry_CdrTbl " _
         & "LEFT JOIN " & sTableName & " ON qry_CdrTbl.ID = " & sTableName & ".ID " _
         & "WHERE " & sTableName & ".ID Is Null;"
    Set rs = db.OpenRecordset(ssql)
    With rs
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF
                Call LogCdrUpd("Added", sTableName, .Fields(0))
                'Debug.Print "Insert ", .Fields("[ID]")
                .MoveNext
            Loop
        End If
    End With
    
    ' Now insert
    ssql = "INSERT INTO " & sTableName & " " _
             & "SELECT qry_CdrTbl.* FROM qry_CdrTbl " _
             & "LEFT JOIN " & sTableName & " ON qry_CdrTbl.ID = " & sTableName & ".ID " _
             & "WHERE " & sTableName & ".ID Is Null;"
    db.Execute ssql
    
    ' 2. Are there deletions?
    ' Same as 1 but reverse order of tables, first report changes
    ssql = "SELECT " & sTableName & ".* FROM " & sTableName & " " _
         & "LEFT JOIN qry_CdrTbl ON " & sTableName & ".ID = qry_CdrTbl.ID " _
         & "WHERE qry_CdrTbl.ID Is Null;"
    Set rs = db.OpenRecordset(ssql)
    With rs
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Do Until .EOF
                Call LogCdrUpd("Deleted", sTableName, .Fields(0))
                'Debug.Print "Delete ", .Fields("ID")
                .MoveNext
            Loop
        End If
    End With
    
    ' Now delete records
    ssql = "DELETE " & sTableName & ".* FROM " & sTableName & " WHERE " & sTableName & ".[ID] " _
         & "IN (SELECT " & sTableName & ".[ID] FROM " & sTableName & " " _
         & "LEFT JOIN qry_CdrTbl ON " & sTableName & ".[ID] = qry_CdrTbl.[ID] " _
         & "WHERE qry_CdrTbl.ID Is Null);"
    db.Execute ssql
    
    db.QueryDefs.Delete "qry_CdrTbl"                                          ' Drop the qdf
    
    ' 3. Are there changes in the fields?
    ssql = "SELECT * FROM " & sTableName & " ORDER BY ID "                    ' Local table (in back-end)
    ssqlC = "SELECT " & sTblPrefix & sTableName & ".* FROM " _
          & sTblPrefix & sTableName & " ORDER BY ID;"                         ' ODBC table
    
    ' Criteria for check if helpqry exists and delete / recreate helpquery
    sCrit = "Name = 'ptq_UpdCdr'"
    If DCount("[Name]", "[msysObjects]", sCrit) > 0 Then
        db.QueryDefs.Delete "ptq_UpdCdr"
    End If
    
    Set qdf = db.CreateQueryDef("ptq_UpdCdr")
    With qdf
        .ReturnsRecords = True
        .Connect = fConnString(1)                   ' Connectionstring 1 = CADER_PRD
        .SQL = ssqlC
    End With
    Set rsC = qdf.OpenRecordset
    Set rs = db.OpenRecordset(ssql)
    
    ' Compare each field and immediate update
    With rs
        While Not .EOF
            For i = 0 To .Fields.Count - 1
                If .Fields(i) <> rsC.Fields(i) Then
                    Call LogCdrUpd("Updated", sTableName, .Fields(0), .Fields(i).Name, .Fields(i), rsC.Fields(i))
                    ' Update from rsC to rs recordset.
                    .Edit
                    .Fields(i) = rsC.Fields(i)
                    .Update
                    'Debug.Print .Fields(i).Name, .Fields(i), rsC.Fields(i)
                End If
            Next i
            'Debug.Print
            .MoveNext
            rsC.MoveNext
        Wend
    End With

fRefreshCDRTable_exit:
    On Error Resume Next
    ' Cleanup
    db.QueryDefs.Delete "ptq_UpdCdr"
    rs.Close
    rsC.Close
    Set rs = Nothing
    Set rsC = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    Exit Sub

fRefreshCDRTable_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fRefreshCDRTable_exit

End Sub

Sub LogCdrUpd(sCrud As String, sTableName As String, iId As Long, Optional vFldName As Variant, Optional vValOld As Variant, Optional vValNw As Variant)
    Dim db As DAO.DATABASE
    
    Set db = CurrentDb()
    
    Open fGetStrSysitem("CDRLogLocation") For Append As #1
    
    If IsMissing(vValOld) Then
        If sCrud = "CR" Then
            Print #1, vbCrLf
        Else
            Print #1, Format(Now, "mm/dd/yyyy, hh:nn:ss") & ", " & _
                      CurrentUser() & ", " & sCrud & ", " & sTableName & ", " & CStr(iId)
        End If
    Else
        Print #1, Format(Now, "mm/dd/yyyy, hh:nn:ss") & ", " & _
                  CurrentUser() & ", " & sCrud & ", Tablename: " & sTableName & ", " & CStr(iId) _
                  & ", Fieldname: " & vFldName & ", " & vValOld & ", " & vValNw
    End If
    Close #1
    db.Close
    Set db = Nothing
End Sub
