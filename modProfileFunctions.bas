Attribute VB_Name = "modProfileFunctions"
Option Compare Database
Option Explicit

Function fGetProfileValue(sTableName As String, sFieldName As String, sAggFunc As String, Optional lReclimit As Variant) As Variant
'******************************************************************************************
'Description:   Function fGetProfileValue(sTableName As String, sFieldName As String, sAggFunc As String) As Variant
'               Get Min/Max or other value from table by creating query dynamically.
'               Works on Netezza ODBC source
'Input:         Tablename & Fieldname, Aggregate function that is used to compute value.
'               Possible values: MIN, MAX, CNULL (count nulls), PNULL (Null Percentage, MINLEN / MAXLEN (Minimal / Maximal Length
'Output:        Max value of field in table
'Example:       fGetProfileValue("Mytable", "Country", "Min", 500000)
'Calls/Uses:    ODBC Connector ID from usystbl_conn
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          29-11-2016     Initial version
'******************************************************************************************

    Dim db As DAO.DATABASE
    Dim qdf As DAO.QueryDef
    Dim iconn As Integer
    Dim ssql, sconn As String

    On Error GoTo fGetProfileValue_error
    
    Set db = CurrentDb()
    ' Put connectionstring in variable
    iconn = fGetintSysitem("ODBCStr_ConnID")
    sconn = fConnString(iconn)          ' Fetch Connection string to correct Environment
    
    If IsMissing(lReclimit) Then
        Select Case UCase(sAggFunc)
            Case "MIN"
                ssql = "SELECT Min(" & sFieldName & ") FROM " & sTableName & ";"
            Case "MAX"
                ssql = "SELECT Max(" & sFieldName & ") FROM " & sTableName & ";"
            Case "CNULL"
                ssql = "SELECT (COUNT(*) - COUNT(" & sFieldName & ")) AS NullCount FROM " & sTableName & ";"
            Case "PNULL"
                ssql = "SELECT (CAST(COUNT(*) AS FLOAT) - CAST(COUNT(" & sFieldName & ") AS FLOAT)) / CAST(COUNT(*) AS FLOAT) AS NullPerc FROM " & sTableName & ";"
            Case "MAXLEN"
                ssql = "SELECT Max(Length(Nvl(CAST (" & sFieldName & " AS VARCHAR(400)),''))) FROM " & sTableName & ";"
            Case "MINLEN"
                ssql = "SELECT Min(Length(Nvl(CAST (" & sFieldName & " AS VARCHAR(400)),''))) FROM " & sTableName & ";"
        End Select
    Else
        Select Case UCase(sAggFunc)
            Case "MIN"
                ssql = "SELECT Min(" & sFieldName & ") FROM " & sTableName & " Limit " & lReclimit & ";"
            Case "MAX"
                ssql = "SELECT Max(" & sFieldName & ") FROM " & sTableName & " Limit " & lReclimit & ";"
            Case "CNULL"
                ssql = "SELECT (COUNT(*) - COUNT(" & sFieldName & ")) AS NullCount FROM " & sTableName & " Limit " & lReclimit & ";"
            Case "PNULL"
                ssql = "SELECT (CAST(COUNT(*) AS FLOAT) - CAST(COUNT(" & sFieldName & ") AS FLOAT)) / CAST(COUNT(*) AS FLOAT) AS NullPerc FROM " & sTableName & " Limit " & lReclimit & ";"
            Case "MAXLEN"
                ssql = "SELECT Max(Length(Nvl(CAST (" & sFieldName & " AS VARCHAR(400)),''))) FROM " & sTableName & " Limit " & lReclimit & ";"
            Case "MINLEN"
                ssql = "SELECT Min(Length(Nvl(CAST (" & sFieldName & " AS VARCHAR(400)),''))) FROM " & sTableName & " Limit " & lReclimit & ";"
        End Select
    End If
    
    ' Delete qdf & create new one
    If DCount("*", "msysObjects", "Name = 'ptqry_ProfileData'") > 0 Then      ' Query exists?
        db.QueryDefs.Delete "ptqry_ProfileData"
    End If
    Set qdf = db.CreateQueryDef("ptqry_ProfileData")              ' Create new querydef
    qdf.Connect = sconn                 ' Set Connection string of the Querydef
    qdf.ReturnsRecords = True
    qdf.SQL = ssql

    ' Openrecordset and retrieve value, put into Function var.
    fGetProfileValue = qdf.OpenRecordset().Fields(0)
    
fGetProfileValue_exit:
    On Error Resume Next
    qdf.Close
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    Exit Function

fGetProfileValue_error:
    If Err.Number = 3146 Then Resume Next       ' Don't stop at odbc error
    MsgBox "fout: " & Err & ", " & Err.DESCRIPTION
    Resume fGetProfileValue_exit
End Function

Function fDomCount(sTableName As String, sFieldName As String) As Boolean
'******************************************************************************************
'Description:   Function fDomCount(sTableName As String, sFieldName As String) As Boolean
'               Shows query with domainvalues for a given Table & Field
'Output:        Boolean True if success.
'Calls/Uses:
'Assumes:       -
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          07-12-2016     Initial version
'******************************************************************************************

    Dim db As DAO.DATABASE
    Dim qdfDomCount As DAO.QueryDef
    Dim i As Integer
    Dim ssql, sconn As String

    On Error GoTo fDomCount_error
    
    fDomCount = False
    Set db = CurrentDb()
    
    ' Recordcount: sql to count records, then
    ssql = "SELECT " & sFieldName & ", COUNT(" & sFieldName & ") FROM " & sTableName & " GROUP BY " & sFieldName & ";"

    If DCount("*", "msysObjects", "Name = 'ptqry_DomCount'") > 0 Then      ' Query exists?
        db.QueryDefs.Delete "ptqry_DomCount"
    End If
    
    ' Fetch connectionstring
    sconn = fConnString(fGetintSysitem("ODBCStr_ConnID"))

    Set qdfDomCount = db.CreateQueryDef("ptqry_DomCount")
    With qdfDomCount
        .Connect = sconn
        .ReturnsRecords = True
        .SQL = ssql
    End With

    DoCmd.OpenQuery qdfDomCount.Name, acViewNormal, acReadOnly
    
    fDomCount = True
    
fDomCount_exit:
    On Error Resume Next
    Set qdfDomCount = Nothing
    db.Close
    Set db = Nothing
    Exit Function

fDomCount_error:
    MsgBox "fout: " & Err & ", " & Err.DESCRIPTION
    Resume fDomCount_exit
End Function
