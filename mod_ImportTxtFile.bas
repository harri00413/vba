Attribute VB_Name = "mod_ImportTxtFile"
Option Compare Database
Option Explicit

Sub fCreateImpTables(sRptName As String)
'******************************************************************************************
'Description:   Sub fCreateImpTables(SRptName as String)
'
'               Loop through the recordtypes to create a table per recordtype
'               In huidige opzet niet nodig omdat de routine met textimport de tabel al aanmaakt
'Input:         SRptName as String
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          21-10-2017     Initial version
'******************************************************************************************

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim i As Integer
    Dim ssql As String
    Dim sTblName As String, sRecType As String
    
    On Error GoTo fCreateImpTables_error

    Set db = CurrentDb()
    ssql = "SELECT a.Column_name, a.Length, a.Position, a.[Record Type] FROM tbl_RCRpt_Attributes AS a " _
         & "WHERE a.Report = '" & sRptName & "' AND a.DEL_REC_IND = 0 ORDER BY a.[Record Type], a.Position; "
    'Debug.Print ssql
    i = 0
    Set rs = db.OpenRecordset(ssql)
    With rs
        .MoveLast
        .MoveFirst
        SysCmd acSysCmdInitMeter, "Creating tables for all recordtypes...", .RecordCount
        Do
            ' Create a table per recordtype
            sRecType = .Fields("[Record Type]")
            ' Tablename is constructed from reportname, recordtype and date
            sTblName = sRptName & sRecType & Format(Date, "yyyymmdd")
            ' All recordtypes have as first field the R_SNA date, so can use this for first field name.
            Call fCreateTable(sTblName, .Fields(0), "VARCHAR", .Fields(1), False)
            .MoveNext       ' Only once per table
            If .EOF Then Exit Do
            
            ' As long as there are records in the set add a field to the table but first check if field length is correct.
            Do Until .Fields("[Record Type]") <> sRecType
                Call fCreateTableField(sTblName, .Fields(0), dbText, .Fields(1))
                .MoveNext
            Loop
        Loop Until .EOF
        i = i + 1
        SysCmd acSysCmdUpdateMeter, i
    End With
    
fCreateImpTables_exit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    db.Close
    Set db = Nothing
    SysCmd acSysCmdRemoveMeter
    Exit Sub

fCreateImpTables_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fCreateImpTables_exit

End Sub

Sub fImportTxtFile(sTblName As String, sPath As String, sTxtFileName As String)
'******************************************************************************************
'Description:   Sub fImportTxtFile(sTblName As String, sPath as string, sTxtFileName As String)
'               Import text file, read an fixed length exported file back into Access
'Assumes        Separate files for separate Recordtypes
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          11-10-2017     Initial version
'******************************************************************************************
    Dim db As DAO.Database
    Dim ssql As String
    
    On Error GoTo fImportTxtFile_error
    Set db = CurrentDb()
    
    ssql = "SELECT * INTO " & sTblName & " FROM [Text;DATABASE=" & sPath & ";].[" & Replace(sTxtFileName, ".", "#") & "]"

    db.Execute ssql, dbFailOnError
    db.TableDefs.Refresh

fImportTxtFile_exit:
    On Error Resume Next
    db.Close
    Set db = Nothing
    Exit Sub

fImportTxtFile_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fImportTxtFile_exit

End Sub

