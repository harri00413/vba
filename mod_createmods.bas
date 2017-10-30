Attribute VB_Name = "mod_createmods"
Option Compare Database
Option Explicit

Sub fCreateProc(BFunction As Boolean, bRS As Boolean, bDB As Boolean, bSSQL As Boolean, bQDF As Boolean, sProcnaam As String, iScope As Integer)
' Sub to create standard functions or subs in VBA with optional db object, Recordset, QueryDef and ssql.
' Is called from form frm_CreateSub

    Dim sProcSrt As String
    Dim sProc As String
    Dim sScope As String
    
    If BFunction Then
        sProcSrt = "Function"
    Else
        sProcSrt = "Sub"
    End If
    
    Select Case iScope
        Case 1
            sScope = "Private "
        Case 2
            sScope = "Public "
        Case Else
            sScope = ""
    End Select
    
    sProc = sScope & sProcSrt & " " & sProcnaam & "()" & vbCrLf _
    & "'" & String(90, "*") & vbCrLf _
    & "'Description:" & vbTab & sProcSrt & " " & sProcnaam & "()" & vbCrLf & "'" & vbCrLf & "'Input:" & vbCrLf & "'" & vbCrLf _
    & "'Uses/Assumes:" & vbCrLf & "'Output:" & vbCrLf & "'Example:" & vbCrLf _
    & "'History:       fh = Frank Harland" & vbCrLf _
    & "'Version:       Owner       Date           Description" & vbCrLf _
    & "'  1.0          fh          " & Format(Date, "dd-mm-yyyy") & "     Initial version " & vbCrLf _
    & "'" & String(90, "*") & vbCrLf & vbCrLf
    If bDB Then sProc = sProc & vbTab & "Dim db As DAO.Database" & vbCrLf
    If bRS Then sProc = sProc & vbTab & "Dim rs As DAO.Recordset" & vbCrLf
    If bQDF Then sProc = sProc & vbTab & "Dim qdf As DAO.QueryDef" & vbCrLf
    sProc = sProc & vbTab & "Dim i As Integer" & vbCrLf
    If bSSQL Then sProc = sProc & vbTab & "Dim ssql As String" & vbCrLf & vbCrLf
    sProc = sProc & vbTab & "On Error goto " & sProcnaam & "_error" & vbCrLf & vbCrLf
    If bDB Then sProc = sProc & vbTab & "Set db = Currentdb() " & vbCrLf
    If bRS Then sProc = sProc & vbTab & "Set rs = db.OpenRecordset(ssql)" & vbCrLf
    If bQDF Then sProc = sProc & vbTab & "Set qdf = db.QueryDefs()" & vbCrLf
    sProc = sProc & vbCrLf
    sProc = sProc & sProcnaam & "_exit: " & vbCrLf
    sProc = sProc & vbTab & "On Error Resume Next" & vbCrLf
    If bRS Then sProc = sProc & vbTab & "rs.Close" & vbCrLf
    If bRS Then sProc = sProc & vbTab & "Set rs = Nothing" & vbCrLf
    If bQDF Then sProc = sProc & vbTab & "Set qdf = Nothing" & vbCrLf
    If bDB Then sProc = sProc & vbTab & "db.Close" & vbCrLf
    If bDB Then sProc = sProc & vbTab & "Set db = Nothing" & vbCrLf
    sProc = sProc & vbTab & "Exit " & sProcSrt & vbCrLf & vbCrLf
    sProc = sProc & sProcnaam & "_error:" & vbCrLf
    sProc = sProc & vbTab & "MsgBox ""fout: "" & Err & "", "" & Err.Description " & vbCrLf
    sProc = sProc & vbTab & "Resume " & sProcnaam & "_exit" & vbCrLf & vbCrLf
    sProc = sProc & "End " & sProcSrt & vbCrLf
    
    Text2Clipboard sProc        ' Op het Clipboard plettere, ter plakking later.
    'Debug.Print sProc
    
End Sub

Public Function fCreateClassFromTable(sTableName As Variant) As String
'******************************************************************************************
'Description:   Function fCreateClassFromTable(sTablename as Variant) as String
'
'Input:         sTablename as Variant (string)
'
'Output:        Classmodule containing property let's and get's for the fields and a
'               private variable that holds the value for the field temporary.
'               The types are derived from the datatypes from the tablefields.
'Example:       fCreateClassFromTable("Users")
'Calls/Uses:    Text2Clipboard
'
'Assumes:       -
'
'History:       fh = Frank Harland
'Version:       Owner       Date        Description
'  1.0          fh          22-01-2003  Initial version
'******************************************************************************************
    Dim sClassname As String, sProc1 As String
    Dim sppg As String
    Dim sppl As String
    Dim sproc2 As String
    Dim stypechr As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim i As Integer
    Dim sType As String
    
    On Error GoTo fCreateClassFromTable_error
    
    sppg = "Public Property Get "
    sppl = "Public Property Let "
    
    If IsNull(sTableName) Then
        MsgBox "Geen tabelnaam gegeven"
        GoTo fCreateClassFromTable_exit
    End If
    
    sClassname = "cls" & sTableName

    sProc1 = "'" & String(90, "*") & vbCrLf _
    & "'Description: " & sClassname & vbCrLf & "'" & vbCrLf _
    & "'Output:" & vbCrLf & "'Example:" & vbCrLf _
    & "'Calls/Uses:" & vbCrLf & "'" & vbCrLf & "'Assumes:       -" & vbCrLf _
    & "'History:       fh = Frank Harland" & vbCrLf _
    & "'Version:       Owner       Date        Description" & vbCrLf _
    & "'  1.0          fh                      Initial version " & vbCrLf _
    & "'" & String(90, "*") & vbCrLf & vbCrLf _
    & "' Variabelen voor de Properties " & vbCrLf
    
    sproc2 = ""                                         ' Variabele waar de propertyprocs inkomen.
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(sTableName)
    For Each fld In tdf.Fields                          ' Get the fieldnames and types
        sType = DLookup("[Type]", "usystbl_dbtypes", "Typenumber = " & fld.Type)
        stypechr = DLookup("[stypechar]", "usystbl_dbtypes", "Typenumber = " & fld.Type)
        'Debug.Print fld.Type, stype, stypechr, fld.Name
        
        sProc1 = sProc1 & "Private " & stypechr & fld.Name & " as " & sType & vbCrLf
        sproc2 = sproc2 & sppg & fld.Name & "() As " & sType & vbCrLf & vbTab _
               & fld.Name & " = " & stypechr & fld.Name & vbCrLf _
               & "End Property" & vbCrLf & vbCrLf _
               & sppl & fld.Name & "(" & stypechr & "NewValue As " & sType & ")" & vbCrLf & vbTab _
               & stypechr & fld.Name & " = " & stypechr & "NewValue" & vbCrLf _
               & "End Property" & vbCrLf & vbCrLf
               
    Next fld
    
    sProc1 = sProc1 & vbCrLf & sproc2          ' Samenvoegen delen 1 en 2
    
    Text2Clipboard sProc1                       ' Op het Clipboard plettere, ter plakking later.
    'Debug.Print sProc1
    
fCreateClassFromTable_exit:
    On Error Resume Next
    db.Close
    Set tdf = Nothing
    Set fld = Nothing
    Set db = Nothing
    Exit Function

fCreateClassFromTable_error:
    On Error Resume Next
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fCreateClassFromTable_exit
    
End Function

