Attribute VB_Name = "mod_excelrecords"
Option Compare Database
Option Explicit

Function fLinkXLRecordsToTable(sTableName As String, sEFileName As String, Optional bHasFldNames As Variant) As Long
'******************************************************************************************
'Description:   Function fLinkXLRecordsToTable(sTablename As String, sEFileName As String, Optional bHasFldNames As Variant) As Long
'               links single Excel file into Table with passed name.
'Input:         sTablename, sEfName as String
'Output:        0 no records imported, or number of loaded records.
'Example:
'Calls/Uses:
'Assumes:       -
'History:       fh = Frank Harland
'Used           Is being used in KM_Registratie
'Version:       Owner       Date           Description
'  1.0          fh          18-10-2011     Initial version
'******************************************************************************************
    Dim i As Integer
    
    On Error GoTo fLinkXLRecordsToTable_error
    i = 0
    
    Call DelTable(sTableName)
    
    If IsMissing(bHasFldNames) Or IsNull(bHasFldNames) Then bHasFldNames = True
    
    DoCmd.TransferSpreadsheet acLink, acSpreadsheetTypeExcel9, sTableName, sEFileName, bHasFldNames
    
    i = DCount("*", sTableName)
        
    fLinkXLRecordsToTable = i
    
fLinkXLRecordsToTable_exit:
    Exit Function

fLinkXLRecordsToTable_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fLinkXLRecordsToTable_exit
End Function

Function fImportXLRecordsToTable(sTableName As String, sEFileName As String, bUseImExSpec As Boolean, iKeepRecords As Integer, iTransferType As Integer, _
                                 sImExSpec As String, Optional vHasFldNames As Variant, Optional vRange As Variant) As Long
'******************************************************************************************
'Description:   fImportXLRecordsToTable(sTableName As String, sEFileName As String, bUseImExSpec As Boolean, iKeepRecords As Integer, iTransferType As Integer, _
'                                       Optional vHasFldNames As Variant, Optional vRange As Variant) As Long
'               Imports a range of rows from an Excel sheet into a given table.
'Input:         sTableName As String, sEFileName As String, Optional bHasFldNames As Variant, _
'               Optional Range As Variant
'Output:        Number of read rows(?)
'Assumes/uses   Function assumes the parameter files and functions from Codelib.mdb are available.
'Example:       fImportXLRecordsToTable("tbl_Excelrecords", "km_stand.xlsx", -1, "RangeKilometers")
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          28-02-2017     Initial version
'******************************************************************************************

    Dim db As DAO.Database
    Dim lRecCnt As Long
    Dim iExcelType As Integer
    Dim sRange As String
    Dim bHasFldNames As Boolean
    
    On Error GoTo fImportXLRecordsToTable_error
    lRecCnt = 0
    Set db = CurrentDb()
    DoCmd.SetWarnings False
    
    ' Skip parameters when True and run a saved import / export spec. with RunSavedImportExport.
    ' Spec to be selected in cboImExSpec.
    ' When false, check on other relevant params and transferspreadsheet.
    If bUseImExSpec Then
        If Len(sImExSpec) = 0 Then
            MsgBox "Specification empty. Please fill and try again.", vbOKOnly, "Missing Specname"
            GoTo fImportXLRecordsToTable_exit
        Else
            DoCmd.RunSavedImportExport sImExSpec
        End If
    Else
        ' Determine spreadsheet type for import from form or paramfile.
        If Nz(fGetintSysitem("ImpSpreadsheetType"), 0) = 0 Then
            iExcelType = 8
        Else
            iExcelType = fGetintSysitem("ImpSpreadsheetType")
        End If
        
        ' Are there fieldnames in the Excel?
        If IsMissing(vHasFldNames) Then
            bHasFldNames = -1
        Else
            bHasFldNames = CBool(vHasFldNames)
        End If
        
        ' A range, when available, is used.
        If IsMissing(vRange) Then
            sRange = ""
        Else
            sRange = CStr(vRange)
        End If
        
        ' Check if table exists, clear if necessary, count records if not to be cleared.
        If fTableExists(sTableName) Then
            Select Case iKeepRecords
                Case 1
                    db.TableDefs.Delete sTableName
                Case 2
                    fZaptbl sTableName
                Case Else
                    ' Keep records, do nothing
            End Select
               lRecCnt = DCount("*", sTableName)
        End If
        
        ' The im/exporting
        DoCmd.TransferSpreadsheet iTransferType, iExcelType, sTableName, sEFileName, bHasFldNames, sRange
    
    End If
    
    ' Give something back to the user...
    lRecCnt = DCount("*", sTableName) - lRecCnt
    fImportXLRecordsToTable = lRecCnt
    
fImportXLRecordsToTable_exit:
    On Error Resume Next
    db.Close
    Set db = Nothing
    DoCmd.SetWarnings True
    Exit Function

fImportXLRecordsToTable_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fImportXLRecordsToTable_exit

End Function


