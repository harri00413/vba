Attribute VB_Name = "mod_StripString"
Option Compare Database
Option Explicit

Function fStripString(sFromString As String, sDelstring As String) As String
'******************************************************************************************
'Description:   Function fStripString(sFromString As String, sDelstring As String) As String
'               Deletes literal strings (if multiple) and returns stripped string.
'Input:         Characterstring to be deleted from string.
'Output:        Stripped string
'Example:       fStripString("Voetbal veld , lijnen"," ,") becomes "Voetbalveld lijnen"
'Calls/Uses:    Mid(), Len(), InStr() a.s.o.
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          15-12-2016     Initial version
'******************************************************************************************

    Dim i, j, l As Integer
    Dim sStr As String

    On Error GoTo fStripString_error
    
    If Len(sDelstring) > 0 Then
        j = Len(sDelstring)                             ' j = lengte van de te verwijderen string
        l = Len(sFromString)                            ' l = lengte van de (rest) string waarin gezocht wordt
        Do
            i = InStr(sFromString, sDelstring)          ' i = Startpunt van gezochte string
            If i > 0 Then
                sFromString = Left(sFromString, i - 1) & Mid(sFromString, i + j)
            End If
            l = l - j
        Loop While i > 0
    End If
    fStripString = sFromString

fStripString_exit:
    On Error Resume Next
    Exit Function

fStripString_error:
    MsgBox "Error: " & Err & ", " & Err.Description
    Resume fStripString_exit
End Function

