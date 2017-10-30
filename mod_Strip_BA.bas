Attribute VB_Name = "mod_Strip_BA"
Option Compare Database
Option Explicit

Function fStripBeforeAfter(sFromString As String, sDivider As String, bAfter As Boolean, bIncluding As Boolean) As String
'******************************************************************************************
'Description:   Function fStripBeforeAfter(sFromString as String, sDivider as String, bAfter as Boolean, bIncluding as Boolean) as String
'               Strip string after or before first occurence of given divider string
'               Boolean bAfter indicates if string after (or before) divider should be deleted
'               Boolean bIncluding indicates if divider string is also deleted.
'Input:         Original String of text, divider character(s), indicator if text Before or After the divider
'               should be stripped
'Output:        Stripped String of text
'Example:       fStripBeforeAfter("A man|woman", "|", True, True) results in "A man"
'Calls/Uses:
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          15-12-2016     Initial version
'******************************************************************************************

    Dim i, j, l As Integer
    Dim sstr As String
    
    On Error GoTo fStripBeforeAfter_error
    
    j = Len(sDivider)
    fStripBeforeAfter = sFromString
    
    ' Find location of (first) divider string
    i = InStr(sFromString, sDivider)
    If i = 0 Then GoTo fStripBeforeAfter_exit       ' Not found, string out = string in
    
    If bAfter Then                  ' Delete part after divider
        If bIncluding Then
            sstr = Left(sFromString, i - 1)
        Else
            sstr = Left(sFromString, i - 1 + j)
        End If
    Else                            ' Delete part before divider
        If bIncluding Then
            sstr = Right(sFromString, i - 1)
        Else
            sstr = Right(sFromString, i - 1 + j)
        End If
    End If
    
    fStripBeforeAfter = sstr

fStripBeforeAfter_exit:
    On Error Resume Next
    Exit Function

fStripBeforeAfter_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fStripBeforeAfter_exit
End Function

