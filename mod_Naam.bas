Attribute VB_Name = "mod_Naam"
Option Compare Database
Option Explicit

Function fNaamcompl(sAchternaam As Variant, sVoornaam As Variant, Optional sTussenvoegsel As Variant, _
                    Optional bVnEerst As Boolean = True) As String
'*****************************************************************************************
'Description:   fNaamcompl() maakt een volledige naam van voornaam + eventueel tussenvoegsel en achternaam
'
'Input:         sAchternaam As Variant, sVoornaam As Variant, Optional sTussenvoegsel As Variant
'               bVnEerst as Boolean.
'Output:        fNaamcompl as string
'Example:       fNaamcompl("Piet", "van", "Dijk") Creeert: "Piet van Dijk"
'History:       fh = Frank Harland
'
'Version:       Owner       Date        Description
'  1.0          fh                      Initial version
'******************************************************************************************

    On Error GoTo fNaamcompl_error
    
    If bVnEerst Then            ' Voornaam eerst
        If IsMissing(sTussenvoegsel) Or IsNull(sTussenvoegsel) Then
            fNaamcompl = LTrim(Trim(sVoornaam) & " " & Trim(sAchternaam))
        Else
            fNaamcompl = LTrim(Trim(sVoornaam) & " " & Trim(sTussenvoegsel) & " " & Trim(sAchternaam))
        End If
    Else                        ' Achternaam, voornaam
        If IsMissing(sTussenvoegsel) Or IsNull(sTussenvoegsel) Then
            fNaamcompl = LTrim(Trim(sAchternaam) & ", " & Trim(sVoornaam))
        Else
            fNaamcompl = LTrim(Trim(sAchternaam) & ", " & Trim(sVoornaam) & " " & Trim(sTussenvoegsel))
        End If
    End If
    
fNaamcompl_exit:
    On Error Resume Next
    Exit Function

fNaamcompl_error:
    On Error Resume Next
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fNaamcompl_exit
End Function

