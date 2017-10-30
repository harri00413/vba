Attribute VB_Name = "mod_db"
Option Compare Database
Option Explicit

Function fStrip_DBO() As Integer

Dim db As DAO.Database
Dim tdf As TableDef
Dim i As Integer
    
On Error GoTo fStrip_DBO_err
    
    Set db = CurrentDb
    i = 0
    For Each tdf In db.TableDefs
        If UCase(left(tdf.Name, 4)) = "DBO_" Then
            i = i + 1
            tdf.Name = Mid(tdf.Name, 5)
        End If
    Next tdf
    fStrip_DBO = i
    
fStrip_DBO_exit:
    On Error Resume Next
    Set tdf = Nothing
    db.Close
    Set db = Nothing
    
fStrip_DBO_err:
    Call ErrorLog("fStrip_DBO", "fStrip_DBO")
    Resume fStrip_DBO_exit

End Function

Function fStrip_Prefix(Optional sPref As Variant) As Integer
' i will contain the number of tables that have been stripped.
Dim db As DAO.Database
Dim tdf As TableDef
Dim i, istrlen As Integer
Dim sStrip As String

On Error GoTo fStrip_Prefix_err

' Indien geen stripstring is opgegeven wordt de string uit de systeemtabellen param: sPrefix gehaald.
If IsMissing(sPref) Then
    sStrip = fGetStrSysitem("sPrefix")
Else
    sStrip = CStr(sPref)
    Call fstorStrSysitem("sPrefix", sStrip)
End If
    istrlen = Len(Trim(sStrip))
    Set db = CurrentDb
    i = 0
    For Each tdf In db.TableDefs        ' for every table in database, check if name is stripstring and cut it off.
        If UCase(left(tdf.Name, istrlen)) = sStrip Then
            i = i + 1
            tdf.Name = Mid(tdf.Name, istrlen + 1)
        End If
    Next tdf
    fStrip_Prefix = i
    
fStrip_Prefix_exit:
    On Error Resume Next
    Set tdf = Nothing
    db.Close
    Set db = Nothing
    
fStrip_Prefix_err:
    Call ErrorLog("fStrip_Prefix", "fStrip_Prefix")
    Resume fStrip_Prefix_exit

End Function
Public Function ErrorLog(objName As String, routineName As String)
Dim db As DAO.Database

Set db = CurrentDb

Open "C:\Error.log" For Append As #1
    
Print #1, Format(Now, "mm/dd/yyyy, hh:nn:ss") & ", " & db.Name & _
    "An error occured in " & objName & ", " & routineName & _
    ", " & CurrentUser() & ", Error#: " & Err.Number & ", " & Err.Description
   
Close #1
db.Close
Set db = Nothing
End Function

Function fSetDBProperty(sPrName As String, vVal As Variant, iType As Integer) As Boolean
'******************************************************************************************
'Description:   Function fSetDBProperty(sPrName As String, vVal As Variant, iType As Integer)
'               Set User Defined property in Database, like version.
'Input:         sPrName as String, Propertyname.
'               vVal as Variant, Value as variant because could be any type
'               iType as Integer, following the regular MS-Access coding:
'               dbBoolean (1); dbLong (4); dbDate (8); dbText (10); dbMemo (12); dbTime (22); dbTimeStamp (23)
'Uses/Assumes:
'Output:        -
'Example:       fSetDBProperty("Version", "7.5.2", dbText)
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          15-05-2017     Initial version
'******************************************************************************************

    Dim db As DAO.Database
    Dim i As Integer
    Dim prpnew As Property
    
    On Error GoTo fSetDBProperty_error
    fSetDBProperty = False
    Set db = CurrentDb()
    
    db.Containers("Databases").Documents("UserDefined").Properties("Version") = vVal
    
    fSetDBProperty = True
    
fSetDBProperty_exit:
    On Error Resume Next
    Set prpnew = Nothing
    db.Close
    Set db = Nothing
    Exit Function

fSetDBProperty_error:
    ' Error 3270 means that the property was not found.
    If DBEngine.Errors(0).Number = 3270 Then
        ' Create property, set its value, and append it to the
        ' Properties collection.
        Set prpnew = db.Containers("Databases").Documents("UserDefined").CreateProperty(sPrName, iType, vVal)
        db.Containers("Databases").Documents("UserDefined").Properties.Append prpnew
        fSetDBProperty = True
    Else
        MsgBox "fout: " & Err & ", " & Err.Description
        fSetDBProperty = False
    End If

    Resume fSetDBProperty_exit

End Function

Function fSetAppTitle(sTitle As String) As Boolean
'******************************************************************************************
'Description:   Function fSetAppTitle(sTitle as string) as Boolean
'               Sets Application title in window
'Input:         Title as string
'Output:        True if succeeded
'Example:       fSetAppTitle("Metatool 7.7.4")
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          31-05-2017     Initial version
'******************************************************************************************

    Dim db As DAO.Database
    Dim i As Integer
    On Error GoTo fSetAppTitle_error
    fSetAppTitle = False
    
    Set db = CurrentDb()
    
    db.Containers("Databases").Documents("MSysDb").Properties("AppTitle") = sTitle
    Application.RefreshTitleBar
    
    fSetAppTitle = True

fSetAppTitle_exit:
    On Error Resume Next
    db.Close
    Set db = Nothing
    Exit Function

fSetAppTitle_error:
    MsgBox "fout: " & Err & ", " & Err.Description
    Resume fSetAppTitle_exit

End Function

Sub ListProps()

Dim db As Database
Dim p As Property
Dim c As Container
Dim d As Document
Set db = CurrentDb()
Set c = db.Containers("Databases")
'For Each c In db.Containers("Databases")
    For Each d In c.Documents
        For Each p In d.Properties
            Debug.Print c.Name, d.Name, p.Name, p.Type, p.Value
        Next p
    Next d
'Next c

'For Each p In CurrentDb.Containers("Databases").Documents("UserDefined")

Set c = Nothing
Set d = Nothing
Set p = Nothing
db.Close
Set db = Nothing
End Sub

