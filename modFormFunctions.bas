Attribute VB_Name = "modFormFunctions"
Option Compare Database
Option Explicit

Sub SetFieldActivation(sMfrm As String, sSfrm As String, sfld As String, sDataSrc As String, iActive As Integer)
    Dim i As Integer
    Dim ofld As TextBox
    
    On Error GoTo SetFieldActivation_Err
    
    ' Set ofld = Forms!frm_DWH_Objects.Form!subfrm_TabCols_IVW2.Form(sfld)
    Set ofld = Forms(sMfrm).Form(sSfrm).Form(sfld)
    
    If iActive = 0 Then
        i = -1
        ofld.ControlSource = ""
    ElseIf iActive = -1 Then
        i = 0
        ofld.ControlSource = sDataSrc
    End If
    
    ' Set subform fields depending on checkboxes on mainform.
    ofld.ColumnHidden = i

SetFieldActivation_Exit:
    On Error Resume Next
    Set ofld = Nothing
    Exit Sub

SetFieldActivation_Err:
    If Err.Number = (2467 Or 91) Then
        Resume Next
    Else
        MsgBox "fout: " & Err & ", " & Err.DESCRIPTION
    End If
    Resume SetFieldActivation_Exit
    
End Sub

Function fSwitchCtrlLock(ofrm As Form, bState As Boolean) As Boolean
'******************************************************************************************
'Description:   Function fSwitchCtrlLock(ofrm as Form, bState as Boolean) as Boolean
'               Switches controls on a form between Locked and Unlocked, based on tag in ctrl
'History:       fh = Frank Harland
'Version:       Owner       Date           Description
'  1.0          fh          07-03-2017     Initial version
'******************************************************************************************

    Dim octrl As Control
    On Error GoTo fSwitchCtrlLock_error
    
    fSwitchCtrlLock = Not bState
    For Each octrl In ofrm.Controls
        If octrl.ControlType = acTextBox Or octrl.ControlType = acComboBox Then
            If InStr("Switch", octrl.Tag) > 0 Then
                octrl.Locked = bState
            End If
        End If
    Next
    fSwitchCtrlLock = bState
    
fSwitchCtrlLock_exit:
    On Error Resume Next
    Exit Function

fSwitchCtrlLock_error:
    MsgBox "fout: " & Err & ", " & Err.DESCRIPTION
    Resume fSwitchCtrlLock_exit

End Function
