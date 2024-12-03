Attribute VB_Name = "Mod_Forms"

Public Sub ErrorMsg(Str_Msg As String)
    MsgBox Str_Msg, vbCritical, "Error"
End Sub

Public Sub InfoMsg(Str_Msg As String)
    MsgBox Str_Msg, vbInformation, "Success"
End Sub

Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub

Public Sub SendTabOnPressEnter(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
    End If
End Sub

Public Sub WarningMsg(Str_Msg As String)
    MsgBox Str_Msg, vbExclamation, "Warning"
End Sub

Public Sub TreatShortcutKeys(FormObj As Form, KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorHandler
    
    If Shift = 2 Then
        If KeyCode = vbKeyN Then
            FormObj.Add
        ElseIf KeyCode = vbKeyC Then
            FormObj.Cancel
        ElseIf KeyCode = vbKeyS Then
            FormObj.Save
        ElseIf KeyCode = vbKeyD Then
            FormObj.Delete
        End If
    End If
    
ErrorHandler:
    Exit Sub
End Sub
