Attribute VB_Name = "Mod_Forms"

Public Property Get ProjectDirPath() As String
    ProjectDirPath = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "")
End Property

Public Sub InfoMsg(Str_Msg As String)
    MsgBox Str_Msg, vbInformation, ""
End Sub

Public Sub WarningMsg(Str_Msg As String)
    MsgBox Str_Msg, vbExclamation, ""
End Sub

Public Sub ErrorMsg(Str_Msg As String)
    MsgBox Str_Msg, vbCritical, ""
End Sub

Public Function QuestionMsg(Str_Msg As String) As VbMsgBoxResult
    QuestionMsgBox = MsgBox(Str_Msg, vbYesNo + vbQuestion, "")
End Function

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

Public Sub CenterForm(FormObj As Form, FormWidth As Integer, FormHeight As Integer)
    FormObj.Width = FormWidth
    FormObj.Height = FormHeight
    FormObj.Left = (MDIMain.ScaleWidth - FormObj.Width) / 2
    FormObj.Top = (MDIMain.ScaleHeight - FormObj.Height) / 2
End Sub
