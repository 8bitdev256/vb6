Attribute VB_Name = "Mod_Forms"

Public Property Get ProjectDirPath() As String
    ProjectDirPath = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "")
End Property

Public Sub InfoMsgBox(Str_Msg As String)
    LoadMsgBox Str_Msg, "info"
End Sub

Public Sub WarningMsgBox(Str_Msg As String)
    LoadMsgBox Str_Msg, "warning"
End Sub

Public Sub ErrorMsgBox(Str_Msg As String)
    LoadMsgBox Str_Msg, "error"
End Sub

Public Sub QuestionMsgBox(Str_Msg As String)
    LoadMsgBox Str_Msg, "question", True
End Sub

Private Sub LoadMsgBox(Str_Msg As String, Str_PictureName As String, Optional Bol_IsQuestionMsg As Boolean = False)
    MessageBox.IsQuestionMsg = Bol_IsQuestionMsg
    MessageBox.Img_MsgType.Picture = LoadPicture(ProjectDirPath & "Icons\" & Str_PictureName & ".ico")
    MessageBox.Lbl_Msg = Str_Msg
    MessageBox.Show 1
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
