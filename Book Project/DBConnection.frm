VERSION 5.00
Begin VB.Form DBConnection 
   Caption         =   "DB Connection"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt_Password 
      Height          =   375
      Left            =   135
      TabIndex        =   8
      Tag             =   "Password"
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox Txt_User 
      Height          =   375
      Left            =   135
      TabIndex        =   6
      Tag             =   "User"
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox Txt_DB 
      Height          =   375
      Left            =   135
      TabIndex        =   4
      Tag             =   "Database"
      Top             =   2400
      Width           =   3975
   End
   Begin VB.TextBox Txt_Server 
      Height          =   375
      Left            =   135
      TabIndex        =   2
      Tag             =   "Server"
      Top             =   1440
      Width           =   3975
   End
   Begin VB.CommandButton Cmd_Save 
      Caption         =   "Save"
      Height          =   735
      Left            =   1695
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Lbl_Password 
      Caption         =   "Password"
      Height          =   255
      Left            =   195
      TabIndex        =   7
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Lbl_User 
      Caption         =   "User"
      Height          =   255
      Left            =   195
      TabIndex        =   5
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Label Lbl_DB 
      Caption         =   "Database"
      Height          =   255
      Left            =   195
      TabIndex        =   3
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Lbl_Server 
      Caption         =   "Server"
      Height          =   255
      Left            =   195
      TabIndex        =   1
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "DBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Save_Click()
    If Save Then
        InfoMsg "Data saved successfully!"
        Unload Me
    End If
End Sub

Private Sub Txt_DB_KeyDown(KeyCode As Integer, Shift As Integer)
    SendTabOnPressEnter KeyCode, Shift
    TreatShortcutKeys Me, KeyCode, Shift
End Sub

Private Sub Txt_Password_KeyDown(KeyCode As Integer, Shift As Integer)
    SendTabOnPressEnter KeyCode, Shift
    TreatShortcutKeys Me, KeyCode, Shift
End Sub

Private Sub Txt_Server_KeyDown(KeyCode As Integer, Shift As Integer)
    SendTabOnPressEnter KeyCode, Shift
    TreatShortcutKeys Me, KeyCode, Shift
End Sub

Private Sub Txt_User_KeyDown(KeyCode As Integer, Shift As Integer)
    SendTabOnPressEnter KeyCode, Shift
    TreatShortcutKeys Me, KeyCode, Shift
End Sub

Public Function Save() As Boolean
    Dim Str_FieldName As String
    Dim Int_DBConfigFileNumber As Integer
    
    Str_Msg = ""
    
    If Txt_Server = "" Then
        ShowBlankFieldMsg Txt_Server
        Exit Function
    ElseIf Txt_DB = "" Then
        ShowBlankFieldMsg Txt_DB
        Exit Function
    ElseIf Txt_User = "" Then
        ShowBlankFieldMsg Txt_User
        Exit Function
    ElseIf Txt_Password = "" Then
        ShowBlankFieldMsg Txt_Password
        Exit Function
    End If
    
    If Not ConnectedToDB(Txt_Server, Txt_DB, Txt_User, Txt_Password) Then Exit Function
    
    If Dir(DBConfigDirPath, vbDirectory) = "" Then
        MkDir DBConfigDirPath
    End If
    
    Int_DBConfigFileNumber = FreeFile

    Open DBConfigFilePath For Output As Int_DBConfigFileNumber

    Print #Int_DBConfigFileNumber, Txt_Server
    Print #Int_DBConfigFileNumber, Txt_DB
    Print #Int_DBConfigFileNumber, Txt_User
    Print #Int_DBConfigFileNumber, Txt_Password
    
    Close #Int_DBConfigFileNumber
    
    Save = True
    
    Exit Function
    
ErrorHandler:
    ErrorMsg Err.Description
End Function

Private Sub ShowBlankFieldMsg(Txt_Box As TextBox)
    WarningMsg Txt_Box.Tag & " field cannot be blank!"
    Txt_Box.SetFocus
    Exit Sub
End Sub

