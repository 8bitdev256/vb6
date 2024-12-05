VERSION 5.00
Begin VB.Form ConnectToDB 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Connect to DB"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   Icon            =   "ConnectToDB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frm_Main 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   4365
      Begin VB.PictureBox Pic_Save 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   550
         Left            =   1920
         Picture         =   "ConnectToDB.frx":3802
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   9
         ToolTipText     =   "Save (Ctrl + S)"
         Top             =   360
         Width           =   550
      End
      Begin VB.TextBox Txt_Password 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   7
         Tag             =   "Password"
         Top             =   3720
         Width           =   3855
      End
      Begin VB.TextBox Txt_User 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Tag             =   "User"
         Top             =   3000
         Width           =   3855
      End
      Begin VB.TextBox Txt_DB 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Tag             =   "Database"
         Text            =   "BooksDB"
         Top             =   2280
         Width           =   3855
      End
      Begin VB.TextBox Txt_Server 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Tag             =   "Server"
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Lbl_Close 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   4030
         TabIndex        =   10
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Lbl_Password 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3480
         Width           =   3855
      End
      Begin VB.Label Lbl_User 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Lbl_DB 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Lbl_Server 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   3855
      End
   End
End
Attribute VB_Name = "ConnectToDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Lbl_Close_Click()
    Unload Me
End Sub

Private Sub Pic_Save_Click()
    Save
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

Public Sub Save()
    Dim Str_FieldName As String
    
    Str_Msg = ""
    
    If Txt_Server = "" Then
        ShowBlankFieldMsg Txt_Server
        Exit Sub
    ElseIf Txt_DB = "" Then
        ShowBlankFieldMsg Txt_DB
        Exit Sub
    ElseIf Txt_User = "" Then
        ShowBlankFieldMsg Txt_User
        Exit Sub
    ElseIf Txt_Password = "" Then
        ShowBlankFieldMsg Txt_Password
        Exit Sub
    End If
    
    SetupDB Txt_Server, Txt_DB, Txt_User, Txt_Password
End Sub

Private Sub ShowBlankFieldMsg(Txt_Box As TextBox)
    WarningMsgBox Txt_Box.Tag & " field cannot be blank!"
    Txt_Box.SetFocus
    Exit Sub
End Sub
