VERSION 5.00
Begin VB.Form ConnectToDB 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect to DB"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   Icon            =   "ConnectToDB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frm_Main 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SQL Server Connection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Pic_Save 
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Tag             =   "User"
         Top             =   3000
         Width           =   3855
      End
      Begin VB.TextBox Txt_DB 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Tag             =   "Server"
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Lbl_Password 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3480
         Width           =   3855
      End
      Begin VB.Label Lbl_User 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label Lbl_DB 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Lbl_Server 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
    WarningMsg Txt_Box.Tag & " field cannot be blank!"
    Txt_Box.SetFocus
    Exit Sub
End Sub
