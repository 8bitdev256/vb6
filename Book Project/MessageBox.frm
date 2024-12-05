VERSION 5.00
Begin VB.Form MessageBox 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
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
      Height          =   2955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4365
      Begin VB.Image Img_MsgType 
         Height          =   480
         Left            =   1920
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Lbl_Msg 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Message"
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
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Lbl_Yes 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Yes"
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
         Left            =   480
         TabIndex        =   3
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label Lbl_No 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "No"
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
         Left            =   2640
         TabIndex        =   2
         Top             =   2400
         Width           =   1065
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
         TabIndex        =   1
         Top             =   120
         Width           =   315
      End
   End
End
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Bol_IsQuestionMsg As Boolean
Private Bol_ClickedOnYes As Boolean

Public Property Get IsQuestionMsg() As Boolean
    IsQuestionMsg = Bol_IsQuestionMsg
End Property

Public Property Let IsQuestionMsg(ByVal vNewValue As Boolean)
    Bol_IsQuestionMsg = vbNewValue
End Property

Public Property Get ClickedOnYes() As Boolean
    ClickedOnYes = Bol_ClickedOnYes
End Property

Public Property Let ClickedOnYes(ByVal vNewValue As Boolean)
    Bol_ClickedOnYes = vNewValue
End Property

Private Sub Form_Load()
    ClickedOnYes = False
    
    If Not IsQuestionMsg Then
        Lbl_No.Visible = False
        Lbl_Yes.Caption = "OK"
        Lbl_Yes.Left = 1650
    End If
End Sub

Private Sub Lbl_No_Click()
    Unload Me
End Sub

Private Sub Lbl_Yes_Click()
    ClickedOnYes = True
    Unload Me
End Sub

