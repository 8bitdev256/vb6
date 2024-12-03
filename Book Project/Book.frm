VERSION 5.00
Begin VB.Form Book 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Registration"
   ClientHeight    =   9135
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   4905
   Icon            =   "Book.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frm_Main 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4695
      Begin VB.PictureBox Pic_Delete 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   550
         Left            =   3240
         Picture         =   "Book.frx":3802
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   23
         ToolTipText     =   "Delete (Ctrl + D)"
         Top             =   240
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.PictureBox Pic_Save 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   550
         Left            =   2520
         Picture         =   "Book.frx":191D8
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   22
         ToolTipText     =   "Save (Ctrl + S)"
         Top             =   240
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.PictureBox Pic_Cancel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   550
         Left            =   1800
         Picture         =   "Book.frx":26FE9
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   21
         ToolTipText     =   "Cancel (Ctrl + C)"
         Top             =   240
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.PictureBox Pic_Add 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   550
         Left            =   1080
         Picture         =   "Book.frx":27536
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   0
         ToolTipText     =   "New (Ctrl + N)"
         Top             =   240
         Width           =   550
      End
      Begin VB.Frame Frm_Genre 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Genre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   0
         TabIndex        =   15
         Top             =   7800
         Width           =   4695
         Begin VB.TextBox Txt_GenreName 
            Enabled         =   0   'False
            Height          =   330
            Left            =   840
            MaxLength       =   40
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox Txt_GenreId 
            Enabled         =   0   'False
            Height          =   330
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Lbl_GenreId 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Id"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Lbl_GenreName 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   840
            TabIndex        =   17
            Top             =   480
            Width           =   3615
         End
      End
      Begin VB.Frame Frm_Book 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Book"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   0
         TabIndex        =   10
         Top             =   1440
         Width           =   4695
         Begin VB.PictureBox Pic_Book 
            BackColor       =   &H00FFFFFF&
            Height          =   1335
            Left            =   1560
            ScaleHeight     =   1275
            ScaleWidth      =   1275
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Txt_BookPrice 
            Enabled         =   0   'False
            Height          =   330
            Left            =   3960
            TabIndex        =   2
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox Txt_BookName 
            Enabled         =   0   'False
            Height          =   330
            Left            =   120
            MaxLength       =   40
            TabIndex        =   1
            Top             =   2520
            Width           =   3735
         End
         Begin VB.TextBox Txt_BookId 
            Enabled         =   0   'False
            Height          =   330
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Lbl_BookPrice 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Price"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4035
            TabIndex        =   14
            Top             =   2280
            Width           =   465
         End
         Begin VB.Label Lbl_BookId 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Id"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Lbl_BookName 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   2280
            Width           =   3735
         End
      End
      Begin VB.Frame Frm_Author 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   0
         TabIndex        =   6
         Top             =   4800
         Width           =   4695
         Begin VB.PictureBox Pic_Author 
            BackColor       =   &H00FFFFFF&
            Height          =   1335
            Left            =   1560
            ScaleHeight     =   1275
            ScaleWidth      =   1275
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox Txt_AuthorId 
            Enabled         =   0   'False
            Height          =   330
            Left            =   120
            TabIndex        =   3
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox Txt_AuthorName 
            Enabled         =   0   'False
            Height          =   330
            Left            =   840
            MaxLength       =   40
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   2160
            Width           =   3735
         End
         Begin VB.Label Lbl_AuthorName 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   840
            TabIndex        =   9
            Top             =   1920
            Width           =   3615
         End
         Begin VB.Label Lbl_AuthorId 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Id"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1920
            Width           =   615
         End
      End
      Begin VB.Image Img_Last 
         Height          =   480
         Left            =   3240
         Picture         =   "Book.frx":27B5C
         ToolTipText     =   "Last"
         Top             =   840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Img_Next 
         Height          =   480
         Left            =   2520
         Picture         =   "Book.frx":2811C
         ToolTipText     =   "Next"
         Top             =   840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Img_Previous 
         Height          =   480
         Left            =   1800
         Picture         =   "Book.frx":286D6
         ToolTipText     =   "Previous"
         Top             =   840
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Img_First 
         Height          =   480
         Left            =   1080
         Picture         =   "Book.frx":28C79
         ToolTipText     =   "First"
         Top             =   840
         Visible         =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "Book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Pic_Add.SetFocus
End Sub

Private Sub Pic_Add_Click()
    Add
End Sub

Private Sub Pic_Add_KeyDown(KeyCode As Integer, Shift As Integer)
    TreatShortcutKeys Me, KeyCode, Shift
End Sub

Private Sub Pic_Cancel_Click()
    Cancel
End Sub

Private Sub Pic_Delete_Click()
    Delete
End Sub

Private Sub Pic_Save_Click()
    Save
End Sub

Private Sub Txt_AuthorId_KeyDown(KeyCode As Integer, Shift As Integer)
    SendTabOnPressEnter KeyCode, Shift
    TreatShortcutKeys Me, KeyCode, Shift
End Sub

Private Sub Txt_BookName_KeyDown(KeyCode As Integer, Shift As Integer)
    SendTabOnPressEnter KeyCode, Shift
    TreatShortcutKeys Me, KeyCode, Shift
End Sub

Private Sub Txt_BookPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    SendTabOnPressEnter KeyCode, Shift
    TreatShortcutKeys Me, KeyCode, Shift
End Sub

Private Sub Txt_GenreId_KeyDown(KeyCode As Integer, Shift As Integer)
    SendTabOnPressEnter KeyCode, Shift
    TreatShortcutKeys Me, KeyCode, Shift
End Sub

Private Sub Add()
    Pic_Add.Visible = False
    Pic_Cancel.Visible = True
    Pic_Save.Visible = True
    Pic_Delete.Visible = False
    
    Txt_BookName.Enabled = True
    Txt_BookPrice.Enabled = True
    
    Txt_AuthorId.Enabled = True
    Txt_GenreId.Enabled = True
    
    Txt_BookName.SetFocus
End Sub

Private Sub Cancel()
    Pic_Add.Visible = True
    Pic_Cancel.Visible = False
    Pic_Save.Visible = False
    Pic_Delete.Visible = False
End Sub

Private Sub Save()
    Pic_Add.Visible = True
    Pic_Cancel.Visible = False
    Pic_Save.Visible = False
    Pic_Delete.Visible = False
End Sub

Private Sub Delete()
    Pic_Add.Visible = True
    Pic_Cancel.Visible = False
    Pic_Save.Visible = False
    Pic_Delete.Visible = False
End Sub
