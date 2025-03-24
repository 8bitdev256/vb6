VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Book Registration System"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Registration 
      Caption         =   "Registration"
      Begin VB.Menu RegistrationGenre 
         Caption         =   "Genre"
      End
      Begin VB.Menu RegistrationAuthor 
         Caption         =   "Author"
      End
      Begin VB.Menu RegistrationBook 
         Caption         =   "Book"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Dim Bol_ConnectedToDB As Boolean
    
    If Not DBConfigFileExists Then
        DBConnection.Show 1
    End If
    
    If Not DBConfigFileExists Then
        Unload Me
        Exit Sub
    End If
    
    If Not ConnectedToDBThroughConfigFile Then
        WarningMsg "DB config file have wrong config!"
        DBConnection.Show 1
    End If
    
    If Not ConnectedToDBThroughConfigFile Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub RegistrationAuthor_Click()
    Author.Show
End Sub

Private Sub RegistrationBook_Click()
    Book.Show
End Sub

Private Sub RegistrationGenre_Click()
    Genre.Show
End Sub

