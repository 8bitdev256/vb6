VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Book 
   Caption         =   "Book"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10890
   ScaleWidth      =   9825
   Begin VB.Frame Frame1 
      Height          =   10575
      Left            =   112
      TabIndex        =   0
      Top             =   120
      Width           =   9600
      Begin VB.Frame Frm_Book 
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
         Height          =   3135
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   9375
         Begin VB.TextBox Txt_BookPrice 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   2520
            Width           =   4215
         End
         Begin VB.CommandButton Cmd_ChangeBookImage 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   6120
            TabIndex        =   19
            Top             =   2520
            Width           =   375
         End
         Begin VB.TextBox Txt_BookName 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   1680
            Width           =   4215
         End
         Begin VB.TextBox Txt_BookID 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   4215
         End
         Begin MSComDlg.CommonDialog Cod_Book 
            Left            =   6120
            Top             =   1920
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Lbl_BookPrice 
            Caption         =   "Price"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Image Img_BookImage 
            Height          =   1575
            Left            =   4440
            Stretch         =   -1  'True
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Lbl_BookName 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Lbl_BookID 
            Caption         =   "ID"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frm_Genre 
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
         Height          =   2175
         Left            =   120
         TabIndex        =   11
         Top             =   6000
         Width           =   9375
         Begin VB.TextBox Txt_GenreID 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   4215
         End
         Begin VB.TextBox Txt_GenreName 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   1560
            Width           =   4215
         End
         Begin VB.Label Lbl_GenreID 
            Caption         =   "ID"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Lbl_GenreName 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame Frm_Author 
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
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   3600
         Width           =   9375
         Begin VB.TextBox Txt_AuthorName 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   1560
            Width           =   4215
         End
         Begin VB.TextBox Txt_AuthorID 
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   4215
         End
         Begin VB.Label Lbl_AuthorName 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Lbl_AuthorID 
            Caption         =   "ID"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
         Begin VB.Image Img_AuthorImage 
            Height          =   1575
            Left            =   4440
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.CommandButton Cmd_Cancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   8760
         Width           =   975
      End
      Begin VB.CommandButton Cmd_New 
         Caption         =   "New"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   8760
         Width           =   975
      End
      Begin VB.CommandButton Cmd_Save 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   8760
         Width           =   975
      End
      Begin VB.CommandButton Cmd_Delete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   8760
         Width           =   975
      End
      Begin VB.ListBox Lst_List 
         Height          =   1035
         ItemData        =   "Book.frx":0000
         Left            =   120
         List            =   "Book.frx":0002
         TabIndex        =   1
         Top             =   9240
         Width           =   9375
      End
   End
End
Attribute VB_Name = "Book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Ado_List As ADODB.Recordset

Private Sub Cmd_Cancel_Click()
    Txt_ID = ""
    Txt_Name = ""
    Txt_Name.Enabled = False
    Cmd_New.Enabled = True
    Cmd_Cancel.Enabled = False
    Cmd_Save.Enabled = False
    Cmd_Delete.Enabled = Lst_List.ListCount > 0 And Lst_List.ListIndex > -1
    Lst_List.Enabled = True
    Cmd_ChangePic.Enabled = False
    LoadBlankProfilePicture
End Sub

Private Sub Cmd_ChangePic_Click()
    CommonDialog1.FileName = ""
    CommonDialog1.Filter = "Images (*.bmp)|*.bmp"
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.DialogTitle = "Select File"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
        Img_Book.Picture = LoadPicture(CommonDialog1.FileName)
    End If
End Sub

Private Sub Cmd_Delete_Click()
    Delete
End Sub

Private Sub Cmd_New_Click()
    LoadBlankProfilePicture
    
    Cmd_ChangePic.Enabled = True
    Txt_ID = ""
    Txt_Name = ""
    Txt_Name.Enabled = True
    Cmd_New.Enabled = False
    Cmd_Cancel.Enabled = True
    Cmd_Save.Enabled = True
    Cmd_Delete.Enabled = False
    Lst_List.Enabled = False
    
    Txt_Name.SetFocus
End Sub

Private Sub Cmd_Save_Click()
    Save
End Sub

Private Sub Form_Load()
    CenterForm Me, 9945, 11400
    
    LoadBlankProfilePicture
    
    PopulateList
End Sub

Private Sub Save()
    Dim Str_Query As String
    
    On Error GoTo ErrorHandler
    
    If Txt_Name = "" Then
        WarningMsg "Name cannot be blank!"
        Exit Sub
    End If
    
    If Txt_ID = "" Then
        Str_Query = "" _
        & vbNewLine & "INSERT INTO Books" _
        & vbNewLine & "(Name, Price, AuthorId, GenreId)" _
        & vbNewLine & "VALUES (" _
        & vbNewLine & "'" & Txt_Name & "'," _
        & vbNewLine & Replace(Txt_BookPrice, ",", ".") & "," _
        & vbNewLine & Txt_AuthorID & "," _
        & vbNewLine & Txt_GenreID _
        & vbNewLine & ")"
    Else
        Str_Query = "" _
        & vbNewLine & "UPDATE Books SET" _
        & vbNewLine & "Name = '" & Txt_Name & "'," _
        & vbNewLine & "Price = '" & Replace(Txt_BookPrice, ",", ".") & "," _
        & vbNewLine & "AuthorId = '" & Txt_AuthorID & "'," _
        & vbNewLine & "GenreId = '" & Txt_GenreID & "'," _
        & vbNewLine & "WHERE Id = " & Txt_ID
    End If
    
    ExecuteOnDB (Str_Query)
    
    SavePictureToDB "Books", Txt_ID
    
    InfoMsg "Data saved successfully!"
    
    Txt_ID = ""
    Txt_Name = ""
    Txt_Name.Enabled = False
    Cmd_New.Enabled = True
    Cmd_Cancel.Enabled = False
    Cmd_Save.Enabled = False
    Cmd_Delete.Enabled = False
    Lst_List.Enabled = True
    Cmd_ChangePic.Enabled = False
    LoadBlankProfilePicture
    
    PopulateList
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err.Description
End Sub

Private Sub PopulateList()
    Dim Str_Query As String
    Dim Str_Book As String
    
    LoadBlankProfilePicture
    Txt_ID = ""
    Txt_Name = ""
    Lst_List.Clear
       
    Str_Query = "" _
    & "SELECT" _
    & vbNewLine & "b.Id BookID, b.Name BookName, b.Price BookPrice, b.Picture BookPicture," _
    & vbNewLine & "a.Id AuthorID, a.Name AuthorName, a.Picture AuthorPicture," _
    & vbNewLine & "g.Id GenreID, g.Name GenreName" _
    & vbNewLine & "FROM Books b" _
    & vbNewLine & "INNER JOIN Authors a ON b.AuthorId = a.Id" _
    & vbNewLine & "INNER JOIN Genres g ON b.GenreId = g.Id"
    
    SelectOnDB Str_Query, Ado_List
    
    With Ado_List
        If Not .EOF Then
            Do While Not .EOF
                Str_Book = "" _
                & .Fields("BookID") _
                & "-" & .Fields("BookName") _
                & "-" & .Fields("BookPrice") _
                & "-" & .Fields("AuthorID") _
                & "-" & .Fields("AuthorName") _
                & "-" & .Fields("GenreID") _
                & "-" & .Fields("GenreName")
                
                If Not IsNull(.Fields("Picture")) Then
                    LoadPictureFromDB Ado_List, False
                End If
                Lst_List.AddItem Str_Book
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub Lst_List_Click()
    If Lst_List.ListCount <= 0 Or Lst_List.ListIndex <= -1 Then Exit Sub
    
    Cmd_Delete.Enabled = Lst_List.Enabled
    
    LoadAuthorData
End Sub

Private Sub Lst_List_DblClick()
    If Lst_List.ListCount <= 0 Or Lst_List.ListIndex <= -1 Then Exit Sub
    
    LoadAuthorData
    
    Cmd_ChangePic.Enabled = True
    Txt_Name.Enabled = True
    Cmd_New.Enabled = False
    Cmd_Cancel.Enabled = True
    Cmd_Save.Enabled = True
    Cmd_Delete.Enabled = False
    Lst_List.Enabled = False
    
    Txt_Name.SetFocus
    Txt_Name.SelStart = 0
    Txt_Name.SelLength = Len(Txt_Name)
End Sub

Private Sub LoadBlankProfilePicture()
    Img_BookImage.Picture = LoadPicture(ProjectDirPath & "Icons\blank-profile-picture.bmp")
End Sub

Private Sub LoadPictureFromDB(Ado_List As ADODB.Recordset, Optional Bol_ReplaceImageControlPicture As Boolean = True)
    Dim Str_ImageFilePath As String
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    
    Str_ImageFilePath = TempImagesDirPath & Ado_List("Id") & ".bmp"
    
    strStream.Write Ado_List.Fields("Picture").Value
    
    If Dir(TempImagesDirPath, vbDirectory) = "" Then
        MkDir TempImagesDirPath
    End If
    
    strStream.SaveToFile Str_ImageFilePath, adSaveCreateOverWrite
    
    If Bol_ReplaceImageControlPicture Then
        Img_Book.Picture = LoadPicture(Str_ImageFilePath)
    End If
End Sub

Private Sub SavePictureToDB(Str_TableName As String, Str_Id As String)
    Dim rs As ADODB.Recordset
    Dim mystream As ADODB.Stream
    Dim Ado_Table As ADODB.Recordset
    Dim Str_Query As String
    Dim Str_FileName As String
    Dim Str_TempImageFilePath As String
            
    If Str_Id = "" Then
        Str_Query = "SELECT MAX(Id) Id FROM " & Str_TableName
         SelectOnDB Str_Query, Ado_Table
         
         With Ado_Table
            If Not .EOF Then
                Str_Id = .Fields(0)
            End If
            .Close
         End With
         Set Ado_Table = Nothing
    End If
       
    Str_TempImageFilePath = TempImagesDirPath & Str_Id & ".bmp"
    
    If Dir(CommonDialog1.FileName) = "" And Dir(Str_TempImageFilePath) = "" Then
        Exit Sub
    End If
    
    Str_FileName = ""
    
    If Dir(CommonDialog1.FileName) <> "" Then
        Str_FileName = CommonDialog1.FileName
    ElseIf Dir(Str_TempImageFilePath) <> "" Then
        Str_FileName = Str_TempImageFilePath
    End If
    
    If Str_FileName = "" Then
        Exit Sub
    End If
    
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    SelectOnDB "SELECT * FROM " & Str_TableName & " WHERE Id = " & Str_Id, rs
    mystream.Open
    mystream.LoadFromFile Str_FileName
    rs("Picture") = mystream.Read
    rs.Update
    mystream.Close
    rs.Close
    Set rs = Nothing
End Sub

Private Sub LoadAuthorData()
    Txt_ID = Split(Lst_List.List(Lst_List.ListIndex), "-")(0)
    Txt_Name = Split(Lst_List.List(Lst_List.ListIndex), "-")(1)
    
    LoadBlankProfilePicture
    Ado_List.MoveFirst
    Ado_List.Find "Id = " & Txt_ID
    
    With Ado_List
        If Not .EOF Then
            If Not IsNull(.Fields("Picture")) Then
                LoadPictureFromDB Ado_List
            End If
        End If
    End With
End Sub

Private Sub Delete()
    Dim Str_Id As String
    Dim Str_TempImageFilePath As String
    
    If Not Lst_List.Enabled Or Lst_List.ListCount <= 0 Or Lst_List.ListIndex <= -1 Then Exit Sub
    If QuestionMsg("Are you sure?") = vbNo Then Exit Sub
    
    Cmd_Delete.Enabled = False
    
    Str_Id = Split(Lst_List.List(Lst_List.ListIndex), "-")(0)
    
    ExecuteOnDB "DELETE FROM Authors WHERE ID = " & Str_Id
    
    Str_TempImageFilePath = TempImagesDirPath & Str_Id & ".bmp"
    
    If Dir(Str_TempImageFilePath) <> "" Then
        Kill Str_TempImageFilePath
    End If
    
    InfoMsg "Data excluded successfully!"
    
    PopulateList
End Sub

Private Sub Lst_List_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not (Shift = 0 And KeyCode = vbKeyDelete) Then Exit Sub
    
    Delete
End Sub

