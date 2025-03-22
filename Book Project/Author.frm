VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Author 
   Caption         =   "Author"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   4635
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   97
      TabIndex        =   0
      Top             =   120
      Width           =   4440
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3840
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Cmd_ChangePic 
         Caption         =   "Change"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.ListBox Lst_List 
         Height          =   1425
         ItemData        =   "Author.frx":0000
         Left            =   120
         List            =   "Author.frx":0002
         TabIndex        =   9
         Top             =   3960
         Width           =   4215
      End
      Begin VB.CommandButton Cmd_Delete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox Txt_Name 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   4215
      End
      Begin VB.TextBox Txt_ID 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   4215
      End
      Begin VB.CommandButton Cmd_Save 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton Cmd_New 
         Caption         =   "New"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton Cmd_Cancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   3480
         Width           =   975
      End
      Begin VB.Image Img_ProfilePicture 
         Height          =   1575
         Left            =   1350
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Lbl_Name 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Lbl_ID 
         Caption         =   "ID"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Author"
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
    CommonDialog1.Filter = "Images (*.bmp)|*.bmp"
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.DialogTitle = "Select File"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
        Img_ProfilePicture.Picture = LoadPicture(CommonDialog1.FileName)
    End If
End Sub

Private Sub Cmd_Delete_Click()
    Dim Str_Id As String
    
    If Not Lst_List.Enabled Or Lst_List.ListCount <= 0 Or Lst_List.ListIndex <= -1 Then Exit Sub
    If Shift <> 0 Or KeyCode <> vbKeyDelete Then Exit Sub
    If QuestionMsg("Are you sure?") = vbNo Then Exit Sub
    
    Str_Id = Split(Lst_List.List(Lst_List.ListIndex), "-")(0)
    
    ExecuteOnDB "DELETE FROM Authors WHERE ID = " & Str_Id
    
    InfoMsg "Data excluded successfully!"
    
    PopulateList
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
    CenterForm Me, 4800, 6255
    
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
        Str_Query = "INSERT INTO Authors (Name) VALUES ('" & Txt_Name & "')"
    Else
        Str_Query = "UPDATE Authors SET Name = '" & Txt_Name & "' WHERE Id = " & Txt_ID
    End If
    
    ExecuteOnDB (Str_Query)
    
    SavePictureToDB "Authors", Txt_ID
    
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
    
    Lst_List.Clear
    
    SelectOnDB "SELECT * FROM Authors", Ado_List
    
    With Ado_List
        If Not .EOF Then
            Do While Not .EOF
                Lst_List.AddItem .Fields(0) & "-" & .Fields(1)
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
    Img_ProfilePicture.Picture = LoadPicture(ProjectDirPath & "\Icons\blank-profile-picture.bmp")
End Sub

Private Sub LoadPictureFromDB(Ado_List As ADODB.Recordset)
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
    Img_ProfilePicture.Picture = LoadPicture(Str_ImageFilePath)
End Sub

Private Sub SavePictureToDB(Str_TableName As String, Str_Id As String)
    Dim rs As ADODB.Recordset
    Dim mystream As ADODB.Stream
    Dim Ado_Table As ADODB.Recordset
    Dim Str_Query As String
    
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
    
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    SelectOnDB "SELECT * FROM " & Str_TableName & " WHERE Id = " & Str_Id, rs
    mystream.Open
    mystream.LoadFromFile TempImagesDirPath & Str_Id & ".bmp"
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
