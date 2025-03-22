VERSION 5.00
Begin VB.Form Genre 
   Caption         =   "Genre"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   4635
   Begin VB.CommandButton Cmd_Delete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4440
      Begin VB.CommandButton Cmd_Cancel 
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Cmd_New 
         Caption         =   "New"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Cmd_Save 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Txt_ID 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   3855
      End
      Begin VB.ListBox Lst_List 
         Height          =   1425
         ItemData        =   "Genre.frx":0000
         Left            =   120
         List            =   "Genre.frx":0002
         TabIndex        =   3
         Top             =   2520
         Width           =   4215
      End
      Begin VB.TextBox Txt_Name 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   3855
      End
      Begin VB.Label Lbl_ID 
         Caption         =   "ID"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Lbl_Name 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Genre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Cancel_Click()
    Txt_ID = ""
    Txt_Name = ""
    Txt_Name.Enabled = False
    Cmd_New.Enabled = True
    Cmd_Cancel.Enabled = False
    Cmd_Save.Enabled = False
    Cmd_Delete.Enabled = False
    Lst_List.Enabled = True
End Sub

Private Sub Cmd_Delete_Click()
    Dim Str_Id As String
    
    If Not Lst_List.Enabled Or Lst_List.ListCount <= 0 Or Lst_List.ListIndex <= -1 Then Exit Sub
    If Shift <> 0 Or KeyCode <> vbKeyDelete Then Exit Sub
    If QuestionMsg("Are you sure?") = vbNo Then Exit Sub
    
    Str_Id = Split(Lst_List.List(Lst_List.ListIndex), "-")(0)
    
    ExecuteOnDB "DELETE FROM Genres WHERE ID = " & Str_Id
    
    InfoMsg "Data excluded successfully!"
    
    PopulateList
End Sub

Private Sub Cmd_New_Click()
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
    CenterForm Me, 4800, 4710
    
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
        Str_Query = "INSERT INTO Genres (Name) VALUES ('" & Txt_Name & "')"
    Else
        Str_Query = "UPDATE Genres SET Name = '" & Txt_Name & "' WHERE ID = " & Txt_ID
    End If
    
        ExecuteOnDB (Str_Query)
    
    InfoMsg "Data saved successfully!"
    
    Txt_ID = ""
    Txt_Name = ""
    Txt_Name.Enabled = False
    Cmd_New.Enabled = True
    Cmd_Cancel.Enabled = False
    Cmd_Save.Enabled = False
    Cmd_Delete.Enabled = False
    Lst_List.Enabled = True
    
    PopulateList
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err.Description
End Sub

Private Sub PopulateList()
    Dim Ado_List As ADODB.Recordset
    
    Lst_List.Clear
    
    SelectOnDB "SELECT * FROM Genres", Ado_List
    
    With Ado_List
        If Not .EOF Then
            Do While Not .EOF
                Lst_List.AddItem .Fields(0) & "-" & .Fields(1)
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set Ado_List = Nothing
End Sub

Private Sub Lst_List_Click()
    Cmd_Delete.Enabled = Lst_List.Enabled And Lst_List.ListCount > 0 And Lst_List.ListIndex > -1
End Sub

Private Sub Lst_List_DblClick()
    If Lst_List.ListCount <= 0 Or Lst_List.ListIndex <= -1 Then Exit Sub
    
    Txt_ID = Split(Lst_List.List(Lst_List.ListIndex), "-")(0)
    Txt_Name = Split(Lst_List.List(Lst_List.ListIndex), "-")(1)
    
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
