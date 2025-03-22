Attribute VB_Name = "Mod_File"

Public Property Get TempImagesDirPath() As String
    TempImagesDirPath = ProjectDirPath & "TempImages\"
End Property

Public Function ReadFile(Str_FilePath As String) As String
    Dim Str_Lines As String
    Dim Int_FileNumber As Integer
    
    Str_Lines = ""
    
    Int_FileNumber = FreeFile
    
    Open Str_FilePath For Input As #Int_FileNumber
    Do While Not EOF(Int_FileNumber)
        Line Input #Int_FileNumber, Str_Line
        Str_Lines = Str_Lines & Str_Line & vbCrLf
    Loop

    Close #Int_FileNumber
    
    ReadFile = Str_Lines
End Function
