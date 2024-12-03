Attribute VB_Name = "Mod_DB"
Public Sub SetupDB(Str_Server As String, Str_DB As String, Str_User As String, Str_Pass As String)
    Dim DBConn As ADODB.Connection
    'Dim ADO_Rs As New ADODB.Recordset
    'Dim Str_SQL As String
    
    On Error GoTo ErrorHandler
    
    Set DBConn = New ADODB.Connection
    
    DBConn.ConnectionString = "Provider=MSOLEDBSQL;Server=" & Str_Server & ";Database=" & Str_DB & ";UID=" & Str_User & ";PWD=" & Str_Pass & ";"
    
    DBConn.Open
    
    
    
    'Set ADO_Rs = New ADODB.Recordset
    
    'Str_SQL = "select * from Employee"
     
    'myRecSet.Open Str_SQL, MyConnObj, adOpenKeyset
    
    InfoMsg "Connection established successfully!"
    
    Exit Sub
    
ErrorHandler:
    ErrorMsg Err.Description
End Sub
