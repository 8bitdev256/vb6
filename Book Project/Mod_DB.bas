Attribute VB_Name = "Mod_DB"
Private DBConfig As New Cls_DBConfig

Public Enum QueryType
    InsertQuery
    UpdateQuery
    DeleteQuery
    SelectQuery
End Enum

Public Property Get DBConfigDirPath() As String
    DBConfigDirPath = ProjectDirPath & "DBConfig"
End Property

Public Property Get DBConfigFilePath() As String
    DBConfigFilePath = DBConfigDirPath & "\DBConfig.txt"
End Property

Public Function ConnectedToDB(Str_Server As String, Str_DB As String, Str_User As String, Str_Pass As String) As Boolean
    Dim DBConn As ADODB.Connection
    
    On Error GoTo ErrorHandler
    
    Set DBConn = New ADODB.Connection
    
    DBConn.ConnectionString = ReturnDBConnectionString(Str_Server, Str_DB, Str_User, Str_Pass)
    
    DBConn.Open
    
    DBConn.Close
    
    DBConfig.DBServer = Str_Server
    DBConfig.DBDatabase = Str_DB
    DBConfig.DBUser = Str_User
    DBConfig.DBPassword = Str_Pass
    
    ConnectedToDB = True
    
    Exit Function
    
ErrorHandler:
    ErrorMsg Err.Description
End Function

Private Function ReturnDBConnectionString(Str_Server As String, Str_DB As String, Str_User As String, Str_Pass As String) As String
    ReturnDBConnectionString = "Provider=SQLOLEDB.1" & ";Server=" & Str_Server & ";Database=" & Str_DB & ";UID=" & Str_User & ";PWD=" & Str_Pass & ";"
End Function

Public Function DBConfigFileExists() As Boolean
    If Dir(DBConfigDirPath, vbDirectory) = "" Or Dir(DBConfigFilePath) = "" Then
        DBConfigFileExists = False
        Exit Function
    End If
    
    DBConfigFileExists = True
End Function

Public Function ConnectedToDBThroughConfigFile() As Boolean
    Dim Str_DBConfigFileLines As String
    Dim Arr_DBConfigFileLines() As String
    Dim Int_Index As Integer
    Dim Str_Server As String
    Dim Str_DB As String
    Dim Str_User As String
    Dim Str_Password As String
    Dim Bol_Connected As Boolean
    
    Str_DBConfigFileLines = ReadFile(DBConfigFilePath)
    
    Arr_DBConfigFileLines = Split(Str_DBConfigFileLines, vbCrLf)
    
    Str_Server = Arr_DBConfigFileLines(0)
    Str_DB = Arr_DBConfigFileLines(1)
    Str_User = Arr_DBConfigFileLines(2)
    Str_Password = Arr_DBConfigFileLines(3)
    
    Bol_Connected = ConnectedToDB(Str_Server, Str_DB, Str_User, Str_Password)
    
    ConnectedToDBThroughConfigFile = Bol_Connected
End Function

Public Sub SelectOnDB(Str_Query As String, Ado_Rs As ADODB.Recordset)
    Dim Ado_Conn As New ADODB.Connection
    
    Ado_Conn.ConnectionString = ReturnDBConnectionString(DBConfig.DBServer, DBConfig.DBDatabase, DBConfig.DBUser, DBConfig.DBPassword)
    Ado_Conn.Open
    
    Set Ado_Rs = New ADODB.Recordset
    Ado_Rs.Open Str_Query, Ado_Conn, adOpenDynamic, adLockOptimistic
End Sub

Public Sub ExecuteOnDB(Str_Query As String)
    Dim Ado_Conn As New ADODB.Connection
    Dim Ado_Comm As New ADODB.Command
       
    Ado_Conn.ConnectionString = ReturnDBConnectionString(DBConfig.DBServer, DBConfig.DBDatabase, DBConfig.DBUser, DBConfig.DBPassword)
    Ado_Conn.Open
    
    Ado_Comm.ActiveConnection = Ado_Conn
    Ado_Comm.CommandType = adCmdText
    Ado_Comm.CommandText = Str_Query
    Ado_Comm.Execute
End Sub
