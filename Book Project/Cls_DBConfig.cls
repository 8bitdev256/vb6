VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_DBConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Str_DBServer As String
Private Str_DBDatabase As String
Private Str_DBUser As String
Private Str_DBPassword As String

Public Property Get DBServer() As String
    DBServer = Str_DBServer
End Property

Public Property Let DBServer(ByVal vNewValue As String)
    Str_DBServer = vNewValue
End Property

Public Property Get DBDatabase() As String
    DBDatabase = Str_DBDatabase
End Property

Public Property Let DBDatabase(ByVal vNewValue As String)
    Str_DBDatabase = vNewValue
End Property

Public Property Get DBUser() As String
    DBUser = Str_DBUser
End Property

Public Property Let DBUser(ByVal vNewValue As String)
    Str_DBUser = vNewValue
End Property

Public Property Get DBPassword() As String
    DBPassword = Str_DBPassword
End Property

Public Property Let DBPassword(ByVal vNewValue As String)
    Str_DBPassword = vNewValue
End Property
