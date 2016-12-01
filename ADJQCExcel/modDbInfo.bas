Attribute VB_Name = "modDbInfo"
Option Explicit



Public Function GetRsBySql(ByVal strSQL As String, ByVal conn As ADODB.Connection) As ADODB.Recordset
    On Error GoTo ErrHandle
    Dim rs As New ADODB.Recordset
'    Dim connTest As New ADODB.Connection
    rs.CursorLocation = adUseClient
    rs.LockType = adLockOptimistic
'    connTest.Open conn.ConnectionString
    rs.Open strSQL, conn
    Set GetRsBySql = rs
    Set rs = Nothing
    Set conn = Nothing
    Exit Function
ErrHandle:
    Err.Raise Err.Number, , Err.Description
    Err.Clear
End Function

'***********************************************************************
'函数名;GetReadOnlyRsBySql(ByVal strSql As String) As ADODB.Recordset
'函数作用：从数据库中取出数据(记录集只具有只读游标)
'***********************************************************************
Public Function GetReadOnlyRsBySql(ByVal strSQL As String, ByVal conn As ADODB.Connection) As ADODB.Recordset
    On Error GoTo ErrHandle
    Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.LockType = adLockReadOnly
    rs.Open strSQL, conn
    Set GetReadOnlyRsBySql = rs
    Set rs = Nothing
    Set conn = Nothing
    Exit Function
ErrHandle:
    Err.Raise Err.Number, , Err.Description
    Err.Clear
End Function

Public Function ConnectDB(ByVal strConnectString As String, Optional ByVal intConnectTimeOut As Integer, Optional intCommandTimeOut As Integer) As ADODB.Connection
    Dim SourceConn As New ADODB.Connection
    SourceConn.ConnectionTimeout = intConnectTimeOut
    SourceConn.CommandTimeout = intCommandTimeOut
    SourceConn.Open strConnectString
    Set ConnectDB = SourceConn
    Set SourceConn = Nothing
End Function

Public Function GetServerTime(ByVal conn As ADODB.Connection) As Date
    '获取服务器时间
     Dim TabStrRs As New ADODB.Recordset
     TabStrRs.Open "SELECT GETDATE() ", conn, adOpenStatic, adLockOptimistic
     GetServerTime = TabStrRs.Fields(0).Value
     Set TabStrRs = Nothing
End Function
