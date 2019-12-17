Attribute VB_Name = "Global"
Option Explicit
Public arrExcelConfig() As String
Public connDsgn As ADODB.Connection
Public iParentApp As DatariverAddin.IApplication

Public rsProjectLoadDataToDB As ADODB.Recordset

Public intExcelConfigTotal As Integer


Public strImportRuleName As String
Public strImportFile As String
Public intStartRow As Integer

Public strSqlUserId As String, strSqlServer As String, strSqlPassword As String
Public strDataBaseName As String, strTableName As String
Public strListSeparator As String
Public blnTableIsExist As Boolean
Public dbeQCRatio As Double
Public strSaveFolder As String

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Function ExportExcelFile(ByVal strStartDate As String, ByVal strEndDate As String, ByVal intCheckFileProject As Integer) As Boolean
  Dim dateStart As Date
  Dim db As ADODB.Connection
  Dim strCsv As String
  Dim strSql As String
  Dim rsCount As ADODB.Recordset
  Dim rsProject As ADODB.Recordset
  Dim strProject As String
  strCsv = Chr(34) & "Workdate" & Chr(34) & "," & Chr(34) & "Count" & Chr(34) & "," & Chr(34) & "Ad3" & Chr(34) & vbCrLf
  Set db = New ADODB.Connection
  Set db = ConnectDB(iParentApp.DBConnection("DR_Adjudication").ConnectionString, 200, 200)
    '
  For dateStart = strStartDate To strEndDate
            strExportDate = Format(dateStart, "yyyy-mm-dd")
         strProcessDate = Format(dateStart, "yyyymmdd")
          strProjecWhere = ""
          If intCheckFileProject = 1 Then
            strSql = "select projectName From  TblClaimManage  WHERE  (CreateDate >= '" & strExportDate & "') AND (CreateDate <= '" & strExportDate & "') AND (Deleted = 0) Group by projectName  "
            Set rsProject = GetReadOnlyRsBySql(strSql, db)
            Set rsProject.ActiveConnection = Nothing
            
            Do While Not rsProject.EOF
                strProject = Trim(rsProject.Fields("projectName").Value & "")
                strProjecWhere = " And  projectName='" & Replace(strProject, "'", "''") & "'"
                strSql = "select Count(0) From  TblClaimManage  WHERE  (CreateDate >= '" & strExportDate & "') AND (CreateDate <= '" & strExportDate & "') AND (Deleted = 0)  " & strProjecWhere
                Set rsCount = GetReadOnlyRsBySql(strSql, db)
                strSaveFileName = strSaveFolder & "\Adj-" & strProject & Format(dateStart, "yyymmdd")
                intTopValue = CInt(rsCount.Fields(0).Value * dbeQCRatio)   ' 72
                strCsv = strCsv & Chr(34) & strExportDate & Chr(34) & "," & Chr(34) & rsCount.Fields(0).Value & Chr(34) & _
                "," & Chr(34) & intTopValue & Chr(34) & vbCrLf
                Call Main
                rsCount.Close
                rsProject.MoveNext
            Loop
            rsProject.Close
        Else
            strSql = "select Count(0) From  TblClaimManage  WHERE  (CreateDate >= '" & strExportDate & "') AND (CreateDate <= '" & strExportDate & "') AND (Deleted = 0) "
            Set rsCount = GetReadOnlyRsBySql(strSql, db)
            strSaveFileName = strSaveFolder & "\Adj" & Format(dateStart, "yyymmdd")
            intTopValue = CInt(rsCount.Fields(0).Value * dbeQCRatio)   ' 72
            strCsv = strCsv & Chr(34) & strExportDate & Chr(34) & "," & Chr(34) & rsCount.Fields(0).Value & Chr(34) & _
            "," & Chr(34) & intTopValue & Chr(34) & vbCrLf
            Call Main
            rsCount.Close
         End If
    
        '2015.09.03
        strSaveFileName = strSaveFolder & "\Adj-SelmanCo" & Format(dateStart, "yyymmdd")
        strSql = "select Count(0) From  TblAsiProduction  WHERE  (ProcessDate >= '" & strProcessDate & "') AND (ProcessDate <= '" & strProcessDate & "') AND (Deleted = 0) "
        Set rsCount = GetReadOnlyRsBySql(strSql, db)
        intTopValue = CInt(rsCount.Fields(0).Value * dbeQCRatio)   ' 72
        strCsv = strCsv & Chr(34) & strProcessDate & Chr(34) & "," & Chr(34) & rsCount.Fields(0).Value & Chr(34) & _
        "," & Chr(34) & intTopValue & Chr(34) & vbCrLf
        rsCount.Close
        Call MainSelmanCo
   Next
   Set db = Nothing
   Call subSaveFile(strCsv, strSaveFolder & "\Total" & Format(strStartDate, "yyyy-mm-dd") & "TO" & Format(strEndDate, "yyyy-mm-dd") & ".csv")
    
End Function
Public Sub subSaveFile(ByVal Message As String, strFileName As String)
        Dim filenum2 As Integer
        On Error Resume Next
        filenum2 = FreeFile()
        Open strFileName For Output As filenum2
          Print #filenum2, Message
        Close filenum2
End Sub

Public Sub initGetSqlInfo()
    Dim strConnSqlString As String
    Dim i As Integer, k As Integer
    Dim j As Integer
    strConnSqlString = iParentApp.DBConnection.ConnectionString
    j = Len(";SERVER=")
    k = InStr(UCase(strConnSqlString), ";SERVER=")
    i = InStr(k + j, strConnSqlString, ";")
    strSqlServer = Mid(strConnSqlString, k + j, i - (k + j))
    
    j = Len(";UID=")
    k = InStr(UCase(strConnSqlString), ";UID=")
    i = InStr(k + j, strConnSqlString, ";")
    strSqlUserId = Mid(strConnSqlString, k + j, i - (k + j))
    
    j = Len(";PWD=")
    k = InStr(UCase(strConnSqlString), ";PWD=")
    i = InStr(k + j, strConnSqlString, ";")
    strSqlPassword = Mid(strConnSqlString, k + j, i - (k + j))
    
End Sub

Public Sub getExcelConfigInfo(ByVal strProjectNameSub As String)
   '获取项目的定制信息
    Dim strProjectNameFrm As String, strSqlTxt As String
    Dim k As Integer, i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    strSqlTxt = "SELECT  FileType, FileTypeAlias, RuleName, ListSeparator, ColName, StartLength,EndLength, TableField, FieldLength,RunScriptFlag,FieldScript From TblCurrencyImportDataRule where  RuleName='" & strProjectNameSub & "' order by StartLength ,ColName "
    Set rs = GetRsBySql(strSqlTxt, connDsgn)
    Set rs.ActiveConnection = Nothing
    k = rs.RecordCount
    intExcelConfigTotal = k
    ReDim arrExcelConfigTmp(k, 10) As String
    i = 1
    Do While Not rs.EOF
        arrExcelConfigTmp(i, 1) = UCase(rs.Fields("FileTypeAlias").Value & "")
        arrExcelConfigTmp(i, 2) = Trim(rs.Fields("ListSeparator").Value & "")
        arrExcelConfigTmp(i, 3) = Trim(rs.Fields("ColName").Value & "")
        arrExcelConfigTmp(i, 4) = Trim(rs.Fields("StartLength").Value & "")
        arrExcelConfigTmp(i, 5) = Trim(rs.Fields("EndLength").Value & "")
        arrExcelConfigTmp(i, 6) = Trim(rs.Fields("TableField").Value & "")
        arrExcelConfigTmp(i, 7) = Trim(rs.Fields("FieldLength").Value & "")
        arrExcelConfigTmp(i, 8) = Trim(rs.Fields("RunScriptFlag").Value & "")
        arrExcelConfigTmp(i, 9) = Trim(rs.Fields("FieldScript").Value & "")
        i = i + 1
        rs.MoveNext
    Loop
    arrExcelConfig = arrExcelConfigTmp
    rs.Close
    Set rs = Nothing
End Sub

'获取当前项目的导库信息
'TblProjectLoadDataToDB
Public Sub GetProjectLoadDataToDB(ByVal strProject As String, strExportNamed As String)
    Dim strSql As String
    strSql = "Select *,newid() as TempTableName From TblProjectLoadDataToDB  " _
        & " Where Project='" & strProject & "' And ExportNamed='" & strExportNamed & "'"
    Set rsProjectLoadDataToDB = GetRsBySql(strSql, connDsgn)
    Set rsProjectLoadDataToDB.ActiveConnection = Nothing
End Sub


Public Function getSwitchListSeparator(ByVal strValue As String) As String
    Dim strTmp As String
    Select Case UCase(strValue)
        Case UCase("vbTab"): getSwitchListSeparator = Chr(9)
        Case UCase("vbCr"): getSwitchListSeparator = Chr(13)
        Case UCase("vbCrLf"): getSwitchListSeparator = Chr(13) & Chr(10)
        Case UCase("vbLf"): getSwitchListSeparator = Chr(10)
        Case UCase("vbNullChar"): getSwitchListSeparator = Chr(0)
        Case UCase("vbVerticalTab"): getSwitchListSeparator = Chr(11)
        Case Else
            getSwitchListSeparator = strValue
    End Select
End Function
Public Sub OrderArrExcelConfigStartAnEnd()
    '对arrExcelConfig数据排序1
    Dim i As Integer, c As Integer
    Dim j As Integer, l As Integer
    Dim intTmpValue As Integer
    Dim arrTmp(1, 8) As String
    c = UBound(arrExcelConfig)
    For i = 1 To c - 1
        intTmpValue = CInt(arrExcelConfig(i, 4))
        For j = i + 1 To c
            If intTmpValue > CInt(arrExcelConfig(j, 4)) Then
                For l = 1 To 7
                    arrTmp(1, l) = arrExcelConfig(i, l)
                Next
                For l = 1 To 7
                    arrExcelConfig(i, l) = arrExcelConfig(j, l)
                Next
                For l = 1 To 7
                    arrExcelConfig(j, l) = arrTmp(1, l)
                Next
                intTmpValue = CInt(arrExcelConfig(i, 4))
            End If
        Next
    Next
End Sub
Public Function FlagDataBaseTableExist(ByVal strDataBase As String, ByVal strTableName As String) As Boolean
    '数据库表是否存在
    Dim strSqlTxt As String
    Dim blnTmp As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    strSqlTxt = "SELECT  name From sysobjects  where name='" & strTableName & "'"
    Set rs = GetRsBySql(strSqlTxt, ConnectDB(iParentApp.DBConnection(strDataBase).ConnectionString))
    If rs.RecordCount >= 1 Then blnTmp = True
    rs.Close
    Set rs = Nothing
    FlagDataBaseTableExist = blnTmp
End Function


   '创建指定的目录
 Public Function CreateForders(ByVal strPath As String) As Long
    Dim fso As FileSystemObject
    Dim i As Integer
    Set fso = New FileSystemObject
    If Not fso.FolderExists(strPath) Then
        For i = 1 To Len(strPath)
            If Mid$(strPath, i, 1) = "\" Or i = Len(strPath) Then
                If Not fso.FolderExists(Left$(strPath, i)) Then
                    
                    fso.CreateFolder Left$(strPath, i)
                End If
            End If
        Next i
    End If
    If fso.FolderExists(strPath) Then
        CreateForders = 1
    End If
End Function
