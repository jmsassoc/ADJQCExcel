Attribute VB_Name = "modDTSSelmanCo"
'****************************************************************
'Microsoft SQL Server 2000
'Visual Basic file generated for DTS Package
'File Name: E:\独立工具11\Adj比例工具\新建包.bas
'Package Name: 新建包
'Package Description: DTS 包描述
'Generated Date: 2011-10-14
'Generated Time: 18:17:16
'****************************************************************

Option Explicit
Public strProcessDate As String
Private goPackageOld As New DTS.Package
Private goPackage As DTS.Package2

Public Sub MainSelmanCo()
        Set goPackage = goPackageOld

        goPackage.Name = "新建包"
        goPackage.Description = "DTS 包描述"
        goPackage.WriteCompletionStatusToNTEventLog = False
        goPackage.FailOnError = False
        goPackage.PackagePriorityClass = 2
        goPackage.MaxConcurrentSteps = 4
        goPackage.LineageOptions = 0
        goPackage.UseTransaction = True
        goPackage.TransactionIsolationLevel = 4096
        goPackage.AutoCommitTransaction = True
        goPackage.RepositoryMetadataOptions = 0
        goPackage.UseOLEDBServiceComponents = True
        goPackage.LogToSQLServer = False
        goPackage.LogServerFlags = 0
        goPackage.FailPackageOnLogFailure = False
        goPackage.ExplicitGlobalVariables = False
        goPackage.PackageType = 0
        

Dim oConnProperty As DTS.OleDBProperty

'---------------------------------------------------------------------------
' create package connection information
'---------------------------------------------------------------------------

Dim oConnection As DTS.Connection2

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("SQLOLEDB")

        oConnection.ConnectionProperties("Persist Security Info") = True
        oConnection.ConnectionProperties("User ID") = strSqlUserId
        oConnection.ConnectionProperties("Initial Catalog") = "DR_Adjudication"
        oConnection.ConnectionProperties("Data Source") = strSqlServer
        oConnection.ConnectionProperties("Application Name") = "DTS 导入/导出向导"
        
        oConnection.Name = "连接1"
        oConnection.ID = 1
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = strSqlServer
        oConnection.UserID = strSqlUserId
        oConnection.Password = strSqlPassword
        oConnection.ConnectionTimeout = 60
        oConnection.Catalog = "DR_Adjudication"
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'------------- a new connection defined below.
'For security purposes, the password is never scripted

Set oConnection = goPackage.Connections.New("Microsoft.Jet.OLEDB.4.0")

        oConnection.ConnectionProperties("Data Source") = strSaveFileName
        oConnection.ConnectionProperties("Extended Properties") = "Excel 8.0;HDR=YES;"
        
        oConnection.Name = "连接2"
        oConnection.ID = 2
        oConnection.Reusable = True
        oConnection.ConnectImmediate = False
        oConnection.DataSource = strSaveFileName
        oConnection.ConnectionTimeout = 60
        oConnection.UseTrustedConnection = False
        oConnection.UseDSL = False
        
        'If you have a password for this connection, please uncomment and add your password below.
        'oConnection.Password = "<put the password here>"

goPackage.Connections.Add oConnection
Set oConnection = Nothing

'---------------------------------------------------------------------------
' create package steps information
'---------------------------------------------------------------------------

Dim oStep As DTS.Step2
Dim oPrecConstraint As DTS.PrecedenceConstraint

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "创建表 结果 步骤"
        oStep.Description = "创建表 结果 步骤"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "创建表 结果 任务"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = False
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a new step defined below

Set oStep = goPackage.Steps.New

        oStep.Name = "Copy Data from 结果 to 结果 步骤"
        oStep.Description = "Copy Data from 结果 to 结果 步骤"
        oStep.ExecutionStatus = 1
        oStep.TaskName = "Copy Data from 结果 to 结果 任务"
        oStep.CommitSuccess = False
        oStep.RollbackFailure = False
        oStep.ScriptLanguage = "VBScript"
        oStep.AddGlobalVariables = True
        oStep.RelativePriority = 3
        oStep.CloseConnection = False
        oStep.ExecuteInMainThread = True
        oStep.IsPackageDSORowset = False
        oStep.JoinTransactionIfPresent = False
        oStep.DisableStep = False
        oStep.FailPackageOnError = False
        
goPackage.Steps.Add oStep
Set oStep = Nothing

'------------- a precedence constraint for steps defined below

Set oStep = goPackage.Steps("Copy Data from 结果 to 结果 步骤")
Set oPrecConstraint = oStep.PrecedenceConstraints.New("创建表 结果 步骤")
        oPrecConstraint.StepName = "创建表 结果 步骤"
        oPrecConstraint.PrecedenceBasis = 0
        oPrecConstraint.Value = 4
        
oStep.PrecedenceConstraints.Add oPrecConstraint
Set oPrecConstraint = Nothing

'---------------------------------------------------------------------------
' create package tasks information
'---------------------------------------------------------------------------

'------------- call Task_Sub1 for task 创建表 结果 任务 (创建表 结果 任务)
Call Task_Sub1(goPackage)

'------------- call Task_Sub2 for task Copy Data from 结果 to 结果 任务 (Copy Data from 结果 to 结果 任务)
Call Task_Sub2(goPackage)

'---------------------------------------------------------------------------
' Save or execute package
'---------------------------------------------------------------------------

'goPackage.SaveToSQLServer "(local)", strSqlUserId, ""
goPackage.Execute
goPackage.UnInitialize
'to save a package instead of executing it, comment out the executing package line above and uncomment the saving package line
Set goPackage = Nothing

Set goPackageOld = Nothing

End Sub


'------------- define Task_Sub1 for task 创建表 结果 任务 (创建表 结果 任务)
Private Sub Task_Sub1(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask1 As DTS.ExecuteSQLTask2
Set oTask = goPackage.Tasks.New("DTSExecuteSQLTask")
Set oCustomTask1 = oTask.CustomTask

        oCustomTask1.Name = "创建表 结果 任务"
        oCustomTask1.Description = "创建表 结果 任务"
        oCustomTask1.SQLStatement = "CREATE TABLE `结果` (" & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`Workdate` VarChar (20) , " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`Project` VarChar (30) , " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`MailBox` VarChar (10) , " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`Folder` VarChar (20) , " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`Groups` VarChar (10) , " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`ClaimID` VarChar (30) , " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`AdjStatus` VarChar (20) , " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`Remark` LongText , " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & "`UserID` VarChar (30) " & vbCrLf
        oCustomTask1.SQLStatement = oCustomTask1.SQLStatement & ")"
        oCustomTask1.ConnectionID = 2
        oCustomTask1.CommandTimeout = 0
        oCustomTask1.OutputAsRecordset = False
        
goPackage.Tasks.Add oTask
Set oCustomTask1 = Nothing
Set oTask = Nothing

End Sub

'------------- define Task_Sub2 for task Copy Data from 结果 to 结果 任务 (Copy Data from 结果 to 结果 任务)
Private Sub Task_Sub2(ByVal goPackage As Object)

Dim oTask As DTS.Task
Dim oLookup As DTS.Lookup

Dim oCustomTask2 As DTS.DataPumpTask2
Set oTask = goPackage.Tasks.New("DTSDataPumpTask")
Set oCustomTask2 = oTask.CustomTask

        oCustomTask2.Name = "Copy Data from 结果 to 结果 任务"
        oCustomTask2.Description = "Copy Data from 结果 to 结果 任务"
        oCustomTask2.SourceConnectionID = 1
        oCustomTask2.SourceSQLStatement = "SELECT     Workdate, Project, MailBox, Folder, Groups, ClaimID, AdjStatus, Remark, UserID" & vbCrLf
        oCustomTask2.SourceSQLStatement = oCustomTask2.SourceSQLStatement & "FROM         (SELECT     TOP " & intTopValue & " ProcessDate AS Workdate, 'SelmanCo' AS Project,UserID as  MailBox, ClaimType as  Folder,'' as  Groups, ClaimNO as  ClaimID,  Status as AdjStatus, NOTE as  Remark, CreateUserID as  UserID" & vbCrLf
        oCustomTask2.SourceSQLStatement = oCustomTask2.SourceSQLStatement & "                       FROM          TblAsiProduction " & vbCrLf
        oCustomTask2.SourceSQLStatement = oCustomTask2.SourceSQLStatement & "                       WHERE      (ProcessDate >= '" & strProcessDate & "') AND (ProcessDate <= '" & strProcessDate & "') AND (Deleted = 0) " & vbCrLf
        oCustomTask2.SourceSQLStatement = oCustomTask2.SourceSQLStatement & "                       ORDER BY GUID) AS derivedtbl_1" & vbCrLf
        oCustomTask2.SourceSQLStatement = oCustomTask2.SourceSQLStatement & "ORDER BY MailBox, Folder, Groups"
        oCustomTask2.DestinationConnectionID = 2
        oCustomTask2.DestinationObjectName = "结果"
        oCustomTask2.ProgressRowCount = 1000
        oCustomTask2.MaximumErrorCount = 0
        oCustomTask2.FetchBufferSize = 1
        oCustomTask2.UseFastLoad = True
        oCustomTask2.InsertCommitSize = 0
        oCustomTask2.ExceptionFileColumnDelimiter = "|"
        oCustomTask2.ExceptionFileRowDelimiter = vbCrLf
        oCustomTask2.AllowIdentityInserts = False
        oCustomTask2.FirstRow = 0
        oCustomTask2.LastRow = 0
        oCustomTask2.FastLoadOptions = 2
        oCustomTask2.ExceptionFileOptions = 1
        oCustomTask2.DataPumpOptions = 0
        
Call oCustomTask2_Trans_Sub1(oCustomTask2)
                
                
goPackage.Tasks.Add oTask
Set oCustomTask2 = Nothing
Set oTask = Nothing

End Sub

Private Sub oCustomTask2_Trans_Sub1(ByVal oCustomTask2 As Object)

        Dim oTransformation As DTS.Transformation2
        Dim oTransProps As DTS.Properties
        Dim oColumn As DTS.Column
        Set oTransformation = oCustomTask2.Transformations.New("DTS.DataPumpTransformCopy")
                oTransformation.Name = "DirectCopyXform"
                oTransformation.TransformFlags = 63
                oTransformation.ForceSourceBlobsBuffered = 0
                oTransformation.ForceBlobsInMemory = False
                oTransformation.InMemoryBlobSize = 1048576
                oTransformation.TransformPhases = 4
                
                Set oColumn = oTransformation.SourceColumns.New("Workdate", 1)
                        oColumn.Name = "Workdate"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 104
                        oColumn.Size = 20
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Project", 2)
                        oColumn.Name = "Project"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 104
                        oColumn.Size = 30
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("MailBox", 3)
                        oColumn.Name = "MailBox"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 104
                        oColumn.Size = 10
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Folder", 4)
                        oColumn.Name = "Folder"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 104
                        oColumn.Size = 20
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Groups", 5)
                        oColumn.Name = "Groups"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 104
                        oColumn.Size = 10
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("ClaimID", 6)
                        oColumn.Name = "ClaimID"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 104
                        oColumn.Size = 30
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("AdjStatus", 7)
                        oColumn.Name = "AdjStatus"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 104
                        oColumn.Size = 20
                        oColumn.DataType = 129
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("Remark", 8)
                        oColumn.Name = "Remark"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 104
                        oColumn.Size = 500
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.SourceColumns.New("UserID", 9)
                        oColumn.Name = "UserID"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 104
                        oColumn.Size = 30
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.SourceColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Workdate", 1)
                        oColumn.Name = "Workdate"
                        oColumn.Ordinal = 1
                        oColumn.Flags = 104
                        oColumn.Size = 20
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Project", 2)
                        oColumn.Name = "Project"
                        oColumn.Ordinal = 2
                        oColumn.Flags = 104
                        oColumn.Size = 30
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("MailBox", 3)
                        oColumn.Name = "MailBox"
                        oColumn.Ordinal = 3
                        oColumn.Flags = 104
                        oColumn.Size = 10
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Folder", 4)
                        oColumn.Name = "Folder"
                        oColumn.Ordinal = 4
                        oColumn.Flags = 104
                        oColumn.Size = 20
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Groups", 5)
                        oColumn.Name = "Groups"
                        oColumn.Ordinal = 5
                        oColumn.Flags = 104
                        oColumn.Size = 10
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("ClaimID", 6)
                        oColumn.Name = "ClaimID"
                        oColumn.Ordinal = 6
                        oColumn.Flags = 104
                        oColumn.Size = 30
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("AdjStatus", 7)
                        oColumn.Name = "AdjStatus"
                        oColumn.Ordinal = 7
                        oColumn.Flags = 104
                        oColumn.Size = 20
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("Remark", 8)
                        oColumn.Name = "Remark"
                        oColumn.Ordinal = 8
                        oColumn.Flags = 104
                        oColumn.Size = 0
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

                Set oColumn = oTransformation.DestinationColumns.New("UserID", 9)
                        oColumn.Name = "UserID"
                        oColumn.Ordinal = 9
                        oColumn.Flags = 104
                        oColumn.Size = 30
                        oColumn.DataType = 130
                        oColumn.Precision = 0
                        oColumn.NumericScale = 0
                        oColumn.Nullable = True
                        
                oTransformation.DestinationColumns.Add oColumn
                Set oColumn = Nothing

        Set oTransProps = oTransformation.TransformServerProperties

                
        Set oTransProps = Nothing

        oCustomTask2.Transformations.Add oTransformation
        Set oTransformation = Nothing

End Sub



