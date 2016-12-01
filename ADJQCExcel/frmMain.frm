VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QC"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   6165
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox CheckFileProject 
      Caption         =   "每天每项目"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   900
      Width           =   1335
   End
   Begin VB.CommandButton cmdQc 
      Caption         =   "QC "
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返 回 "
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPickerF 
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   49676289
      CurrentDate     =   40833
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox TxtSaveFolder 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "C:\ADJReport\QcExcel"
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox QcValue 
      Height          =   285
      Left            =   5280
      TabIndex        =   1
      Text            =   "0.03"
      Top             =   480
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPickerT 
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   49676289
      CurrentDate     =   40833
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Workdate："
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "保存文件夹："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "比率："
      Height          =   255
      Left            =   4680
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQc_Click()
    On Error GoTo cmdQc_ClickErr
    If DTPickerT.Value < DTPickerF.Value Then
        MsgBox "ToWorkdate不能比FromWorkdate小", , ""
        Exit Sub
    End If
    Me.MousePointer = 13
    cmdQc.Enabled = False
    dbeQCRatio = QcValue.Text
    strSaveFolder = TxtSaveFolder.Text ' "C:\ADJReport\QcExcel"
    Call CreateForders(strSaveFolder)
    Call ExportExcelFile(Format(DTPickerF.Value, "yyyy-mm-dd"), Format(DTPickerT.Value, "yyyy-mm-dd"), CheckFileProject.Value)
    Me.MousePointer = 1
    cmdQc.Enabled = True
    ShellExecute frmMain.hWnd, vbNullString, strSaveFolder, vbNullString, "C:\", SW_SHOWNORMAL
    Exit Sub
cmdQc_ClickErr:
    Me.MousePointer = 1
    cmdQc.Enabled = True
    MsgBox Err.Description, , ""
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub




Private Sub Form_Load()
    DTPickerF.Value = Format(Now - 1, "yyyy-mm-dd")
    DTPickerT.Value = Format(Now, "yyyy-mm-dd")
End Sub
