VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm880011 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "印表機設定"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   3780
      TabIndex        =   6
      Top             =   1980
      Width           =   1095
   End
   Begin VB.ComboBox cboPrinters 
      Height          =   300
      Left            =   1710
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   1620
      Width           =   4305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4905
      TabIndex        =   0
      Top             =   1980
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   135
      Top             =   1980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblDefaultPrinter 
      Height          =   285
      Left            =   1755
      TabIndex        =   5
      Top             =   1290
      Width           =   4245
   End
   Begin VB.Label Label3 
      Caption         =   "控制台預設印表機"
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   1290
      Width           =   1545
   End
   Begin VB.Label lblMemo 
      Caption         =   "本設定將會更改控制台的預設印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1080
      Left            =   135
      TabIndex        =   3
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "系統使用印表機"
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   1620
      Width           =   1545
   End
End
Attribute VB_Name = "frm880011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/15 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
'Create by Morgan 2006/10/13
Option Explicit

Public bolAppOnly As Boolean '設定系統預設印表機
Dim iDefault As Integer

'Removed by Morgan 2017/12/13 沒用了
'Public Function GetPrinterIndex() As Integer
'   GetPrinterIndex = iDefault
'End Function

Private Sub cmdCancel_Click()
   Unload Me
End Sub
'Add by Morgan 2010/2/3
Private Sub SavePrinter()
   Dim stSQL As String, iR As Integer
   
   stSQL = "update PrintStartPoint set PSP06='" & Printer.DeviceName & "'" & _
      " where PSP01='" & pub_HostName & "' and PSP02='" & App.EXEName & "' and PSP03='APP'"
   
   cnnConnection.Execute stSQL, iR
   If iR = 0 Then
      stSQL = "insert into PrintStartPoint(PSP01,PSP02,PSP03,PSP06)" & _
         " values('" & pub_HostName & "','" & App.EXEName & "','APP','" & Printer.DeviceName & "')"
      cnnConnection.Execute stSQL, iR
   End If
End Sub

Private Sub cmdOK_Click()
   'Modify by Morgan 2010/2/3
   'If cboPrinters.ListIndex <> iDefault Then
   '   Printer.TrackDefault = False
   '   CreateObject("WScript.Network").SetDefaultPrinter Printers(cboPrinters.ListIndex).DeviceName
   'End If
   'PUB_SetWordActivePrinter
   If bolAppOnly Then
      'Modified by Morgan 2017/12/13
      'Set Printer = Printers(cboPrinters.ListIndex)
      PUB_RestorePrinter cboPrinters
      'end 2017/12/13
      Printer.TrackDefault = False
      SavePrinter
   Else
      'Modified by Morgan 2017/12/13
      'PUB_SetOsDefaultPrinter Printers(cboPrinters.ListIndex).DeviceName
      PUB_SetOsDefaultPrinter cboPrinters
      'end 2017/12/13
      PUB_SetWordActivePrinter
   End If
   'end 2010/2/3
   Unload Me
End Sub

Private Sub Form_Activate()
   'Modify by Morgan 2010/2/3
   'lblDefaultPrinter = Printer.DeviceName
   lblDefaultPrinter = PUB_GetOsDefaultPrinter
   If bolAppOnly Then
      lblMemo = "本功能將不會更改控制台預設印表機，僅設定系統使用中印表機！"
      cmdCancel.Visible = True
   Else
      lblMemo = "本設定將會暫時更改控制台的預設印表機直到列印完成後才會恢復為原來預設；" & vbCrLf & vbCrLf & "為確保列印順利，請暫停所有其他操作!!"
      cmdCancel.Visible = False
   End If
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   MoveFormToCenter Me
   
   'Modified by Morgan 2017/12/13 Win7的Printers會抓到已經不存在的印表機
   If Val(pub_OS_Ver) > 6 Then
      CollectPrinters
   Else
      For i = 0 To Printers.Count - 1
         cboPrinters.AddItem Printers(i).DeviceName, i
         If Printers(i).DeviceName = Printer.DeviceName Then iDefault = i
      Next i
   End If
   'end 2017/12/13
   cboPrinters.ListIndex = iDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm880011 = Nothing
End Sub
'Added by Morgan 2017/12/13
Private Sub CollectPrinters()

    Dim strComputer As String
    Dim objWMIService As Object
    Dim colInstalledPrinters
    Dim objPrinter
    Dim idx As Integer
    
    strComputer = "."
    'Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    'Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer")
    'Pub_SetPrinter 暫時不可使用此法，除非確認系統都不再使用 printers 的索引值操作。Ex:還原印表機 Set Printer = Printers(SeekPrint)
    Set colInstalledPrinters = Interaction.GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2").ExecQuery("Select * from Win32_Printer")
    
    cboPrinters.Clear
    idx = -1
    For Each objPrinter In colInstalledPrinters
      idx = idx + 1
      cboPrinters.AddItem objPrinter.Name, idx
      If objPrinter.Name = Printer.DeviceName Then iDefault = idx
    Next

End Sub

