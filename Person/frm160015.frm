VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160015 
   BorderStyle     =   1  '單線固定
   Caption         =   "考勤機設定"
   ClientHeight    =   4644
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   6360
   Begin TabDlg.SSTab SSTab1 
      Height          =   2745
      Left            =   120
      TabIndex        =   4
      Top             =   450
      Width           =   6045
      _ExtentX        =   10668
      _ExtentY        =   4847
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本"
      TabPicture(0)   =   "frm160015.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblDateTime"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(40)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cboBranch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboHtaIp"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command1(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command1(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "時區表"
      TabPicture(1)   =   "frm160015.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command2(5)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboHtaIp2(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MSHFlexGrid1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command2(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command2(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command2(2)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command2(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command2(4)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtInput"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command2(6)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Text1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "時段表"
      TabPicture(2)   =   "frm160015.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(1)"
      Tab(2).Control(1)=   "MSHFlexGrid2"
      Tab(2).Control(2)=   "Command3(4)"
      Tab(2).Control(3)=   "Command3(3)"
      Tab(2).Control(4)=   "Command3(2)"
      Tab(2).Control(5)=   "Command3(1)"
      Tab(2).Control(6)=   "Command3(0)"
      Tab(2).Control(7)=   "txtInput2"
      Tab(2).Control(8)=   "cboHtaIp2(1)"
      Tab(2).Control(9)=   "Command3(5)"
      Tab(2).Control(10)=   "Text2"
      Tab(2).Control(11)=   "Command3(6)"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "假日表"
      TabPicture(3)   =   "frm160015.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command4(0)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command4(1)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "List2(1)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "List2(0)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cboHtaIp2(2)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Command4(2)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "lblAlert"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label6(1)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label6(0)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label5"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label4"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label1(2)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).ControlCount=   12
      Begin VB.CommandButton Command1 
         Caption         =   "考勤機校時"
         Height          =   345
         Index           =   0
         Left            =   -71112
         TabIndex        =   52
         Top             =   1128
         Width           =   1185
      End
      Begin VB.TextBox Text3 
         Height          =   276
         Left            =   -73320
         TabIndex        =   50
         Top             =   1560
         Width           =   2148
      End
      Begin VB.CommandButton Command1 
         Caption         =   "考勤機初始化"
         Height          =   345
         Index           =   4
         Left            =   -70536
         TabIndex        =   49
         Top             =   2304
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "讀取"
         Height          =   345
         Index           =   6
         Left            =   -69780
         TabIndex        =   48
         Top             =   2310
         Width           =   685
      End
      Begin VB.TextBox Text2 
         Height          =   1575
         Left            =   -70560
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   47
         Top             =   690
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   4440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   46
         Top             =   690
         Width           =   1485
      End
      Begin VB.CommandButton Command2 
         Caption         =   "讀取"
         Height          =   345
         Index           =   6
         Left            =   5220
         TabIndex        =   45
         Top             =   2310
         Width           =   685
      End
      Begin VB.CommandButton Command1 
         Caption         =   "下載刷卡紀錄"
         Height          =   345
         Index           =   3
         Left            =   -74856
         TabIndex        =   44
         Top             =   2304
         Width           =   1545
      End
      Begin VB.CommandButton Command4 
         Caption         =   "重整"
         Height          =   345
         Index           =   0
         Left            =   -73440
         TabIndex        =   41
         Top             =   360
         Width           =   915
      End
      Begin VB.CommandButton Command4 
         Caption         =   "讀取"
         Height          =   345
         Index           =   1
         Left            =   -70050
         TabIndex        =   40
         Top             =   360
         Width           =   915
      End
      Begin VB.ListBox List2 
         Height          =   1068
         Index           =   1
         Left            =   -71400
         Style           =   1  '項目包含核取方塊
         TabIndex        =   37
         Top             =   720
         Width           =   2265
      End
      Begin VB.ListBox List2 
         Height          =   1068
         Index           =   0
         Left            =   -74760
         Style           =   1  '項目包含核取方塊
         TabIndex        =   36
         Top             =   720
         Width           =   2265
      End
      Begin VB.ComboBox cboHtaIp2 
         Height          =   276
         Index           =   2
         Left            =   -74040
         Style           =   2  '單純下拉式
         TabIndex        =   34
         Top             =   2340
         Width           =   3945
      End
      Begin VB.CommandButton Command4 
         Caption         =   "回寫"
         Height          =   345
         Index           =   2
         Left            =   -69990
         TabIndex        =   33
         Top             =   2310
         Width           =   915
      End
      Begin VB.CommandButton Command3 
         Caption         =   "回寫"
         Height          =   345
         Index           =   5
         Left            =   -70500
         TabIndex        =   31
         Top             =   2310
         Width           =   685
      End
      Begin VB.ComboBox cboHtaIp2 
         Height          =   276
         Index           =   1
         Left            =   -74040
         Style           =   2  '單純下拉式
         TabIndex        =   30
         Top             =   2340
         Width           =   3500
      End
      Begin VB.TextBox txtInput2 
         Appearance      =   0  '平面
         Height          =   375
         Left            =   -73380
         TabIndex        =   29
         Top             =   1650
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton Command3 
         Caption         =   "新增"
         Height          =   285
         Index           =   0
         Left            =   -74880
         TabIndex        =   28
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "修改"
         Height          =   285
         Index           =   1
         Left            =   -74130
         TabIndex        =   27
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "刪除"
         Height          =   285
         Index           =   2
         Left            =   -73380
         TabIndex        =   26
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "存檔"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -72630
         TabIndex        =   25
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "取消"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -71880
         TabIndex        =   24
         Top             =   390
         Width           =   735
      End
      Begin VB.TextBox txtInput 
         Appearance      =   0  '平面
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   1770
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取消"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3120
         TabIndex        =   21
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "存檔"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2370
         TabIndex        =   20
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "刪除"
         Height          =   285
         Index           =   2
         Left            =   1620
         TabIndex        =   19
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "修改"
         Height          =   285
         Index           =   1
         Left            =   870
         TabIndex        =   18
         Top             =   390
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "新增"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   390
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1605
         Left            =   120
         TabIndex        =   16
         Top             =   690
         Width           =   4300
         _ExtentX        =   7599
         _ExtentY        =   2836
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         AllowBigSelection=   0   'False
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.ComboBox cboHtaIp2 
         Height          =   276
         Index           =   0
         Left            =   960
         Style           =   2  '單純下拉式
         TabIndex        =   14
         Top             =   2340
         Width           =   3500
      End
      Begin VB.CommandButton Command2 
         Caption         =   "回寫"
         Height          =   345
         Index           =   5
         Left            =   4500
         TabIndex        =   13
         Top             =   2310
         Width           =   685
      End
      Begin VB.CommandButton Command1 
         Caption         =   "所有在職員工卡號回寫考勤機"
         Height          =   345
         Index           =   1
         Left            =   -74856
         TabIndex        =   9
         Top             =   1896
         Width           =   2715
      End
      Begin VB.CommandButton Command1 
         Caption         =   "時間回寫考勤機"
         Height          =   345
         Index           =   5
         Left            =   -74832
         TabIndex        =   8
         Top             =   1524
         Width           =   1452
      End
      Begin VB.ComboBox cboHtaIp 
         Height          =   276
         Left            =   -73845
         TabIndex        =   7
         Text            =   "cboHtaIp"
         Top             =   750
         Width           =   4455
      End
      Begin VB.ComboBox cboBranch 
         Height          =   276
         Left            =   -71370
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "讀取考勤機時間"
         Height          =   345
         Index           =   2
         Left            =   -74844
         TabIndex        =   5
         Top             =   1110
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1605
         Left            =   -74880
         TabIndex        =   23
         Top             =   690
         Width           =   4300
         _ExtentX        =   7599
         _ExtentY        =   2836
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         AllowBigSelection=   0   'False
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.Label lblAlert 
         AutoSize        =   -1  'True
         Caption         =   "新機要用HAMS的管理系統設定(清除)一次假日表才會有作用!!!"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   204
         Left            =   -74784
         TabIndex        =   53
         Top             =   1824
         Width           =   5556
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "(YYYY/MM/DD HH:MM:SS)"
         Height          =   180
         Left            =   -71160
         TabIndex        =   51
         Top             =   1608
         Width           =   2136
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "共 0 筆"
         Height          =   180
         Index           =   1
         Left            =   -71400
         TabIndex        =   43
         Top             =   2070
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "共 0 筆"
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   42
         Top             =   2070
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "門禁機："
         Height          =   180
         Left            =   -71400
         TabIndex        =   39
         Top             =   450
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "系統："
         Height          =   180
         Left            =   -74730
         TabIndex        =   38
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "門禁機："
         Height          =   180
         Index           =   2
         Left            =   -74850
         TabIndex        =   35
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "門禁機："
         Height          =   180
         Index           =   1
         Left            =   -74850
         TabIndex        =   32
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "門禁機："
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "所別："
         Height          =   180
         Left            =   -71952
         TabIndex        =   12
         Top             =   1980
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "考勤機："
         Height          =   180
         Index           =   40
         Left            =   -74745
         TabIndex        =   11
         Top             =   810
         Width           =   720
      End
      Begin VB.Label lblDateTime 
         AutoSize        =   -1  'True
         Caption         =   "YYYY/MM/DD HH:MI:SS"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   -73296
         TabIndex        =   10
         Top             =   1188
         Width           =   2076
      End
   End
   Begin VB.ListBox List1 
      Height          =   588
      Left            =   120
      TabIndex        =   1
      Top             =   3810
      Width           =   6000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   345
      Left            =   5190
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   3210
      Width           =   5970
      _ExtentX        =   10520
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0/0 )"
      Height          =   255
      Left            =   150
      TabIndex        =   3
      Top             =   3510
      Width           =   5970
   End
End
Attribute VB_Name = "frm160015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/27 Form2.0不用改
'Created by Morgan 2013/7/18
Option Explicit

Const clColorSel As Long = &HFFC0C0
Dim iLstRow1 As Integer '前次點選列數1
Dim iLstRow2 As Integer '前次點選列數2
Dim iRow As Integer, iCol As Integer '本次點選列數,行數

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   If cboHtaIp = "" Then
      MsgBox "請點選考勤機IP！", vbCritical
   Else
      Select Case Index
      Case 0 '考勤機校時
         If UpdateTime() = True Then
            ReadTime
            MsgBox "校時成功！", vbInformation
         End If
         
      Case 1 '所有在職員工卡號回寫考勤機
         BatchWrite
      
      'Added by Morgan 2016/2/23
      Case 2 '讀取考勤機時間
         ReadTime
         
      'Added by Morgan 2020/11/10
      Case 3
         PollingData
      
      'Added by Morgan 2024/3/4
      Case 4
         InitDevice
      
      'Added by Morgan 2024/9/30
      Case 5 '時間回寫考勤機
         If Text3 = "" Then
            MsgBox "請依格式輸入要回寫的時間!!", vbExclamation
            Text3.SetFocus
         ElseIf UpdateTime2(Text3) = True Then
            ReadTime
            MsgBox "回寫時間成功！", vbInformation
         End If
         'end 2024/9/30
      End Select
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click(Index As Integer)
   
   Select Case Index
      Case 0 '新增
         GridAddRow
      Case 1 '修改
         If MSHFlexGrid1.TextMatrix(1, 0) = "" Then
            GridAddRow
         End If
      Case 2 '刪除
         GridDelRow
      Case 3 '存檔
         If SaveGrid1 = False Then
            Exit Sub
         End If
      Case 4 '取消
         LoadGrid1
      Case 5 '回寫
         If cboHtaIp2(0) = "" Then
            MsgBox "請點選門禁機！", vbCritical
            Exit Sub
         ElseIf WritTimeSheet() = False Then
            Exit Sub
         End If
         
      'Added by Morgan 2020/11/16
      Case 6 '讀取
         If cboHtaIp2(0) = "" Then
            MsgBox "請點選門禁機！", vbCritical
            Exit Sub
         End If
         ReadTimeSheet
   End Select
   CmdEnable Index
End Sub

Private Sub CmdEnable(pIdx As Integer)
   Select Case pIdx
   Case 0, 1, 2
      Command2(1).Enabled = False
      Command2(3).Enabled = True
      Command2(4).Enabled = True
      Command2(5).Enabled = False
      TabEnable False, SSTab1.Tab
      
   Case 3, 4
      Command2(1).Enabled = True
      Command2(3).Enabled = False
      Command2(4).Enabled = False
      Command2(5).Enabled = True
      txtInput.Visible = False
      TabEnable True
   End Select
End Sub

Private Sub CmdEnable2(pIdx As Integer)
   Select Case pIdx
   Case 0, 1, 2
      Command3(1).Enabled = False
      Command3(3).Enabled = True
      Command3(4).Enabled = True
      Command3(5).Enabled = False
      TabEnable False, SSTab1.Tab
      
   Case 3, 4
      Command3(1).Enabled = True
      Command3(3).Enabled = False
      Command3(4).Enabled = False
      Command3(5).Enabled = True
      txtInput2.Visible = False
      TabEnable True
   End Select
End Sub

Private Sub TabEnable(pEnable As Boolean, Optional pActTab As Integer)
   Dim ii As Integer
   If pEnable Then
      SSTab1.TabEnabled(0) = True
      SSTab1.TabEnabled(1) = True
      SSTab1.TabEnabled(2) = True
      SSTab1.TabEnabled(3) = True
   Else
      For ii = 0 To SSTab1.Tabs - 1
         If ii <> pActTab Then
            SSTab1.TabEnabled(ii) = False
         End If
      Next
   End If
End Sub

Private Sub GridAddRow()
   Dim iNo As Integer
   Dim ii As Integer
   Dim oGrid As MSHFlexGrid
   
   If SSTab1.Tab = 1 Then
      Set oGrid = MSHFlexGrid1
   Else
      Set oGrid = MSHFlexGrid2
   End If
   
   With oGrid
   If .TextMatrix(1, 0) = "" Then
      iNo = 0
   Else
      iNo = Val(.TextMatrix(.Rows - 1, 0)) + 1
      'ID 開放至 9 (10筆)，若要增加則
      If iNo > 9 Then
         MsgBox "目前只 ID 開放至 9 (10筆) ！" & vbCrLf & "若要增加請通知電腦中心！", vbExclamation
         Exit Sub
      End If
      .Rows = .Rows + 1
   End If
   .row = .Rows - 1
   .TextMatrix(.row, 0) = iNo
   If SSTab1.Tab = 1 Then
      .TextMatrix(.row, 1) = 0
      .TextMatrix(.row, 2) = 0
      SetGridColor oGrid, iLstRow1
      iLstRow1 = .row
   Else
      SetGridColor oGrid, iLstRow2
      iLstRow2 = .row
   End If
   .TopRow = .row
   .Refresh
   End With
   
End Sub

Private Sub GridDelRow()
   Dim iRow As Integer
   Dim oGrid As MSHFlexGrid
   
   If SSTab1.Tab = 1 Then
      Set oGrid = MSHFlexGrid1
   Else
      Set oGrid = MSHFlexGrid2
   End If
   With oGrid
   If .row = 0 Then
      MsgBox "請點選要刪除的資料！", vbExclamation
   ElseIf .Rows = 2 Then
      MsgBox "最後一筆資料不可刪除！", vbExclamation
   Else
      .RemoveItem .row
      .row = 0
      If SSTab1.Tab = 1 Then
         iLstRow1 = 0
      Else
         iLstRow2 = 0
      End If
      .Refresh
   End If
   End With
End Sub

Private Sub SetGridColor(pGrid As MSHFlexGrid, pLstRow As Integer)
   Dim ii As Integer
   Dim lColor As Long
   Dim iRow As Integer
   
   With pGrid
   If pLstRow <> .row Then
      iRow = .row
      For ii = 0 To .Cols - 1
         .col = ii
         .CellBackColor = clColorSel
      Next
      If pLstRow > 0 Then
         .row = pLstRow
         For ii = 0 To .Cols - 1
            .col = ii
            .CellBackColor = .BackColor
         Next
      End If
      .row = iRow
   End If
   .Refresh
   End With
End Sub

Private Sub Command3_Click(Index As Integer)
   Select Case Index
      Case 0 '新增
         GridAddRow
      Case 1 '修改
         If MSHFlexGrid2.TextMatrix(1, 0) = "" Then
            GridAddRow
         End If
      Case 2 '刪除
         GridDelRow
      Case 3 '存檔
         If SaveGrid2 = False Then
            Exit Sub
         End If
      Case 4 '取消
         LoadGrid2
      Case 5 '回寫
         If cboHtaIp2(1) = "" Then
            MsgBox "請點選門禁機！", vbCritical
            Exit Sub
         ElseIf WritTimeZone() = False Then
            Exit Sub
         End If
         
      Case 6 '讀取
         If cboHtaIp2(1) = "" Then
            MsgBox "請點選門禁機！", vbCritical
            Exit Sub
         End If
         ReadTimeZone
   End Select
   CmdEnable2 Index
End Sub

Private Sub Command4_Click(Index As Integer)
   Select Case Index
   Case 0 '重整
      LoadHoliday
   Case 1 '讀取門禁機假日表
      If cboHtaIp2(2).ListIndex < 1 Then
         MsgBox "請點選門禁機！", vbCritical
         Exit Sub
      Else
         ReadHoliday
      End If
   Case 2 '假日表回寫門禁機
      If cboHtaIp2(2).ListIndex < 0 Then
         MsgBox "請點選門禁機！", vbCritical
         Exit Sub
      Else
         If MsgBox(lblAlert & vbCrLf & "是否繼續？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
            WriteHoliday
         End If
      End If
   End Select
End Sub

Private Sub Form_Load()
   Dim arrIP() As String
   Dim ii As Integer, jj As Integer, stTmp As String
   
   
   MoveFormToCenter Me
   
   '讀取考勤機IP
   cboHtaIp.Clear
   HTAips = GetHtaIP()
   If HTAips <> "" Then
      arrIP = Split(HTAips, ";")
      'Added by Morgan 2024/9/30
      '以IP排序
      For ii = LBound(arrIP) To UBound(arrIP)
         For jj = ii + 1 To UBound(arrIP)
            If arrIP(ii) <> "" And arrIP(jj) <> "" Then
               If arrIP(jj) < arrIP(ii) Then
                  stTmp = arrIP(ii)
                  arrIP(ii) = arrIP(jj)
                  arrIP(jj) = stTmp
               End If
            End If
         Next
      Next
      'end 2024/9/30
      
      For intI = LBound(arrIP) To UBound(arrIP)
         If arrIP(intI) <> "" Then
            strExc(1) = arrIP(intI) & " " & Pub_GetSpecMan(arrIP(intI))
            cboHtaIp.AddItem strExc(1)
         End If
      Next
      cboHtaIp.ListIndex = -1
   End If
   
   '設定所別
   cboBranch.Clear
   cboBranch.AddItem "高", 0
   cboBranch.ITEMDATA(0) = "4"
   cboBranch.AddItem "南", 0
   cboBranch.ITEMDATA(0) = "3"
   cboBranch.AddItem "中", 0
   cboBranch.ITEMDATA(0) = "2"
   cboBranch.AddItem "北", 0
   cboBranch.ITEMDATA(0) = "1"
   
   'Added by Morgan 2018/1/4
   Label3 = "( 0/0 )"
   List1.Clear
   'end 2018/1/4
   
   'Added by Morgan 2020/9/3
   SSTab1_Click SSTab1.Tab
   HTAips2 = GetHtaIP(2) '門禁機
   cboHtaIp2(0).Clear
   cboHtaIp2(1).Clear
   cboHtaIp2(2).Clear
   cboHtaIp2(2).AddItem "全部"
   If HTAips2 <> "" Then
      arrIP = Split(HTAips2, ";")
      
      '以IP排序
      For ii = LBound(arrIP) To UBound(arrIP)
         For jj = ii + 1 To UBound(arrIP)
            If arrIP(ii) <> "" And arrIP(jj) <> "" Then
               If arrIP(jj) < arrIP(ii) Then
                  stTmp = arrIP(ii)
                  arrIP(ii) = arrIP(jj)
                  arrIP(jj) = stTmp
               End If
            End If
         Next
      Next
      
      For intI = LBound(arrIP) To UBound(arrIP)
         If arrIP(intI) <> "" Then
            strExc(1) = arrIP(intI) & " " & Pub_GetSpecMan(arrIP(intI))
            cboHtaIp2(0).AddItem strExc(1)
            cboHtaIp2(1).AddItem strExc(1)
            cboHtaIp2(2).AddItem strExc(1)
         End If
      Next
      
      cboHtaIp2(0).ListIndex = -1
      cboHtaIp2(1).ListIndex = -1
      cboHtaIp2(2).ListIndex = -1
   End If
   
   SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160015 = Nothing
End Sub

'寫記錄
Private Function WriteLog(oStrLog As String)
   Dim ffa As Integer
   ffa = FreeFile
   Open App.path & "\" & App.EXEName & ".log" For Append As ffa
   Print #ffa, Trim(Now) & "  ==>  " & oStrLog
   Close ffa
End Function

Private Function GetDomain(pIP) As String
   Dim iPos As Integer
   iPos = InStrRev(pIP, ".")
   If iPos > 1 Then
      GetDomain = Left(pIP, iPos - 1)
   End If
End Function

Private Sub ReadTime()
   Dim arrIP() As String
   Dim strDate As String, strTime As String
   
   If cboHtaIp = "" Then
      MsgBox "請輸入考勤機IP！", vbCritical
   Else
      arrIP = Split(cboHtaIp, " ")
      HTAip = arrIP(0)
      If HTAReadTime(strDate, strTime) = True Then
         Me.lblDateTime = Format(Left(strDate, 8), ADFormat) & " " & Format(strTime, "00:00:00")
      End If
   End If
End Sub

Private Function UpdateTime() As Boolean
   Dim arrIP() As String
   Dim bolErr As Boolean
   
   arrIP = Split(cboHtaIp, " ")
   HTAip = arrIP(0)
   If HTAWriteTime(True) = False Then
      bolErr = True
      MsgBox "考勤機 ( " & HTAip & " ) 校時失敗！", vbCritical
   Else
      UpdateTime = True
   End If
End Function

Private Sub BatchWrite()
   Dim bolFail As Boolean
   Dim arrIP() As String
   Dim bolResult As Boolean
   
   If cboHtaIp = "" Then
      MsgBox "請輸入考勤機IP！", vbCritical
   ElseIf cboBranch = "" Then
      MsgBox "請選擇所別！", vbCritical
   Else
      'Modified by Morgan 2019/4/26 +st60(若有設定時改用此欄位寫入考勤機)
      'Modified by Morgan 2020/8/28 +st73
      'Modified by Morgan 2021/6/9 排除空白的指紋資料
      strExc(0) = "select st02,nvl(st60,st02) st60,st73,a.* from staffcarddata a,staff where st01(+)=scd01 and st04='1' and st06='" & cboBranch.ITEMDATA(cboBranch.ListIndex) & "' and (scd02<>scd01 or scd03||scd04 is not null) order by st01"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("將寫入 " & RsTemp.RecordCount & " 筆資料至 (" & cboHtaIp & ") 是否確定要繼續？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
         
         arrIP = Split(cboHtaIp, " ")
         HTAip = arrIP(0)
         If ghComm > 0 Then
            If HTAclose() = False Then Exit Sub
         End If
         
         If ghComm = 0 Then HTAconnect
         If ghComm > 0 Then
            If HTAdeleteAllCard = True Then
               With RsTemp
               'Added by Morgan 2018/1/4
               ProgressBar1.max = .RecordCount
               ProgressBar1.Value = 0
               Label3 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
               DoEvents
               'end 2018/1/4
               Do While Not .EOF
                  'Added by Morgan 2020/8/25
                  '門禁機無批次刪除指令，改逐筆刪除
                  If pubIsNewDevice Then
                     'Modified by Morgan 2024/8/1 +失敗選擇
                     'if HTAdeleteCard(.Fields("scd02")) = False then
                     '   bolFail = False
                     'end if
                     bolFail = False
                     Do While HTAdeleteCard(.Fields("scd02"), True) = False
                        intI = MsgBox("指紋/卡片刪除失敗！" & vbCrLf & vbCrLf & " 是:繼續 否:重試 取消:結束", vbYesNoCancel)
                        If intI = vbYes Then
                           Exit Do
                        ElseIf intI = vbCancel Then
                           bolFail = True
                           Exit Do
                        End If
                     Loop
                     'end 2024/8/1
                     
                     If bolFail = True Then
                        Exit Do
                     End If
                  End If
                  'end 2020/8/25
                  
                  '指紋
                  If .Fields("scd01") = .Fields("scd02") Then
                     'Added by Morgan 2018/1/4
                     List1.AddItem time & " --> " & .AbsolutePosition & ":" & .Fields("st02") & " ( " & .Fields("scd01") & " )", 0
                     ProgressBar1.Value = ProgressBar1.Value + 1
                     Label3 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                     DoEvents
                     'end 2018/1/4
                     'Modified by Morgan 2023/10/6 原全時段(不檢查)改為設0/4時段
                     'If HTAaddFingerPrinter(.Fields("scd02"), .Fields("st60"), "" & .Fields("scd03"), "" & .Fields("scd04"), , , , , , IIf("" & .Fields("st73") = "Y", True, False)) = False Then
                     'Modified by Morgan 2024/8/1 +失敗選擇
                     'If HTAaddFingerPrinter(.Fields("scd02"), .Fields("st60"), "" & .Fields("scd03"), "" & .Fields("scd04"), , , , , IIf("" & .Fields("st73") = "S", 0, IIf("" & .Fields("st73") = "Y", 4, 1))) = False Then
                     '   MsgBox .Fields("st02") & " ( " & .Fields("scd01") & " ) 指紋回寫失敗！作業中斷！", vbCritical
                     '   bolFail = True
                     '   Exit Do
                     'End If
                     bolFail = False
                     'Modified by Morgan 2025/4/8
                     'Do While HTAaddFingerPrinter(.Fields("scd02"), .Fields("st60"), "" & .Fields("scd03"), "" & .Fields("scd04"), , , , , IIf("" & .Fields("st73") = "S", 0, IIf("" & .Fields("st73") = "Y", 4, 1))) = False
                     Do While HTAaddFingerPrinter(.Fields("scd02"), .Fields("st60"), "" & .Fields("scd03"), "" & .Fields("scd04"), , , , , IIf("" & .Fields("st73") = "S", 0, IIf("" & .Fields("st73") = "Y", 4, 1)), IIf("" & .Fields("st73") <> "", True, False)) = False
                        intI = MsgBox(.Fields("st02") & " ( " & .Fields("scd01") & " ) 指紋回寫失敗！" & vbCrLf & vbCrLf & " 是:繼續 否:重試 取消:結束", vbYesNoCancel)
                        If intI = vbYes Then
                           Exit Do
                        ElseIf intI = vbCancel Then
                           bolFail = True
                           Exit Do
                        End If
                     Loop
                     If bolFail = True Then
                        Exit Do
                     End If
                     'end 2024/8/1
                  '卡片
                  Else
                     'Added by Morgan 2018/1/4
                     List1.AddItem time & " --> " & .AbsolutePosition & ":" & .Fields("st02") & " ( " & .Fields("scd01") & " ) 卡片 ( " & .Fields("scd02") & " )", 0
                     ProgressBar1.Value = ProgressBar1.Value + 1
                     Label3 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                     DoEvents
                     'end 2018/1/4
                     'Modified by Morgan 2023/10/6 原全時段(不檢查)改為設0/4時段
                     'If HTAaddCard(.Fields("scd02"), .Fields("st60"), , , , , , IIf("" & .Fields("st73") = "Y", True, False)) = False Then
                     'Modified by Morgan 2024/8/1 +失敗選擇
                     'If HTAaddCard(.Fields("scd02"), .Fields("st60"), , , , , IIf("" & .Fields("st73") = "S", 0, IIf("" & .Fields("st73") = "Y", 4, 1))) = False Then
                     '   MsgBox .Fields("st02") & " ( " & .Fields("scd01") & " ) 卡片 ( " & .Fields("scd02") & " ) 回寫失敗！作業中斷！", vbCritical
                     '   bolFail = True
                     '   Exit Do
                     'End If
                     bolFail = False
                     'Modified by Morgan 2025/4/8
                     'Do While HTAaddCard(.Fields("scd02"), .Fields("st60"), , , , , , IIf("" & .Fields("st73") = "Y", True, False)) = False
                     Do While HTAaddCard(.Fields("scd02"), .Fields("st60"), , , , , IIf("" & .Fields("st73") = "S", 0, IIf("" & .Fields("st73") = "Y", 4, 1)), IIf("" & .Fields("st73") <> "", True, False)) = False
                     'end 2025/4/8
                        intI = MsgBox(.Fields("st02") & " ( " & .Fields("scd01") & " ) 卡片 ( " & .Fields("scd02") & " ) 回寫失敗！" & vbCrLf & vbCrLf & " 是:繼續 否:重試 取消:結束", vbYesNoCancel)
                        If intI = vbYes Then
                           Exit Do
                        ElseIf intI = vbCancel Then
                           bolFail = True
                           Exit Do
                        End If
                     Loop
                     If bolFail = True Then
                        Exit Do
                     End If
                     'end 2024/8/1
                  End If
                  .MoveNext
               Loop
               End With
               
               If bolFail = False Then
                  MsgBox "回寫完成！", vbInformation
               End If
            End If
         End If
      End If
      
      If ghComm > 0 Then HTAclose
   End If
End Sub

Private Sub MSHFlexGrid1_Click()
   With MSHFlexGrid1
   .row = .MouseRow
   .col = .MouseCol
   If Command2(3).Enabled = True Then
      If .col = 3 Then
         SetBox MSHFlexGrid1, txtInput
      Else
         SetBox MSHFlexGrid1, txtInput, Replace(.TextMatrix(.row, .col), ":", "")
      End If
   End If
   If .TextMatrix(.row, 0) <> "" Then
      SetGridColor MSHFlexGrid1, iLstRow1
      iLstRow1 = .row
   End If
   End With
End Sub

Private Sub MSHFlexGrid2_Click()
   With MSHFlexGrid2
   .row = .MouseRow
   .col = .MouseCol
   If Command3(3).Enabled = True Then
      SetBox MSHFlexGrid2, txtInput2
   End If
   If .TextMatrix(.row, 0) <> "" Then
      SetGridColor MSHFlexGrid2, iLstRow2
      iLstRow2 = .row
   End If
   End With
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Static flg1 As Byte, flg2 As Byte, flg3 As Byte
   
   If SSTab1.Tab = 1 Then
      If flg1 = 0 Then
         LoadGrid1
         flg1 = 1
      End If
   ElseIf SSTab1.Tab = 2 Then
      If flg2 = 0 Then
         LoadGrid2
         flg2 = 1
      End If
   ElseIf SSTab1.Tab = 3 Then
      If flg3 = 0 Then
         LoadHoliday
         List2(1).Clear
         flg3 = 1
      End If
   End If
End Sub

Private Sub LoadGrid1()
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   SetGridHead1 True
   iLstRow1 = 0
   stSQL = "select TS01,trim(replace(to_char(TS02/100,'00.00'),'.',':')) TS02,trim(replace(to_char(TS03/100,'00.00'),'.',':')) TS03" & _
      ", TS04 from timesheet order by 1"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      Set MSHFlexGrid1.Recordset = rsQuery
      SetGridHead1
      MSHFlexGrid1_Click
   End If
End Sub

Private Sub LoadGrid2()
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   SetGridHead2 True
   iLstRow2 = 0
   stSQL = "select TZ01,TZ02,TZ03,TZ02,TZ05,TZ06,TZ07,TZ08,TZ09 from timezone order by 1"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      Set MSHFlexGrid2.Recordset = rsQuery
      SetGridHead2
      MSHFlexGrid2_Click
   End If
End Sub

Private Sub LoadHoliday()
   Dim rsQuery As ADODB.Recordset
   
   List2(0).Clear
   Label6(0) = "共 0 筆"
   If PUB_GetHoliday(rsQuery, True) = True Then
      With rsQuery
      .MoveFirst
      Do While Not .EOF
         List2(0).AddItem .Fields("td") & " (" & PUB_ChgNumber2Chinese(.Fields("wd")) & ")", 0
         List2(0).Selected(0) = True
         .MoveNext
      Loop
      List2(0).ListIndex = -1
      Label6(0) = "共 " & List2(0).ListCount & " 筆"
      End With
   End If
End Sub

Private Function SaveGrid1() As Boolean
   Dim ii As Integer, stSQL As String
      
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd1
   cnnConnection.Execute "delete timesheet"
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open "select * from timesheet", cnnConnection, adOpenDynamic, adLockBatchOptimistic
   With MSHFlexGrid1
   For ii = 1 To .Rows - 1
      If .TextMatrix(ii, 0) <> "" Then
         RsTemp.AddNew
         RsTemp.Fields("TS01") = .TextMatrix(ii, 0)
         RsTemp.Fields("TS02") = Replace(.TextMatrix(ii, 1), ":", "")
         RsTemp.Fields("TS03") = Replace(.TextMatrix(ii, 2), ":", "")
         RsTemp.Fields("TS04") = .TextMatrix(ii, 3)
      End If
   Next
   End With
   RsTemp.UpdateBatch
   cnnConnection.CommitTrans
   SaveGrid1 = True
   Exit Function
   
ErrHnd1:
   cnnConnection.RollbackTrans
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Function SaveGrid2() As Boolean
   Dim ii As Integer, stSQL As String
      
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd1
   cnnConnection.Execute "delete timezone"
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open "select * from timezone", cnnConnection, adOpenDynamic, adLockBatchOptimistic
   With MSHFlexGrid2
   For ii = 1 To .Rows - 1
      If .TextMatrix(ii, 0) <> "" Then
         RsTemp.AddNew
         RsTemp.Fields("TZ01") = .TextMatrix(ii, 0)
         RsTemp.Fields("TZ02") = .TextMatrix(ii, 1)
         RsTemp.Fields("TZ03") = .TextMatrix(ii, 2)
         RsTemp.Fields("TZ04") = .TextMatrix(ii, 3)
         RsTemp.Fields("TZ05") = .TextMatrix(ii, 4)
         RsTemp.Fields("TZ06") = .TextMatrix(ii, 5)
         RsTemp.Fields("TZ07") = .TextMatrix(ii, 6)
         RsTemp.Fields("TZ08") = .TextMatrix(ii, 7)
         RsTemp.Fields("TZ09") = .TextMatrix(ii, 8)
      End If
   Next
   End With
   RsTemp.UpdateBatch
   cnnConnection.CommitTrans
   SaveGrid2 = True
   Exit Function
   
ErrHnd1:
   cnnConnection.RollbackTrans
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub SetGridHead1(Optional bolReset As Boolean = False)
   With MSHFlexGrid1
   If bolReset Then
      .Clear
      .Rows = 2
      .Cols = 4
   End If
   .TextMatrix(0, 0) = "ID"
   .ColWidth(0) = 500
   .ColAlignmentFixed(0) = flexAlignCenterCenter
   .ColAlignment(0) = flexAlignCenterCenter
   .TextMatrix(0, 1) = "起始時間"
   .ColWidth(1) = 1000
   .ColAlignmentFixed(1) = flexAlignCenterCenter
   .ColAlignment(1) = flexAlignCenterCenter
   .TextMatrix(0, 2) = "結束時間"
   .ColWidth(2) = 1000
   .ColAlignmentFixed(2) = flexAlignCenterCenter
   .ColAlignment(2) = flexAlignCenterCenter
   .TextMatrix(0, 3) = "說明"
   .ColWidth(3) = 1500
   .ColAlignmentFixed(3) = flexAlignLeftCenter
   .ColAlignment(3) = flexAlignLeftCenter
   End With
End Sub

Private Sub SetGridHead2(Optional bolReset As Boolean = False)
   With MSHFlexGrid2
   If bolReset Then
      .Clear
      .Rows = 2
      .Cols = 9
   End If
   .TextMatrix(0, 0) = "ID"
   .ColWidth(0) = 350
   .ColAlignmentFixed(0) = flexAlignCenterCenter
   .ColAlignment(0) = flexAlignCenterCenter
   .TextMatrix(0, 1) = "一"
   .ColWidth(1) = 350
   .ColAlignmentFixed(1) = flexAlignCenterCenter
   .ColAlignment(1) = flexAlignCenterCenter
   .TextMatrix(0, 2) = "二"
   .ColWidth(2) = 350
   .ColAlignmentFixed(2) = flexAlignCenterCenter
   .ColAlignment(2) = flexAlignCenterCenter
   .TextMatrix(0, 3) = "三"
   .ColWidth(3) = 350
   .ColAlignmentFixed(3) = flexAlignCenterCenter
   .ColAlignment(3) = flexAlignCenterCenter
   .TextMatrix(0, 4) = "四"
   .ColWidth(4) = 350
   .ColAlignmentFixed(4) = flexAlignCenterCenter
   .ColAlignment(4) = flexAlignCenterCenter
   .TextMatrix(0, 5) = "五"
   .ColWidth(5) = 350
   .ColAlignmentFixed(5) = flexAlignCenterCenter
   .ColAlignment(5) = flexAlignCenterCenter
   .TextMatrix(0, 6) = "六"
   .ColWidth(6) = 350
   .ColAlignmentFixed(6) = flexAlignCenterCenter
   .ColAlignment(6) = flexAlignCenterCenter
   .TextMatrix(0, 7) = "日"
   .ColWidth(7) = 350
   .ColAlignmentFixed(7) = flexAlignCenterCenter
   .ColAlignment(7) = flexAlignCenterCenter
   .TextMatrix(0, 8) = "說明"
   .ColWidth(8) = 1250
   .ColAlignmentFixed(8) = flexAlignLeftCenter
   .ColAlignment(8) = flexAlignLeftCenter
   End With
End Sub

Private Sub SetBox(pGrid As MSHFlexGrid, pText As TextBox, Optional pValue As String = "")
   Dim ii As Integer
   Dim lngLeft As Long, lngTop As Long
   
   With pGrid
      If .row > 0 And .col > 0 Then
         pText.FontName = .CellFontName
         pText.FontSize = .CellFontSize
         pText.Alignment = .CellAlignment \ 5
         If pValue <> "" Then
            pText.Text = pValue
         Else
            pText.Text = .TextMatrix(.row, .col)
         End If
         pText.Tag = pText.Text
         pText.Width = .ColWidth(.col)
         pText.Height = .RowHeight(.row)
         
         If .CellAlignment < 3 Then
            pText.Alignment = 0
         ElseIf .CellAlignment < 6 Then
            pText.Alignment = 2
         Else
            pText.Alignment = 1
         End If
         lngLeft = .Left + 25
         lngTop = .Top + .RowHeight(0) + 25
         For ii = 0 To .col - 1
            lngLeft = lngLeft + .ColWidth(ii)
         Next
         For ii = .TopRow To .row - 1
            lngTop = lngTop + .RowHeight(ii)
         Next
         pText.Left = lngLeft: pText.Top = lngTop
         pText.Visible = True
         pText.SetFocus
         TextInverse pText
         iRow = .row: iCol = .col
      End If
   End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
      
   If iCol <> 3 And Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   Else
      If KeyAscii = vbKeyReturn Then
         With MSHFlexGrid1
         If iCol = 3 Then
            .TextMatrix(iRow, iCol) = txtInput.Text
         Else
            .TextMatrix(iRow, iCol) = Format(txtInput.Text, "00:00")
         End If
         If iCol > 0 And iCol < 3 Then
            .col = iCol + 1
            If .col = 3 Then
               SetBox MSHFlexGrid1, txtInput
            Else
               SetBox MSHFlexGrid1, txtInput, Replace(.TextMatrix(.row, .col), ":", "")
            End If
         Else
            txtInput.Visible = False
         End If
         End With
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput = txtInput.Tag
         TextInverse txtInput
      End If
   End If
End Sub
'時區表回寫門禁機
Private Function WritTimeSheet() As Boolean
   Dim ii As Integer
   Dim arrIP() As String
   Dim bolOK As Boolean
   Dim sID As String, sFromTime As String, sToTime As String
   
   With MSHFlexGrid1
   If .TextMatrix(1, 0) <> "" Then
      arrIP = Split(cboHtaIp2(0), " ")
      HTAip = arrIP(0)
      '關閉舊連線
      If ghComm > 0 Then
         If HTAclose() = False Then Exit Function
      End If
      HTAconnect
      If ghComm > 0 Then
         If HTAClearTimeSheet = True Then '清除時區表
            bolOK = True
            For ii = 1 To .Rows - 1
               If .TextMatrix(ii, 0) <> "" Then
                  sID = .TextMatrix(ii, 0)
                  sFromTime = Replace(.TextMatrix(ii, 1), ":", "")
                  sToTime = Replace(.TextMatrix(ii, 2), ":", "")
                  If HTAWriteTimeSheet(sID, sFromTime, sToTime) = False Then
                     If ghComm > 0 Then HTAclose
                     MsgBox "時區表寫入失敗！", vbCritical
                     bolOK = False
                     Exit For
                  End If
               End If
            Next
         End If
         If bolOK Then
            If ghComm > 0 Then HTAclose
            WritTimeSheet = True
            MsgBox "時區表回寫完成！", vbOKOnly + vbInformation
         Else
            MsgBox "時區表回寫失敗！", vbCritical
         End If
      End If
   End If
   End With
End Function

'時段表回寫門禁機
Private Function WritTimeZone() As Boolean
   Dim ii As Integer
   Dim arrIP() As String
   Dim bolOK As Boolean
   
   With MSHFlexGrid2
   If .TextMatrix(1, 0) <> "" Then
      arrIP = Split(cboHtaIp2(1), " ")
      HTAip = arrIP(0)
      '關閉舊連線
      If ghComm > 0 Then
         If HTAclose() = False Then Exit Function
      End If
      HTAconnect
      If ghComm > 0 Then
         If HTAClearTimeZone = True Then '清除時段表
            bolOK = True
            For ii = 1 To .Rows - 1
               If .TextMatrix(ii, 0) <> "" Then
                  If HTAWriteTimeZone(.TextMatrix(ii, 0), .TextMatrix(ii, 1), .TextMatrix(ii, 2), .TextMatrix(ii, 3), .TextMatrix(ii, 4), .TextMatrix(ii, 5), .TextMatrix(ii, 6), .TextMatrix(ii, 7)) = False Then
                     If ghComm > 0 Then HTAclose
                     MsgBox "時段表寫入失敗！", vbCritical
                     bolOK = False
                     Exit For
                  End If
               End If
            Next
         End If
         If bolOK Then
            If ghComm > 0 Then HTAclose
            WritTimeZone = True
            MsgBox "時段表回寫完成！", vbOKOnly + vbInformation
         Else
            MsgBox "時段表回寫失敗！", vbCritical
         End If
      End If
   End If
   End With
End Function

Private Sub txtInput_Validate(Cancel As Boolean)
   If iCol = 3 Then
      MSHFlexGrid1.TextMatrix(iRow, iCol) = txtInput.Text
   Else
      MSHFlexGrid1.TextMatrix(iRow, iCol) = Format(txtInput.Text, "00:00")
   End If
End Sub

Private Sub txtInput2_KeyPress(KeyAscii As Integer)
      
   If iCol <> 8 And Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
      KeyAscii = 0
      Beep
   Else
      If KeyAscii = vbKeyReturn Then
         MSHFlexGrid2.TextMatrix(iRow, iCol) = txtInput2.Text
         If iCol > 0 And iCol < 8 Then
            MSHFlexGrid2.col = iCol + 1
            SetBox MSHFlexGrid2, txtInput2
         Else
            txtInput2.Visible = False
         End If
      ElseIf KeyAscii = vbKeyEscape Then
         txtInput2 = txtInput2.Tag
         TextInverse txtInput2
      End If
   End If
End Sub

Private Sub txtInput2_Validate(Cancel As Boolean)
   MSHFlexGrid2.TextMatrix(iRow, iCol) = txtInput2.Text
End Sub

'Added by Morgan 2020/11/16
'讀取時區表
Private Sub ReadTimeSheet()
   Dim ii As Integer
   Dim arrIP() As String
   Dim bolOK As Boolean
   Dim arrDate(100) As String
   Dim stData As String
   
   Text1 = ""
   arrIP = Split(cboHtaIp2(0), " ")
   HTAip = arrIP(0)
   '關閉舊連線
   If ghComm > 0 Then
      If HTAclose() = False Then Exit Sub
   End If
   HTAconnect
   If ghComm > 0 Then
      Erase arrDate
      If HTAReadTimeSheet(arrDate) = True Then
         For ii = 1 To 40
            If ii Mod 4 = 1 Then
               If ii > 1 Then stData = stData & vbCrLf
               stData = stData & ii \ 4 & " "
               stData = stData & arrDate(ii) & ":"
            ElseIf ii Mod 4 = 2 Then
               stData = stData & arrDate(ii) & " "
               
            ElseIf ii Mod 4 = 3 Then
               stData = stData & arrDate(ii) & ":"
            
            ElseIf ii Mod 4 = 0 Then
               stData = stData & arrDate(ii) & " "
            End If
         Next
         Text1 = stData
         bolOK = True
      End If
   End If
   
   If ghComm > 0 Then HTAclose
   If bolOK Then MsgBox "時區表讀取完成！", vbOKOnly + vbInformation
End Sub

'讀取時段表
Private Sub ReadTimeZone()
   Dim ii As Integer
   Dim arrIP() As String
   Dim bolOK As Boolean
   Dim arrDate(100) As String
   Dim stData As String
   
   Text2 = ""
   arrIP = Split(cboHtaIp2(1), " ")
   HTAip = arrIP(0)
   '關閉舊連線
   If ghComm > 0 Then
      If HTAclose() = False Then Exit Sub
   End If
   HTAconnect
   If ghComm > 0 Then
      Erase arrDate
      If HTAReadTimeZone(arrDate) = True Then
         For ii = 1 To 70
            If arrDate(ii) = "FF" Then Exit For 'Added by Morgan 2023/10/11 預設為FF(不是00)
            If ii Mod 7 = 1 Then
               If ii > 1 Then stData = stData & vbCrLf
               stData = stData & ii \ 7 & " "
            End If
            stData = stData & Val("&H" & arrDate(ii)) & " "
         Next
         Text2 = stData
         bolOK = True
      End If
   End If
   
   If ghComm > 0 Then HTAclose
   If bolOK Then MsgBox "時段表讀取完成！", vbOKOnly + vbInformation
End Sub

'自門禁機讀取假日表
'注意:假日設定需跨日才有效(應該是門禁機只有在跨日時才會讀取設定)
Private Sub ReadHoliday()
   Dim ii As Integer
   Dim arrIP() As String
   Dim bolOK As Boolean
   Dim arrDate(200) As String
   
   If cboHtaIp2(2).ListIndex <= 0 Then Exit Sub
   
   List2(1).Clear
   Label6(1) = "共 0 筆"
   arrIP = Split(cboHtaIp2(2), " ")
   HTAip = arrIP(0)
   '關閉舊連線
   If ghComm > 0 Then
      If HTAclose() = False Then Exit Sub
   End If
   HTAconnect
   If ghComm > 0 Then
      Erase arrDate
      If HTAReadHoliday(arrDate) = True Then
         For ii = 1 To 200
            'Debug.Print ii & ">>" & arrDate(ii)
            If Left(arrDate(ii), 2) <> "FF" And arrDate(ii) <> "" Then
               List2(1).AddItem Format(Left(arrDate(ii), 4), "@@/@@") & " " & Val(Right(arrDate(ii), 2))
            End If
         Next
         List2(1).ListIndex = -1
         Label6(1) = "共 " & List2(1).ListCount & " 筆"
         MsgBox "假日表讀取完成！" & vbCrLf & Label6(1), vbOKOnly + vbInformation
      End If
   End If
   
   If ghComm > 0 Then HTAclose
End Sub

'假日表回寫門禁機
'注意:假日設定需跨日才有效(應改是門禁機只有在跨日時才會讀取設定)
Private Function WriteHoliday() As Boolean
   Dim ii As Integer
   Dim iCount As Integer, iCount2 As Integer
   Dim arrIP() As String
   Dim bolOK As Boolean
   Dim arrDate(200) As String
   Dim strErr As String, StrOk As String
   
   'Added by Morgan 2023/4/26
   If cboHtaIp2(2).ListIndex = 0 Then
      If MsgBox("是否確定全部回寫？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   'end 2023/4/26
   
   iCount = 0
   iCount2 = 0
   For ii = 0 To List2(0).ListCount - 1
      If List2(0).Selected(ii) = True Then
         If Left(List2(0).List(ii), 4) = Left(strSrvDate(1), 4) Then
            iCount = iCount + 1
            arrDate(iCount) = Replace(Mid(List2(0).List(ii), 6, 5), "/", "")
         Else
            iCount2 = iCount2 + 1
            arrDate(100 + iCount2) = Replace(Mid(List2(0).List(ii), 6, 5), "/", "")
         End If
      End If
   Next
   
   iCount = cboHtaIp2(2).ListIndex
   For ii = 1 To cboHtaIp2(2).ListCount - 1
      If iCount = 0 Or iCount = ii Then
         arrIP = Split(cboHtaIp2(2).List(ii), " ")
         HTAip = arrIP(0)
         '關閉舊連線
         If ghComm > 0 Then
            If HTAclose() = False Then Exit Function
         End If
         HTAconnect
         If ghComm > 0 Then
            bolOK = HTAWriteHoliday(arrDate)
            If ghComm > 0 Then HTAclose
            If bolOK Then
               StrOk = StrOk & vbCrLf & cboHtaIp2(2).List(ii)
            Else
               strErr = strErr & vbCrLf & cboHtaIp2(2).List(ii)
            End If
         End If
      End If
   Next
   
   If strErr = "" Then
      WriteHoliday = True
      MsgBox "假日表回寫完成！" & vbCrLf & StrOk, vbOKOnly + vbInformation
      
      If cboHtaIp2(2).ListIndex > 0 Then ReadHoliday 'Added by Morgan 2024/12/3
   Else
      MsgBox "假日表回寫失敗！" & vbCrLf & IIf(StrOk <> "", "成功：" & vbCrLf & StrOk, "") & "失敗：" & vbCrLf & strErr, vbCritical
   End If
   
End Function


Private Function PollingData() As Boolean
   Dim arrIP() As String
   Dim iRecs As Integer, iRecTot As Integer
   Dim arrIpList
   Dim ii As Integer
   Dim bResult As Boolean
   Dim iRtn As Integer
   Dim iTimes As Integer
   
   arrIP = Split(cboHtaIp, " ")
   HTAip = arrIP(0)
   If HTAip <> "" Then
      Pub_WriteSysLog "開始下載...(" & HTAip & ")"
      iTimes = 1
      bResult = False
      bResult = HTAPolling(iRecs, True)
      Do While (bResult = False And iTimes < 3)
         Sleep 3000
         iTimes = iTimes + 1
         bResult = HTAPolling(iRecs, True)
      Loop
      
      If bResult = True Then
         iRecTot = iRecTot + iRecs
      Else
         MsgBox "考勤機(" & HTAip & ") 刷卡紀錄接收失敗！", vbCritical
      End If
   End If
   
   PollingData = bResult
   MsgBox "已接收 " & iRecTot & " 筆!", vbInformation
   
End Function

'Added by Morgan 2024/3/4
Private Sub InitDevice()
   Dim arrIP() As String
   Dim strMsg As String
   
   If cboHtaIp = "" Then
      MsgBox "請輸入考勤機IP！", vbCritical
   Else
      arrIP = Split(cboHtaIp, " ")
      HTAip = arrIP(0)
      
      If CheckLevel(HTAip & ";", "門禁機IP") Then
         MsgBox "本功能尚未支援【門禁機】初始化！", vbCritical
      ElseIf MsgBox("是否確定要將【考勤機 " & cboHtaIp & "】初始化？" & vbCrLf & vbCrLf & "初始化後記得用廠商的APP重新回寫【語言、響鈴時間...等】相關設定!!!", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
      
         If ghComm > 0 Then
            If HTAclose() = False Then Exit Sub
         End If
         
         If ghComm = 0 Then HTAconnect
         If ghComm > 0 Then
            If HTAInitial() = True Then
               MsgBox "【考勤機 " & cboHtaIp & "】 系統初始化成功！", vbCritical
            Else
               MsgBox "【考勤機 " & cboHtaIp & "】 系統初始化失敗！", vbCritical
            End If
         End If
      End If
   End If
End Sub

'Added by Morgan 2024/5/24 (目前還不能用)
'讀取所有合法卡
Private Sub ReadAllCard()
   Dim ii As Integer
   Dim arrIP() As String
   Dim bolOK As Boolean
   Dim arrDate(100) As String
   Dim stData As String
   
   Text1 = ""
   arrIP = Split(cboHtaIp, " ")
   HTAip = arrIP(0)
   '關閉舊連線
   If ghComm > 0 Then
      If HTAclose() = False Then Exit Sub
   End If
   HTAconnect
   If ghComm > 0 Then
      Erase arrDate
      If HTAReadAllCard(arrDate) = True Then
         For ii = 1 To 40
            If ii Mod 4 = 1 Then
               If ii > 1 Then stData = stData & vbCrLf
               stData = stData & ii \ 4 & " "
               stData = stData & arrDate(ii) & ":"
            ElseIf ii Mod 4 = 2 Then
               stData = stData & arrDate(ii) & " "
               
            ElseIf ii Mod 4 = 3 Then
               stData = stData & arrDate(ii) & ":"
            
            ElseIf ii Mod 4 = 0 Then
               stData = stData & arrDate(ii) & " "
            End If
         Next
         Text1 = stData
         bolOK = True
      End If
   End If
   
   If ghComm > 0 Then HTAclose
   If bolOK Then MsgBox "合法卡讀取完成！", vbOKOnly + vbInformation
End Sub

'Added by Morgan 2024/5/24
'讀取所有合法卡
Public Function HTAReadAllCard(ByRef rDate() As String, Optional NoErrMsg As Boolean, Optional pRetry As Integer = 3) As Boolean
   Dim iReturnCode, iReturn
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   Dim rByte(256) As Byte
   Dim iGetLen As Integer
   Dim iBank As Integer
   Dim iCompress As Integer
   Dim iDataLen As Integer
   Dim ii As Integer, jj As Integer
   Const cNodeID As Integer = 1
   Const cHTATimOut2 As Integer = 6000
   
On Error GoTo ErrHnd
   
   If ghComm = 0 Then
      HTAconnect NoErrMsg
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function

   iErrCount = 0
   Do While iErrCount < pRetry
      iGetLen = 40
      iBank = 4
      iCompress = 0
      iReturn = htaGetCardData(ghComm, cNodeID, rByte(0), iDataLen, iBank, iCompress, cHTATimOut2)
      If iReturn = 0 And iDataLen > 0 Then
         For ii = 1 To iDataLen
            rDate(ii) = Right("0" & Hex(rByte(ii - 1)), 2)
         Next
         HTAReadAllCard = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
   
   If iReturn <> 0 And NoErrMsg = False Then
      MsgBox "合法卡讀取失敗！", vbExclamation
   End If
   
   If bolNewComm = True Then
      HTAclose NoErrMsg
   End If
   Exit Function
      
ErrHnd:
   If NoErrMsg = False Then
      MsgBox Err.Description, vbCritical
   End If
   
End Function

'Added by Morgan 2024/9/30
'回寫指定時間到考勤機/門禁機
Private Function UpdateTime2(pDateTime As String, Optional pRetry As Integer = 3) As Boolean
   Dim arrIP() As String
   Dim bolErr As Boolean
   
   Dim iReturnCode, iReturn, iELID As Integer
   Dim iweek As Integer
   Dim Sdate As String
   Dim sTime As String, sTime1 As String
   Dim dtNow As Date
   Dim bolNewComm As Boolean
   Dim iErrCount As Integer
   
On Error GoTo ErrHnd

   arrIP = Split(cboHtaIp, " ")
   HTAip = arrIP(0)
   
   dtNow = pDateTime
   
   If ghComm = 0 Then
      HTAconnect True
      bolNewComm = True
   End If
   If ghComm = 0 Then Exit Function

   iErrCount = 0
   Do While iErrCount < pRetry
      
      sTime = Format(dtNow, "HHmmss")
      iweek = Weekday(dtNow) - 1
      If iweek = 0 Then iweek = 7
      Sdate = Format(dtNow, "yyyymmdd")
      Sdate = Sdate & Trim(str(iweek))
      
      iReturn = 0
      iReturnCode = 0
      
      If pubIsNewDevice Then
         iReturn = hacSetDateTime(1, Sdate, sTime, ghComm, 6000)
      Else
         iReturn = hsHTA850WriteTime(ghComm, Sdate, sTime, iReturnCode, 3000)
      End If
      
      If iReturn = 0 Then
         UpdateTime2 = True
         Exit Do
      End If
      iErrCount = iErrCount + 1
   Loop
      
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
   If bolNewComm = True Then
      HTAclose True
   End If
End Function
