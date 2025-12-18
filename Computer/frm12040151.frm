VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm12040151 
   BorderStyle     =   1  '單線固定
   Caption         =   "其他特殊信函"
   ClientHeight    =   5604
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5604
   ScaleWidth      =   8280
   Begin VB.CommandButton Command7 
      Caption         =   "補美金請款金額"
      Height          =   525
      Left            =   3960
      TabIndex        =   32
      Top             =   2850
      Width           =   1245
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Default         =   -1  'True
      Height          =   465
      Left            =   5910
      TabIndex        =   25
      Top             =   30
      Width           =   800
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3225
      Left            =   30
      TabIndex        =   3
      Top             =   240
      Width           =   8085
      _ExtentX        =   14266
      _ExtentY        =   5694
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "特殊信函"
      TabPicture(0)   =   "frm12040151.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCust"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdok(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdok(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdok(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdok(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdok(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdok(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdok(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCust"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdok(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdok(9)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdok(10)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdok(11)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdok(12)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdok(13)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdok(14)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdok(15)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "聯絡單"
      TabPicture(1)   =   "frm12040151.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CommandButton cmdok 
         Caption         =   "110/9/1緬甸商標法通知"
         Height          =   615
         Index           =   15
         Left            =   6690
         TabIndex        =   35
         Top             =   1110
         Width           =   1335
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "110/6/1大陸新法-台灣設計案通函"
         Height          =   705
         Index           =   14
         Left            =   6690
         TabIndex        =   34
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "通知英國再註冊"
         Height          =   495
         Index           =   13
         Left            =   2430
         TabIndex        =   33
         Top             =   2610
         Width           =   1485
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "109英國脫歐通知2(CFP,CFT)"
         Height          =   465
         Index           =   12
         Left            =   5220
         TabIndex        =   31
         Top             =   2670
         Width           =   1425
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "109緬甸商標通知重新申請"
         Height          =   375
         Index           =   11
         Left            =   90
         TabIndex        =   30
         Top             =   2490
         Width           =   2295
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "109英國脫歐通知(CFP,CFT)"
         Height          =   495
         Index           =   10
         Left            =   5220
         TabIndex        =   29
         Top             =   2160
         Width           =   1425
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "T109專利年費退費通知"
         Height          =   315
         Index           =   9
         Left            =   90
         TabIndex        =   28
         Top             =   690
         Width           =   2175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "執行"
         Height          =   375
         Left            =   -69510
         TabIndex        =   27
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -74610
         Style           =   2  '單純下拉式
         TabIndex        =   26
         Top             =   600
         Width           =   4935
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "T100專利年費退費通知"
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtCust 
         Height          =   285
         Left            =   1050
         MaxLength       =   9
         TabIndex        =   18
         Top             =   990
         Width           =   1050
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "設定T109業務"
         Height          =   315
         Index           =   2
         Left            =   2310
         TabIndex        =   17
         Top             =   690
         Width           =   1365
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   735
         TabIndex        =   16
         Top             =   1320
         Width           =   285
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "設定T100業務"
         Height          =   315
         Index           =   3
         Left            =   2310
         TabIndex        =   15
         Top             =   360
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "新式樣改部分設計通知"
         Height          =   525
         Left            =   3990
         TabIndex        =   14
         Top             =   360
         Width           =   1230
      End
      Begin VB.CommandButton Command2 
         Caption         =   "聯合改衍生設計通知"
         Height          =   525
         Left            =   3960
         TabIndex        =   13
         Top             =   900
         Width           =   1230
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   870
         MaxLength       =   7
         TabIndex        =   12
         Top             =   1620
         Width           =   1005
      End
      Begin VB.CommandButton Command3 
         Caption         =   "設定NP業務"
         Height          =   465
         Left            =   3960
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "102年美國發明維持費規費調漲通知"
         Height          =   495
         Index           =   4
         Left            =   60
         TabIndex        =   10
         Top             =   1950
         Width           =   2355
      End
      Begin VB.CommandButton Command4 
         Caption         =   "匯入Word表格"
         Height          =   435
         Left            =   2430
         TabIndex        =   9
         Top             =   1590
         Width           =   1470
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "105年英國脫歐之CFP,CFT案通知"
         Height          =   495
         Index           =   5
         Left            =   2430
         TabIndex        =   8
         Top             =   2070
         Width           =   1485
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "108台灣設計P案專用期延長通知"
         Height          =   495
         Index           =   6
         Left            =   5220
         TabIndex        =   7
         Top             =   390
         Width           =   1425
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "108台灣設計P案專用期延長通知(已消滅)"
         Height          =   615
         Index           =   7
         Left            =   5220
         TabIndex        =   6
         Top             =   900
         Width           =   1425
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "108台灣設計P案專用期延長通知(最後1年未繳)"
         Height          =   615
         Index           =   8
         Left            =   5220
         TabIndex        =   5
         Top             =   1530
         Width           =   1425
      End
      Begin VB.CommandButton Command5 
         Caption         =   "108台灣設計案專用期延長補掛NP年費期限"
         Height          =   645
         Left            =   3930
         TabIndex        =   4
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶編號："
         Height          =   180
         Left            =   105
         TabIndex        =   24
         Top             =   1050
         Width           =   900
      End
      Begin VB.Label lblCust 
         Height          =   225
         Left            =   2220
         TabIndex        =   23
         Top             =   1050
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "條件：          ( 1:只可退當次修法 2:可退過去修法)"
         Height          =   180
         Left            =   150
         TabIndex        =   22
         Top             =   1380
         Width           =   3810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "發文日：                         以後"
         Height          =   180
         Left            =   105
         TabIndex        =   21
         Top             =   1680
         Width           =   2205
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "......"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   90
         TabIndex        =   20
         Top             =   2820
         Width           =   360
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   3510
      Width           =   6630
      _ExtentX        =   11705
      _ExtentY        =   614
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   1455
      Left            =   60
      TabIndex        =   2
      Top             =   4110
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   2561
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0/0 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   3870
      Width           =   6630
   End
End
Attribute VB_Name = "frm12040151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/5 改成Form2.0 (grdDataList)
'Memo By Sonia 2012/12/6 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'sonia 2010/8/19 日期欄已修改
'Create by Morgan 2010/1/18
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 1
         Process1
      Case 2
         Process2
      Case 3
         Process3
      '2013/2/6 add by sonia 102年美國發明維持費規費調漲通知
      Case 4
         Process4
      '2013/2/6 end
      'add by sonia 2016/7/13 105年英國脫歐之CFP,CFT案通知
      Case 5
         Process5
      'end 2016/7/13
      
      'Added by Morgan 2019/10/8 108台灣設計P案專用期延長通知
      Case 6
         Process6
      'Added by Morgan 2019/10/14 108台灣設計P案專用期延長通知(已消滅)
      Case 7
         Process7
      'Added by Morgan 2019/10/14 108台灣設計P案專用期延長通知(最後一年未繳)
      Case 8
         Process8
      'Added by Morgan 2020/8/21 109減免退費通知
      Case 9
         Process9
      'Added by Morgan 2020/11/17 109英國脫歐通知
      Case 10
         Process10
      'Added by Morgan 2020/12/9 109緬甸商標通知重新申請
      Case 11
         Process11
      'Added by Morgan 2020/12/16 109英國脫歐通知
      Case 12
         Process12
      'Added by Morgan 2021/1/20 通知英國再註冊
      Case 13
         Process13
         
      'Added by Morgan 2021/6/8 台灣設計案通函(大陸修法開拓)
      Case 14
         Process14
         
      'Added by Lydia 2021/08/31 110緬甸商標法通知
      Case 15
         Process15
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub Process2()
   Dim adoRst As ADODB.Recordset
   Dim strSNo As String
   
   'Modified by Morgan 2020/7/27 T100->T109
   strExc(0) = "select t01,t02,t03,t04 from t109"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRst
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      DoEvents
      
      .MoveFirst
      Do While Not .EOF
         strSNo = PUB_GetAKindSalesNo(.Fields("t01"), .Fields("t02"), .Fields("t03"), .Fields("t04"))
         strSql = "update t109 set t17='" & strSNo & "' where t01='" & .Fields("t01") & "'" & _
            " and t02='" & .Fields("t02") & "' and t03='" & .Fields("t03") & "'" & _
            " and t04='" & .Fields("t04") & "'"
         cnnConnection.Execute strSql, intI
         
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         .MoveNext
      Loop
      End With
      MsgBox "更新結束！"
   Else
      MsgBox "無資料可更新！"
   End If
   Set adoRst = Nothing
End Sub

Private Sub Process3()
   Dim adoRst As ADODB.Recordset
   Dim strSNo As String
   
   strExc(0) = "select t01,t02,t03,t04 from t100"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRst
      .MoveFirst
      Do While Not .EOF
         strSNo = PUB_GetAKindSalesNo(.Fields("t01"), .Fields("t02"), .Fields("t03"), .Fields("t04"))
         strSql = "update t100 set t11='" & strSNo & "' where t01='" & .Fields("t01") & "'" & _
            " and t02='" & .Fields("t02") & "' and t03='" & .Fields("t03") & "'" & _
            " and t04='" & .Fields("t04") & "'"
         cnnConnection.Execute strSql, intI
         .MoveNext
      Loop
      End With
      MsgBox "更新結束！"
   Else
      MsgBox "無資料可更新！"
   End If
   Set adoRst = Nothing
End Sub

Private Sub Command2_Click()
   Dim stCon As String
   Dim adoRst As ADODB.Recordset
   
   If Text2 <> "" Then
      stCon = stCon & " and cp27>=" & DBDATE(Text2)
   End If
   
   strExc(0) = "select cp01,cp02,cp03,cp04,cp09" & _
      " from patent,caseprogress c1" & _
      " where pa01='P' and pa08='3' and pa09='000' and pa57||pa108||pa16 is null and substr(pa11,10,1)='U'" & stCon & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10 in ('105','305') and cp24 is null and cp27>19221111" & _
      " and not exists(select * from caseprogress c2 where c2.cp01=pa01 and c2.cp02=pa02 and c2.cp03=pa03 and c2.cp04=pa04 and c2.cp10='306' and c2.cp57 is null)" & _
      " and not exists(select * from caseprogress c2 where c2.cp01=pa01 and c2.cp02=pa02 and c2.cp03=pa03 and c2.cp04=pa04 and c2.cp10='1202')" & _
      " order by 1,2,3,4"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRst
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      
      strUserNum = "81002"
      Do While Not .EOF
         EndLetter "21", .Fields("cp09"), "05", strUserNum
         strExc(1) = Format(.AbsolutePosition, "0000")
         strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & .Fields("cp09") & "','05','" & strUserNum & _
                     "','發文流水號' ,'" & strExc(1) & "')"
         cnnConnection.Execute strExc(0), intI
         
         NowPrint .Fields("cp09"), "21", "05", False, strUserNum
         
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         .MoveNext
      Loop
      strUserNum = strUser1Num
      End With
   Else
      MsgBox "無符合資料！"
   End If
   Set adoRst = Nothing
End Sub

Private Sub Command1_Click()
   Dim stCon As String
   Dim adoRst As ADODB.Recordset
   
   If Text2 <> "" Then
      stCon = stCon & " and cp27>=" & DBDATE(Text2)
   End If
   
   strExc(0) = "select cp01,cp02,cp03,cp04,cp09" & _
      " from patent,caseprogress c1" & _
      " where pa01='P' and pa08='3' and pa09='000' and pa57||pa108||pa16 is null and length(pa11)=9" & stCon & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10 in ('103','303') and cp24 is null and cp27>19221111" & _
      " and not exists(select * from caseprogress c2 where c2.cp01=pa01 and c2.cp02=pa02 and c2.cp03=pa03 and c2.cp04=pa04 and c2.cp10='305' and c2.cp57 is null)" & _
      " and not exists(select * from caseprogress c2 where c2.cp01=pa01 and c2.cp02=pa02 and c2.cp03=pa03 and c2.cp04=pa04 and c2.cp10='1202')" & _
      " order by 1,2,3,4"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRst
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      strUserNum = "81002"
      Do While Not .EOF
         EndLetter "21", .Fields("cp09"), "04", strUserNum
         strExc(1) = Format(.AbsolutePosition, "0000")
         strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & .Fields("cp09") & "','04','" & strUserNum & _
                     "','發文流水號' ,'" & strExc(1) & "')"
         cnnConnection.Execute strExc(0), intI
         
         NowPrint .Fields("cp09"), "21", "04", False, strUserNum
         
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         .MoveNext
      Loop
      strUserNum = strUser1Num
      End With
   Else
      MsgBox "無符合資料！"
   End If
   Set adoRst = Nothing
End Sub

Private Sub Command3_Click()
   Dim adoRst As ADODB.Recordset
   Dim strSNo As String
   
   strExc(0) = "select np01,np02,np03,np04,np05,np22 from nextprogress where np17=to_char(sysdate,'yyyymmdd') and np02='P' and np07='606' and np16='PGMID'"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoRst
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      
      .MoveFirst
      Do While Not .EOF
         strSNo = PUB_GetAKindSalesNo(.Fields("np02"), .Fields("np03"), .Fields("np04"), .Fields("np05"))
         strSql = "update nextprogress set np10='" & strSNo & "' where np01='" & .Fields("np01") & "' and np22=" & .Fields("np22")
         cnnConnection.Execute strSql, intI
         
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         .MoveNext
      Loop
      End With
      MsgBox "更新結束！"
   Else
      MsgBox "無資料可更新！"
   End If
   Set adoRst = Nothing
End Sub

Private Sub Command4_Click()
   Dim iCols As Integer, iRows As Integer
   Dim iCol As Integer, iRow As Integer
   Dim strPath As String
   
   Screen.MousePointer = vbHourglass
   strPath = PUB_Getdesktop & "\111.doc"
   
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   With g_WordAp
      .Documents.Open FileName:=strPath, ReadOnly:=True
      '.Visible = True
      .Selection.HomeKey Unit:=wdStory
      '執行尋找
      .Selection.Find.ClearFormatting
      .Selection.Find.Text = "欄位1"
      .Selection.Find.Replacement.Text = ""
      .Selection.Find.Forward = True
      .Selection.Find.Wrap = wdFindContinue
      .Selection.Find.Format = False
      .Selection.Find.MatchCase = False
      .Selection.Find.MatchWholeWord = False
      .Selection.Find.MatchWildcards = False
      .Selection.Find.MatchSoundsLike = False
      .Selection.Find.MatchAllWordForms = False
      .Selection.Find.MatchByte = True
      .Selection.Find.Execute
      iCols = .Selection.Tables(1).Columns.Count
      iRows = .Selection.Tables(1).Rows.Count
      If iCols > 0 Then
         SetDataListWidth iCols, iRows
         For iRow = 1 To iRows
            If iRow > 5 And iRow <= iRows - 5 Then
               grdDataList.TopRow = iRow
            End If
            .Selection.Tables(1).Rows(iRow).Select
            For iCol = 1 To iCols
               .Selection.SelectRow
               .Selection.Cells(iCol).Select
               grdDataList.TextMatrix(iRow - 1, iCol - 1) = .Selection.Text
            Next
         Next
      End If
   End With
   g_WordAp.Quit wdDoNotSaveChanges
   Set g_WordAp = Nothing
   MsgBox "匯入完成！"
   Screen.MousePointer = vbDefault
End Sub

Private Sub SetDataListWidth(pCols As Integer, pRows As Integer)
   Dim iCol As Integer
   With grdDataList
   .Visible = False
   .Clear
   .Rows = pRows: .Cols = pCols: .FixedRows = 1: .FixedCols = 0
   .row = 0
   For iCol = 0 To .Cols - 1
   .col = iCol
   .ColAlignment(.col) = flexAlignCenterCenter
   .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
   Next
   .Visible = True
   End With
End Sub
'Added by Morgan 2019/10/31
'108台灣設計案專用期延長補掛NP年費期限
Private Sub Command5_Click()
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim bInTrans As Boolean
   Dim iRecCnt As Long
   Dim stNP01 As String, stNP02 As String, stNP03 As String, stNP04 As String, stnp05 As String
   Dim stNP08 As String, stNP09 As String, stNP10 As String, stNP15 As String, stNP22 As String
   Dim stPA26 As String, stPA75 As String
   Dim stPA27 As String, stPA28 As String, stPA29 As String, stPA30 As String 'Added by Lydia 2022/08/02
   
On Error GoTo ErrHnd

   If strUserNum <> "QPGMR" Then MsgBox "要先切換 User 為 QPGMR!!", vbCritical: Exit Sub

   '53筆
   'Modified by Lydia 2022/08/02 +pa27~pa30
   stSQL = "select pa01,pa02,pa03,pa04,pa26,pa75,cp09,pa27,pa28,pa29,pa30" & _
      ",to_char(to_date(pa14+LASTYEAR(pa72)*10000,'yyyymmdd')-1,'yyyymmdd') np09" & _
      " From patent a,caseprogress b where pa09='000' and pa08='3' and pa25>=20191101 and pa24<20191101" & _
      " and pa14+LASTYEAR(pa72)*10000>pa25 and pa57 is null" & _
      " and not exists(select * from nextprogress where np02=pa01 and np03=pa02 and np04=pa03" & _
      " and np05=pa04 and np07='605' and np09=to_char(to_date(pa14+LASTYEAR(pa72)*10000,'yyyymmdd')-1,'yyyymmdd'))" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10(+)='605'" & _
      " and not exists(select * from caseprogress x where x.cp01=b.cp01 and x.cp02=b.cp02 and x.cp03=b.cp03" & _
      " and x.cp04=b.cp04 and x.cp10=b.cp10 and x.cp27>b.cp27)" & _
      " order by 1,2,3,4"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Label4.Caption = "新增年費期限中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         cnnConnection.BeginTrans
         bInTrans = True
         stNP01 = .Fields("cp09")
         stNP02 = .Fields("pa01")
         stNP03 = .Fields("pa02")
         stNP04 = .Fields("pa03")
         stnp05 = .Fields("pa04")
         stNP09 = .Fields("np09")
         stNP22 = GetNextProgressNo
         stPA26 = "" & .Fields("pa26")
         stPA75 = "" & .Fields("pa75")
         'Added by Lydia 2022/08/02
         stPA27 = "" & .Fields("pa27")
         stPA28 = "" & .Fields("pa28")
         stPA29 = "" & .Fields("pa29")
         stPA30 = "" & .Fields("pa30")
         'end 2022/08/02
         
         If .Fields("pa01") = "FCP" Then
            stNP08 = PUB_GetFCPOurDeadline(stNP09, 2)
            stNP10 = PUB_GetFCPSalesNo(stNP02, stNP03, stNP04, stnp05)
            'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
            'stNP15 = PUB_GetNpMemo(stNP02 & stNP03 & stNP04 & stnp05, "605", stPA75, stPA26)
            stNP15 = PUB_GetNpMemo2("1", stNP02 & stNP03 & stNP04 & stnp05, "605", stPA75, stPA26 & "," & stPA27 & "," & stPA28 & "," & stPA29 & "," & stPA30)
         Else
            stNP08 = PUB_GetOurDeadline(stNP09)
            stNP10 = PUB_GetAKindSalesNo(stNP02, stNP03, stNP04, stnp05)
            stNP15 = ""
         End If
         stSQL = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22) " & _
               "VALUES ('" & stNP01 & "','" & stNP02 & "','" & stNP03 & "','" & stNP04 & _
               "','" & stnp05 & "',605," & stNP08 & "," & stNP09 & ",'" & stNP10 & "','" & ChgSQL(stNP15) & "'," & stNP22 & ")"
         cnnConnection.Execute stSQL, intQ
         cnnConnection.CommitTrans
         bInTrans = False
         iRecCnt = iRecCnt + 1
         .MoveNext
      Loop
      End With
      
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆!!"
   Else
      MsgBox "無案件須新增年費期限!!", vbExclamation
   End If
   
ErrHnd:
   
   If Err.Number <> 0 Then
      If bInTrans Then cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   Set rsQuery = Nothing
End Sub

Private Sub Command6_Click()
   Dim stSQL As String, iQ As Integer, ii As Integer
   Dim RsQ As ADODB.Recordset
   Dim xlsFileName As String
   Dim xlsApp, xlsWks
   'Dim xlsApp As New Excel.Application
   'Dim xlsWks As Worksheet
   Dim stTmp, intWidth
   
On Error GoTo ErrHnd
    
   Select Case Combo1.ListIndex
      Case 0 'Columbia
         
         xlsFileName = PUB_Getdesktop & "\X69455_" & strSrvDate(1) & ".xls"
         If Dir(xlsFileName) <> "" Then Kill xlsFileName
         
         stSQL = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) c01" & _
            ",pa77 c02,decode(pa09,'000',cpm03,cpm04) c03,decode(cp27,null,sqldatew(cp07)) c04" & _
            ",decode(cp27,null,'未發文','未請款') c06" & _
            " from patent,caseprogress,casepropertymap m1" & _
            " where instr(pa26||pa27||pa28||pa29||pa30,'X69455000')>0" & _
            " and pa57||pa108 is null" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp159=0" & _
            " and cpm01(+)=cp01 and cpm02(+)=cp10" & _
            " and (cp158=0 or (cp27>19221111 and cp20||cp60 is null and ( cp01<>'P' or cp61 is not null" & _
            " or (cp01='P' and exists(select * from casepropertymap m2 where cpm01='FCP' and m2.cpm02=cp10 and m2.cpm18>0))))" & _
            ") order by 1"
         
         stTmp = Array("本所案號", "彼號", "案件性質", "法定期限", "狀態")
         intWidth = Array(0, 0, 0, 0, 0)
        
      Case 1
      
      Case 2
      Case 3
   End Select
   If stSQL <> "" Then
      iQ = 1
      Set RsQ = ClsLawReadRstMsg(iQ, stSQL)
      If iQ = 1 Then
         Set xlsApp = CreateObject("Excel.Application")
         With xlsApp
         .SheetsInNewWorkbook = 1
         .Workbooks.add
         Set xlsWks = .Worksheets(1)
         
         '.Visible = True
         iQ = 1
         Do While Not RsQ.EOF
            iQ = iQ + 1
            For ii = LBound(stTmp) To UBound(stTmp)
               xlsWks.Range(Chr(ii + 65) & iQ).Value = "" & RsQ.Fields(ii)
            Next ii
            RsQ.MoveNext
         Loop
         
         '設定欄位名稱及欄寬
         xlsWks.Rows("1:1").Select
         .Selection.Font.Bold = True
         For ii = LBound(stTmp) To UBound(stTmp)
             xlsWks.Range(Chr(ii + 65) & "1").Value = stTmp(ii)
             If intWidth(ii) > 0 Then
                xlsWks.Columns(Chr(ii + 65) & ":" & Chr(ii + 65)).ColumnWidth = intWidth(ii)
             Else
                xlsWks.Columns(Chr(ii + 65) & ":" & Chr(ii + 65)).EntireColumn.AutoFit
             End If
         Next ii

         If Val(.Version) < 12 Then
            .Workbooks(1).SaveAs FileName:=xlsFileName, FileFormat:=-4143
         Else
            .Workbooks(1).SaveAs FileName:=xlsFileName, FileFormat:=56
         End If
         
         If MsgBox("檔案已匯出！是否開啟？" & vbCrLf & vbCrLf & xlsFileName, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            .Visible = True
         End If
         End With
         
      End If
   End If
   
ErrHnd:

On Error Resume Next
   If TypeName(xlsApp) = "Application" Then
      If xlsApp.Visible = False Then
         xlsApp.Workbooks(1).Close savechanges:=False
         xlsApp.Quit
      End If
   End If
   
   Set xlsApp = Nothing
   Set xlsWks = Nothing
   Set RsQ = Nothing
End Sub

Private Sub SetCombo1()
   Combo1.Clear
   Combo1.AddItem "Columbia(X69455) 已收未發文或已發未請款", 0
   'Combo1.AddItem "Metis IP(Y54339、Y54339B10) 每個月月底", 1
   
End Sub

Private Sub Command7_Click()
   Dim rsQuery As ADODB.Recordset
      
   strExc(0) = "select a1k01,a1k32 from acc1k0 where a1k29 is null and a1k12 is null and a1k25 is null and a1k33='4' and a1k38 is null"
   'strExc(0) = strExc(0) & " and a1k32='Y'"
   strExc(0) = strExc(0) & " and a1k32 is null"
   strExc(0) = strExc(0) & " order by a1k01 desc"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With rsQuery
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         'PUB_PrintBill .Fields("a1k01"), Printer.DeviceName, False, False, , , 1, "1"
         Load Frmacc2480
         With Frmacc2480
         .Visible = False
         .Text1.Text = rsQuery("a1k01")
         .Text2.Text = .Text1.Text
'         .Combo1.Text = Printer.DeviceName
         .m_bEditDoc = True
         .Command2_Click
         g_WordAp.Quit wdDoNotSaveChanges
         End With
         Unload Frmacc2480
         If .AbsolutePosition Mod 20 = 0 Then
            If MsgBox("繼續？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
               Exit Do
            End If
         End If
         .MoveNext
      Loop
      End With
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   SetCombo1
End Sub

'專利修法預繳年費退費通知
Private Sub Process1()
   Dim stCon As String, stLstCustNo As String, stET02 As String
   Dim strTmp As String, StrCaseList(5) As String, iNo As Integer, idx As Integer
   Dim stT09 As String, stT10 As String, stT11 As String, stT13 As String, stT16 As String, stLY As String
   Dim bolSave As Boolean, iSNo As Integer, stT18 As String
   Dim strUserNo As String
   
   stCon = ""
   If txtCust <> "" Then
      stCon = " and pa26='" & txtCust & "'"
   End If
   
   If Text1 = "1" Then
      stCon = stCon & " and t15+(t09-1)*10000>=20110701"
      bolSave = True
   End If
   
   If Text1 = "2" Then
      stCon = stCon & " and t15+(t09-1)*10000<20110701"
      bolSave = True
   End If
   
   strExc(0) = "select t01,t02,t03,t04,t09,t10,t11,t13,t16,pa05,pa11,pa26,lastyear(t06) LY,t18 from t100,patent,staff,customer" & _
      " where t01='P' and t14 is null and t13>0 and pa01(+)=t01 and pa02(+)=t02 and pa03(+)=t03 and pa04(+)=t04" & stCon & _
      " and st01(+)=t17 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " order by st15,t17,cu04,pa26,t01,t02,t03,t04"
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strUserNo = strUserNum
      strUserNum = "81002" '用程序編號跑才會有發文字且方便列印及維護--郭說用"81002" 100/6/29
      With RsTemp
      stLstCustNo = "" & .Fields("pa26")
      stET02 = "" & .Fields("t01") & .Fields("t02") & .Fields("t03") & .Fields("t04") & "&000"
      Erase StrCaseList
      iNo = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      iSNo = 1
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         If .Fields("pa26") <> stLstCustNo Then
            '只有1案
            If iNo = 1 Then
               EndLetter "21", stET02, "03", strUserNum
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                  "','已繳年度' ,'" & stLY & "')"
               cnnConnection.Execute strExc(0), intI
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                  "','可退年度' ,'" & stT16 & "')"
               cnnConnection.Execute strExc(0), intI
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                  "','已繳規費' ,'" & stT11 & "')"
               cnnConnection.Execute strExc(0), intI
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                  "','可退年度起年' ,'" & stT09 & "')"
               cnnConnection.Execute strExc(0), intI
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                  "','可退年度迄年' ,'" & stT10 & "')"
               cnnConnection.Execute strExc(0), intI
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                  "','可退規費' ,'" & stT13 & "')"
               cnnConnection.Execute strExc(0), intI
               
               If Not bolSave Then
                  If stT18 <> "" Then
                     strTmp = stT18
                  Else
                     strTmp = Format(iSNo, "0000")
                  End If
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                     "','發文流水號' ,'" & strTmp & "')"
                  cnnConnection.Execute strExc(0), intI
               End If
               
               NowPrint stET02, "21", "03", False, strUserNum, , , , , 1, , , bolSave
               iSNo = iSNo + 1
            Else
               '最後一頁要保留兩筆否則跳頁印
               If iNo Mod 8 = 7 Or iNo Mod 8 = 0 Then
                  StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
                     "  本所案號   申請號    案件名稱            預繳年度        可退金額" & vbCrLf & _
                     "  --------------------------------------------------------------------------------------------------" & vbCrLf
               End If
               StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
               
               EndLetter "21", stET02, "01", strUserNum
               For idx = 1 To 5
                  If StrCaseList(idx) <> "" Then
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                        "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                     cnnConnection.Execute strExc(0), intI
                  End If
               Next
               
               If Not bolSave Then
                  If stT18 <> "" Then
                     strTmp = stT18
                  Else
                     strTmp = Format(iSNo, "0000")
                  End If
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                     "','發文流水號' ,'" & strTmp & "')"
                  cnnConnection.Execute strExc(0), intI
               End If
               
               NowPrint stET02, "21", "01", False, strUserNum, , , , , 1, , , bolSave
               iSNo = iSNo + 1
            End If
            
'            If iNo > 60 Then
'               NowPrint stET02, "21", "01", True, strUserNum, , , , , 1
'               Debug.Print stET02
'            End If
            
            stLstCustNo = "" & .Fields("pa26")
            stET02 = "" & .Fields("t01") & .Fields("t02") & .Fields("t03") & .Fields("t04") & "&000"
            Erase StrCaseList
            iNo = 0
         Else
         
            '跳頁控制
            If iNo > 0 And iNo Mod 8 = 0 Then
               StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
                  "  本所案號   申請號    案件名稱            預繳年度        可退金額" & vbCrLf & _
                  "  --------------------------------------------------------------------------------------------------" & vbCrLf
            End If
            StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
         End If
         
         stT09 = "" & .Fields("t09")
         stT10 = "" & .Fields("t10")
         stT11 = "" & .Fields("t11")
         stT13 = "" & .Fields("t13")
         stT16 = "" & .Fields("t16")
         stT18 = "" & .Fields("t18")
         stLY = "" & .Fields("LY")
         
         '案件清單
         iNo = iNo + 1
         idx = 1 + iNo \ 16
         strTmp = .Fields("t01") & "-" & .Fields("t02") & IIf(.Fields("t03") & .Fields("t04") = "000", "", "-" & .Fields("t03") & "-" & .Fields("t04"))
         StrCaseList(0) = Format(iNo, "@@") & "." & PUB_StrToStr(strTmp, 9, True) '本所案號
         strTmp = "" & .Fields("pa11")
         StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 8, True) '申請號
         strTmp = "" & .Fields("pa05")
         StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 18, True) '案件名稱
         'Modify by Morgan 2010/3/15 改列舉
         'If .Fields("t09") = .Fields("t10") Then
         '   strTmp = "   " & .Fields("t09")
         'Else
         '   strTmp = "  " & .Fields("t09") & "-" & .Fields("t10")
         'End If
         strTmp = stT16
         'end 2010/3/15
         StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 14, True) '預繳年度
         strTmp = Format("" & .Fields("t13"), DDollar)
         StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 8, True, True) '可退金額
         StrCaseList(0) = StrCaseList(0) & vbCrLf & _
            vbCrLf & "□ 本人／本公司同意將差額款項移作下一年年費。" & _
            vbCrLf & "□ 本人／本公司同意辦理差額退費，並將款項直接退本人／本公司。" & vbCrLf & vbCrLf
         
         If stT18 = "" Then
            strSql = "update t100 set t18='" & Format(iSNo, "0000") & "' where t01='" & .Fields("t01") & "' and t02='" & .Fields("t02") & "' and t03='" & .Fields("t03") & "' and t04='" & .Fields("t04") & "'"
            cnnConnection.Execute strSql, intI
         End If
         .MoveNext
      Loop
      
      '最後一頁要保留兩筆否則跳頁印
      If iNo Mod 8 = 7 Or iNo Mod 8 = 0 Then
         StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
            "  本所案號   申請號    案件名稱            預繳年度        可退金額" & vbCrLf & _
            "  --------------------------------------------------------------------------------------------------" & vbCrLf
      End If
      StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
      '只有1案
      If iNo = 1 Then
         EndLetter "21", stET02, "03", strUserNum
         strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
            "','已繳年度' ,'" & stLY & "')"
         cnnConnection.Execute strExc(0), intI
         
         strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
            "','可退年度' ,'" & stT16 & "')"
         cnnConnection.Execute strExc(0), intI
         
         strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
               "','已繳規費' ,'" & stT11 & "')"
         cnnConnection.Execute strExc(0), intI
         
         strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
            "','可退年度起年' ,'" & stT09 & "')"
         cnnConnection.Execute strExc(0), intI
         
         strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
            "','可退年度迄年' ,'" & stT10 & "')"
         cnnConnection.Execute strExc(0), intI
         
         strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
            "','可退規費' ,'" & stT13 & "')"
         cnnConnection.Execute strExc(0), intI
         
         If Not bolSave Then
            If stT18 <> "" Then
               strTmp = stT18
            Else
               strTmp = Format(iSNo, "0000")
            End If
            strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
               "','發文流水號' ,'" & strTmp & "')"
            cnnConnection.Execute strExc(0), intI
         End If
         
         NowPrint stET02, "21", "03", False, strUserNum, , , , , 1, , , bolSave
      Else
         EndLetter "21", stET02, "01", strUserNum
         For idx = 1 To 5
            If StrCaseList(idx) <> "" Then
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                  "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
               cnnConnection.Execute strExc(0), intI
            End If
         Next
         
         If Not bolSave Then
            If stT18 <> "" Then
               strTmp = stT18
            Else
               strTmp = Format(iSNo, "0000")
            End If
            strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
               "','發文流水號' ,'" & strTmp & "')"
            cnnConnection.Execute strExc(0), intI
         End If
         
         NowPrint stET02, "21", "01", False, strUserNum, , , , , 1, , , bolSave
      End If
      End With
      strUserNum = strUserNo
      MsgBox "定稿已產生完畢！"
   Else
      MsgBox "無資料可作業！"
   End If
End Sub

'Added by Morgan 2019/10/8 108台灣設計P案專用期延長通知
Private Sub Process6()
   
   Dim stCon As String, stLstCustNo As String, stLstFagentNo As String, stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP12 As String, stCP13 As String, stCaseNo As String
   Dim iRecCnt As Integer
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
         
   'Modified by Morgan 2019/10/21 要切系統的 User (因 LP 的判發人會抓 CP 的發文人)
   'strUserNo = strUserNum
   'strUserNum = "A3014" '用程序編號跑才會有發文字且方便列印及維護
   If strUserNum <> "A3014" Then MsgBox "要先切換 User 為 A3014!!", vbCritical: Exit Sub
         
   stCon = ""
   If txtCust <> "" Then
      stCon = " and pa26='" & txtCust & "'"
   End If
   
stCon = " and pa75 is not null" '重跑大->台定稿
   
   '未閉卷且下次繳費日未逾期
   'Modified by Morgan 2019/10/25 +補 pa23='1' 條件 -->P-117059 誤通知
   strExc(0) = "select * from patent,customer,staff" & _
      " where pa01='P' and pa08='3' and pa09='000' and pa24<20191101 and pa25>=20191101 and pa23='1'" & _
      " and pa57 is null and (pa14+LASTYEAR(pa72)*10000)>to_char(sysdate,'yyyymmdd')" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and st01(+)=cu13" & stCon & _
      " ORDER BY CU12,CU13,PA26,PA09,PA01,PA02,PA03,PA04"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Label4.Caption = "本所信函進度檔建立中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         '先收文本所信函1999
         stSQL = "update caseprogress set cp28=cp28 where cp01='" & .Fields("pa01") & "' and cp02='" & .Fields("pa02") & "' and cp03='" & .Fields("pa03") & "' and cp04='" & .Fields("pa04") & "' and instr(cp64,'設計延長專用期間通知')>0"
         cnnConnection.Execute stSQL, intI
         If intI = 0 Then
            stCP13 = PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
            stCP12 = GetSalesArea(stCP13)
            
            cnnConnection.BeginTrans
            bolInTrans = True
            
            stCP09 = AutoNo("D", 6)
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp64 ) values ('" & .Fields("pa01") & "'" & _
               ",'" & .Fields("pa02") & "','" & .Fields("pa03") & "','" & .Fields("pa04") & "'," & strSrvDate(1) & _
               ",'" & stCP09 & "','1999','" & stCP12 & "'" & _
               ",'" & stCP13 & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N','設計延長專用期間通知(正常)')"
            cnnConnection.Execute stSQL, intI
            
            PUB_AddLetterProgress stCP09, 0, True, , True, "" & .Fields("pa26"), "1999", "" & .Fields("pa75"), , , True

            cnnConnection.CommitTrans
            bolInTrans = False
            iRecCnt = iRecCnt + 1
         End If

         .MoveNext
      Loop
      End With
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
      
      If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
      
      strExc(0) = "select cp09,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo,pa26,pa75" & _
         " from caseprogress,patent,customer" & _
         " where cp05>20191000 and cp01='P' and cp10='1999' and instr(cp64,'設計延長專用期間通知(正常)')>0" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & stCon & _
         " ORDER BY cp12,cp13,PA26,PA09,PA01,PA02,PA03,PA04"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         
         stLstCustNo = "" & .Fields("pa26")
         stLstFagentNo = "" & .Fields("pa75")
         stCP09 = .Fields("cp09")
         stCaseNo = .Fields("CNo")
         iRecCnt = 0
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            If stLstCustNo <> "" & .Fields("pa26") Then
               If iRecCnt = 1 Then
                  If stLstFagentNo <> "" Then
                     stET03 = "07" '一案
                  Else
                     stET03 = "01" '一案
                  End If
               Else
                  If stLstFagentNo <> "" Then
                     stET03 = "08" '多案
                  Else
                     stET03 = "02" '多案
                  End If
               End If
               
               EndLetter "21", stCP09, stET03, strUserNum
               NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09
               
               stLstCustNo = "" & .Fields("pa26")
               stLstFagentNo = "" & .Fields("pa75")
               stCP09 = .Fields("cp09")
               stCaseNo = .Fields("CNo")
               iRecCnt = 0
            End If
            
            iRecCnt = iRecCnt + 1
            
            If stCP09 <> .Fields("cp09") Then
               strExc(1) = "已併入" & stCaseNo & "案(" & stCP09 & ")告知客戶"
               stSQL = "update letterprogress set lp06='" & strUserNum & "',lp07=to_char(sysdate,'yyyymmdd'),lp10='N',lp12='" & strExc(1) & "',lp42='" & stCP09 & "'" & _
                  " where lp01='" & .Fields("cp09") & "'"
               cnnConnection.Execute stSQL, intI
            End If
            
            .MoveNext
         Loop
         
         If iRecCnt = 1 Then
            If stLstFagentNo <> "" Then
               stET03 = "07" '一案
            Else
               stET03 = "01" '一案
            End If
         Else
            If stLstFagentNo <> "" Then
               stET03 = "08" '多案
            Else
               stET03 = "02" '多案
            End If
         End If
         EndLetter "21", stCP09, stET03, strUserNum
         NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09
         End With
         Label4.Caption = "通知信產生完成！"
      End If
   End If
   
ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
End Sub

'Added by Morgan 2019/10/14 108台灣設計P案專用期延長通知(已消滅)
Private Sub Process7()
   
   Dim stCon As String, stLstCustNo As String, stLstFagentNo As String, stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP12 As String, stCP13 As String, stCaseNo As String
   Dim iRecCnt As Integer
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
         
   'Modified by Morgan 2019/10/21 要切系統的 User (因 LP 的判發人會抓 CP 的發文人)
   'strUserNo = strUserNum
   'strUserNum = "A3014" '用程序編號跑才會有發文字且方便列印及維護
   If strUserNum <> "A3014" Then MsgBox "要先切換 User 為 A3014!!", vbCritical: Exit Sub
   
   stCon = ""
   If txtCust <> "" Then
      stCon = " and pa26='" & txtCust & "'"
   End If
   
stCon = " and pa75 is not null" '重跑大->台定稿

   '已閉卷且下次繳費日>=20180501
   '１．下次繳費期限在１０７年５月１日以後，已閉卷並收到專利權消滅通知
   strExc(0) = "select * from patent,customer,staff" & _
      " where pa01='P' and pa08='3' and pa09='000' and pa24<20191101 and pa25>=20191101 and pa23='1'" & _
      " and pa57 is not null and pa14+LASTYEAR(pa72)*10000>20180501" & _
      " and exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='1604')" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and st01(+)=cu13" & stCon & _
      " ORDER BY CU12,CU13,PA26,PA09,PA01,PA02,PA03,PA04"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Label4.Caption = "本所信函進度檔建立中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         '先收文本所信函1999
         stSQL = "update caseprogress set cp28=cp28 where cp01='" & .Fields("pa01") & "' and cp02='" & .Fields("pa02") & "' and cp03='" & .Fields("pa03") & "' and cp04='" & .Fields("pa04") & "' and instr(cp64,'設計延長專用期間通知')>0"
         cnnConnection.Execute stSQL, intI
         If intI = 0 Then
            stCP13 = PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
            stCP12 = GetSalesArea(stCP13)
            
            cnnConnection.BeginTrans
            bolInTrans = True
            
            stCP09 = AutoNo("D", 6)
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp64 ) values ('" & .Fields("pa01") & "'" & _
               ",'" & .Fields("pa02") & "','" & .Fields("pa03") & "','" & .Fields("pa04") & "'," & strSrvDate(1) & _
               ",'" & stCP09 & "','1999','" & stCP12 & "'" & _
               ",'" & stCP13 & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N','設計延長專用期間通知(已消滅)')"
            cnnConnection.Execute stSQL, intI
            
            PUB_AddLetterProgress stCP09, 0, True, , True, "" & .Fields("pa26"), "1999", "" & .Fields("pa75"), , , True

            cnnConnection.CommitTrans
            bolInTrans = False
            iRecCnt = iRecCnt + 1
         End If

         .MoveNext
      Loop
      End With
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
      
      If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
      
      strExc(0) = "select cp09,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo,pa26,pa75" & _
         " from caseprogress,patent,customer" & _
         " where cp05>20191000 and cp01='P' and cp10='1999' and instr(cp64,'設計延長專用期間通知(已消滅)')>0" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & stCon & _
         " ORDER BY cp12,cp13,PA26,PA09,PA01,PA02,PA03,PA04"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         
         stLstCustNo = "" & .Fields("pa26")
         stLstFagentNo = "" & .Fields("pa75")
         stCP09 = .Fields("cp09")
         stCaseNo = .Fields("CNo")
         iRecCnt = 0
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            If stLstCustNo <> "" & .Fields("pa26") Then
               If iRecCnt = 1 Then
                  If stLstFagentNo <> "" Then
                     stET03 = "09" '一案
                  Else
                     stET03 = "03" '一案
                  End If
               Else
                  If stLstFagentNo <> "" Then
                     stET03 = "10" '多案
                  Else
                     stET03 = "04" '多案
                  End If
               End If
               EndLetter "21", stCP09, stET03, strUserNum
               NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09
               
               stLstCustNo = "" & .Fields("pa26")
               stLstFagentNo = "" & .Fields("pa75")
               stCP09 = .Fields("cp09")
               stCaseNo = .Fields("CNo")
               iRecCnt = 0
            End If
            
            iRecCnt = iRecCnt + 1
            
            If stCP09 <> .Fields("cp09") Then
               strExc(1) = "已併入" & stCaseNo & "案(" & stCP09 & ")告知客戶"
               stSQL = "update letterprogress set lp06='" & strUserNum & "',lp07=to_char(sysdate,'yyyymmdd'),lp10='N',lp12='" & strExc(1) & "',lp42='" & stCP09 & "'" & _
                  " where lp01='" & .Fields("cp09") & "'"
               cnnConnection.Execute stSQL, intI
            End If
            
            .MoveNext
         Loop
         
         If iRecCnt = 1 Then
            If stLstFagentNo <> "" Then
               stET03 = "09" '一案
            Else
               stET03 = "03" '一案
            End If
         Else
            If stLstFagentNo <> "" Then
               stET03 = "10" '多案
            Else
               stET03 = "04" '多案
            End If
         End If
         
         EndLetter "21", stCP09, stET03, strUserNum
         NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09
         End With
         
         Label4.Caption = "通知信產生完成！"
      End If
   End If
   
ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
End Sub

'Added by Morgan 2019/10/18 108台灣設計P案專用期延長通知(未消滅，最後一年未繳)
Private Sub Process8()
   
   Dim stCon As String, stLstCustNo As String, stLstFagentNo As String, stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP12 As String, stCP13 As String, stCaseNo As String
   Dim iRecCnt As Integer
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
         
   'Modified by Morgan 2019/10/21 要切系統的 User (因 LP 的判發人會抓 CP 的發文人)
   'strUserNo = strUserNum
   'strUserNum = "A3014" '用程序編號跑才會有發文字且方便列印及維護
   If strUserNum <> "A3014" Then MsgBox "要先切換 User 為 A3014!!", vbCritical: Exit Sub
         
   stCon = ""
   If txtCust <> "" Then
      stCon = " and pa26='" & txtCust & "'"
   End If
      
   '２．下次繳費期限在１０７年５月１日以後，繳費年度為最後一年，已閉卷，尚未收到專利權消滅通知
   strExc(0) = "select * from patent,customer,staff" & _
      " where pa01='P' and pa08='3' and pa09='000' and pa24<20191101 and pa25>=20191101 and pa23='1'" & _
      " and pa57 is not null and (pa14+(LASTYEAR(pa72)+1)*10000)>pa25 and (pa14+(LASTYEAR(pa72))*10000)<pa25" & _
      " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='1604')" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and st01(+)=cu13" & stCon & _
      " ORDER BY CU12,CU13,PA26,PA09,PA01,PA02,PA03,PA04"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Label4.Caption = "本所信函進度檔建立中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         '先收文本所信函1999
         stSQL = "update caseprogress set cp28=cp28 where cp01='" & .Fields("pa01") & "' and cp02='" & .Fields("pa02") & "' and cp03='" & .Fields("pa03") & "' and cp04='" & .Fields("pa04") & "' and instr(cp64,'設計延長專用期間通知')>0"
         cnnConnection.Execute stSQL, intI
         If intI = 0 Then
            stCP13 = PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
            stCP12 = GetSalesArea(stCP13)
            
            cnnConnection.BeginTrans
            bolInTrans = True
            
            stCP09 = AutoNo("D", 6)
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp64 ) values ('" & .Fields("pa01") & "'" & _
               ",'" & .Fields("pa02") & "','" & .Fields("pa03") & "','" & .Fields("pa04") & "'," & strSrvDate(1) & _
               ",'" & stCP09 & "','1999','" & stCP12 & "'" & _
               ",'" & stCP13 & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N','設計延長專用期間通知(最後一年未繳)')"
            cnnConnection.Execute stSQL, intI
            
            PUB_AddLetterProgress stCP09, 0, True, , True, "" & .Fields("pa26"), "1999", "" & .Fields("pa75"), , , True

            cnnConnection.CommitTrans
            bolInTrans = False
            iRecCnt = iRecCnt + 1
         End If

         .MoveNext
      Loop
      End With
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
      
      If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
      
      strExc(0) = "select cp09,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo,pa26,pa75" & _
         " from caseprogress,patent,customer" & _
         " where cp05>20191000 and cp01='P' and cp10='1999' and instr(cp64,'設計延長專用期間通知(最後一年未繳)')>0" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & stCon & _
         " ORDER BY cp12,cp13,PA26,PA09,PA01,PA02,PA03,PA04"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         
         stLstCustNo = "" & .Fields("pa26")
         stLstFagentNo = "" & .Fields("pa75")
         stCP09 = .Fields("cp09")
         stCaseNo = .Fields("CNo")
         iRecCnt = 0
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            If stLstCustNo <> "" & .Fields("pa26") Then
               If iRecCnt = 1 Then
                  If stLstFagentNo <> "" Then
                     stET03 = "11" '一案
                  Else
                     stET03 = "05" '一案
                  End If
               Else
                  If stLstFagentNo <> "" Then
                     stET03 = "12" '多案
                  Else
                     stET03 = "06" '多案
                  End If
               End If
               EndLetter "21", stCP09, stET03, strUserNum
               NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09
               
               '不印,玲玲要自行處理
               stSQL = "update letterdemand set ld16='*' where ld18='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intI
               
               stLstCustNo = "" & .Fields("pa26")
               stLstFagentNo = "" & .Fields("pa75")
               stCP09 = .Fields("cp09")
               stCaseNo = .Fields("CNo")
               iRecCnt = 0
            End If
            
            iRecCnt = iRecCnt + 1
            
            If stCP09 <> .Fields("cp09") Then
               strExc(1) = "已併入" & stCaseNo & "案(" & stCP09 & ")告知客戶"
               stSQL = "update letterprogress set lp06='" & strUserNum & "',lp07=to_char(sysdate,'yyyymmdd'),lp10='N',lp12='" & strExc(1) & "',lp42='" & stCP09 & "'" & _
                  " where lp01='" & .Fields("cp09") & "'"
               cnnConnection.Execute stSQL, intI
            End If
            
            .MoveNext
         Loop
               
         If iRecCnt = 1 Then
            If stLstFagentNo <> "" Then
               stET03 = "11" '一案
            Else
               stET03 = "05" '一案
            End If
         Else
            If stLstFagentNo <> "" Then
               stET03 = "12" '多案
            Else
               stET03 = "06" '多案
            End If
         End If
         EndLetter "21", stCP09, stET03, strUserNum
         NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09
         
         '不印,玲玲要自行處理
         stSQL = "update letterdemand set ld16='*' where ld18='" & stCP09 & "'"
         cnnConnection.Execute stSQL, intI
         End With
         
         Label4.Caption = "通知信產生完成！"
      End If
   End If
   
ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
End Sub

'Added by Morgan 2020/8/21 109減免退費通知
Private Sub Process9()
   
   Dim stCon As String, stLstCustNo As String, stLstFagentNo As String, stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP12 As String, stCP13 As String, stCaseNo As String
   Dim rsQuery As ADODB.Recordset
   Dim strTmp As String, StrCaseList(5) As String, idx As Integer, strHead As String, strTail As String
   Dim iRecCnt As Integer, stLstYear As String, stRtnYear As String, stRtnFee As String
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
         
   '要切系統的 User (因 LP 的判發人會抓 CP 的發文人且用程序的員工號跑才會有發文字)
   If strUserNum <> "A3014" Then MsgBox "要先切換 User 為 A3014!!", vbCritical: Exit Sub
         
   stCon = ""
   If txtCust <> "" Then
      stCon = " and pa26='" & txtCust & "'"
   End If
   
   strExc(0) = "select * from T109,patent,customer,staff" & _
      " where T20='Y' and pa01(+)=T01 and pa02(+)=T02 and pa03(+)=T03 and pa04(+)=T04" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and st01(+)=cu13" & stCon & _
      " ORDER BY CU12,CU13,PA26,PA09,PA01,PA02,PA03,PA04"
      
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With rsQuery
      Label4.Caption = "本所信函進度檔建立中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         '先收文本所信函1999
         stSQL = "update caseprogress set cp28=cp28 where cp01='" & .Fields("pa01") & "' and cp02='" & .Fields("pa02") & "' and cp03='" & .Fields("pa03") & "' and cp04='" & .Fields("pa04") & "' and instr(cp64,'109減免退費通知')>0"
         cnnConnection.Execute stSQL, intI
         If intI = 0 Then
            stCP13 = PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
            stCP12 = GetSalesArea(stCP13)
            
            cnnConnection.BeginTrans
            bolInTrans = True
            
            stCP09 = AutoNo("D", 6)
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp64 ) values ('" & .Fields("pa01") & "'" & _
               ",'" & .Fields("pa02") & "','" & .Fields("pa03") & "','" & .Fields("pa04") & "'," & strSrvDate(1) & _
               ",'" & stCP09 & "','1999','" & stCP12 & "'" & _
               ",'" & stCP13 & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N','109減免退費通知')"
            cnnConnection.Execute stSQL, intI
            
            stSQL = "update T109 set T18='" & stCP09 & "' where T01='" & .Fields("pa01") & "' and T02='" & .Fields("pa02") & "' and T03='" & .Fields("pa03") & "' and T04='" & .Fields("pa04") & "'"
            cnnConnection.Execute stSQL, intI
            
            PUB_AddLetterProgress stCP09, 0, True, , , "" & .Fields("pa26"), "1999", "" & .Fields("pa75")

            cnnConnection.CommitTrans
            bolInTrans = False
            iRecCnt = iRecCnt + 1
         End If

         .MoveNext
      Loop
      End With
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
      
      If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
      
      strExc(0) = "select cp09,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo,pa05,pa11,pa26,pa75,lastyear(pa72) LstYr,T09,T10,T13" & _
         " from caseprogress,patent,customer,T109" & _
         " where cp05>20191000 and cp01='P' and cp10='1999' and instr(cp64,'109減免退費通知')>0" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & stCon & _
         " and T01(+)=pa01 and T02(+)=pa02 and T03(+)=pa03 and T04(+)=pa04 ORDER BY cp12,cp13,PA26,PA09,PA01,PA02,PA03,PA04"
      intI = 1
      Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With rsQuery
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         
         stLstCustNo = "" & .Fields("pa26")
         stLstFagentNo = "" & .Fields("pa75")
         stCP09 = .Fields("cp09")
         stCaseNo = .Fields("CNo")
         
         iRecCnt = 0
         strHead = "  本所案號   申請號    案件名稱            可減免退費年度  減免規費" & vbCrLf & _
                  "------------------------------------------------------------------------------------------------------" & vbCrLf
               
         strTail = "□ 本人／本公司同意將減免規費款項於下一年年費扣抵。" & vbCrLf & _
                  "□ 本人／本公司同意辦理退費，並將款項直接退本人／本公司。" & vbCrLf & vbCrLf
                  
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            If stLstCustNo <> "" & .Fields("pa26") Then
               If iRecCnt = 1 Then
                  stET03 = "13" '一案
                  EndLetter "21", stCP09, stET03, strUserNum
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stCP09 & "','" & stET03 & "','" & strUserNum & _
                     "','已繳年度' ,'" & stLstYear & "')"
                  cnnConnection.Execute strExc(0), intI
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stCP09 & "','" & stET03 & "','" & strUserNum & _
                     "','可退費年度' ,'" & stRtnYear & "')"
                  cnnConnection.Execute strExc(0), intI
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stCP09 & "','" & stET03 & "','" & strUserNum & _
                     "','退費金額' ,'" & stRtnFee & "')"
                  cnnConnection.Execute strExc(0), intI
               Else
                  '最後一頁要保留兩筆否則跳頁印
                  If iRecCnt Mod 8 = 7 Or iRecCnt Mod 8 = 0 Then
                     StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & strHead
                  End If
                  StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
               
                  stET03 = "14" '多案
                  EndLetter "21", stCP09, stET03, strUserNum
                  For idx = 1 To 5
                     If StrCaseList(idx) <> "" Then
                        strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('21','" & stCP09 & "','" & stET03 & "','" & strUserNum & _
                           "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                        cnnConnection.Execute strExc(0), intI
                     End If
                  Next
               End If
               NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09
               
               '不印,自行處理
               stSQL = "update letterdemand set ld16='*' where ld18='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intI
               
               stLstCustNo = "" & .Fields("pa26")
               stLstFagentNo = "" & .Fields("pa75")
               stCP09 = .Fields("cp09")
               stCaseNo = .Fields("CNo")
               iRecCnt = 0
               Erase StrCaseList
            Else
               '跳頁控制
               If iRecCnt > 0 Then
                  If iRecCnt Mod 8 = 0 Then
                     StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & strHead
                  End If
                  StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
               End If
            End If
            
            '案件清單
            iRecCnt = iRecCnt + 1
            idx = 1 + iRecCnt \ 8 '第1頁7案,第2頁起8案,最後1頁最多6案(保留簽名行)
            
            If iRecCnt = 1 Then
               StrCaseList(idx) = StrCaseList(idx) & strHead
            End If
            
            stLstYear = .Fields("LstYr")
            stRtnFee = .Fields("T13")
            If .Fields("T09") = .Fields("T10") Then
               stRtnYear = .Fields("T09")
            Else
               stRtnYear = .Fields("T09") & "-" & .Fields("T10")
            End If
            
            strTmp = "" & .Fields("CNo")
            StrCaseList(0) = Format(iRecCnt, "@@") & "." & PUB_StrToStr(strTmp, 9, True) '本所案號
            strTmp = "" & .Fields("pa11")
            StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 8, True) '申請號
            strTmp = "" & .Fields("pa05")
            StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 18, True) '案件名稱
           
            StrCaseList(0) = StrCaseList(0) & "  " & String(7 - Len(stRtnYear) \ 2, " ") & PUB_StrToStr(stRtnYear, 7 + Len(stRtnYear) \ 2, True) '預繳年度
            strTmp = Format("" & .Fields("T13"), DDollar)
            StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 8, True, True) '減免規費
            StrCaseList(0) = StrCaseList(0) & vbCrLf & vbCrLf & strTail
            
            If stCP09 <> .Fields("cp09") Then
               strExc(1) = "已併入" & stCaseNo & "案(" & stCP09 & ")告知客戶"
               stSQL = "update letterprogress set lp06='" & strUserNum & "',lp07=to_char(sysdate,'yyyymmdd'),lp10='N',lp12='" & strExc(1) & "',lp42='" & stCP09 & "'" & _
                  " where lp01='" & .Fields("cp09") & "'"
               cnnConnection.Execute stSQL, intI
            End If
            
            .MoveNext
         Loop
               
         If iRecCnt = 1 Then
            stET03 = "13" '一案
            EndLetter "21", stCP09, stET03, strUserNum
            strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('21','" & stCP09 & "','" & stET03 & "','" & strUserNum & _
               "','已繳年度' ,'" & stLstYear & "')"
            cnnConnection.Execute strExc(0), intI
            strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('21','" & stCP09 & "','" & stET03 & "','" & strUserNum & _
               "','可退費年度' ,'" & stRtnYear & "')"
            cnnConnection.Execute strExc(0), intI
            strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('21','" & stCP09 & "','" & stET03 & "','" & strUserNum & _
               "','退費金額' ,'" & stRtnFee & "')"
            cnnConnection.Execute strExc(0), intI
         Else
            '最後一頁要保留兩筆否則跳頁印
            If iRecCnt Mod 8 = 7 Or iRecCnt Mod 8 = 0 Then
               StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & strHead
            End If
            StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
         
            stET03 = "14" '多案
            EndLetter "21", stCP09, stET03, strUserNum
            For idx = 1 To 5
               If StrCaseList(idx) <> "" Then
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stCP09 & "','" & stET03 & "','" & strUserNum & _
                     "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                  cnnConnection.Execute strExc(0), intI
               End If
            Next
         End If
         NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09
         
         '不印,自行處理
         stSQL = "update letterdemand set ld16='*' where ld18='" & stCP09 & "'"
         cnnConnection.Execute stSQL, intI
         End With
         
         Label4.Caption = "通知信產生完成！"
      End If
   End If
   
ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
End Sub

'Added by Morgan 2020/11/17 109英國脫歐通知
Private Sub Process10()
      
   Dim stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP10 As String, stCP12 As String, stCP13 As String, stNP08 As String, stNP09 As String
   Dim rsQuery As ADODB.Recordset, intQ As Integer, iRecCnt As Integer
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
   
   strUserNo = strUserNum
   
   '重跑商標定稿
   If 1 = 1 Then
   
      stSQL = "select decode(substr(st15,1,1),'S',st06,'1')||CU12||CU13||cu01||cu02||NVL(TM123,CU127) Srt,CP01,CP09,np07,np08,np09" & _
         " from caseprogress,trademark,customer,staff,nextprogress" & _
         " where cp05>=20201118 and cp01='CFT' and cp10='1799' and instr(cp64,'英國脫歐通知')>0" & _
         " and TM01(+)=np02 and TM02(+)=np03 and TM03(+)=np04 and TM04(+)=np05 and TM10='239'" & _
         " and cu01(+)=substr(tm23,1,8) and cu02(+)='0' and st01(+)=cu13" & _
         " and np02(+)=cp01 and np03(+)=cp02 and np04(+)=cp03 and np05(+)=cp04 and np01(+)=cp43 and np22(+)=cp30 and (cp12 like 'S%' or tm44 is null)"
         
      stSQL = stSQL & " order by 1,2,3,4,5"
      
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With rsQuery
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                           
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            stCP09 = .Fields("cp09")
            strUserNum = "78028"
            stET03 = "03"
            
            m_DocSNo = Format(.AbsolutePosition, "000")
            Debug.Print m_DocSNo & " --> " & Now
            NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , True, , True, , , stCP09
            m_DocSNo = ""
            Sleep 3000
            .MoveNext
         Loop
         End With
         Label4.Caption = "通知信產生完成！"
      End If
      
   Else
      
      'CFP延展費期限109/7/1-12/31(15)
      stSQL = "select ST06||CU12||CU13||cu01||cu02||NVL(PA149,CU127) Srt,np02,np03,np04,np05,np07,np08,np09,np01,np22,pa26 CuNo" & _
         " from nextprogress,patent,customer,staff" & _
         " where np09>20200700 and np09<20210000 and np02='CFP' and np07='607' and nvl(np06,'N')='N'" & _
         " and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05 and pa09='239'" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)='0' and st01(+)=cu13"
         
      'CFP延展費(英國)期限,排除歐盟已繳2021以後延展費案件(384)
      stSQL = stSQL & " union all " & _
         " select ST06||CU12||CU13||cu01||cu02||NVL(PA149,CU127) Srt,np02,np03,np04,np05,np07,np08,np09,np01,np22,pa26 CuNo" & _
         " from nextprogress a,patent,customer,staff" & _
         " where np09>20210000 and np02='CFP' and np07='613' and np06 is null" & _
         " and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05 and pa09='239'" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)='0' and st01(+)=cu13" & _
         " and not exists(select * from nextprogress b where np02=a.np02 and np03=a.np03 and np04=a.np04 and np05=a.np05 and np07='607' and np09>a.np09)"
      
      
      'CFT延展(英國)期限,排除已歐盟已繳2021以後延展案件(564)
      stSQL = stSQL & " union all " & _
         " select ST06||CU12||CU13||cu01||cu02||NVL(TM123,CU127) Srt,np02,np03,np04,np05,np07,np08,np09,np01,np22,tm23 CuNo" & _
         " from nextprogress a,trademark,customer,staff" & _
         " where np09>20210000 and np02='CFT' and np07='110' and np06 is null" & _
         " and TM01(+)=np02 and TM02(+)=np03 and TM03(+)=np04 and TM04(+)=np05 and TM10='239'" & _
         " and cu01(+)=substr(tm23,1,8) and cu02(+)='0' and st01(+)=cu13" & _
         " and not exists(select * from nextprogress b where np02=a.np02 and np03=a.np03 and np04=a.np04 and np05=a.np05 and np07='102' and np09>a.np09)"
      stSQL = stSQL & " order by 1,2,3,4,5"
      
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With rsQuery
         Label4.Caption = "本所信函進度檔建立中..."
         iRecCnt = 0
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            '先收文本所信函CFP1999/CFT1799
            stSQL = "update caseprogress set cp28=cp28 where cp01='" & .Fields("np02") & "' and cp02='" & .Fields("np03") & "' and cp03='" & .Fields("np04") & "' and cp04='" & .Fields("np05") & "' and instr(cp64,'英國脫歐通知')>0 and cp05>=20201118"
            cnnConnection.Execute stSQL, intQ
            If intQ = 0 Then
               If .Fields("np02") = "CFP" Then
                  stCP10 = "1999"
                  strUserNum = "99043"
               Else
                  stCP10 = "1799"
                  strUserNum = "78028"
               End If
               stCP13 = PUB_GetAKindSalesNo(.Fields("np02"), .Fields("np03"), .Fields("np04"), .Fields("np05"))
               stCP12 = GetSalesArea(stCP13)
               
               cnnConnection.BeginTrans
               bolInTrans = True
               stCP09 = AutoNo("D", 6)
               stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
                  ",cp12,cp13,cp14,cp20,cp26,cp27,cp30,cp32,cp43,cp64 ) values ('" & .Fields("np02") & "'" & _
                  ",'" & .Fields("np03") & "','" & .Fields("np04") & "','" & .Fields("np05") & "'," & strSrvDate(1) & _
                  ",'" & stCP09 & "','" & stCP10 & "','" & stCP12 & "'" & _
                  ",'" & stCP13 & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'" & .Fields("np22") & "','N','" & .Fields("np01") & "','英國脫歐通知" & IIf(.Fields("np07") = "607", "(未繳)", "") & "')"
               cnnConnection.Execute stSQL, intQ
               
               PUB_AddLetterProgress stCP09, 0, True, , True, .Fields("CuNo"), stCP10
               
               stSQL = "update caseprogress set cp28=cp09,cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intQ
               If intQ = 1 Then
                  'Trigger 會寫發文人故要另外更新
                  stSQL = "update caseprogress set cp154='QPGMR' where cp09='" & stCP09 & "'"
                  cnnConnection.Execute stSQL, intQ
               End If
               '自動確認
               stSQL = "update letterprogress set lp06='QPGMR',lp07=to_char(sysdate,'YYYYMMDD') where lp01='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intQ
               
               cnnConnection.CommitTrans
               bolInTrans = False
               iRecCnt = iRecCnt + 1
            End If
   
            .MoveNext
         Loop
         End With
         Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
         
         If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            Exit Sub
         End If
         
         stSQL = "select decode(substr(st15,1,1),'S',st06,'1')||CU12||CU13||cu01||cu02||NVL(PA149,CU127) Srt,CP01,CP09,np07,np08,np09" & _
            " from caseprogress,patent,customer,staff,nextprogress" & _
            " where cp05>=20201118 and cp01='CFP' and cp10='1999' and instr(cp64,'英國脫歐通知')>0" & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
            " and cu01(+)=substr(pa26,1,8) and cu02(+)='0' and st01(+)=cu13" & _
            " and np02(+)=cp01 and np03(+)=cp02 and np04(+)=cp03 and np05(+)=cp04 and np01(+)=cp43 and np22(+)=cp30"
         
         '排除外對外(自行處理)
         stSQL = stSQL & " union all select decode(substr(st15,1,1),'S',st06,'1')||CU12||CU13||cu01||cu02||NVL(TM123,CU127) Srt,CP01,CP09,np07,np08,np09" & _
            " from caseprogress,trademark,customer,staff,nextprogress" & _
            " where cp05>=20201118 and cp01='CFT' and cp10='1799' and instr(cp64,'英國脫歐通知')>0" & _
            " and TM01(+)=np02 and TM02(+)=np03 and TM03(+)=np04 and TM04(+)=np05 and TM10='239'" & _
            " and cu01(+)=substr(tm23,1,8) and cu02(+)='0' and st01(+)=cu13" & _
            " and np02(+)=cp01 and np03(+)=cp02 and np04(+)=cp03 and np05(+)=cp04 and np01(+)=cp43 and np22(+)=cp30 and (cp12 like 'S%' or tm44 is null)"
            
         stSQL = stSQL & " order by 1,2,3,4,5"
         
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            With rsQuery
            Label4.Caption = "通知信產生中..."
            ProgressBar1.max = .RecordCount
            ProgressBar1.Value = 0
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                              
            Do While Not .EOF
               ProgressBar1.Value = ProgressBar1.Value + 1
               lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
               DoEvents
               
               stCP09 = .Fields("cp09")
               If .Fields("cp01") = "CFP" Then
                  strUserNum = "99043"
                  If .Fields("np07") = "607" Then
                     stET03 = "08"
                     stNP09 = CompDate(1, 6, .Fields("np09"))
                     stNP08 = CompDate(2, -14, stNP09)
                  Else
                     stET03 = "09"
                     stNP09 = .Fields("np09")
                     stNP08 = .Fields("np08")
                  End If
                  EndLetter "21", stCP09, stET03, strUserNum
               
                  stSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stCP09 & "','" & stET03 & "','" & strUserNum & _
                        "','法定期限' ,'" & stNP09 & "')"
                  cnnConnection.Execute stSQL, intQ
                     
                  stSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stCP09 & "','" & stET03 & "','" & strUserNum & _
                        "','本所期限' ,'" & stNP08 & "')"
                  cnnConnection.Execute stSQL, intQ
               Else
                  strUserNum = "78028"
                  stET03 = "03"
               End If
               m_DocSNo = Format(.AbsolutePosition, "000")
               Debug.Print m_DocSNo & " --> " & Now
               NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , True, , True, , , stCP09
               m_DocSNo = ""
               Sleep 3000
               .MoveNext
            Loop
            End With
            
            Label4.Caption = "通知信產生完成！"
         End If
      End If
      
   End If
   
ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
   
End Sub

'Added by Morgan 2020/12/15 109英國脫歐通知2
Private Sub Process12()
      
   Dim stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP10 As String, stCP12 As String, stCP13 As String, stNP08 As String, stNP09 As String
   Dim rsQuery As ADODB.Recordset, intQ As Integer, iRecCnt As Integer
   Dim stSrt As String, stCP28 As String, stCaseNo As String
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
   
   strUserNo = strUserNum
       
   'CFP延展費(英國)期限
   stSQL = " select decode(substr(st15,1,1),'S',st06,'0')||CU12||CU13||cu01||cu02||NVL(PA149,CU127)||NP02 Srt" & _
      ",np02,np03,np04,np05,np07,np08,np09,np01,np22,pa26 CuNo,cp09" & _
      " from nextprogress a,patent,customer,staff,caseprogress" & _
      " where np09>20210000 and np02='CFP' and np07='613' and np06 is null" & _
      " and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05 and pa09='239' and pa57||pa108 is null" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)='0' and st01(+)=cu13" & _
      " and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp10(+)='1999' and cp05(+)>=20201216 and instr(cp64(+),'英國脫歐通知2')>0"
   
   'CFT延展(英國)期限
   stSQL = stSQL & " union select decode(substr(st15,1,1),'S',st06,'0')||CU12||CU13||cu01||cu02||NVL(TM123,CU127)||NP02 Srt" & _
      ",np02,np03,np04,np05,np07,np08,np09,np01,np22,tm23 CuNo,cp09" & _
      " from nextprogress a,trademark,customer,staff,caseprogress" & _
      " where np09>20210000 and np02='CFT' and np07='110' and np06 is null" & _
      " and TM01(+)=np02 and TM02(+)=np03 and TM03(+)=np04 and TM04(+)=np05 and TM10='239' and tm29||tm57 is null" & _
      " and cu01(+)=substr(tm23,1,8) and cu02(+)='0' and st01(+)=cu13" & _
      " and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp10(+)='1799' and cp05(+)>=20201216 and instr(cp64(+),'英國脫歐通知2')>0"
      
   stSQL = stSQL & " order by 1,2,3,4,5"
   
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Label4.Caption = "本所信函進度檔建立中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         '先收文本所信函CFP1999/CFT1799
         If Not IsNull(.Fields("cp09")) Then
            stCP09 = .Fields("cp09")
         Else
            If .Fields("np02") = "CFP" Then
               stCP10 = "1999"
               strUserNum = "99043"
            Else
               stCP10 = "1799"
               strUserNum = "78028"
            End If
            stCP13 = PUB_GetAKindSalesNo(.Fields("np02"), .Fields("np03"), .Fields("np04"), .Fields("np05"))
            stCP12 = GetSalesArea(stCP13)
            
            cnnConnection.BeginTrans
            bolInTrans = True
            stCP09 = AutoNo("D", 6)
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp64 ) values ('" & .Fields("np02") & "'" & _
               ",'" & .Fields("np03") & "','" & .Fields("np04") & "','" & .Fields("np05") & "'," & strSrvDate(1) & _
               ",'" & stCP09 & "','" & stCP10 & "','" & stCP12 & "'" & _
               ",'" & stCP13 & "','" & strUserNum & "','N','N',19221111,'N','英國脫歐通知2(業務自行列印及處理);')"
            cnnConnection.Execute stSQL, intQ
            
            PUB_AddLetterProgress stCP09, 0, True, , , .Fields("CuNo"), stCP10
            
            '第一案
            If stSrt <> .Fields("Srt") Then
               '自動確認
               stSQL = "update letterprogress set lp06='QPGMR',lp07=to_char(sysdate,'YYYYMMDD') where lp01='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intQ
               
               stSQL = "update caseprogress set cp28=cp09,cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intQ
               
               'Trigger 會寫發文人故要另外更新
               stSQL = "update caseprogress set cp154='QPGMR' where cp09='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intQ
            Else
               '自動確認
               stSQL = "update letterprogress set lp06='QPGMR',lp07=to_char(sysdate,'YYYYMMDD'),lp12='信函存於" & stCaseNo & "(多案併函);' where lp01='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intQ
               
               stSQL = "update caseprogress set cp64=cp64||'信函存於" & stCaseNo & "(多案併函);',cp28='" & stCP28 & "' where cp09='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intQ
            End If
            cnnConnection.CommitTrans
            
            bolInTrans = False
            iRecCnt = iRecCnt + 1
         End If
         
         If stSrt <> .Fields("Srt") Then
            stSrt = .Fields("Srt")
            stCP28 = stCP09
            stCaseNo = .Fields("np02") & "-" & .Fields("np03") & IIf(.Fields("np04") & .Fields("np05") = "000", "", "-" & .Fields("np04") & "-" & .Fields("np05"))
         End If
         .MoveNext
      Loop
      End With
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
      
      If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         If strUserNo <> "" Then strUserNum = strUserNo
         Exit Sub
      End If
      
      stSQL = "select cp01,cp28,count(*) cnt from caseprogress where cp05>=20201216 and cp01='CFP' and cp10='1999' and instr(cp64,'英國脫歐通知2')>0 group by cp01,cp28"
      stSQL = stSQL & " union all select cp01,cp28,count(*) cnt from caseprogress where cp05>=20201216 and cp01='CFT' and cp10='1799' and instr(cp64,'英國脫歐通知2')>0 group by cp01,cp28"
      'stSQL = " select cp01,cp28,count(*) cnt from caseprogress where cp05>=20201216 and cp01='CFT' and cp10='1799' and instr(cp64,'英國脫歐通知2')>0 group by cp01,cp28"
      stSQL = stSQL & " order by cp28"
      
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With rsQuery
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                           
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            stCP09 = .Fields("cp28")
            If .Fields("cp01") = "CFP" Then
               strUserNum = "99043"
               If .Fields("cnt") > 1 Then
                  stET03 = "10"
               Else
                  stET03 = "11"
               End If
            Else
               strUserNum = "78028"
               If .Fields("cnt") > 1 Then
                  stET03 = "05"
               Else
                  stET03 = "06"
               End If
            End If
            'm_DocSNo = Format(.AbsolutePosition, "000")
            'Debug.Print m_DocSNo & " --> " & Now
            'NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , True, , True, , , stCP09
            'm_DocSNo = ""
            'Sleep 3000
            NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09
            .MoveNext
         Loop
         End With
         
         Label4.Caption = "通知信產生完成！"
      End If
   End If

ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
   
End Sub
'Added by Morgan 2021/1/20 通知英國再註冊
Private Sub Process13()
      
   Dim stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP01 As String, stCP09 As String, stCP10 As String, stCP12 As String, stCP13 As String, stNP08 As String, stNP09 As String
   Dim rsQuery As ADODB.Recordset, intQ As Integer, iRecCnt As Integer
   Dim stSrt As String, stCP28 As String, stCaseNo As String, iCount As Integer
   Dim bolInTrans As Boolean
   
   Dim stFile As String
   Dim m_PdfReader As String
   Dim strPrinter As String
   Dim process_id As Long
   Dim process_handle As Long
   
On Error GoTo ErrHnd
   
   strUserNo = strUserNum
   
   '補未閉卷條件以免下次又漏了
   'CFP依照程序排序
   'stSQL = "select decode(substr(st15,1,1),'S',st06,'0')||CU12||CU13||cu01||cu02||NVL(PA149,CU127) Srt" & _
      ",cp01,cp02,cp03,cp04,cp09,cp14,lp01,cpp02,pa09 NaNo,pa26 CuNo,NVL(PA149,CU127) cuc,st02,c.*" & _
      " from caseprogress,patent,customer c,staff,letterprogress,casepaperpdf" & _
      " where cp05>=20210000 and cp01='CFP' and cp10='1608'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa57 is null" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)='0' and st01(+)=cu13" & _
      " and lp01(+)=cp09 and cpp01(+)=cp09 and instr(upper(cpp02(+)),'.1608.ATT.PDF')>0" & _
      " order by cp14,1,2,3,4,5"
   
   'CFT依照業務區客戶排序
   stSQL = "select decode(substr(st15,1,1),'S',st06,'0')||CU12||CU13||cu01||cu02||NVL(TM123,CU127) Srt" & _
      ",cp01,cp02,cp03,cp04,cp09,cp14,lp01,cpp02,tm10 NaNo,tm23 CuNo,NVL(TM123,CU127) cuc,st02,c.*" & _
      " from caseprogress,trademark,customer c,staff,letterprogress,casepaperpdf" & _
      " where cp05>=20210000 and cp01='CFT' and cp10='1730'" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm29 is null" & _
      " and cu01(+)=substr(tm23,1,8) and cu02(+)='0' and st01(+)=cu13" & _
      " and lp01(+)=cp09 and cpp01(+)=cp09 and instr(upper(cpp02(+)),'.1730.ATT.PDF')>0" & _
      " order by 1,2,3,4,5"
   
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Label4.Caption = "通知信產生中..."
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
                        
      strPrinter = PUB_GetOsDefaultPrinter
      '載入Reader
      m_PdfReader = PUB_SetFileAssociation
      process_id = Shell("""" & m_PdfReader & """", vbNormalNoFocus)
      process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
      
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         '測試英國案
         '.Find "NaNo='201'"
         '.Find "cp09='CB0003670'"
         'If .EOF Then Exit Do
         
         stSrt = .Fields("Srt")
         stCP01 = .Fields("cp01")
         stCP09 = .Fields("cp09")
         strUserNum = .Fields("cp14")
         
         If .Fields("NaNo") = "239" Then
            stET03 = "01"
         Else
            stET03 = "02"
         End If
         
         If IsNull(.Fields("lp01")) Then
            cnnConnection.BeginTrans
            bolInTrans = True
            
            PUB_AddLetterProgress stCP09, 0, True, , True, .Fields("CuNo")
         
            '自動確認
            stSQL = "update letterprogress set lp06='QPGMR',lp07=to_char(sysdate,'YYYYMMDD') where lp01='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ
            
            stSQL = "update caseprogress set cp27=" & strSrvDate(1) & ",cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ
            
            'Trigger 會寫發文人故要另外更新
            stSQL = "update caseprogress set cp154='QPGMR' where cp09='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ
            
            cnnConnection.CommitTrans
            bolInTrans = False
         End If
         
         m_DocSNo = Format(.AbsolutePosition, "000")
         Debug.Print m_DocSNo & " --> " & Now
         
         '下載附件
         stFile = App.path & "\$" & m_DocSNo & Replace(.Fields("cpp02"), " ", "_")
         If Dir(stFile) = "" Then
            If PUB_GetAttachFile_CPP(.Fields("cp09"), .Fields("cpp02"), stFile, True) = False Then
               If MsgBox("附件( " & .Fields("cpp02") & ")下載失敗!!" & vbCrLf & vbCrLf & "是否要繼續?", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Do
               End If
            End If
         End If
         
         NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , True, , True, , , stCP09
         Sleep 1000
         
         '列印附件
         If PrintOnePdf(m_PdfReader, " /n /t """ & stFile & """ """ & strPrinter & """") = False Then
            If MsgBox("附件( " & .Fields("cpp02") & ")附件失敗!!" & vbCrLf & vbCrLf & "是否要繼續?", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Do
            End If
         End If
         'If .AbsolutePosition Mod 10 = 0 Then
         '   If MsgBox("是否要繼續？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
         '      Exit Do
         '   End If
         'End If
         m_DocSNo = ""
         .MoveNext
      Loop
      End With
      
      Label4.Caption = "通知信產生完成！"
   End If

ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
   
   If process_handle <> 0 Then
      TerminateProcess process_handle, 0&
      CloseHandle process_handle
   End If
   
   
End Sub

'Added by Morgan 2020/12/9 緬甸商標通知重新申請
Private Sub Process11()
      
   Dim stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP10 As String, stCP12 As String, stCP13 As String, stCP14 As String, stCP64 As String
   Dim rsQuery As ADODB.Recordset, intQ As Integer, iRecCnt As Integer
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
   
   strUserNo = strUserNum

   
   '緬甸有註冊號數且未銷卷之商標案件
   'stSQL = "select decode(substr(st15,1,1),'S',st06,'1')||CU12||CU13||cu01||cu02||NVL(TM123,CU127) Srt" & _
      ",tm01,tm02,tm03,tm04,tm23 from trademark t,customer c,staff" & _
      " where tm10='048' and tm57 is null and tm15 is not null" & _
      " and cu01(+)=substr(tm23,1,8) and cu02='0' and st01(+)=cu13" & _
      " order by 1,2,3,4,5"
   '改抓Excel匯入(因有人工剔除)
   stSQL = "select decode(substr(st15,1,1),'S',st06,'1')||CU12||CU13||cu01||cu02||NVL(TM123,CU127) Srt" & _
      ",tm01,tm02,tm03,tm04,tm23,tm09 from morgan,trademark t,customer c,staff" & _
      " where tm01(+)='CFT' and tm02(+)=m02 and tm03(+)=m03 and tm04(+)='00'" & _
      " and cu01(+)=substr(tm23,1,8) and cu02='0' and st01(+)=cu13" & _
      " order by 1,2,3,4,5"
   
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Label4.Caption = "本所信函進度檔建立中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         '先收文本所信函CFT1799
         stSQL = "update caseprogress set cp28=cp28 where cp01='" & .Fields("tm01") & "' and cp02='" & .Fields("tm02") & "' and cp03='" & .Fields("tm03") & "' and cp04='" & .Fields("tm04") & "' and instr(cp64,'緬甸商標通知重新申請')>0 and cp05>=20210120"
         cnnConnection.Execute stSQL, intQ
         If intQ = 0 Then
            stCP10 = "1799"
            stCP14 = "A6034"
            stCP13 = PUB_GetAKindSalesNo(.Fields("tm01"), .Fields("tm02"), .Fields("tm03"), .Fields("tm04"))
            stCP12 = GetSalesArea(stCP13)
            stCP64 = "緬甸商標通知重新申請;30,000(12);"
            If InStr(.Fields("tm09"), ",") > 0 Then
               stCP64 = stCP64 & " 25,000(9);"
            End If
            cnnConnection.BeginTrans
            bolInTrans = True
            stCP09 = AutoNo("D", 6)
            
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp64 ) values ('" & .Fields("tm01") & "'" & _
               ",'" & .Fields("tm02") & "','" & .Fields("tm03") & "','" & .Fields("tm04") & "'," & strSrvDate(1) & _
               ",'" & stCP09 & "','" & stCP10 & "','" & stCP12 & "'" & _
               ",'" & stCP13 & "','" & stCP14 & "','N','N',20210122,'N','" & stCP64 & "')"
            cnnConnection.Execute stSQL, intQ
            
            PUB_AddLetterProgress stCP09, 0, True, , True, .Fields("tm23"), stCP10
            
            stSQL = "update caseprogress set cp28=cp09,cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ
            If intQ = 1 Then
               'Trigger 會寫發文人故要另外更新
               stSQL = "update caseprogress set cp154='QPGMR' where cp09='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intQ
            End If
            '自動確認
            stSQL = "update letterprogress set lp06='QPGMR',lp07=to_char(sysdate,'YYYYMMDD') where lp01='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ
            
            cnnConnection.CommitTrans
            bolInTrans = False
            iRecCnt = iRecCnt + 1
         End If

         .MoveNext
      Loop
      End With
      
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
      
      If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
         
         
      stSQL = "select decode(substr(st15,1,1),'S',st06,'1')||CU12||CU13||cu01||cu02||NVL(TM123,CU127) Srt" & _
         ",tm01,tm02,tm03,tm04,cp09,tm09 from caseprogress,trademark,customer,staff" & _
         " where cp05>=20210120 and cp01='CFT' and cp10='1799' and instr(cp64,'緬甸商標通知重新申請')>0" & _
         " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04" & _
         " and cu01(+)=substr(tm23,1,8) and cu02(+)='0' and st01(+)=cu13" & _
         " order by 1,2,3,4,5"
      
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With rsQuery
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         strUserNum = "A6034"
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            stCP09 = .Fields("cp09")
            stET03 = "04"
            If InStr(.Fields("tm09"), ",") > 0 Then  '一類以上
               stET03 = "07"
            End If
            m_DocSNo = Format(.AbsolutePosition, "000")
            Debug.Print m_DocSNo & " --> " & Now
            NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , True, , True, , , stCP09
            m_DocSNo = ""
            Sleep 2000
            .MoveNext
         Loop
         End With
         Label4.Caption = "通知信產生完成！"
      End If
   End If
   
ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
   
End Sub

'Added by Morgan 2021/6/8 台灣設計案通函(大陸修法開拓)
Private Sub Process14()
      
   Dim stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP10 As String, stCP12 As String, stCP13 As String, stCP14 As String, stCP64 As String
   Dim rsQuery As ADODB.Recordset, intQ As Integer, iRecCnt As Integer
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
   
   If strUserNum <> "79075" Then MsgBox "要先切換 User 為 79075!!", vbCritical: Exit Sub
   
   strUserNo = strUserNum
   
   '改抓Excel匯入(內專提供)
   stSQL = "select m00,pa01,pa02,pa03,pa04,pa26 from morgan,patent" & _
      " where pa01(+)=substr(m01,1,1) and pa02(+)=substr(m01,2,6) and pa03='0' and pa04='00' order by m00"
   
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Label4.Caption = "本所信函進度檔建立中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         
         '先收文本所信函 P1999
         stSQL = "update caseprogress set cp28=cp28 where cp01='" & .Fields("pa01") & "' and cp02='" & .Fields("pa02") & "' and cp03='" & .Fields("pa03") & "' and cp04='" & .Fields("pa04") & "' and cp10='1999' and instr(cp64,'開拓辦理大陸外觀設計')>0 and cp05>=20210608"
         cnnConnection.Execute stSQL, intQ
         If intQ = 0 Then
            stCP10 = "1999"
            stCP14 = "79075"
            stCP13 = PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
            stCP12 = GetSalesArea(stCP13)
            stCP64 = "開拓辦理大陸外觀設計"
            cnnConnection.BeginTrans
            bolInTrans = True
            stCP09 = AutoNo("D", 6)
            
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp64 ) values ('" & .Fields("pa01") & "'" & _
               ",'" & .Fields("pa02") & "','" & .Fields("pa03") & "','" & .Fields("pa04") & "'," & strSrvDate(1) & _
               ",'" & stCP09 & "','" & stCP10 & "','" & stCP12 & "'" & _
               ",'" & stCP13 & "','" & stCP14 & "','N','N'," & strSrvDate(1) & ",'N','" & stCP64 & "')"
            cnnConnection.Execute stSQL, intQ
            
            PUB_AddLetterProgress stCP09, 0, True, , , .Fields("pa26"), stCP10
            
            'stSQL = "update caseprogress set cp28=cp09,cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & stCP09 & "'"
            'cnnConnection.Execute stSQL, intQ
            'If intQ = 1 Then
            '   'Trigger 會寫發文人故要另外更新
            '   stSQL = "update caseprogress set cp154='QPGMR' where cp09='" & stCP09 & "'"
            '   cnnConnection.Execute stSQL, intQ
            'End If
            ''自動確認
            'stSQL = "update letterprogress set lp06='QPGMR',lp07=to_char(sysdate,'YYYYMMDD') where lp01='" & stCP09 & "'"
            'cnnConnection.Execute stSQL, intQ
            
            cnnConnection.CommitTrans
            bolInTrans = False
            iRecCnt = iRecCnt + 1
         End If

         .MoveNext
      Loop
      End With
      
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
      
      If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
     
      stSQL = "select cp09 from caseprogress where cp05>=20210608 and cp01='P' and cp10='1999' and instr(cp64,'開拓辦理大陸外觀設計')>0 order by cp09"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With rsQuery
         .MoveFirst
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         strUserNum = "79075"
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            stCP09 = .Fields("cp09")
            stET03 = "15"
            'm_DocSNo = Format(.AbsolutePosition, "000")
            'Debug.Print m_DocSNo & " --> " & Now
            NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09
            'm_DocSNo = ""
            'Sleep 2000
            .MoveNext
         Loop
         End With
         Label4.Caption = "通知信產生完成！"
      End If
      
   End If
   
ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
   
End Sub

'2013/2/6 add by sonia
'102年美國發明維持費規費調漲通知
Private Sub Process4()
Dim stCon As String, stLstCustNo As String, stET02 As String
Dim strTmp As String, StrCaseList(5) As String, iNo As Integer, idx As Integer
Dim stLY As String, stNextLY As String, stOldAmt As String, stNewAmt As String
Dim bolSave As Boolean, iSNo As Integer, stPA91 As String
Dim strUserNo As String
   
   strSql = "delete CFP606"
   cnnConnection.Execute strSql, intI

   stCon = "": bolSave = False
   If txtCust <> "" Then
      txtCust = ChangeCustomerL(txtCust)
      stCon = " and pa26='" & txtCust & "'"
   End If
   
   strExc(0) = "select pa01,pa02,pa03,pa04,pa05,sqldatew(np09) np09,pa26,lastyear(pa72) LY,pa91 from nextprogress,patent where np02='CFP' and np07='606' " & _
               "and np09>=20130209 and np09<=20130918 and np06 is null " & _
               "and np02=pa01(+) and np03=pa02(+) and np04=pa03(+) and np05=pa04(+) and '101'=pa09(+) and '1'=pa08 and pa57 is null " & stCon & _
               " order by pa26,pa02"
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strUserNo = strUserNum
      strUserNum = "79017" '用程序編號跑才會有發文字且方便列印及維護
      With RsTemp
         stLstCustNo = "" & .Fields("pa26")
         stET02 = "" & .Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04") & "&000"
         Erase StrCaseList
         iNo = 0
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         iSNo = 1
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            If .Fields("pa26") <> stLstCustNo Then
               '只有1案
               If iNo = 1 Then
                  EndLetter "21", stET02, "03", strUserNum
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                     "','維持費次數' ,'" & stNextLY & "')"
                  cnnConnection.Execute strExc(0), intI
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                     "','舊報價' ,'" & Format(stOldAmt, DDollar) & "')"
                  cnnConnection.Execute strExc(0), intI
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                     "','新報價' ,'" & Format(stNewAmt, DDollar) & "')"
                  cnnConnection.Execute strExc(0), intI
                  
                  If Not bolSave Then
                     strTmp = Format(iSNo, "0000")
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                        "','發文流水號' ,'" & strTmp & "')"
                     cnnConnection.Execute strExc(0), intI
                  End If
                  
                  NowPrint stET02, "21", "03", False, strUserNum, , , , , 2, , , bolSave
                  iSNo = iSNo + 1
               Else
                  '最後一頁
                  'If iNo Mod 3 = 2 Or iNo Mod 3 = 0 Then
                  If iNo Mod 3 = 1 Then
                     StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
                        "  本所案號    案件名稱   維持費次數 法定期限     費用(新台幣)" & vbCrLf & _
                        " -----------------------------------------------------------------------------------------------------" & vbCrLf
                  End If
                  StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
                  
                  EndLetter "21", stET02, "01", strUserNum
                  For idx = 1 To 5
                     If StrCaseList(idx) <> "" Then
                        strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                           "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                        cnnConnection.Execute strExc(0), intI
                     End If
                  Next
                  
                  If Not bolSave Then
                     strTmp = Format(iSNo, "0000")
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                        "','發文流水號' ,'" & strTmp & "')"
                     cnnConnection.Execute strExc(0), intI
                  End If
                  
                  NowPrint stET02, "21", "01", False, strUserNum, , , , , 2, , , bolSave
                  iSNo = iSNo + 1
               End If
               
               stLstCustNo = "" & .Fields("pa26")
               stET02 = "" & .Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04") & "&000"
               Erase StrCaseList
               iNo = 0
            Else
            
               '跳頁控制
               'If iNo > 0 And iNo Mod 4 = 0 Then
               If iNo = 4 Or iNo = 7 Or iNo = 10 Or iNo = 13 Then
                  StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
                     "  本所案號    案件名稱   維持費次數 法定期限     費用(新台幣)" & vbCrLf & _
                     " -----------------------------------------------------------------------------------------------------" & vbCrLf
               End If
               StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
            End If
            
            stPA91 = "" & .Fields("pa91")
            stLY = "" & .Fields("LY")   '已繳次數
            stNextLY = ""               '下次次數
            stOldAmt = 0: stNewAmt = 0  '新舊報價
            Select Case stLY
               Case ""       '第1次
                  stNextLY = 1
                  If InStr(stPA91, "小個體") > 0 Then
                     stOldAmt = 26000: stNewAmt = 33000
                  ElseIf InStr(stPA91, "大個體") > 0 Then
                     stOldAmt = 44000: stNewAmt = 58000
                     '特殊報價
                     If "" & .Fields("pa26") = "X01677000" Or "" & .Fields("pa26") = "X01677030" Then        '謝武弘及功學社教育用品股份有限公司
                        stOldAmt = 53000: stNewAmt = 67000
                     ElseIf "" & .Fields("pa26") = "X14843010" Or "" & .Fields("pa26") = "X14843040" Then    '新日興股份有限公司及呂勝男
                        stOldAmt = 58000: stNewAmt = 58000
                     End If
                  End If
               Case "3.5"    '第2次
                  stNextLY = 2
                  If InStr(stPA91, "小個體") > 0 Then
                     stOldAmt = 53000: stNewAmt = 64000
                  ElseIf InStr(stPA91, "大個體") > 0 Then
                     stOldAmt = 98000: stNewAmt = 120000
                     '特殊報價
                     If "" & .Fields("pa26") = "X01677000" Or "" & .Fields("pa26") = "X01677030" Then        '謝武弘及功學社教育用品股份有限公司
                        stOldAmt = 107000: stNewAmt = 129000
                     ElseIf "" & .Fields("pa26") = "X14843010" Or "" & .Fields("pa26") = "X14843040" Then    '新日興股份有限公司及呂勝男
                        stOldAmt = 116000: stNewAmt = 120000
                     End If
                  End If
               Case "7.5"    '第3次
                  stNextLY = 3
                   If InStr(stPA91, "小個體") > 0 Then
                     stOldAmt = 83000: stNewAmt = 123000
                  ElseIf InStr(stPA91, "大個體") > 0 Then
                     stOldAmt = 157000: stNewAmt = 238000
                     '特殊報價
                     If "" & .Fields("pa26") = "X01677000" Or "" & .Fields("pa26") = "X01677030" Then        '謝武弘及功學社教育用品股份有限公司
                        stOldAmt = 166000: stNewAmt = 247000
                     ElseIf "" & .Fields("pa26") = "X14843010" Or "" & .Fields("pa26") = "X14843040" Then    '新日興股份有限公司及呂勝男
                        stOldAmt = 175000: stNewAmt = 238000
                     End If
                  End If
           End Select
            
            '案件清單
            iNo = iNo + 1
            idx = 1 + iNo \ 8
            strTmp = .Fields("pa01") & "-" & .Fields("pa02") & IIf(.Fields("pa03") = "0", "", "-" & .Fields("pa03"))
            StrCaseList(0) = Format(iNo, "@@") & "." & PUB_StrToStr(strTmp, 12, True) '本所案號
            strTmp = "" & .Fields("pa05")
            StrCaseList(0) = StrCaseList(0) & PUB_StrToStr(strTmp, 16, True)          '案件名稱
            strTmp = stNextLY
            StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 3, True)    '維持費次數
            strTmp = "" & .Fields("np09")
            StrCaseList(0) = StrCaseList(0) & PUB_StrToStr(strTmp, 10, True)          '法定期限
            strTmp = Format(stOldAmt, DDollar)
            If Len(PUB_StrToStr("" & .Fields("pa05"), 16, False)) <= 4 Then
               StrCaseList(0) = StrCaseList(0) & " " & PUB_StrToStr(strTmp, 8, True, True) & " (3/18前(含當日))"  '金額(新台幣)2013/3/18前(含當日)
            ElseIf Len(PUB_StrToStr("" & .Fields("pa05"), 16, True)) >= 12 Then
               StrCaseList(0) = StrCaseList(0) & " " & PUB_StrToStr(strTmp, 7, True, True) & " (3/18前(含當日))"  '金額(新台幣)2013/3/18前(含當日)
            Else
               StrCaseList(0) = StrCaseList(0) & " " & PUB_StrToStr(strTmp, 8, True, True) & " (3/18前(含當日))"  '金額(新台幣)2013/3/18前(含當日)
            End If
            strTmp = Format(stNewAmt, DDollar)
            StrCaseList(0) = StrCaseList(0) & vbCrLf & "                                              " & PUB_StrToStr(strTmp, 8, True, True) & " (3/19後(含當日))"  '金額(新台幣)2013/3/19後(含當日)
            
            StrCaseList(0) = StrCaseList(0) & vbCrLf & _
               vbCrLf & "□ 同意辦理，請　貴所代為管制本案後續發展。" & _
               vbCrLf & "□ 本人／本公司／本單位自行處理本案之後續作業，並同意　貴所不需作本" & _
               vbCrLf & "　 案之後續追蹤及通知。" & _
               vbCrLf & "□ 放棄本案。" & _
               vbCrLf & "□ 其他，請說明ˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍ" & vbCrLf & vbCrLf & vbCrLf
            
            strSql = "insert into CFP606 (t01,t02,t03,t04,t05,told,tnew,t06) " & _
                     "values ('" & .Fields("pa01") & "','" & .Fields("pa02") & "','" & .Fields("pa03") & "','" & .Fields("pa04") & "'," & Val(stNextLY) & "," & Val(stOldAmt) & "," & Val(stNewAmt) & "," & Val(DBDATE("" & .Fields("np09"))) & ") "
            cnnConnection.Execute strSql, intI
            .MoveNext
         Loop
         
         '最後一頁
         'If iNo Mod 3 = 1 Or iNo Mod 3 = 0 Then
         If iNo Mod 3 = 1 Then
            StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
               "  本所案號    案件名稱   維持費次數 法定期限     費用(新台幣)" & vbCrLf & _
               " -----------------------------------------------------------------------------------------------------" & vbCrLf
         End If
         
         StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
         '只有1案
         If iNo = 1 Then
            EndLetter "21", stET02, "03", strUserNum
            strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
               "','維持費次數' ,'" & stNextLY & "')"
            cnnConnection.Execute strExc(0), intI
            strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
               "','舊報價' ,'" & Format(stOldAmt, DDollar) & "')"
            cnnConnection.Execute strExc(0), intI
            strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
               "','新報價' ,'" & Format(stNewAmt, DDollar) & "')"
            cnnConnection.Execute strExc(0), intI
            
            If Not bolSave Then
               strTmp = Format(iSNo, "0000")
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','03','" & strUserNum & _
                  "','發文流水號' ,'" & strTmp & "')"
               cnnConnection.Execute strExc(0), intI
            End If
            
            NowPrint stET02, "21", "03", False, strUserNum, , , , , 2, , , bolSave
         Else
            EndLetter "21", stET02, "01", strUserNum
            For idx = 1 To 5
               If StrCaseList(idx) <> "" Then
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                     "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                  cnnConnection.Execute strExc(0), intI
               End If
            Next
            
            If Not bolSave Then
               strTmp = Format(iSNo, "0000")
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                  "','發文流水號' ,'" & strTmp & "')"
               cnnConnection.Execute strExc(0), intI
            End If
            
            NowPrint stET02, "21", "01", False, strUserNum, , , , , 2, , , bolSave
         End If
      End With
      strUserNum = strUserNo
      MsgBox "定稿已產生完畢！"
   Else
      MsgBox "無資料可作業！"
   End If
End Sub
'2013/2/6 end

'add by sonia 2016/7/13
'105年英國脫歐之CFP,CFT案通知
'1.CFP:先執行秀玲電腦上C:\83002\案件系統文件\雜文\專利處\CFP\英國脫歐CFP抓資料語法.txt先產生PATENT_EU_SONIA20160707
'2.CFT
Private Sub Process5()
Dim stCon As String, stLstCustNo As String, stET02 As String
Dim strTmp As String, StrCaseList(26) As String, iNo As Integer, idx As Integer
Dim stSQL As String
Dim bolSave As Boolean, iSNo As Integer
Dim strUserNo As String
   
'CFP之EPC
   stCon = "and y.pa09='221'": bolSave = False
   
   '依申請人案件筆數由少至多排序
   strExc(0) = "select PA26,X.PA01 PA01,X.PA02 PA02,X.PA03 PA03,X.PA04 PA04,PA05,PA11,sqldatew(PA25) PA25 from PATENT_EU_SONIA20160707 X,PATENT Y, " & _
               "  (select PA26 CUNO,COUNT(*) CNT from PATENT_EU_SONIA20160707 X,PATENT Y" & _
               "    where X.PA01=Y.PA01(+) AND X.PA02=Y.PA02(+) AND X.PA03=Y.PA03(+) AND X.PA04=Y.PA04(+) " & stCon & " GROUP BY PA26) Z " & _
               " where X.PA01=Y.PA01(+) AND X.PA02=Y.PA02(+) AND X.PA03=Y.PA03(+) AND X.PA04=Y.PA04(+) AND Y.PA26=Z.CUNO(+) " & stCon & _
               " ORDER BY Z.CNT,Y.PA26,Y.PA02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strUserNo = strUserNum
      strUserNum = "85037" '用程序編號跑才會有發文字且方便列印及維護
      iSNo = 1             '三定稿用一組發文流水號
      With RsTemp
         stLstCustNo = "" & .Fields("pa26")
         stET02 = "" & .Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04") & "&000"
         Erase StrCaseList
         iNo = 0
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            If .Fields("pa26") <> stLstCustNo Then
               '只有1案(05)
               If iNo = 1 Then
                  EndLetter "21", stET02, "05", strUserNum
                  If Not bolSave Then
                     strTmp = Format(iSNo, "0000")
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stET02 & "','05','" & strUserNum & _
                        "','發文流水號' ,'" & strTmp & "')"
                     cnnConnection.Execute strExc(0), intI
                  End If
                  
                  NowPrint stET02, "21", "05", False, strUserNum, , , , , 1, , , bolSave
                  iSNo = iSNo + 1
               Else  '多案(04)
                  '最後一頁
                  If iNo Mod 25 = 1 Then
                     StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
                        "   本所案號    案件名稱              申請案號       專用期止日" & vbCrLf & _
                        " --------------------------------------------------------------------------------------------------" & vbCrLf
                  End If
                  StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
                  
                  EndLetter "21", stET02, "04", strUserNum
                  For idx = 1 To 25
                     If StrCaseList(idx) <> "" Then
                        strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('21','" & stET02 & "','04','" & strUserNum & _
                           "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                        cnnConnection.Execute strExc(0), intI
                     End If
                  Next
                  
                  If Not bolSave Then
                     strTmp = Format(iSNo, "0000")
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stET02 & "','04','" & strUserNum & _
                        "','發文流水號' ,'" & strTmp & "')"
                     cnnConnection.Execute strExc(0), intI
                  End If
                  
                  NowPrint stET02, "21", "04", False, strUserNum, , , , , 1, , , bolSave
                  iSNo = iSNo + 1
               End If
               
               stLstCustNo = "" & .Fields("pa26")
               stET02 = "" & .Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04") & "&000"
               Erase StrCaseList
               iNo = 0
            Else
            
               '跳頁控制
               If iNo = 26 Then
                  StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
                     "   本所案號    案件名稱              申請案號       專用期止日" & vbCrLf & _
                     " --------------------------------------------------------------------------------------------------" & vbCrLf
               End If
               StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
            End If
            
            '案件清單
            iNo = iNo + 1
            idx = 1 + iNo \ 8
            strTmp = .Fields("pa01") & "-" & .Fields("pa02") & IIf(.Fields("pa03") = "0", "", "-" & .Fields("pa03"))
            StrCaseList(0) = Format(iNo, "@@") & "." & PUB_StrToStr(strTmp, 12, True) '本所案號
            strTmp = "" & .Fields("PA05")
            StrCaseList(0) = StrCaseList(0) & PUB_StrToStr(strTmp, 20, True)          '案件名稱
            strTmp = "" & .Fields("pa11")
            StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 16, True)   '申請案號
            strTmp = "" & .Fields("PA25")
            StrCaseList(0) = StrCaseList(0) & PUB_StrToStr(strTmp, 10, True)          '專用期止日
            StrCaseList(0) = StrCaseList(0) & vbCrLf
            '新增進度檔
            strTmp = Format(iSNo, "0000")
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp28,cp64 ) values ('" & .Fields("pa01") & "'" & _
               ",'" & .Fields("pa02") & "','" & .Fields("pa03") & "','" & .Fields("pa04") & "'," & strSrvDate(1) & _
               ",'" & AutoNo("D", 6) & "','1999','" & GetSalesArea(PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))) & "'" & _
               ",'" & PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04")) & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N','" & strTmp & "','英國脫歐通知')"
            cnnConnection.Execute stSQL, intI
            
            .MoveNext
         Loop
         
         '最後一頁
         If iNo Mod 25 = 1 Then
            StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
               "   本所案號    案件名稱              申請案號       專用期止日" & vbCrLf & _
               " --------------------------------------------------------------------------------------------------" & vbCrLf
         End If
         
         StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
         '只有1案(05)
         If iNo = 1 Then
            EndLetter "21", stET02, "05", strUserNum
           
            If Not bolSave Then
               strTmp = Format(iSNo, "0000")
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','05','" & strUserNum & _
                  "','發文流水號' ,'" & strTmp & "')"
               cnnConnection.Execute strExc(0), intI
            End If
            
            NowPrint stET02, "21", "05", False, strUserNum, , , , , 1, , , bolSave
         Else  '多案(04)
            EndLetter "21", stET02, "04", strUserNum
            For idx = 1 To 25
               If StrCaseList(idx) <> "" Then
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stET02 & "','04','" & strUserNum & _
                     "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                  cnnConnection.Execute strExc(0), intI
               End If
            Next
            
            If Not bolSave Then
               strTmp = Format(iSNo, "0000")
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','04','" & strUserNum & _
                  "','發文流水號' ,'" & strTmp & "')"
               cnnConnection.Execute strExc(0), intI
            End If
            
            NowPrint stET02, "21", "04", False, strUserNum, , , , , 1, , , bolSave
         End If
      End With
      strUserNum = strUserNo
      'MsgBox "CFP之EPC定稿已產生完畢！"
   Else
      'MsgBox "CFP之EPC無資料可作業！"
   End If

'CFP之EU
   stCon = "and y.pa09='239'": bolSave = False
   
   '依申請人案件筆數由少至多排序
   strExc(0) = "select PA26,X.PA01 PA01,X.PA02 PA02,X.PA03 PA03,X.PA04 PA04,PA05,PA11,sqldatew(PA25) PA25 from PATENT_EU_SONIA20160707 X,PATENT Y, " & _
               "  (select PA26 CUNO,COUNT(*) CNT from PATENT_EU_SONIA20160707 X,PATENT Y" & _
               "    where X.PA01=Y.PA01(+) AND X.PA02=Y.PA02(+) AND X.PA03=Y.PA03(+) AND X.PA04=Y.PA04(+) " & stCon & " GROUP BY PA26) Z " & _
               " where X.PA01=Y.PA01(+) AND X.PA02=Y.PA02(+) AND X.PA03=Y.PA03(+) AND X.PA04=Y.PA04(+) AND Y.PA26=Z.CUNO(+) " & stCon & _
               " ORDER BY Z.CNT,Y.PA26,Y.PA02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strUserNo = strUserNum
      strUserNum = "85037" '用程序編號跑才會有發文字且方便列印及維護
      iSNo = iSNo + 1      '三定稿用一組發文流水號
      With RsTemp
         stLstCustNo = "" & .Fields("pa26")
         stET02 = "" & .Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04") & "&000"
         Erase StrCaseList
         iNo = 0
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            If .Fields("pa26") <> stLstCustNo Then
               '只有1案(07)
               If iNo = 1 Then
                  EndLetter "21", stET02, "07", strUserNum
                  If Not bolSave Then
                     strTmp = Format(iSNo, "0000")
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stET02 & "','07','" & strUserNum & _
                        "','發文流水號' ,'" & strTmp & "')"
                     cnnConnection.Execute strExc(0), intI
                  End If
                  
                  NowPrint stET02, "21", "07", False, strUserNum, , , , , 1, , , bolSave
                  iSNo = iSNo + 1
               Else  '多案(06)
                  '最後一頁
                  If iNo Mod 25 = 1 Then
                     StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
                        "   本所案號    案件名稱              申請案號       專用期止日" & vbCrLf & _
                        " --------------------------------------------------------------------------------------------------" & vbCrLf
                  End If
                  StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
                  
                  EndLetter "21", stET02, "06", strUserNum
                  For idx = 1 To 25
                     If StrCaseList(idx) <> "" Then
                        strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('21','" & stET02 & "','06','" & strUserNum & _
                           "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                        cnnConnection.Execute strExc(0), intI
                     End If
                  Next
                  
                  If Not bolSave Then
                     strTmp = Format(iSNo, "0000")
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stET02 & "','06','" & strUserNum & _
                        "','發文流水號' ,'" & strTmp & "')"
                     cnnConnection.Execute strExc(0), intI
                  End If
                  
                  NowPrint stET02, "21", "06", False, strUserNum, , , , , 1, , , bolSave
                  iSNo = iSNo + 1
               End If
               
               stLstCustNo = "" & .Fields("pa26")
               stET02 = "" & .Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04") & "&000"
               Erase StrCaseList
               iNo = 0
            Else
            
               '跳頁控制
               If iNo = 26 Then
                  StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
                     "   本所案號    案件名稱              申請案號       專用期止日" & vbCrLf & _
                     " --------------------------------------------------------------------------------------------------" & vbCrLf
               End If
               StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
            End If
            
            '案件清單
            iNo = iNo + 1
            idx = 1 + iNo \ 8
            strTmp = .Fields("pa01") & "-" & .Fields("pa02") & IIf(.Fields("pa03") = "0", "", "-" & .Fields("pa03"))
            StrCaseList(0) = Format(iNo, "@@") & "." & PUB_StrToStr(strTmp, 12, True) '本所案號
            strTmp = "" & .Fields("PA05")
            StrCaseList(0) = StrCaseList(0) & PUB_StrToStr(strTmp, 20, True)          '案件名稱
            strTmp = "" & .Fields("pa11")
            StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 16, True)   '申請案號
            strTmp = "" & .Fields("PA25")
            StrCaseList(0) = StrCaseList(0) & PUB_StrToStr(strTmp, 10, True)          '專用期止日
            StrCaseList(0) = StrCaseList(0) & vbCrLf
            '新增進度檔
            strTmp = Format(iSNo, "0000")
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp28,cp64 ) values ('" & .Fields("pa01") & "'" & _
               ",'" & .Fields("pa02") & "','" & .Fields("pa03") & "','" & .Fields("pa04") & "'," & strSrvDate(1) & _
               ",'" & AutoNo("D", 6) & "','1999','" & GetSalesArea(PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))) & "'" & _
               ",'" & PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04")) & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N','" & strTmp & "','英國脫歐通知')"
            cnnConnection.Execute stSQL, intI
            
            .MoveNext
         Loop
         
         '最後一頁
         If iNo Mod 25 = 1 Then
            StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
               "   本所案號    案件名稱              申請案號       專用期止日" & vbCrLf & _
               " --------------------------------------------------------------------------------------------------" & vbCrLf
         End If
         
         StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
         '只有1案(07)
         If iNo = 1 Then
            EndLetter "21", stET02, "07", strUserNum
           
            If Not bolSave Then
               strTmp = Format(iSNo, "0000")
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','07','" & strUserNum & _
                  "','發文流水號' ,'" & strTmp & "')"
               cnnConnection.Execute strExc(0), intI
            End If
            
            NowPrint stET02, "21", "07", False, strUserNum, , , , , 1, , , bolSave
         Else  '多案(06)
            EndLetter "21", stET02, "06", strUserNum
            For idx = 1 To 25
               If StrCaseList(idx) <> "" Then
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stET02 & "','06','" & strUserNum & _
                     "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                  cnnConnection.Execute strExc(0), intI
               End If
            Next
            
            If Not bolSave Then
               strTmp = Format(iSNo, "0000")
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','06','" & strUserNum & _
                  "','發文流水號' ,'" & strTmp & "')"
               cnnConnection.Execute strExc(0), intI
            End If
            
            NowPrint stET02, "21", "06", False, strUserNum, , , , , 1, , , bolSave
         End If
      End With
      strUserNum = strUserNo
      'MsgBox "CFP之EU定稿已產生完畢！"
   Else
      'MsgBox "CFP之EU無資料可作業！"
   End If

'CFT
   bolSave = False
   
   '依申請人案件筆數由少至多排序
   strExc(0) = "SELECT TM23,TM01,TM02,TM03,TM04,TM05,TM15,TM09,SQLDATEW(TM22) TM22 FROM TRADEMARK, " & _
               "  (SELECT TM23 CUNO,COUNT(*) CNT FROM TRADEMARK WHERE TM01='CFT' AND TM10='239' AND TM28='1' " & _
               "      AND TM29 IS NULL AND NVL(TM22,0)>=20160701 AND NVL(TM57,0)=0 GROUP BY TM23) Z " & _
               " WHERE TM01='CFT' AND TM10='239' AND TM28='1' AND TM29 IS NULL AND NVL(TM22,0)>=20160701 AND NVL(TM57,0)=0 AND TM23=Z.CUNO(+) " & _
               " ORDER BY Z.CNT,TM23,TM02"
'補寄指定申請人
'   strExc(0) = "SELECT TM23,TM01,TM02,TM03,TM04,TM05,TM15,TM09,SQLDATEW(TM22) TM22 FROM TRADEMARK, " & _
'               "  (SELECT TM23 CUNO,COUNT(*) CNT FROM TRADEMARK WHERE TM01='CFT' AND TM10='239' AND TM28='1' " & _
'               "      AND TM29 IS NULL AND NVL(TM22,0)>=20160701 AND NVL(TM57,0)=0 GROUP BY TM23) Z " & _
'               " WHERE tm23='X30998080' and TM01='CFT' AND TM10='239' AND TM28='1' AND TM29 IS NULL AND NVL(TM22,0)>=20160701 AND NVL(TM57,0)=0 AND TM23=Z.CUNO(+) " & _
'               " ORDER BY Z.CNT,TM23,TM02"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strUserNo = strUserNum
      strUserNum = "72012" '用程序編號跑才會有發文字且方便列印及維護
      iSNo = iSNo + 1      '三定稿用一組發文流水號
      With RsTemp
         stLstCustNo = "" & .Fields("TM23")
         stET02 = "" & .Fields("TM01") & .Fields("TM02") & .Fields("TM03") & .Fields("TM04") & "&000"
         Erase StrCaseList
         iNo = 0
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            If .Fields("TM23") <> stLstCustNo Then
               '只有1案(02)
               If iNo = 1 Then
                  EndLetter "21", stET02, "02", strUserNum
                  If Not bolSave Then
                     strTmp = Format(iSNo, "0000")
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stET02 & "','02','" & strUserNum & _
                        "','發文流水號' ,'" & strTmp & "')"
                     cnnConnection.Execute strExc(0), intI
                  End If
                  
                  NowPrint stET02, "21", "02", False, strUserNum, , , , , 1, , , bolSave
                  iSNo = iSNo + 1
               Else  '多案(01)
                  '最後一頁
                  If iNo Mod 25 = 1 Then
                     StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
                     "  本所案號    審定號     專用期止日  商品類別" & vbCrLf & _
                     "              案件名稱    " & vbCrLf & _
                        " --------------------------------------------------------------------------------------------------" & vbCrLf
                  End If
                  StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
                  
                  EndLetter "21", stET02, "01", strUserNum
                  For idx = 1 To 25
                     If StrCaseList(idx) <> "" Then
                        strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                           "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                        cnnConnection.Execute strExc(0), intI
                     End If
                  Next
                  
                  If Not bolSave Then
                     strTmp = Format(iSNo, "0000")
                     strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                        "','發文流水號' ,'" & strTmp & "')"
                     cnnConnection.Execute strExc(0), intI
                  End If
                  
                  NowPrint stET02, "21", "01", False, strUserNum, , , , , 1, , , bolSave
                  iSNo = iSNo + 1
               End If
               
               stLstCustNo = "" & .Fields("TM23")
               stET02 = "" & .Fields("TM01") & .Fields("TM02") & .Fields("TM03") & .Fields("TM04") & "&000"
               Erase StrCaseList
               iNo = 0
            Else
            
               '跳頁控制
               If iNo = 26 Then
                  StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
                     "  本所案號    審定號     專用期止日  商品類別" & vbCrLf & _
                     "              案件名稱    " & vbCrLf & _
                     " -----------------------------------------------------------------------------------------------------" & vbCrLf
               End If
               StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
            End If
            
            '案件清單 Replace(pa06, "'", "''")
            iNo = iNo + 1
            idx = iNo
            strTmp = .Fields("TM01") & "-" & .Fields("TM02") & IIf(.Fields("TM03") = "0", "", "-" & .Fields("TM03"))
            StrCaseList(0) = Format(iNo, "@@") & "." & PUB_StrToStr(strTmp, 12, True)             '本所案號
            strTmp = "" & .Fields("TM15")
            StrCaseList(0) = StrCaseList(0) & "" & PUB_StrToStr(strTmp, 9, True)                  '審定號
            strTmp = "" & .Fields("TM22")
            StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 10, True)               '專用期止日
            strTmp = "" & .Fields("TM09")
            StrCaseList(0) = StrCaseList(0) & "  " & PUB_StrToStr(strTmp, 18, True)               '商品類別
            StrCaseList(0) = StrCaseList(0) & vbCrLf
            strTmp = "" & .Fields("TM05")
            StrCaseList(0) = StrCaseList(0) & "               " & PUB_StrToStr(Replace(strTmp, "'", "''"), 50, True)  '案件名稱
            '新增進度檔
            strTmp = Format(iSNo, "0000")
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp28,cp64 ) values ('" & .Fields("TM01") & "'" & _
               ",'" & .Fields("TM02") & "','" & .Fields("TM03") & "','" & .Fields("TM04") & "'," & strSrvDate(1) & _
               ",'" & AutoNo("D", 6) & "','1799','" & GetSalesArea(PUB_GetAKindSalesNo(.Fields("TM01"), .Fields("TM02"), .Fields("TM03"), .Fields("TM04"))) & "'" & _
               ",'" & PUB_GetAKindSalesNo(.Fields("TM01"), .Fields("TM02"), .Fields("TM03"), .Fields("TM04")) & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'N','" & strTmp & "','英國脫歐通知')"
            cnnConnection.Execute stSQL, intI
            
            .MoveNext
         Loop
         
         '最後一頁
         If iNo Mod 25 = 1 Then
            StrCaseList(idx) = StrCaseList(idx) & Chr(12) & vbCrLf & vbCrLf & _
               "  本所案號    審定號     專用期止日  商品類別" & vbCrLf & _
               "              案件名稱    " & vbCrLf & _
               " -----------------------------------------------------------------------------------------------------" & vbCrLf
         End If
         
         StrCaseList(idx) = StrCaseList(idx) & StrCaseList(0)
         '只有1案(02)
         If iNo = 1 Then
            EndLetter "21", stET02, "02", strUserNum
           
            If Not bolSave Then
               strTmp = Format(iSNo, "0000")
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','02','" & strUserNum & _
                  "','發文流水號' ,'" & strTmp & "')"
               cnnConnection.Execute strExc(0), intI
            End If
            
            NowPrint stET02, "21", "02", False, strUserNum, , , , , 1, , , bolSave
         Else  '多案(01)
            EndLetter "21", stET02, "01", strUserNum
            For idx = 1 To 25
               If StrCaseList(idx) <> "" Then
                  strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                     "','案件清單" & idx & "' ,'" & StrCaseList(idx) & "')"
                  cnnConnection.Execute strExc(0), intI
               End If
            Next
            
            If Not bolSave Then
               strTmp = Format(iSNo, "0000")
               strExc(0) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('21','" & stET02 & "','01','" & strUserNum & _
                  "','發文流水號' ,'" & strTmp & "')"
               cnnConnection.Execute strExc(0), intI
            End If
            
            NowPrint stET02, "21", "01", False, strUserNum, , , , , 1, , , bolSave
         End If
      End With
      strUserNum = strUserNo
      MsgBox "CFT之EU定稿已產生完畢！"
   Else
      'MsgBox "CFT之EU無資料可作業！"
   End If

End Sub
'end 2016/7/13

'Added by Lydia 2021/08/31 110緬甸商標法通知
Private Sub Process15()
      
   Dim stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP10 As String, stCP12 As String, stCP13 As String, stCP14 As String, stCP64 As String
   Dim rsQuery As ADODB.Recordset, intQ As Integer, iRecCnt As Integer
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
   
   strUserNo = strUserNum

'----Excel
'Select Decode(Substr(St15,1,1),'S',St06,'1')||Cu12||Cu13||Cu01||Cu02||Nvl(Tm123,Cu127) Srt, --Tm01,Tm02,Tm03,Tm04,Tm23,cp13,st06
'Tm01||'-'||Tm02||Decode(Tm03,'0',Null,'-'||Tm03)||Decode(Tm04,'00',Null,'-'||Tm04) As 本所案號,
'Nvl(Tm05,Nvl(Tm06,Tm07)) As 案件名稱,
'Tm23 As 客戶編號, Nvl(Cu04,Nvl(Cu05,Cu06)) As 客戶名稱,
'cp13||st02 as 智權人員
'From Caseprogress, Trademark, customer, staff
'Where Cp158>=20210101 and cp159=0 And Cp01='CFT' And Cp10='101' And Cp01=Tm01(+) And Cp02=Tm02(+) And Cp03=Tm03(+) And Cp04=Tm04(+)
'And Tm10='048' And Tm57 Is Null And Cu01(+)=Substr(Tm23,1,8) And Cu02='0' And St01(+)=Cu13
'order by 1,2

   'CFT緬甸申請101案於2021年發文之案件
   stSQL = "Select Decode(Substr(St15,1,1),'S',St06,'1')||Cu12||Cu13||Cu01||Cu02||Nvl(Tm123,Cu127) Srt, " & _
                 "tm01 , tm02, tm03, tm04, Tm23, cp13, st06 " & _
                 "From Caseprogress, Trademark, Customer, Staff " & _
                 "Where Cp158>=20210101 and cp159=0 And Cp01='CFT' And Cp10='101' And Cp01=Tm01(+) And Cp02=Tm02(+) And Cp03=Tm03(+) And Cp04=Tm04(+) " & _
                 "And Tm10='048' And Tm57 Is Null And Cu01(+)=Substr(Tm23,1,8) And Cu02='0' And St01(+)=Cu13 " & _
                 "order by 1,2,3,4,5 "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Label4.Caption = "本所信函進度檔建立中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents

         '先收文本所信函CFT1799
         stSQL = "update caseprogress set cp28=cp28 where cp01='" & .Fields("tm01") & "' and cp02='" & .Fields("tm02") & "' and cp03='" & .Fields("tm03") & "' and cp04='" & .Fields("tm04") & "' and instr(cp64,'緬甸商標法通知')>0 and cp05>=20210120"
         cnnConnection.Execute stSQL, intQ
         If intQ = 0 Then
            stCP10 = "1799"
            stCP14 = "A6034"
            stCP13 = PUB_GetAKindSalesNo(.Fields("tm01"), .Fields("tm02"), .Fields("tm03"), .Fields("tm04"))
            stCP12 = GetSalesArea(stCP13)
            stCP64 = "緬甸商標法通知;"
            cnnConnection.BeginTrans
            bolInTrans = True
            stCP09 = AutoNo("D", 6)

            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp64 ) values ('" & .Fields("tm01") & "'" & _
               ",'" & .Fields("tm02") & "','" & .Fields("tm03") & "','" & .Fields("tm04") & "'," & strSrvDate(1) & _
               ",'" & stCP09 & "','" & stCP10 & "','" & stCP12 & "'" & _
               ",'" & stCP13 & "','" & stCP14 & "','N','N',20210122,'N','" & stCP64 & "')"
            cnnConnection.Execute stSQL, intQ

            PUB_AddLetterProgress stCP09, 0, True, , True, .Fields("tm23"), stCP10

            stSQL = "update caseprogress set cp28=cp09,cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ
            If intQ = 1 Then
               'Trigger 會寫發文人故要另外更新
               stSQL = "update caseprogress set cp154='QPGMR' where cp09='" & stCP09 & "'"
               cnnConnection.Execute stSQL, intQ
            End If
            '自動確認
            stSQL = "update letterprogress set lp06='QPGMR',lp07=to_char(sysdate,'YYYYMMDD') where lp01='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ

            cnnConnection.CommitTrans
            bolInTrans = False
            iRecCnt = iRecCnt + 1
         End If

         .MoveNext
      Loop
      End With
      
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
      
      If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Exit Sub
      End If
         
         
      stSQL = "select decode(substr(st15,1,1),'S',st06,'1')||CU12||CU13||cu01||cu02||NVL(TM123,CU127) Srt" & _
         ",tm01,tm02,tm03,tm04,cp09,tm09 from caseprogress,trademark,customer,staff" & _
         " where cp05>=20210901 and cp01='CFT' and cp10='1799' and instr(cp64,'緬甸商標法通知')>0" & _
         " and TM01(+)=cp01 and TM02(+)=cp02 and TM03(+)=cp03 and TM04(+)=cp04" & _
         " and cu01(+)=substr(tm23,1,8) and cu02(+)='0' and st01(+)=cu13" & _
         " order by 1,2,3,4,5"
      
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With rsQuery
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         strUserNum = "A6034"
         Do While Not .EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            stCP09 = .Fields("cp09")
            stET03 = "08"

            m_DocSNo = Format(.AbsolutePosition, "000")
            Debug.Print m_DocSNo & " --> " & Now
            NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , True, , True, , , stCP09
            'Memo by Lydia 2021/09/03 需要將LD27改為CUS，才會自動存PDF轉卷宗區
            m_DocSNo = ""
            Sleep 2000
            .MoveNext
         Loop
         End With
         Label4.Caption = "通知信產生完成！"
      End If
   End If
   
ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
   
End Sub

'Added by Morgan 2023/2/21 112UP通知
'定稿加|#自動公司章#|時不可用<信函下款>否則全E化案件會重複蓋章，要改用<出名公司>
'掛號直寄信非E化案件(lp11='Y' and lp26 null)自動確認，半E(lp11 isnull)或全E(lp26='E')仍由智權人員確認(下次類似的通知需再確認全E是否也應自動確認)
Private Sub Process16()
      
   Dim stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP10 As String, stCP12 As String, stCP13 As String, stCP14 As String
   Dim rsQuery As ADODB.Recordset, intQ As Integer, iRecCnt As Integer
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
   
   strUserNo = strUserNum
       
   'EPC申請中案件
   stSQL = " select a0916,decode(substr(st15,1,1),'S',st06,'0')||CU12||CU13||cu01||cu02||NVL(PA149,CU127) Srt" & _
      ",pa01,pa02,pa03,pa04,pa26 CuNo,cp09" & _
      " from patent,customer,staff,acc090,caseprogress" & _
      " where pa09='221' and pa23='1' and pa57||pa108 is null and nvl(pa16,'2')='2' and pa11 is not null" & _
      " and exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='101' and cp05>20170000)" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)='0' and st01(+)=cu13 and a0901(+)=st15" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
      " and cp10(+)='1999' and cp05(+)>=20230221 and instr(cp64(+),'UP制度簡介')>0" & _
      " order by 1,2,3,4,5,6"
   
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Label4.Caption = "本所信函進度檔建立中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         '先收文本所信函CFP1999
         If Not IsNull(.Fields("cp09")) Then
            stCP09 = .Fields("cp09")
         Else
            stCP10 = "1999"
            stCP13 = PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
            stCP12 = GetSalesArea(stCP13)
            strUserNum = PUB_GetCFPHandler(.Fields("pa01") & "-" & .Fields("pa02") & "-" & .Fields("pa03") & "-" & .Fields("pa04"))
            
            cnnConnection.BeginTrans
            bolInTrans = True
            stCP09 = AutoNo("D", 6)
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp64 ) values ('" & .Fields("pa01") & "'" & _
               ",'" & .Fields("pa02") & "','" & .Fields("pa03") & "','" & .Fields("pa04") & "'," & strSrvDate(1) & _
               ",'" & stCP09 & "','" & stCP10 & "','" & stCP12 & "'" & _
               ",'" & stCP13 & "','" & strUserNum & "','N','N',20230223,'N','UP制度簡介;')"
            cnnConnection.Execute stSQL, intQ
            
            '掛號
            PUB_AddLetterProgress stCP09, 0, True, , True, .Fields("CuNo"), stCP10

            '自動發文室發文
            stSQL = "update caseprogress set cp28=cp09,cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ
               
            '更新建立人,發文人,發文室發文人
            stSQL = "update caseprogress set cp65='QPGMR',cp83='QPGMR',cp154='QPGMR' where cp09='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ
            
            '直寄且非E化的上自動確認
            stSQL = "update letterprogress set lp06='QPGMR',lp07=to_char(sysdate,'YYYYMMDD') where lp01='" & stCP09 & "' and lp11='Y' and lp26 is null"
            cnnConnection.Execute stSQL, intQ
            
            cnnConnection.CommitTrans
            
            bolInTrans = False
            iRecCnt = iRecCnt + 1
         End If
         .MoveNext
      Loop
      End With
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
      
      If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         If strUserNo <> "" Then strUserNum = strUserNo
         Exit Sub
      End If
      
      stSQL = "select * from caseprogress where cp05>=20230221 and cp01='CFP' and cp10='1999' and instr(cp64,'UP制度簡介')>0" & _
         " order by cp14,cp09"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With rsQuery
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         stCP14 = .Fields("cp14")
         Do While Not .EOF
            '不同承辦停下來存檔
            If stCP14 <> .Fields("cp14") Then
               MsgBox stCP14 & "案件已列印完成！"
               stCP14 = .Fields("cp14")
            End If
            
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            stCP09 = .Fields("cp09")
            strUserNum = .Fields("cp14")
            stET03 = "12"
            NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , True, , , stCP09
            .MoveNext
         Loop
         End With
         
         Label4.Caption = "通知信產生完成！"
      End If
   End If

ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
   
End Sub

'Added by Morgan 2023/5/16 UPC選擇退出通知函
'定稿加|#自動公司章#|時不可用<信函下款>否則全E化案件會重複蓋章，要改用<出名公司>
'自動上發文室發文但不自動確認
'半E(lp11 isnull)或全E(lp26='E')不印紙本並改設定為非掛號(LP52=null)，由智權自行決定寄送方式
'確認定稿無誤後再整批上發文
Private Sub Process17()
      
   Dim stET03 As String
   Dim strUserNo As String
   Dim stSQL As String, stCP09 As String, stCP10 As String, stCP12 As String, stCP13 As String, stCP14 As String
   Dim rsQuery As ADODB.Recordset, intQ As Integer, iRecCnt As Integer
   Dim bolInTrans As Boolean
   
On Error GoTo ErrHnd
   
   strUserNo = strUserNum
       
   '已領證的EPC專利案，排除專利權期間在2030年6月1日前屆滿及子案不包含UP會員國的案件
   stSQL = " select a0916,decode(substr(st15,1,1),'S',st06,'0')||CU12||CU13||cu01||cu02||NVL(PA149,CU127) Srt" & _
      ",pa01,pa02,pa03,pa04,pa26 CuNo,cp09" & _
      " from nextprogress,patent,customer,staff,acc090,caseprogress" & _
      " where np02='CFP' and np07='250'" & _
      " and pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05" & _
      " and pa09='221' and pa25>20300601 and pa57||pa108 is null" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)='0' and st01(+)=cu13 and a0901(+)=st15" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
      " and cp10(+)='1999' and cp05(+)>=20230516 and instr(cp64(+),'UPC選擇退出通知函')>0" & _
      " order by 1,2,3,4,5,6"
   
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Label4.Caption = "本所信函進度檔建立中..."
      iRecCnt = 0
      ProgressBar1.max = .RecordCount
      ProgressBar1.Value = 0
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      Do While Not .EOF
         ProgressBar1.Value = ProgressBar1.Value + 1
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         DoEvents
         '先收文本所信函CFP1999
         If Not IsNull(.Fields("cp09")) Then
            stCP09 = .Fields("cp09")
         Else
            stCP10 = "1999"
            stCP13 = PUB_GetAKindSalesNo(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
            stCP12 = GetSalesArea(stCP13)
            strUserNum = PUB_GetCFPHandler(.Fields("pa01") & "-" & .Fields("pa02") & "-" & .Fields("pa03") & "-" & .Fields("pa04"))
            
            cnnConnection.BeginTrans
            bolInTrans = True
            stCP09 = AutoNo("D", 6)
            stSQL = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
               ",cp12,cp13,cp14,cp20,cp26,cp32,cp64 ) values ('" & .Fields("pa01") & "'" & _
               ",'" & .Fields("pa02") & "','" & .Fields("pa03") & "','" & .Fields("pa04") & "'," & strSrvDate(1) & _
               ",'" & stCP09 & "','" & stCP10 & "','" & stCP12 & "'" & _
               ",'" & stCP13 & "','" & strUserNum & "','N','N','N','UPC選擇退出通知函;')"
            cnnConnection.Execute stSQL, intQ
            
            '掛號
            PUB_AddLetterProgress stCP09, 0, True, , True, .Fields("CuNo"), stCP10

            '自動發文室發文
            stSQL = "update caseprogress set cp28=cp09,cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ
               
            '更新建立人,發文室發文人
            stSQL = "update caseprogress set cp65='QPGMR',cp154='QPGMR' where cp09='" & stCP09 & "'"
            cnnConnection.Execute stSQL, intQ
            
            '改為非掛號(不確收)
            stSQL = "update letterprogress set lp52=null where lp01='" & stCP09 & "' and lp52='Y'"
            cnnConnection.Execute stSQL, intQ
            
            '半E/全E都由智權自行決定寄送方式
            stSQL = "update letterprogress set lp11=null where lp01='" & stCP09 & "' and lp26 is not null and lp11='Y'"
            cnnConnection.Execute stSQL, intQ
            
            cnnConnection.CommitTrans
            
            bolInTrans = False
            iRecCnt = iRecCnt + 1
         End If
         .MoveNext
      Loop
      End With
      Label4.Caption = "已完成共新增 " & iRecCnt & " 筆本所信函進度!!"
      
      If MsgBox(Label4.Caption & vbCrLf & "是否要繼續產生通知信??", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         If strUserNo <> "" Then strUserNum = strUserNo
         Exit Sub
      End If
      
      stSQL = "select * from caseprogress,letterprogress where cp05>=20230516 and cp01='CFP' and cp10='1999' and instr(cp64,'UPC選擇退出通知函')>0 and lp01(+)=cp09" & _
         " order by cp14,cp09"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         With rsQuery
         Label4.Caption = "通知信產生中..."
         ProgressBar1.max = .RecordCount
         ProgressBar1.Value = 0
         lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
         stCP14 = .Fields("cp14")
         Do While Not .EOF
            '不同承辦停下來存檔
            If stCP14 <> .Fields("cp14") Then
               MsgBox stCP14 & "案件已列印完成！"
               stCP14 = .Fields("cp14")
            End If
            
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
            DoEvents
            
            stCP09 = .Fields("cp09")
            strUserNum = .Fields("cp14")
            stET03 = "13"
            If Not IsNull(.Fields("lp26")) Then
               NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , , , , stCP09, , , , , True
            Else
               NowPrint stCP09, "21", stET03, False, strUserNum, 0, , , , , , , , , True, , , stCP09
            End If
            .MoveNext
         Loop
         End With
         MsgBox stCP14 & "案件已列印完成！"
         Label4.Caption = "通知信產生完成！"
      End If
   End If

ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   If strUserNo <> "" Then strUserNum = strUserNo
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040151 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 And Not IsNumeric(KeyAscii) Then
   KeyAscii = 0
   Beep
End If
End Sub

Private Sub txtCust_GotFocus()
   TextInverse txtCust
End Sub

Private Sub txtCust_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCust_Validate(Cancel As Boolean)
   If txtCust <> "" Then
      If ClsPDGetCustomer(txtCust, strExc(1)) Then
         lblCust = strExc(1)
      Else
         lblCust = ""
      End If
   Else
      lblCust = ""
   End If
End Sub

Private Function PrintOnePdf(ByVal program_name As String, parameters As String) As Boolean

   Dim process_id As Long
   Dim process_handle As Long
    ' Start the program.
    On Error GoTo ShellError

    'Modified by Morgan 2017/5/15 路徑可能含空白,改加雙引號
    process_id = Shell("""" & program_name & """ " & parameters, vbHide)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
   
    PrintOnePdf = True
    Exit Function

ShellError:
    MsgBox Err.Number & ":" & Err.Description & "(" & program_name & ")", vbCritical
End Function
