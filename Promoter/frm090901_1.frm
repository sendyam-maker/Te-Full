VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090901_1 
   BorderStyle     =   1  '單線固定
   Caption         =   " 工作進度資料維護"
   ClientHeight    =   6440
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   6910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6440
   ScaleWidth      =   6910
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   105
      TabIndex        =   37
      Top             =   3660
      Width           =   6735
      _ExtentX        =   11889
      _ExtentY        =   4904
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   176
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frm090901_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(11)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDST05"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDST05Old"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPA162"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frm090901_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtAMD05"
      Tab(1).Control(1)=   "Label6"
      Tab(1).ControlCount=   2
      Begin VB.CommandButton Command1 
         Caption         =   "複製前次內容"
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         Top             =   1680
         Width           =   1365
      End
      Begin VB.TextBox txtPA162 
         Height          =   270
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   315
         Width           =   255
      End
      Begin MSForms.TextBox txtAMD05 
         Height          =   2025
         Left            =   -74880
         TabIndex        =   9
         Top             =   630
         Width           =   6405
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "11298;3572"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDST05Old 
         Height          =   705
         Left            =   120
         TabIndex        =   45
         Top             =   1980
         Width           =   6405
         VariousPropertyBits=   -1466941409
         BackColor       =   -2147483638
         ScrollBars      =   2
         Size            =   "11298;1244"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtDST05 
         Height          =   705
         Left            =   120
         TabIndex        =   7
         Top             =   930
         Width           =   6405
         VariousPropertyBits=   -1466941413
         MaxLength       =   1000
         ScrollBars      =   2
         Size            =   "11298;1244"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         Caption         =   "中說請款修正定稿文字"
         Height          =   255
         Left            =   -74880
         TabIndex        =   41
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "前次內容:"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "核准分割建議定稿文字:"
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否提供核准分割建議:        ( Y: 是 N:否 )"
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   3195
      End
   End
   Begin VB.TextBox txtEP33 
      Height          =   300
      Left            =   5490
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "txtEP33"
      Top             =   3255
      Width           =   915
   End
   Begin VB.TextBox txtEP08 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   5490
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   32
      Text            =   "txtEP08"
      Top             =   2960
      Width           =   915
   End
   Begin VB.TextBox txtCP48 
      Height          =   300
      Left            =   5490
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "txtCP48"
      Top             =   2280
      Width           =   915
   End
   Begin VB.TextBox txtEP09 
      Height          =   300
      Left            =   5490
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "txtEP09"
      Top             =   2620
      Width           =   915
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5850
      TabIndex        =   5
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   4605
      TabIndex        =   4
      Top             =   90
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3780
      TabIndex        =   3
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   300
      Index           =   1
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   270
      Width           =   495
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   300
      Index           =   2
      Left            =   1980
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   270
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   300
      Index           =   3
      Left            =   2820
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   270
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   300
      Index           =   4
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   270
      Width           =   375
   End
   Begin MSForms.Label lblCP64 
      Height          =   1275
      Left            =   1500
      TabIndex        =   48
      Top             =   2280
      Width           =   2505
      BackColor       =   -2147483638
      Size            =   "4419;2249"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblEP04C 
      Height          =   255
      Left            =   5460
      TabIndex        =   47
      Top             =   1980
      Width           =   1395
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Size            =   "2461;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCP14C 
      Height          =   255
      Left            =   2100
      TabIndex        =   46
      Top             =   1980
      Width           =   1365
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Size            =   "2408;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   255
      Index           =   3
      Left            =   1500
      TabIndex        =   44
      Top             =   1146
      Width           =   5200
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Size            =   "9172;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   255
      Index           =   2
      Left            =   1500
      TabIndex        =   43
      Top             =   869
      Width           =   5200
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Size            =   "9172;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   255
      Index           =   1
      Left            =   1500
      TabIndex        =   42
      Top             =   592
      Width           =   5200
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Size            =   "9172;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "核稿人:"
      Height          =   255
      Index           =   7
      Left            =   4185
      TabIndex        =   36
      Top             =   1980
      Width           =   585
   End
   Begin VB.Label lblEP04 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   4860
      TabIndex        =   35
      Top             =   1980
      Width           =   585
   End
   Begin VB.Label lblEP33 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "核稿完成日:"
      Height          =   255
      Left            =   4500
      TabIndex        =   34
      Top             =   3278
      Width           =   945
   End
   Begin VB.Label lblEP08 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "核稿期限:"
      Height          =   255
      Left            =   4680
      TabIndex        =   33
      Top             =   2953
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "承辦期限:"
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   31
      Top             =   2303
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   255
      Index           =   10
      Left            =   645
      TabIndex        =   30
      Top             =   2310
      Width           =   765
   End
   Begin VB.Label lblPA08T 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   4860
      TabIndex        =   29
      Top             =   1700
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   255
      Index           =   9
      Left            =   4005
      TabIndex        =   28
      Top             =   1700
      Width           =   765
   End
   Begin VB.Label lblCP14 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1500
      TabIndex        =   27
      Top             =   1980
      Width           =   585
   End
   Begin VB.Label lblCP09 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   4860
      TabIndex        =   26
      Top             =   1423
      Width           =   1305
   End
   Begin VB.Label lblCP10C 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1500
      TabIndex        =   25
      Top             =   1700
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "總收文號:"
      Height          =   255
      Index           =   8
      Left            =   4005
      TabIndex        =   24
      Top             =   1423
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   255
      Index           =   4
      Left            =   825
      TabIndex        =   23
      Top             =   1980
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   255
      Index           =   3
      Left            =   645
      TabIndex        =   22
      Top             =   1700
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "收文日:"
      Height          =   255
      Index           =   2
      Left            =   825
      TabIndex        =   21
      Top             =   1423
      Width           =   585
   End
   Begin VB.Label lblCP05T 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1500
      TabIndex        =   20
      Top             =   1423
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "完稿日:"
      Height          =   255
      Index           =   1
      Left            =   4860
      TabIndex        =   19
      Top             =   2628
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   200
      Left            =   120
      TabIndex        =   18
      Top             =   619
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   200
      Left            =   1065
      TabIndex        =   17
      Top             =   619
      Width           =   345
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   200
      Left            =   1065
      TabIndex        =   16
      Top             =   869
      Width           =   345
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "(外):"
      Height          =   180
      Index           =   0
      Left            =   1065
      TabIndex        =   15
      Top             =   1140
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   255
      Index           =   0
      Left            =   645
      TabIndex        =   14
      Top             =   293
      Width           =   765
   End
End
Attribute VB_Name = "frm090901_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/23 改成Form2.0 ;lblCaseName(index)、lblCP14C、lblEP04C、lblCP64、txtDST05、txtDST05Old、txtAMD05
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Public NextFormName As String
Dim bolActive As Boolean
Dim m_CP10 As String
Dim bolSetData As Boolean
'Added by Lydia 2015/06/04
Public bolAMD As Boolean '是否經由"補輸中說"
Dim bolAMDset As Boolean '已發文之"補輸中說"

Private Sub cmdBack_Click()
   'Add by Morgan 2008/10/15
   If NextFormName = "frm060204" Then
      cmdExit_Click
   Else
   'end 2008/10/15
      frm090901.m_AMD = bolAMD 'Added by Lydia 2015/06/04
      Call frm090901.SetGrid(False)
      frm090901.Show
      Unload Me
   End If
End Sub

Public Sub SetData(ByRef rstGrid As ADODB.Recordset, ByVal iRow As Integer)
    
    Dim ii As Integer

    With frm090901
      For ii = 1 To 4
          txtCaseNo(ii) = .txtCaseNo(ii)
      Next ii
      For ii = 1 To 3
          lblCaseName(ii) = .lblCaseName(ii)
      Next ii
     'memo by Lydia 2015/04/24 預設表單寬度
      '.Width = 7020 .Height=6135
    End With
    
    With rstGrid
      .Move iRow - 1, adBookmarkFirst
      lblCP05T = "" & .Fields("CP05T")
      lblCP09 = "" & .Fields("CP09")
      lblCP10C = "" & .Fields("CP10C")
      lblPA08T = "" & .Fields("PA08T")
      m_CP10 = "" & .Fields("cp10")
      
      lblCP14 = "" & .Fields("CP14")
      lblCP14C = "" & .Fields("CP14C")
      '承辦期限
      txtCP48 = TransDate("" & .Fields("CP48"), 1)
      '完稿日
      txtEP09 = TransDate("" & .Fields("EP09"), 1)
      txtCP48.Tag = txtCP48
      txtEP09.Tag = txtEP09
      
      '進度備註
      lblCP64 = "" & .Fields("CP64")
      If m_CP10 = "201" Then
         'Modified by Morgan 2012/5/18
         '欄位分開
         'Label1(4) = "核稿人:"
         'lblCP14 = "" & .Fields("EP04")
         'lblCP14C = "" & .Fields("EP04C")
         'Label1(12) = "核稿期限:"
         'txtCP48 = TransDate("" & .Fields("EP08"), 1)
         'Label1(1) = "核稿完成日:"
         'txtEP09 = TransDate("" & .Fields("EP33"), 1)
         'txtCP48.Tag = txtCP48
         'txtEP09.Tag = txtEP09
         txtCP48.Enabled = False
         txtEP09.Enabled = False
         Label1(7).Visible = True
         lblEP04.Visible = True
         lblEP04C.Visible = True
         lblEP08.Visible = True
         txtEP08.Visible = True
         lblEP33.Visible = True
         txtEP33.Visible = True
         lblEP04 = "" & .Fields("EP04")
         lblEP04C = "" & .Fields("EP04C")
         txtEP08 = TransDate("" & .Fields("EP08"), 1)
         'Modify By Sindy 2023/10/30
         If strSrvDate(1) >= FCP核完日改用EP39 Then
            txtEP33 = TransDate("" & .Fields("EP39"), 1)
         Else
         '2023/10/30 END
            txtEP33 = TransDate("" & .Fields("EP33"), 1)
         End If
         'end 2012/5/18
      Else
         'Added by Morgan 2012/5/18
         Label1(7).Visible = False
         lblEP04.Visible = False
         lblEP04C.Visible = False
         lblEP08.Visible = False
         txtEP08.Visible = False
         lblEP33.Visible = False
         txtEP33.Visible = False
         'end 2012/5/18
         
         If txtEP09 <> "" Then
            txtCP48.Enabled = False
            'Add by Morgan 2008/10/21 管制人才能改
            If strUserNum = lblCP14 Then
               txtEP09.Enabled = False
            End If
         ElseIf m_CP10 <> "926" Then
            txtCP48.Enabled = False
         End If
      End If
      
      txtEP08.Tag = txtEP08
      txtEP33.Tag = txtEP33
    End With
    
   'Added by Morgan 2012/11/30
   txtPA162.Enabled = False
   txtPA162.Tag = "" 'Added by Morgan 2022/8/1
   txtDST05.Locked = True
   Command1.Enabled = False
   'Added by Lydia 2015/04/24 +中說請款修正定稿文字
   'Modified by Lydia 2015/08/27 為了能拉動卷軸,改成locked
   'txtAMD05.Enabled = False
   txtAMD05.Locked = True
   
   Me.Height = 3885  'lydia 無初審的高度
   'Modified by Morgan 2019/12/30 +107,203
   If m_CP10 = "204" Or m_CP10 = "205" Or m_CP10 = "1001" Or m_CP10 = "203" Or m_CP10 = "107" Then
      intI = 1
      'Added by Morgan 2012/12/13 日文定稿不可輸分割建議
      'Removed by Morgan 2022/10/11 取消,改也可輸入
      'If strUserNum = lblCP14 Then
      '   strExc(1) = PUB_GetLanguage(txtCaseNo(1), txtCaseNo(2), txtCaseNo(3), txtCaseNo(4))
      '   If strExc(1) = "3" Then
      '      intI = 0
      '   End If
      'End If
      'end 2022/10/11
      'end 2012/12/13
      If intI = 1 Then
         Me.SSTab1.TabVisible(1) = False: Me.SSTab1.TabCaption(0) = ""  'Added by Lydia 2015/04/24
         'Modified by Morgan 2019/8/2 開放發明新型都可輸入分割建議,不必限定發明初審--淑華 Ex:FCP-50474申復
         'If m_CP10 = "1001" Then
         '   '2013/9/16 modify by sonia 加入b.cp10='307' or b.cp10='301'(FCP-045402)
         '   strExc(0) = "select pa162,DST05 from caseprogress a,patent,divsugtext" & _
         '      " where cp09='" & lblCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa08='1'" & _
         '      " and dst01(+)=pa01 and dst01(+)=pa01 and dst02(+)=pa02 and dst03(+)=pa03 and dst04(+)=pa04" & _
         '      " and exists(select * from caseprogress b where b.cp09=a.cp43 and (b.cp10='101' or b.cp10='307' or b.cp10='301'))"
         '
         'Else
         '   '2013/9/16 modify by sonia 加入b.cp10='307' or b.cp10='301'(FCP-045402)
         '   strExc(0) = "select pa162,DST05 from caseprogress a,patent,divsugtext" & _
         '      " where cp09='" & lblCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa08='1'" & _
         '      " and dst01(+)=pa01 and dst01(+)=pa01 and dst02(+)=pa02 and dst03(+)=pa03 and dst04(+)=pa04" & _
         '      " and exists(select * from caseprogress b where b.cp01=pa01 and b.cp02=pa02 and b.cp03=pa03 and b.cp04=pa04 and (b.cp10='101' or b.cp10='307' or b.cp10='301') and b.cp27>0)" & _
         '      " and exists(select * from caseprogress b where b.cp01=pa01 and b.cp02=pa02 and b.cp03=pa03 and b.cp04=pa04 and b.cp10='416' and b.cp27>0)" & _
         '      " and not exists(select * from caseprogress b where b.cp01=pa01 and b.cp02=pa02 and b.cp03=pa03 and b.cp04=pa04 and b.cp10='107')"
         'End If
            strExc(0) = "select pa162,DST05 from caseprogress a,patent,divsugtext" & _
               " where cp09='" & lblCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa08 in ('1','2')" & _
               " and dst01(+)=pa01 and dst01(+)=pa01 and dst02(+)=pa02 and dst03(+)=pa03 and dst04(+)=pa04"
         'end 2019/8/2
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            txtPA162 = "" & RsTemp(0)
            txtPA162.Tag = txtPA162.Text 'Added by Morgan 2022/8/1
            If strUserNum = lblCP14 Then
               'Memo by Lydia 承辦人可輸入建議定稿文字,欄位清空,下方保留上次記錄文字
               txtDST05Old = "" & RsTemp(1)
               txtPA162.Enabled = True
               txtDST05.Locked = False
               Command1.Enabled = True
               'Modified by Lydia 2015/04/24
               'Me.Height = 6135
               Me.Height = 6810
               If m_CP10 = "1001" Then
                  txtEP09.Enabled = False
               End If
            Else
               'Memo by Lydia 非承辦人只show建議定稿文字
               txtDST05 = "" & RsTemp(1)
               'Modified by Lydia 2015/04/24
               'Me.Height = 5100
                Me.Height = 5640
            End If
            
         End If
      End If
      
   'Added by Lydia 2015/04/24 +中說請款修正定稿文字
   'Added by Lydia 2015/06/25 +主動修正203
   ElseIf InStr("201,209,210,235,203", m_CP10) > 0 Then
      'Added by Lydia 2015/06/25 判斷未經過"補輸中說"的主動修正,是否符合
      If frm090901_1.bolAMD = False And m_CP10 = "203" Then
         strExc(0) = "select count(*) from caseprogress where cp01='" & txtCaseNo(1) & "' and cp02='" & txtCaseNo(2) & "' and cp03='" & txtCaseNo(3) & "' and cp04='" & txtCaseNo(4) & "' " & _
               " and cp57 is null and cp10 in ('201','209','210','235') "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then bolSetData = True: Exit Sub
      End If
      'end 2015/06/25
      
      Me.SSTab1.TabVisible(0) = False: Me.SSTab1.TabCaption(1) = ""
      'Modified by Lydia 2015/11/26 AMD05長度已達2000字,Text高度拉長
      'Me.Height = 5400
      Me.Height = 6810
      'Modified by Lydia 2015/06/04 +CP27
      strExc(0) = "select AMD05,nvl(CP27,0) CP27 from caseprogress a,patent,Amendedtext" & _
                  " where cp09='" & lblCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 " & _
                  " and amd01(+)=pa01 and amd01(+)=pa01 and amd02(+)=pa02 and amd03(+)=pa03 and amd04(+)=pa04"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtAMD05.Tag = "" & RsTemp(0)
         txtAMD05.Text = txtAMD05.Tag
         'Modified by Lydia 2015/06/04 已發文只可更改-修正定稿文字
         If RsTemp.Fields("CP27") > 0 Then
            txtEP09.Enabled = False
            txtEP33.Enabled = False
            bolAMDset = True
         End If
      End If
      'Modified by Lydia 2015/06/12 除了承辦人外,核稿人也可輸入中說請款備註
      'If strUserNum = lblCP14 Or Pub_StrUserSt03 = "M51" Then
      'MODIFY BY SONIA 2015/6/17 再加下班翻譯的所內編號也可以輸 FCP-051573
      'If strUserNum = lblCP14 Or strUserNum = lblEP04 Or Pub_StrUserSt03 = "M51" Then
      If strUserNum = lblCP14 Or strUserNum = lblEP04 Or PUB_GetMapID(strUserNum, 0) = lblCP14 Or Pub_StrUserSt03 = "M51" Then
         'Modified by Lydia 2015/08/27 為了能拉動卷軸,改成locked
         ' txtAMD05.Enabled = True
          txtAMD05.Locked = False
      End If
   'end 2015/04/24
   End If
   'end 2012/11/30
    
    bolSetData = True
    
End Sub

Private Function FormSave() As Boolean
Dim strToM As String 'Added by Lydia 2020/08/24 外專工程師主管
Dim strTo As String, strSub As String, strContent As String '收件人,主旨,內文 'Added by Morgan 2022/8/1

On Error GoTo flgError

cnnConnection.BeginTrans
    
   strToM = PUB_GetFCPEngSup(strUserNum) 'Added by Lydia 2020/08/24 外專工程師主管
 
   If txtCP48 <> txtCP48.Tag Then
      strSql = " Update CASEPROGRESS Set CP48=" & CNULL(DBDATE(txtCP48), True) & " Where CP09='" & lblCP09 & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   
   If txtEP09 <> txtEP09.Tag Then
      'Modified by Morgan 2012/5/18
      '欄位分開
      'If m_CP10 = "201" Then
      '   strSql = " Update engineerPROGRESS Set ep33=" & CNULL(DBDATE(txtEP09), True) & " Where ep02='" & lblCP09 & "'"
      'Else
         strSql = " Update engineerPROGRESS Set ep09=" & CNULL(DBDATE(txtEP09), True) & " Where ep02='" & lblCP09 & "'"
      'End If
      'end 2012/5/18
      
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      
      'Added by Morgan 2012/5/17
      '電話聯絡單完搞日輸入自動發Mail給主管
      If txtEP09.Tag = "" Then
         If m_CP10 = "945" Then
             'Modified by Lydia 2020/08/24 改用模組取得
             'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " select '" & strUserNum & "' mc01,oMan mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
               ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'('||cp09||')電話聯絡單已完稿請發文!(完搞日：" & txtEP09 & ") ' mc07" & _
               ",'如旨' mc08 from caseprogress,staff,SetSpecMan" & _
               " where cp09='" & lblCP09 & "' and cp27 is null and st01(+)=cp14 and OCODE=decode(st16,'1','T','2','R','3','S','4','T1')"
             strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " select '" & strUserNum & "' mc01,'" & strToM & "' mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
               ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'('||cp09||')電話聯絡單已完稿請發文!(完搞日：" & txtEP09 & ") ' mc07" & _
               ",'如旨' mc08 from caseprogress " & _
               " where cp09='" & lblCP09 & "' and cp27 is null "
            cnnConnection.Execute strSql, intI
         End If
      End If
   End If
   
   'Added by Morgan 2012/5/18
   If txtEP33 <> txtEP33.Tag Then
      'Modify By Sindy 2023/10/30 EP33要回歸用在英文核完日,改抓EP39.核稿完成日
      If strSrvDate(1) >= FCP核完日改用EP39 Then
         strSql = " Update engineerPROGRESS Set ep39=" & CNULL(DBDATE(txtEP33), True) & " Where ep02='" & lblCP09 & "'"
      Else
      '2023/10/30 END
         strSql = " Update engineerPROGRESS Set ep33=" & CNULL(DBDATE(txtEP33), True) & " Where ep02='" & lblCP09 & "'"
      End If
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   'end 2012/5/18
   
   'Added by Morgan 2012/12/3
   If txtPA162.Enabled = True Then
      strSql = "update patent set pa162='" & txtPA162 & "' where pa01='" & txtCaseNo(1) & "' and pa02='" & txtCaseNo(2) & "' and pa03='" & txtCaseNo(3) & "' and pa04='" & txtCaseNo(4) & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      
      strSql = "delete divsugtext where dst01='" & txtCaseNo(1) & "' and dst02='" & txtCaseNo(2) & "' and dst03='" & txtCaseNo(3) & "' and dst04='" & txtCaseNo(4) & "'"
      cnnConnection.Execute strSql, intI
      
      strSql = "insert into divsugtext(dst01,dst02,dst03,dst04,dst05,dst06,dst07,dst08,dst09) values " & _
         "('" & txtCaseNo(1) & "','" & txtCaseNo(2) & "','" & txtCaseNo(3) & "','" & txtCaseNo(4) & "'" & _
         ",'" & ChgSQL(txtDST05) & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & lblCP09 & "')"
         
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
      
      If txtDST05 <> "" Then
         strExc(1) = "'本所案號：'||pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)||chr(13)||chr(10)" & _
                     "||'案件名稱：'||pa05||chr(13)||chr(10)" & _
                     "||'申請人：'||cu04||chr(13)||chr(10)" & _
                     "||'承辦期限：'||sqldatet(cp48)||chr(13)||chr(10)" & _
                     "||'核准分割建議定稿文字：'||dst05||chr(13)||chr(10)"
         
         If m_CP10 = "1001" Then
            strExc(2) = "cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'('||cp09||')已輸入核准分割建議定稿文字，請審核後至系統上完稿日，再將卷宗交各區程序上發文日!'"
         Else
            strExc(2) = "cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'('||cp09||')已輸入核准分割建議定稿文字，請審核!'"
         End If
         
         'Modified by Lydia 2020/08/24 改用模組取得
         'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " select '" & strUserNum & "' mc01,decode(oMan,st01,B0102,oMan) mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
            "," & strExc(2) & " mc07," & strExc(1) & " mc08" & _
            " from caseprogress,patent,customer,divsugtext,staff,SetSpecMan,ABS001" & _
            " where cp09='" & lblCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
            " and dst01(+)=cp01 and dst02(+)=cp02 and dst03(+)=cp03 and dst04(+)=cp04" & _
            " and st01(+)=cp14 and OCODE=decode(st16,'1','T','2','R','3','S','4','T1') and B0101(+)=st01"
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " select '" & strUserNum & "' mc01,'" & strToM & "' mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
            "," & strExc(2) & " mc07," & strExc(1) & " mc08" & _
            " from caseprogress,patent,customer,divsugtext " & _
            " where cp09='" & lblCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
            " and dst01(+)=cp01 and dst02(+)=cp02 and dst03(+)=cp03 and dst04(+)=cp04"
            
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2012/12/3
   
   'Added by Lydia 2015/04/24 +中說請款修正定稿文字
   'Modified by Lydia 2015/08/27 為了能拉動卷軸,改成locked
   'If txtAMD05.Enabled = True And txtAMD05.Text <> txtAMD05.Tag Then
   If txtAMD05.Locked = False And txtAMD05.Text <> txtAMD05.Tag Then
      strSql = "delete AmendedText where AMD01='" & txtCaseNo(1) & "' and AMD02='" & txtCaseNo(2) & "' and AMD03='" & txtCaseNo(3) & "' and AMD04='" & txtCaseNo(4) & "'"
      cnnConnection.Execute strSql, intI
      
      strSql = "insert into AmendedText(AMD01,AMD02,AMD03,AMD04,AMD05,AMD06,AMD07,AMD08,AMD09) values " & _
         "('" & txtCaseNo(1) & "','" & txtCaseNo(2) & "','" & txtCaseNo(3) & "','" & txtCaseNo(4) & "'" & _
         ",'" & ChgSQL(txtAMD05) & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & lblCP09 & "')"
         
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql, intI
   End If
   
   'Added by Morgan 2022/8/1--Winfrey
   If m_CP10 = "1001" Then
      strSub = ""
      '1.不需加註分割建議，email通知各區程序上核准發文。
      '2.分割建議主管上完稿日，email通知各區程序上核准發文。
      If (txtPA162.Enabled = True And txtPA162 = "N") Then
         strSub = "【工程師已確認不須分割加註】請進行告准 Our Ref: "
      ElseIf (txtPA162 = "Y" And txtEP09.Enabled And txtEP09 <> "") Then
         'Memo by Morgan 2022/10/11 因為日文定稿還是要給工程師核稿,主旨保留以作識別
         If PUB_GetLanguage(txtCaseNo(1), txtCaseNo(2), txtCaseNo(3), txtCaseNo(4)) = "3" Then
            strSub = "【工程師已完成分割加註(日文定稿)】請進行告准 Our Ref: "
         Else
            strSub = "【工程師已完成分割加註】請進行告准 Our Ref: "
         End If
      End If
      If strSub <> "" Then
         strTo = PUB_GetFCPHandler(txtCaseNo(1), txtCaseNo(2), txtCaseNo(3), txtCaseNo(4))
         strContent = "1.請程序上發文日。" & vbCrLf & "2.請告准人員進行後續告准，感謝您。"

         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " select '" & strUserNum & "' mc01,'" & strTo & "' mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
            ",'" & strSub & "'||cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) mc07,'" & strContent & "' mc08,cp14 mc09" & _
            " from caseprogress  where cp43='" & lblCP09 & "' and cp10='1917' and cp27 is null"
            
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2022/8/1

   cnnConnection.CommitTrans
   FormSave = True

flgError:
   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If

End Function

Private Sub cmdExit_Click()
   Unload frm090901
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If TxtValidate = True Then
      If FormSave() = True Then
         PUB_SendMailCache 'Add by Morgan 2012/5/17
         'Add by Morgan 2008/9/22
         If NextFormName = "frm060204" Then
            If PUB_IsFormExist(NextFormName) = False Then
               cmdExit_Click
            Else
               bolSetData = False
               frm060204.PubShowNextData True
               If bolSetData = False Then
                  cmdExit_Click
               Else
                  bolActive = False
                  Form_Activate
               End If
            End If
         Else
         'end 2008/9/22
            cmdBack_Click
         End If
      Else
         MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
      End If
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   txtCP48_Validate bCancel
   If bCancel = True Then
      txtCP48.SetFocus
      txtCP48_GotFocus
      Exit Function
   End If
   
   txtEP09_Validate bCancel
   If bCancel = True Then
      txtEP09.SetFocus
      txtEP09_GotFocus
      Exit Function
   End If
   
   'Added by Morgan 2012/12/3
   If txtPA162.Enabled = True Then
      If txtPA162 = "" Then
         MsgBox "請設定是否要加註核准分割建議！", vbExclamation
         txtPA162.SetFocus
         Exit Function
      ElseIf txtPA162 = "Y" And Trim(txtDST05) = "" Then
         MsgBox "請輸入建議定稿文字！", vbExclamation
         txtDST05.SetFocus
         Exit Function
      ElseIf txtPA162 = "N" And Trim(txtDST05) <> "" Then
         MsgBox "當設定為""不要""加註核准分割建議時，不可輸入建議定稿文字！", vbExclamation
         txtDST05.SetFocus
         Exit Function
      End If
      
      'Added by Morgan 2012/12/26
      If txtDST05 <> "" Then
         strExc(0) = PUB_StringFilter(txtDST05)
         If strExc(0) <> txtDST05 Then
            If MsgBox("建議定稿文字內發現有跳行符號，存檔時將自動清除。是否要繼續??", vbYesNo + vbDefaultButton2) = vbYes Then
               txtDST05 = strExc(0)
            Else
               txtDST05.SetFocus
               Exit Function
            End If
         End If
      End If
      'end 2012/12/26
      
   End If
   If m_CP10 = "1001" And txtEP09 = "" And txtEP09.Enabled = True Then
      MsgBox "請輸入完稿日！", vbExclamation
      txtEP09.SetFocus
      Exit Function
   End If
   'end 2012/12/3
   
    'Added by Lydia 2021/09/23 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If
   
   TxtValidate = True
End Function

Private Sub Command1_Click()
   txtDST05.Text = txtDST05Old.Text
End Sub

Private Sub Form_Activate()
   If bolActive = False Then
      bolActive = True
      If txtEP33.Visible And txtEP33.Enabled Then
         txtEP33.SetFocus
         txtEP33_GotFocus
      ElseIf txtCP48.Enabled = True Then
         txtCP48.SetFocus
         txtCP48_GotFocus
      ElseIf txtEP09.Enabled = True Then
         txtEP09.SetFocus
         txtEP09_GotFocus
      'Added by Lydia 2015/06/04
      'Added by Lydia 2015/06/12 +判斷是否可輸入
      'Modified by Lydia 2015/08/27
      'ElseIf bolAMDset = True And txtAMD05.Enabled = True Then
      ElseIf bolAMDset = True And txtAMD05.Locked = False Then
           txtAMD05.SetFocus
      End If
   End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm090901_1 = Nothing
End Sub

Private Sub txtAMD05_GotFocus()
   TextInverse txtAMD05
End Sub

Private Sub txtCP48_GotFocus()
   TextInverse txtCP48
   CloseIme
End Sub

Private Sub txtCP48_Validate(Cancel As Boolean)
   If txtCP48 <> "" Then
      If txtCP48 <> txtCP48.Tag Then
         MsgBox "承辦期限只能取消不可修改！"
         txtCP48 = txtCP48.Tag
         txtCP48_GotFocus
         Cancel = True
      End If
   End If
End Sub

'Added by Lydia 2021/09/23 Form 2.0的TextBox增加右鍵選單功能
Private Sub txtDST05_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 txtDST05
End Sub

'Added by Lydia 2021/09/23 Form 2.0的TextBox增加右鍵選單功能
Private Sub txtAMD05_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then Forms(0).PopupMenu2 txtAMD05
End Sub

Private Sub txtEP09_GotFocus()
   TextInverse txtEP09
   CloseIme
End Sub

Private Sub txtEP09_Validate(Cancel As Boolean)
   If txtEP09 <> "" Then
      If Not ChkDate(txtEP09) Then
         txtEP09_GotFocus
         Cancel = True
      ElseIf Val(DBDATE(txtEP09)) > Val(strSrvDate(1)) Then
         MsgBox Left(Label1(1), Len(Label1(1)) - 1) & "不可大於系統日！"
         txtEP09_GotFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub txtEP33_GotFocus()
   TextInverse txtEP33
   CloseIme
End Sub

Private Sub txtEP33_Validate(Cancel As Boolean)
   If txtEP33 <> "" Then
      If Not ChkDate(txtEP33) Then
         txtEP33_GotFocus
         Cancel = True
      ElseIf Val(DBDATE(txtEP33)) > Val(strSrvDate(1)) Then
         MsgBox "核稿完成日不可大於系統日！"
         txtEP33_GotFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub txtPA162_GotFocus()
   CloseIme
   TextInverse txtPA162
End Sub

Private Sub txtPA162_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
