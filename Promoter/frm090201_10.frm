VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_10 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護－准駁案件明細資料"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1515
   ClientWidth     =   9315
   ControlBox      =   0   'False
   FillColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   8130
      TabIndex        =   11
      Top             =   105
      Width           =   850
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "繼續(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6301
      TabIndex        =   10
      Top             =   105
      Width           =   850
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "卷宗區(&O)"
      Height          =   400
      Index           =   1
      Left            =   7167
      TabIndex        =   9
      Top             =   105
      Width           =   945
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "全部選取"
      Height          =   400
      Index           =   3
      Left            =   5100
      TabIndex        =   8
      Top             =   105
      Width           =   1185
   End
   Begin VB.TextBox txtWorkDays 
      Alignment       =   2  '置中對齊
      Height          =   285
      Left            =   3150
      TabIndex        =   7
      Top             =   578
      Width           =   375
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢"
      Height          =   315
      Index           =   4
      Left            =   4365
      TabIndex        =   4
      Top             =   563
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   4668
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5784
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4590
      Left            =   90
      TabIndex        =   1
      Top             =   960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8096
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin MSForms.Label lblName 
      Height          =   255
      Left            =   1170
      TabIndex        =   12
      Top             =   330
      Width           =   1875
      VariousPropertyBits=   27
      Caption         =   "lblName"
      Size            =   "3307;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "個工作天)"
      Height          =   180
      Index           =   1
      Left            =   3570
      TabIndex        =   6
      Top             =   630
      Width           =   780
   End
   Begin VB.Label Label7 
      Caption         =   "("
      Height          =   180
      Left            =   3060
      TabIndex        =   5
      Top             =   630
      Width           =   75
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人： "
      Height          =   180
      Index           =   35
      Left            =   150
      TabIndex        =   2
      Top             =   360
      Width           =   1020
   End
   Begin VB.Label lblDate 
      Caption         =   "lblDate"
      Height          =   180
      Left            =   90
      TabIndex        =   3
      Top             =   630
      Width           =   3015
   End
End
Attribute VB_Name = "frm090201_10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/28 改成Form2.0 ; grd1改字型=新細明體-ExtB、lblName
'Create by Sindy 2021/2/22
Option Explicit

Public TextOk As Boolean
Public m_bolIsDrawer As Boolean '是否為繪圖人員
Public strContinue As Boolean   '判斷為繼續或結束

Dim k As Integer, i As Integer
Dim strTemp(0 To 21) As String
Dim StrGrp090201 As String, StrSQL6 As String, ChkNoData As Boolean
Dim m_dblWDBegin As Double '工作天起
Dim m_dblWDEnd As Double '工作天迄


Public Sub cmdOK_Click(Index As Integer)
Dim KeyWord As String

   Select Case Index
      Case 0 '繼續
         strContinue = True
         Unload Me
         
      Case 1 '卷宗區
         With GRD1
            For i = 1 To .Rows - 1
               .Visible = False
               If .TextMatrix(i, 0) = "V" Then
                  Screen.MousePointer = vbHourglass
                  frm100101_L.m_strKey = .TextMatrix(i, 4)
                  frm100101_L.SetParent Me
                  If frm100101_L.QueryData = True Then
                     GRD1.Tag = "Y"
                     .TextMatrix(i, 0) = ""
                     frm100101_L.Show
                     frm100101_L.Tag = .TextMatrix(i, 22) '總收文號
'                     strSql = "Update CaseProgress Set CP143=" & strSrvDate(1) & " Where CP09='" & .TextMatrix(i, 22) & "' "
'                     cnnConnection.Execute strSql
                     Me.Hide
                     Exit Sub
'                     Call RefreshData
                  Else
                     Unload frm100101_L
                  End If
                  Screen.MousePointer = vbDefault
               End If
            Next
            .Visible = True
         End With
         
         If GRD1.Tag = "Y" Then
            Call RefreshData
         Else
            MsgBox "請點選欲查看的資料"
            Exit Sub
         End If
         
      Case 2 '結束
         strContinue = False
         Unload Me
         
      Case 3 '全部選取
         GRD1.col = 0
         Screen.MousePointer = vbHourglass
         If Trim(cmdOK(3).Caption) = "全部選取" Then
             KeyWord = "V"
             cmdOK(3).Caption = "全部取消"
         Else
             KeyWord = ""
             cmdOK(3).Caption = "全部選取"
         End If
         For i = 1 To GRD1.Rows - 1
             GRD1.row = i
             GRD1.Text = KeyWord
         Next i
         Screen.MousePointer = vbDefault
         
      Case 4 '查詢
         Screen.MousePointer = vbHourglass
         StrMenu1
         Screen.MousePointer = vbDefault
   End Select
End Sub

Private Sub Form_Load()
   Me.Hide
   Screen.MousePointer = vbHourglass
   MoveFormToCenter Me

   '專利處改列本所期限前3個工作天提醒
   If Left(Pub_StrUserSt03, 2) = "P1" Then
      Me.Caption = Replace(Me.Caption, "已達", "3個工作天內達")
      txtWorkDays = 3
   Else
      txtWorkDays = 1
   End If
   If m_bolIsDrawer Then
      Label1(35) = "繪圖人員："
   End If
   
   StrMenu
   SetGrd1
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strContinue = True Then Nextstep
   Set frm090201_10 = Nothing
End Sub

Private Sub Nextstep()
   frm090201_9.TextOk = False
   frm090201_9.StrMenu1  '已完稿超過2天或已屆預定會稿日且尚未會稿案件資料,,無資料時由frm090201_9的nextstep執行下一畫面
   If frm090201_9.TextOk = True Then
      frm090201_9.Show
   End If
End Sub

Public Sub StrMenu1()
   Me.Hide
   strContinue = True: GRD1.Tag = ""
   DoEvents
   
On Error Resume Next '若 table 不存在時跳過
   adoEng.Execute "drop table R090614 "
On Error GoTo 0 '還原錯誤控制
   adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text,R110006 text,R110007 text,R110008 text,R110009 text,R110010 text,R110011 text,R110012 text,R110013 text,R110014 text,R110015 text,R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo,R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text,R110026 double,R110027 double,R110028 double,R110029 text,R110030 text)"
   
   If m_dblWDBegin = 0 Then
      m_dblWDEnd = strSrvDate(1)
      '3個工作天內的資料
      m_dblWDBegin = CompWorkDay(3, strSrvDate(1), 1)
   End If
   
   StrGrp090201 = ""
   StrSQL6 = ""
   
   '抓准駁案件明細資料:
   '在進入承辦人工作進度時，增加第三畫面，顯示准駁案件明細資料，
   '查出P及CFP最新進度為1001.核准或1002.核駁或1006最終核駁
   '且其相關總收文號為(101發明申請,102新型申請,103設計申請,104追加申請,105聯合申請,107再審申請,113.CIP申請,114.CPA申請,501訴願,424請求繼續審查,126期末拋棄,805復審, 301改請發明,302改請新型,303改請設計,304改請追加,305改請聯合,306改請獨立,307分割,
   '122.CA申請,438再考量試行計畫(AFCP2.0))
   '3個工作天內的資料供人員查看。看過卷宗區後自動消失。
   
   'Modified by Morgan 2016/4/18 +繪圖人員判斷
   StrSQL6 = StrSQL6 + " and " & IIf(m_bolIsDrawer, "CP29", "CP14") & "='" & strUserNum & "' and CP143 is null and cp57 is null And (CP05>=" & m_dblWDBegin & " And CP05<=" & m_dblWDEnd & ") "
   
   CheckOC
   strSql = "SELECT " & IIf(m_bolIsDrawer, "CP29", "CP14") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),CP26,DECODE(PA09,'000',PTM03,PTM04),nvl(DECODE(PA09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,PA09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112" & _
                  " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION" & _
                  " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 1) & ")" & StrSQL6 & " AND CP09=EP02(+) AND CP01=pa01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND PA09=NA01(+)" & _
                  " AND PA57 IS NULL AND CP10 in('1001','1002','1006')" & _
                  " AND exists(select * from caseprogress cp2 where cp2.cp09=cp43 and cp2.cp10 in('101','102','103','104','105','107','113','114','501','424','126','805','301','302','303','304','305','306','307','122','438'))"
'   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "CP14") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112" & _
'                  " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION" & _
'                  " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ")" & StrSQL6 & " AND CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+)" & _
'                  " AND TM29 IS NULL"
'   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "CP14") & ",ep01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112" & _
'                  " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION" & _
'                  " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ")" & StrSQL6 & " AND CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+)" & _
'                  " AND LC08 IS NULL"
'   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "CP14") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112" & _
'                  " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION" & _
'                  " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ")" & StrSQL6 & " AND CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+)" & _
'                  " AND HC09 IS NULL"
'   strSql = strSql + " UNION all  SELECT " & IIf(m_bolIsDrawer, "CP29", "CP14") & ",EP01,SUBSTR(CP09,1,1),substrb(' '||sqldatet(cp05),-9),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10)," & SQLDate("CP48") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP06") & "," & SQLDate("EP09") & "," & SQLDate("EP07") & ",nvl(S3.ST02,ep04)," & SQLDate("EP08") & "," & SQLDate("CP27") & ",Nvl(EP35,0),EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04)," & SQLDate("CP57") & ", S1.ST02,cp97,cp98,cp111,ep34,cp112" & _
'                  " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION" & _
'                  " WHERE CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") " & StrSQL6 & " AND CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+)" & _
'                  " AND SP15 IS NULL"
   strSql = strSql + " ORDER BY 1,4 "
   CheckOC
   TextOk = True
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           .MoveFirst
           k = 0
           DoEvents
           '判斷等級是否屬於專利
           Do While .EOF = False
               For i = 0 To 21
                  If i = 3 Then
                     strTemp(3) = "" & .Fields(3)
                  Else
                     strTemp(i) = CheckStr(.Fields(i))
                  End If
               Next i
               strSql = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "',0,0,0,0,0) "
               adoEng.Execute strSql
               .MoveNext
               DoEvents
           Loop
           StrMenu
           SetGrd1
           Me.Show
       Else
           TextOk = False
           Unload Me 'Nextstep
       End If
       CheckOC
   End With
End Sub

Private Sub StrMenu()
   Me.lblName.Caption = strUserName
   Me.lblDate.Caption = ChangeTStringToTDateString(m_dblWDBegin - 19110000) & " >= 審定來函日 <= " & ChangeTStringToTDateString(m_dblWDEnd - 19110000)
   
   strSql = "SELECT ' ',R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110010,R110011,R110012,R110013,R110014,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024 FROM R090614 " & _
                 " WHERE ID='" & strUserNum & "' AND R110001='" & strUserNum & "'"
   '+繪圖
   If m_bolIsDrawer Then
      strSql = strSql & " ORDER BY R110012 ASC, R110005 Desc "
   Else
      strSql = strSql & " ORDER BY R110002 Desc, R110005 Desc "
   End If
   CheckOC
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, adoEng, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
           Set GRD1.Recordset = adoRecordset
           ChkNoData = False
       Else
           ChkNoData = True
           GRD1.Clear
           GRD1.Rows = 2
           Screen.MousePointer = vbDefault
           Exit Sub
       End If
   End With
   CheckOC
End Sub

Private Sub SetGrd1()
   With GRD1
      '日期欄寬改800
      .Visible = False
      .Cols = 24 '23
      .row = 0
      .col = 0:   .Text = "V"
      .ColWidth(0) = 200
      .CellAlignment = flexAlignCenterCenter
      
      .col = 1:   .Text = "目次"
      If m_bolIsDrawer Then
        .ColWidth(1) = 0
      Else
        .ColWidth(1) = 300
      End If
      .CellAlignment = flexAlignCenterCenter
       
      .col = 2:   .Text = "收文類別"
      .ColWidth(2) = 200
      .CellAlignment = flexAlignCenterCenter
      .col = 3:   .Text = "收文日"
      .ColWidth(3) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 4:   .Text = "本所案號"
      .ColWidth(4) = 1000
      .CellAlignment = flexAlignCenterCenter
      .col = 5:   .Text = "案件名稱"
      .ColWidth(5) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 6:   .Text = "國家"
      .ColWidth(6) = 450
      .CellAlignment = flexAlignCenterCenter
      .col = 7:   .Text = "種類"
      .ColWidth(7) = 400
      .CellAlignment = flexAlignCenterCenter
      .col = 8:   .Text = "案件性質"
      .ColWidth(8) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 9:   .Text = "Y/N"
      .ColWidth(9) = 300
      .CellAlignment = flexAlignCenterCenter
      
      .col = 10:   .Text = "承辦期限"
      .ColWidth(10) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 11:   .Text = "本所期限"
      .ColWidth(11) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 12:  .Text = "法定期限"
      .ColWidth(12) = 0
      .CellAlignment = flexAlignCenterCenter
      .col = 13:  .Text = "齊備日"
      .ColWidth(13) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 14:  .Text = "完稿日"
      .ColWidth(14) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 15:  .Text = "會稿日"
      .ColWidth(15) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 16:  .Text = "核稿人"
      .ColWidth(16) = 700
      .CellAlignment = flexAlignCenterCenter
      .col = 17:  .Text = "會稿完成日"
      .ColWidth(17) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 18:  .Text = "發文日"
      .ColWidth(18) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 19:  .Text = "承辦天數"
      .ColWidth(19) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 20:  .Text = "備註"
      .ColWidth(20) = 2000
      .CellAlignment = flexAlignCenterCenter
      .col = 21:  .Text = "智權人員"
      .ColWidth(21) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 22:  .Text = ""
      .ColWidth(22) = 0
      .CellAlignment = flexAlignCenterCenter
      .col = 23:  .Text = ""
      .ColWidth(23) = 0
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub grd1_SelChange()
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   If GRD1.MouseRow <> 0 Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
   GRD1.Visible = True
End Sub

Private Sub txtWorkDays_GotFocus()
   TextInverse txtWorkDays
   CloseIme
End Sub

Private Sub txtWorkDays_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(KeyAscii) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtWorkDays_Validate(Cancel As Boolean)
   If Val(txtWorkDays) > 0 Then
      'm_dblWDEnd = CompWorkDay(Val(txtWorkDays) + 1, strSrvDate(1))
      m_dblWDBegin = CompWorkDay(Val(txtWorkDays), strSrvDate(1), 1)
      Me.lblDate.Caption = ChangeTStringToTDateString(m_dblWDBegin - 19110000) & " >= 審定來函日 <= " & ChangeTStringToTDateString(m_dblWDEnd - 19110000)
   End If
End Sub

' 回到該畫面再重新查詢一次
Public Sub RefreshData()
   TextOk = False
   StrMenu1
End Sub
