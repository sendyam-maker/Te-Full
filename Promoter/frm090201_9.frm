VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_9 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護－已完稿超過2天或已屆預定會稿日且尚未會稿案件資料"
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
      Caption         =   "全部選取"
      Height          =   400
      Index           =   3
      Left            =   5070
      TabIndex        =   6
      Top             =   135
      Width           =   900
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   2
      Left            =   8225
      TabIndex        =   5
      Top             =   135
      Width           =   850
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "E-mail核稿人(&S)"
      Height          =   400
      Index           =   1
      Left            =   6825
      TabIndex        =   4
      Top             =   135
      Width           =   1400
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   4668
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5784
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "繼續(&C)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5970
      TabIndex        =   0
      Top             =   135
      Width           =   850
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4900
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8652
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
      Left            =   930
      TabIndex        =   7
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
      Caption         =   "承辦人： "
      Height          =   180
      Index           =   35
      Left            =   150
      TabIndex        =   3
      Top             =   360
      Width           =   795
   End
End
Attribute VB_Name = "frm090201_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/21 改成Form2.0 ; grd1改字型=新細明體-ExtB、lblName
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'2009/11/10 CREATE BY SONIA COPY FROM frm090201_5
Option Explicit
Public TextOk As Boolean
Public strContinue As Boolean    '2009/11/13 add by sonia 判斷為繼續或結束
Dim strSql As String, i As Integer, j As Integer
Dim dblDeadLine As Double

Private Sub cmdOK_Click(Index As Integer)
Dim KeyWord As String
Dim m_mailCNT As Integer

   Select Case Index
      Case 0 '繼續
         strContinue = True   '2009/11/13 add by sonia
         Unload Me
      Case 1  '寄mail
         With grd1
            For i = 1 To .Rows - 1
               .Visible = False
               If .TextMatrix(i, 0) = "V" Then
                  Exit For
               Else
                  If i = .Rows - 1 Then
                     MsgBox "請點選欲處理的資料"
                     .Visible = True
                     Exit Sub
                  End If
               End If
            Next
            .Visible = True
         End With
   
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         m_mailCNT = 0
         With grd1
            For i = 1 To .Rows - 1
               .Visible = False
               If .TextMatrix(i, 0) = "V" Then
                  If CheckStr(.TextMatrix(i, 7)) <> "" Then
                     PUB_SendMail strUserNum, CheckStr(.TextMatrix(i, 8)), "", "催函", vbCrLf + "本所案號： " + .TextMatrix(i, 1) + vbCrLf + "完稿日　： " + .TextMatrix(i, 10) + vbCrLf + "案件名稱： " + .TextMatrix(i, 4) + " " + .TextMatrix(i, 5) + " " + .TextMatrix(i, 6) + vbCrLf + "案件性質： " + .TextMatrix(i, 7) + vbCrLf + "本所期限： " + .TextMatrix(i, 11) + vbCrLf + "法定期限： " + .TextMatrix(i, 12) + vbCrLf + "預定會稿日： " + .TextMatrix(i, 13) + vbCrLf + "智權人員　： " + .TextMatrix(i, 14) + vbCrLf + vbCrLf + "核稿天數超過 或 已屆預定會稿日且尚未會稿案件 請儘速核稿!", ""
                     m_mailCNT = m_mailCNT + 1
                  End If
               End If
               .TextMatrix(i, 0) = ""
               For j = 0 To .Cols - 1
                  .row = i
                  .col = j
                  .CellBackColor = QBColor(15)
               Next j
            Next
            .Visible = True
         End With
         Me.Enabled = True
         Screen.MousePointer = vbDefault
         If m_mailCNT > 0 Then MsgBox "寄件成功!!", vbInformation
      Case 2 '結束
         strContinue = False   '2009/11/13 add by sonia
         Unload Me
      Case 3
         grd1.col = 0
         Screen.MousePointer = vbHourglass
         If Trim(cmdok(3).Caption) = "全部選取" Then
            KeyWord = "V"
            cmdok(3).Caption = "全部取消"
         Else
            KeyWord = ""
            cmdok(3).Caption = "全部選取"
         End If
         For i = 1 To grd1.Rows - 1
            grd1.row = i
            grd1.Text = KeyWord
         Next i
         Screen.MousePointer = vbDefault
   End Select
End Sub

Private Sub Form_Load()
   Me.Hide
   Screen.MousePointer = vbHourglass
   MoveFormToCenter Me
'   Me.Enabled = False
'   StrMenu1
'   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If strContinue = True Then Nextstep
   Set frm090201_9 = Nothing
End Sub
'2009/11/13 add by sonia
Private Sub Nextstep()
   'Modify by Morgan 2010/11/16
   'frm090201_2.Show   '工作進度維護
   frm090201_a.TextOk = False
   frm090201_a.StrMenu1
   If frm090201_a.TextOk = True Then
      frm090201_a.Show
   End If
   'end 2010/11/16
End Sub
'2009/11/13 end

Public Sub StrMenu1()
Dim rsA As New ADODB.Recordset

   DoEvents
   
   strContinue = True   '2009/11/13 add by sonia
   rsA.MaxRecords = 0
   Set rsA = Nothing
   dblDeadLine = PUB_GetWorkDay(-2)
                       strSql = "SELECT '',CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(PA57,'Y','＊',''),DECODE(EP28,NULL,'已完稿超過2天',DECODE(SIGN(EP28-TO_CHAR(SYSDATE,'YYYYMMDD')),'1','已完稿超過2天','已屆預定會稿日')),PA05,PA06,PA07,NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),ep04,nvl(A.ST02,EP04)," & SQLDate("EP09") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP28") & ",NVL(B.ST02,CP13)," & SQLDate("CP48") & ",CP64,cp09 FROM CASEPROGRESS,PATENT,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and ep02=CP09(+) AND (EP07 IS NULL OR EP07=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' and cp57 is null and cp27 is null AND CP01 IN (" & SQLGrpStr("", 1) & ") And ( EP09 <" & dblDeadLine & " or (EP28<=" & strSrvDate(1) & " and NVL(EP34,'Y')='Y'))"
   strSql = strSql & " UNION all SELECT '',CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(TM29,'Y','＊',''),DECODE(EP28,NULL,'已完稿超過2天',DECODE(SIGN(EP28-TO_CHAR(SYSDATE,'YYYYMMDD')),'1','已完稿超過2天','已屆預定會稿日')),TM05,TM06,TM07,NVL(DECODE(TM10,'000',CPM03,CPM04),CP10),ep04,nvl(A.ST02,EP04)," & SQLDate("EP09") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP28") & ",NVL(B.ST02,CP13)," & SQLDate("CP48") & ",CP64,cp09 FROM CASEPROGRESS,TRADEMARK,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) and ep02=CP09(+) AND (EP07 IS NULL OR EP07=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' and cp57 is null and cp27 is null AND CP01 IN (" & SQLGrpStr("", 2) & ") And ( EP09 <" & dblDeadLine & " or (EP28<=" & strSrvDate(1) & " and NVL(EP34,'Y')='Y'))"
   strSql = strSql & " UNION all SELECT '',CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(LC08,'Y','＊',''),DECODE(EP28,NULL,'已完稿超過2天',DECODE(SIGN(EP28-TO_CHAR(SYSDATE,'YYYYMMDD')),'1','已完稿超過2天','已屆預定會稿日')),LC05,LC06,LC07,NVL(DECODE(LC15,'000',CPM03,CPM04),CP10),ep04,nvl(A.ST02,EP04)," & SQLDate("EP09") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP28") & ",NVL(B.ST02,CP13)," & SQLDate("CP48") & ",CP64,cp09 FROM CASEPROGRESS,LAWCASE,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) and ep02=CP09(+) AND (EP07 IS NULL OR EP07=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' and cp57 is null and cp27 is null AND CP01 IN (" & SQLGrpStr("", 3) & ") And ( EP09 <" & dblDeadLine & " or (EP28<=" & strSrvDate(1) & " and NVL(EP34,'Y')='Y'))"
   strSql = strSql & " UNION all SELECT '',CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(HC09,'Y','＊',''),DECODE(EP28,NULL,'已完稿超過2天',DECODE(SIGN(EP28-TO_CHAR(SYSDATE,'YYYYMMDD')),'1','已完稿超過2天','已屆預定會稿日')),HC06,' ',' ',nvl(CPM03,cp10),ep04,nvl(A.ST02,EP04)," & SQLDate("EP09") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP28") & ",NVL(B.ST02,CP13)," & SQLDate("CP48") & ",CP64,cp09 FROM CASEPROGRESS,HIRECASE,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A, STAFF B WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) and ep02=CP09(+) AND (EP07 IS NULL OR EP07=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' and cp57 is null and cp27 is null AND CP01 IN (" & SQLGrpStr("", 4) & ") And ( EP09 <" & dblDeadLine & " or (EP28<=" & strSrvDate(1) & " and NVL(EP34,'Y')='Y'))"
   strSql = strSql & " UNION all SELECT '',CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(SP15,'Y','＊',''),DECODE(EP28,NULL,'已完稿超過2天',DECODE(SIGN(EP28-TO_CHAR(SYSDATE,'YYYYMMDD')),'1','已完稿超過2天','已屆預定會稿日')),SP05,SP06,SP07,NVL(DECODE(SP09,'000',CPM03,CPM04),CP10),ep04,nvl(A.ST02,EP04)," & SQLDate("EP09") & "," & SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("EP28") & ",NVL(B.ST02,CP13)," & SQLDate("CP48") & ",CP64,cp09 FROM CASEPROGRESS,SERVICEPRACTICE,ENGINEERPROGRESS,CASEPROPERTYMAP,STAFF A,STAFF B WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) and ep02=CP09(+) AND (EP07 IS NULL OR EP07=0) and cp01=cpm01(+) and cp10=cpm02(+) AND EP04=A.ST01(+) AND CP13=B.ST01(+) and CP26 IS NULL  and ep05='" & strUserNum & "' and cp57 is null and cp27 is null AND CP01 IN (" & SQLGrpStr("", 5) & ") And ( EP09 <" & dblDeadLine & " or (EP28<=" & strSrvDate(1) & " and NVL(EP34,'Y')='Y'))"
   strSql = strSql + " ORDER BY 2 "
   CheckOC
   TextOk = True
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 And .RecordCount > 0 Then
         grd1.Clear
         Set grd1.Recordset = adoRecordset
         SetGrd
      Else
         TextOk = False
         Nextstep        '2009/11/13 end
      End If
      CheckOC
   End With
End Sub

Private Sub SetGrd()
   Me.lblName.Caption = strUserName
   
   With grd1
      .Visible = False
      .Cols = 18
      .row = 0
      .col = 0:   .Text = "V"
      .ColWidth(0) = 200
      .CellAlignment = flexAlignCenterCenter
      .col = 1:   .Text = "本所案號"
      .ColWidth(1) = 1000
      .CellAlignment = flexAlignCenterCenter
      .col = 2:  .Text = "是否閉卷"
      .ColWidth(2) = 0
      .CellAlignment = flexAlignCenterCenter
      .col = 3:   .Text = "提醒類別"
      .ColWidth(3) = 1350
      .CellAlignment = flexAlignCenterCenter
      .col = 4:   .Text = "案件名稱"
      .ColWidth(4) = 1500
      .CellAlignment = flexAlignCenterCenter
      .col = 5:   .Text = "案件名稱－英"
      .ColWidth(5) = 0
      .CellAlignment = flexAlignCenterCenter
      .col = 6:   .Text = "案件名稱－日"
      .ColWidth(6) = 0
      .CellAlignment = flexAlignCenterCenter
      .col = 7:   .Text = "案件性質"
      .ColWidth(7) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 8:  .Text = "核稿人編號"
      .ColWidth(8) = 0
      .CellAlignment = flexAlignCenterCenter
      .col = 9:  .Text = "核稿人"
      .ColWidth(9) = 700
      .CellAlignment = flexAlignCenterCenter
      .col = 10:  .Text = "完稿日"
      .ColWidth(10) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 11:  .Text = "本所期限"
      .ColWidth(11) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 12:  .Text = "法定期限"
      .ColWidth(12) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 13:  .Text = "預定會稿日"
      .ColWidth(13) = 1000
      .CellAlignment = flexAlignCenterCenter
      .col = 14:  .Text = "智權人員"
      .ColWidth(14) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 15:   .Text = "承辦期限"
      .ColWidth(15) = 800
      .CellAlignment = flexAlignCenterCenter
      .col = 16:  .Text = "備註"
      .ColWidth(16) = 3000
      .CellAlignment = flexAlignCenterCenter
      .col = 17:  .Text = "總收文號"
      .ColWidth(17) = 1000
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub grd1_SelChange()
   grd1.Visible = False
   grd1.col = 0
   grd1.row = grd1.MouseRow
   If grd1.MouseRow <> 0 Then
      If grd1.Text = "V" Then
         grd1.Text = ""
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = QBColor(15)
         Next i
      Else
         grd1.Text = "V"
         For i = 0 To grd1.Cols - 1
            grd1.col = i
            grd1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
   grd1.Visible = True
End Sub
