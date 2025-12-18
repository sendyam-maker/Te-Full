VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090901 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Begin VB.CheckBox Check1 
      Caption         =   "發文後補分割建議"
      Height          =   225
      Left            =   6690
      TabIndex        =   14
      Top             =   1470
      Width           =   2085
   End
   Begin VB.CommandButton cmdAMD 
      Caption         =   "補輸中說請款函修正內容(&U)"
      Height          =   400
      Left            =   120
      TabIndex        =   13
      Top             =   90
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7185
      TabIndex        =   6
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8025
      TabIndex        =   7
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3780
      TabIndex        =   4
      Top             =   570
      Width           =   800
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   1
      Left            =   1770
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCP"
      Top             =   630
      Width           =   495
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   2
      Left            =   2250
      MaxLength       =   6
      TabIndex        =   1
      Top             =   630
      Width           =   855
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   3
      Left            =   3090
      MaxLength       =   1
      TabIndex        =   2
      Top             =   630
      Width           =   255
   End
   Begin VB.TextBox txtCaseNo 
      Height          =   270
      Index           =   4
      Left            =   3330
      MaxLength       =   2
      TabIndex        =   3
      Top             =   630
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3660
      Left            =   90
      TabIndex        =   5
      Top             =   1950
      Width           =   8715
      _ExtentX        =   15363
      _ExtentY        =   6456
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label lblCaseName 
      Height          =   255
      Index           =   3
      Left            =   1770
      TabIndex        =   17
      Top             =   1620
      Width           =   4500
      VariousPropertyBits=   27
      Size            =   "7937;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   255
      Index           =   2
      Left            =   1770
      TabIndex        =   16
      Top             =   1290
      Width           =   4500
      VariousPropertyBits=   27
      Size            =   "7937;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseName 
      Height          =   255
      Index           =   1
      Left            =   1770
      TabIndex        =   15
      Top             =   960
      Width           =   4500
      VariousPropertyBits=   27
      Size            =   "7937;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   450
      TabIndex        =   12
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   1245
      TabIndex        =   11
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   1245
      TabIndex        =   10
      Top             =   1290
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(外):"
      Height          =   180
      Index           =   0
      Left            =   1245
      TabIndex        =   9
      Top             =   1620
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   825
      TabIndex        =   8
      Top             =   630
      Width           =   765
   End
End
Attribute VB_Name = "frm090901"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/23 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、lblCaseName(Index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
'Create by Morgan 2008/9/4
Option Explicit

Public CallFormName As String
Dim intLastRow As Integer
'Added by Lydia 2015/06/04 補輸中說-查詢
Public m_AMD As Boolean
Dim bolAMD As Boolean


Private Sub SetGridHead()
    Dim i As Integer
    FixGrid MSHFlexGrid1
    With MSHFlexGrid1
        .Visible = False
        .row = 0
        .col = 0: .ColWidth(.col) = 200: .Text = "v"
        .CellAlignment = flexAlignCenterCenter
        .col = 1: .ColWidth(.col) = 900: .Text = "收文日"
        .CellAlignment = flexAlignCenterCenter
        .col = 2: .ColWidth(.col) = 1300: .Text = "收文號"
        .CellAlignment = flexAlignCenterCenter
        .col = 3: .ColWidth(.col) = 1400: .Text = "案件性質"
        .CellAlignment = flexAlignCenterCenter
        .col = 4: .ColWidth(.col) = 1200: .Text = "承辦人"
        .CellAlignment = flexAlignCenterCenter
        .col = 5: .ColWidth(.col) = 1200: .Text = "承辦期限"
        .CellAlignment = flexAlignCenterCenter
        .col = 6: .ColWidth(.col) = 1200: .Text = "核稿人"
        .CellAlignment = flexAlignCenterCenter
        .col = 7: .ColWidth(.col) = 1200: .Text = "核稿期限"
        For i = 8 To .Cols - 1
         .col = i: .ColWidth(i) = 0
      Next
      .Visible = True
    End With
End Sub

Private Sub ClearGrid()
    Dim rstGrid As New ADODB.Recordset, stSQL As String
    
    stSQL = "SELECT 0,1,2,3,4,5,6,7,8,9,10,11 FROM DUAL WHERE ROWNUM<1"
    rstGrid.CursorLocation = adUseClient
    rstGrid.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
    Set MSHFlexGrid1.Recordset = rstGrid
    SetGridHead
    Set rstGrid = Nothing
End Sub

'Added by Morgan 2020/2/27
Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      cmdAMD.Enabled = False
   Else
      cmdAMD.Enabled = True
   End If
End Sub

'Added by Lydia 2015/06/04 補輸中說請款函修正內容
Private Sub cmdAMD_Click()
   m_AMD = True
   Call SetGrid
End Sub

Private Sub cmdExit_Click()
    blnIsFormBack = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim ii As Integer
    
    If MSHFlexGrid1.Rows < 2 Then Exit Sub
    
    With MSHFlexGrid1
        .Visible = False
        For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 0) = "v" Then Exit For
        Next ii
        .Visible = True
        If ii = .Rows Then
            MsgBox "請點選欲輸入資料！"
        Else
            frm090901_1.bolAMD = bolAMD 'Added by Lydia 2015/06/04
            frm090901_1.Show
            Call frm090901_1.SetData(MSHFlexGrid1.Recordset, ii)
            Me.Hide
        End If
    End With
End Sub

Public Function SetGrid(Optional ByVal bolMsg As Boolean = True) As Boolean

   Dim stSQL As String
   Dim stCon926 As String, stConCP As String, stCon201 As String, stCon1001 As String
   Dim arrCaseNo(1 To 4) As String
   Dim strAMD As String 'Added by Lydia 2015/06/04
On Error GoTo flgErr
   
   
   If Pub_StrUserSt03 <> "M51" Then
      'Modify by Morgan 2008/10/21 本人只能輸，管制人才能改
      'Modified by Morgan 2012/5/24 +38,42 權限第五級主管(副總)可清除926承辦期限
      'modify by sonia 2013/11/20 因Tammy留職停薪,化學組暫由79034王俊傑及94006宗家澔代理,何副總指示二人核對已准專利案件由對方消承辦期限
      'stCon926 = stCon926 & " AND CP01 IN ('FCP','FG') AND (INSTR(S1.ST52||','||S1.ST53||','||S1.ST54,'" & strUserNum & "')>0 or ((S1.st05='38' or S1.st05='42') and S1.ST55='" & strUserNum & "'))"
      'modify by sonia 2016/5/5 94006宗家澔離職,改A0022任政宏
      stCon926 = stCon926 & " AND CP01 IN ('FCP','FG') AND (INSTR(S1.ST52||','||S1.ST53||','||S1.ST54,'" & strUserNum & "')>0 or ((S1.st05='38' or S1.st05='42') and S1.ST55='" & strUserNum & "') or (s1.st16='2' and cp14='79034' and 'A0022'='" & strUserNum & "')or (s1.st16='2' and cp14='A0022' and '79034'='" & strUserNum & "'))"
      '2013/11/20
      stConCP = stConCP & " AND CP01 IN ('FCP','FG') AND ( CP14='" & strUserNum & "' or INSTR(S1.ST52||','||S1.ST53||','||S1.ST54,'" & strUserNum & "')>0 )"
      stCon201 = stCon201 & " AND CP01 IN ('FCP','FG') AND ( EP04 in ('" & strUserNum & "','" & PUB_GetMapID(strUserNum, 0) & "') or INSTR(S2.ST52||','||S2.ST53||','||S2.ST54,'" & strUserNum & "')>0 )"
      'Added by Morgan 2012/12/4
      '核准函可由所屬組別主管(承辦人為主管則由職代)或各級管制人(考慮主管請假)上完稿日,自己則只能輸入分割建議
      stCon1001 = stCon1001 & " AND CP01 IN ('FCP','FG') AND ( CP14='" & strUserNum & "' or INSTR(S1.ST52||','||S1.ST53||','||S1.ST54,'" & strUserNum & "')>0  or decode(oMan,CP14,B0102,oMan)='" & strUserNum & "')"
   End If
   
   'Add by Morgan 2008/9/22
   If Me.Tag <> "" Then
      stConCP = stConCP & " and cp09='" & Me.Tag & "'"
      stCon926 = stCon926 & " and cp09='" & Me.Tag & "'"
   End If
   
   'Added by Lydia 2015/06/04
   If m_AMD = True Then
      strAMD = " AND CP10 IN ('201','209','210','235') "
   Else
      strAMD = " AND CP27 IS NULL "
   End If
   
   arrCaseNo(1) = txtCaseNo(1)
   arrCaseNo(2) = Right("000000" & txtCaseNo(2), 6)
   arrCaseNo(3) = Right("0" & txtCaseNo(3), 1)
   arrCaseNo(4) = Right("00" & txtCaseNo(4), 2)
   
   txtCaseNo(1) = arrCaseNo(1)
   txtCaseNo(2) = arrCaseNo(2)
   txtCaseNo(3) = arrCaseNo(3)
   txtCaseNo(4) = arrCaseNo(4)
   
   'Added by Morgan 2020/2/27
   'Modify By Sindy 2023/10/30 EP33要回歸用在英文核完日,改抓EP39.核稿完成日
   If Me.Check1.Value = vbChecked Then
      stSQL = stSQL & " SELECT '' V" & _
        ", sqldatet(cp05) CP05T" & _
        ", CP09,NVL(CPM03,CP10) CP10C" & _
        ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
        ", '' EP04C, '' EP08T" & _
        ", PTM03 PA08T, PA08,PA05,PA06,PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08" & IIf(strSrvDate(1) >= FCP核完日改用EP39, ",ep39", ",ep33") & _
        " FROM CASEPROGRESS C, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1,PatentTrademarkMap" & _
        " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
        " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
        " AND CP10 in ('107','203','204','205')" & stConCP & _
        " AND CP57 IS NULL and cp27>19221111" & _
        " AND not exists(select * from caseprogress X where X.cp01=C.cp01" & _
        " AND X.cp02=C.cp02 AND X.cp03=C.cp03 AND X.cp04=C.cp04 and X.cp10 in ('107','203','204','205') and X.cp27>C.cp27)" & _
        " AND EP02(+)=CP09" & _
        " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
        " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
        " AND S1.ST01(+)=CP14" & _
        " AND PTM02=PA08 AND PTM01='1' "
   'end 2020/2/27
   
   ElseIf arrCaseNo(1) = "FG" Or arrCaseNo(1) = "PS" Or arrCaseNo(1) = "CPS" Then
      'Modified by Lydia 2015/06/04 未發文(CP27 IS NULL)改傳入條件
      'stSQL = " SELECT '' V" & _
           ", sqldatet(cp05) CP05T" & _
           ", CP09,NVL(CPM03,CP10) CP10C" & _
           ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
           ", '' EP04C, '' EP08T" & _
           ", '' PA08T,'' PA08,SP05 PA05,SP06 PA06,SP07 PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08,ep33" & _
           " FROM CASEPROGRESS, ENGINEERPROGRESS, servicepractice, CASEPROPERTYMAP, STAFF S1" & _
           " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
           " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
           " AND CP10<>'201'" & stConCP & _
           " AND CP27 IS NULL AND CP57 IS NULL" & _
           " AND EP02=CP09" & _
           " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
           " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
           " AND S1.ST01(+)=CP14"
      If m_AMD = True Then
         strExc(5) = strAMD & stConCP & " AND CP57 IS NULL"
      Else
         strExc(5) = " AND CP10<>'201' " & stConCP & strAMD & " AND CP57 IS NULL"
      End If
      stSQL = " SELECT '' V" & _
           ", sqldatet(cp05) CP05T" & _
           ", CP09,NVL(CPM03,CP10) CP10C" & _
           ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
           ", '' EP04C, '' EP08T" & _
           ", '' PA08T,'' PA08,SP05 PA05,SP06 PA06,SP07 PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08" & IIf(strSrvDate(1) >= FCP核完日改用EP39, ",ep39", ",ep33") & _
           " FROM CASEPROGRESS, ENGINEERPROGRESS, servicepractice, CASEPROPERTYMAP, STAFF S1" & _
           " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
           " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
           strExc(5) & " AND EP02=CP09" & _
           " AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04" & _
           " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
           " AND S1.ST01(+)=CP14"
   Else
      'Modified by Lydia 2015/06/04 只抓中說請款'201','209','210','235'
      If m_AMD = True Then
        'Added by Lydia 2015/06/25 再開放主動修正的承辦人可選擇主動修正程序去輸入修正定稿
        strExc(1) = "SELECT '' V" & _
             ", sqldatet(cp05) CP05T" & _
             ", CP09,NVL(CPM03,CP10) CP10C" & _
             ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
             ", '' EP04C, '' EP08T" & _
             ", PTM03 PA08T, PA08,PA05,PA06,PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08" & IIf(strSrvDate(1) >= FCP核完日改用EP39, ",ep39", ",ep33") & _
             " FROM CASEPROGRESS c1, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1,PatentTrademarkMap" & _
             " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
             " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "' AND cp10='203' " & _
             " and exists (select * from caseprogress c2 where c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) " & Replace(strAMD, "CP10", "c2.CP10") & " and cp57 is null) " & _
             " AND CP57 IS NULL" & _
             " AND EP02(+)=CP09" & _
             " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
             " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
             " AND S1.ST01(+)=CP14" & _
             " AND PTM02=PA08 AND PTM01='1' "
       '------------------
        stSQL = strExc(1) & " UNION ALL SELECT '' V" & _
             ", sqldatet(cp05) CP05T" & _
             ", CP09,NVL(CPM03,CP10) CP10C" & _
             ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
             ", '' EP04C, '' EP08T" & _
             ", PTM03 PA08T, PA08,PA05,PA06,PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08" & IIf(strSrvDate(1) >= FCP核完日改用EP39, ",ep39", ",ep33") & _
             " FROM CASEPROGRESS, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1,PatentTrademarkMap" & _
             " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
             " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
             strAMD & " AND CP57 IS NULL" & _
             " AND EP02(+)=CP09" & _
             " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
             " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
             " AND S1.ST01(+)=CP14" & _
             " AND PTM02=PA08 AND PTM01='1' "
         GoTo JUMPSQL
      End If
      '核對已准專利銷承辦期限
      stSQL = "SELECT '' V" & _
           ", sqldatet(cp05) CP05T" & _
           ", CP09,NVL(CPM03,CP10) CP10C" & _
           ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
           ", '' EP04C, '' EP08T" & _
           ", PTM03 PA08T, PA08,PA05,PA06,PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08" & IIf(strSrvDate(1) >= FCP核完日改用EP39, ",ep39", ",ep33") & _
           " FROM CASEPROGRESS, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1,PatentTrademarkMap" & _
           " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
           " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
           " AND CP10='926' AND CP27 IS NULL AND CP57 IS NULL and cp48>0" & _
           " AND EP02(+)=CP09" & _
           " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
           " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
           " AND S1.ST01(+)=CP14" & stCon926 & _
           " AND PTM02=PA08 AND PTM01='1' "
           
      '上完稿日(未發文)
      'Modify by Morgan 2011/8/31 +控制926除外,否則上面管制語法無效
      stSQL = stSQL & " UNION SELECT '' V" & _
           ", sqldatet(cp05) CP05T" & _
           ", CP09,NVL(CPM03,CP10) CP10C" & _
           ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
           ", '' EP04C, '' EP08T" & _
           ", PTM03 PA08T, PA08,PA05,PA06,PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08" & IIf(strSrvDate(1) >= FCP核完日改用EP39, ",ep39", ",ep33") & _
           " FROM CASEPROGRESS, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1,PatentTrademarkMap" & _
           " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
           " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
           " AND CP10<>'201' AND CP10<>'926' AND CP10<>'1001'" & stConCP & _
           " AND CP27 IS NULL AND CP57 IS NULL" & _
           " AND EP02(+)=CP09" & _
           " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
           " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
           " AND S1.ST01(+)=CP14" & _
           " AND PTM02=PA08 AND PTM01='1' "
      
      'Added by Morgan 2012/12/4 +核准
      'Modified by Lydia 2015/10/05 + 1008
      stSQL = stSQL & " UNION SELECT '' V" & _
           ", sqldatet(cp05) CP05T" & _
           ", CP09,NVL(CPM03,CP10) CP10C" & _
           ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
           ", '' EP04C, '' EP08T" & _
           ", PTM03 PA08T, PA08,PA05,PA06,PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08" & IIf(strSrvDate(1) >= FCP核完日改用EP39, ",ep39", ",ep33") & _
           " FROM CASEPROGRESS, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1,PatentTrademarkMap,SetSpecMan,ABS001" & _
           " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
           " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & _
           " AND CP10 in ('1001','1008') AND CP01='FCP'" & _
           " AND CP27 IS NULL AND CP57 IS NULL" & _
           " AND EP02(+)=CP09" & _
           " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
           " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
           " AND S1.ST01(+)=CP14 and OCODE(+)=decode(s1.st16,'1','T','2','R','3','S','4','T1') and B0101(+)=st01" & stCon1001 & _
           " AND PTM02=PA08 AND PTM01='1' "

      '上核稿完成日(已完稿,未發文)
      stSQL = stSQL & " UNION SELECT '' V" & _
           ", sqldatet(cp05) CP05T" & _
           ", CP09,NVL(CPM03,CP10) CP10C" & _
           ", S1.ST02 CP14C, sqldatet(cp48) CP48T" & _
           ", S2.ST02 EP04C, sqldatet(EP08) EP08T" & _
           ", PTM03 PA08T, PA08,PA05,PA06,PA07,CP05,CP14,CP48,cp64,ep09,cp10,ep04,ep08" & IIf(strSrvDate(1) >= FCP核完日改用EP39, ",ep39", ",ep33") & _
           " FROM CASEPROGRESS, ENGINEERPROGRESS, PATENT, CASEPROPERTYMAP, STAFF S1, STAFF S2,PatentTrademarkMap" & _
           " WHERE CP01='" & arrCaseNo(1) & "' AND CP02='" & arrCaseNo(2) & "'" & _
           " AND CP03='" & arrCaseNo(3) & "' AND CP04='" & arrCaseNo(4) & "'" & stCon201 & _
           " AND CP10='201'" & _
           " AND CP27 IS NULL AND CP57 IS NULL" & _
           " AND EP02(+)=CP09 and EP09>0" & _
           " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
           " AND CPM01(+)=CP01 AND CPM02(+)=CP10" & _
           " AND S1.ST01(+)=CP14 AND S2.ST01(+)=EP04" & stCon201 & _
           " AND PTM02=PA08 AND PTM01='1' "
   End If
   
'Added by Lydia 2015/06/04
JUMPSQL:
   '保留到下一畫面的判斷
   If m_AMD = True Then
      bolAMD = True
   Else
      bolAMD = False
   End If
   m_AMD = False
'end 2015/06/04

   stSQL = stSQL & " ORDER BY CP05, CP09"
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   Set MSHFlexGrid1.Recordset = RsTemp.Clone
   SetGridHead
   If intI = 1 Then
      With RsTemp
         lblCaseName(1) = "" & .Fields("PA05")
         lblCaseName(2) = "" & .Fields("PA06")
         lblCaseName(3) = "" & .Fields("PA07")
      End With
      'Add by Morgan 2008/9/22
      If Me.Tag <> "" Then
         Me.Tag = ""
         With frm090901_1
            .Show
            .ZOrder
            .SetData RsTemp, 1
            .NextFormName = CallFormName
         End With
         Me.Hide
      End If
      'end 2008/9/22
      SetGrid = True
   ElseIf bolMsg Then
      ShowNoData
      txtCaseNo(2).SetFocus
   End If
   Exit Function
   
flgErr:
   MsgBox Err.Description, vbCritical

End Function

'Modify By Sindy 2023/10/17
'Private Sub cmdSearch_Click()
Public Sub cmdSearch_Click()
    Call SetGrid
'    If Me.MSHFlexGrid1.Rows = 2 And Me.Visible = True Then
'        MSHFlexGrid1.row = 1
'        GridClick MSHFlexGrid1, intLastRow, 0
'        cmdok_Click
'   End If
End Sub

Private Sub Form_Activate()
   If txtCaseNo(2).Visible = True Then txtCaseNo(2).SetFocus
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    ClearGrid
   'Added by Morgan 2012/5/21
   If Pub_StrUserSt03 = "F22" Then
      txtCaseNo(1) = "P"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Set frm090901 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK.SetFocus
End Sub

Private Sub MSHFlexGrid1_KeyPress(KeyAscii As Integer)
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK.SetFocus
End Sub

Private Sub txtCaseNo_Change(Index As Integer)
    lblCaseName(1) = ""
    lblCaseName(2) = ""
    lblCaseName(3) = ""
    ClearGrid
End Sub

Private Sub txtCaseNo_GotFocus(Index As Integer)
    TextInverse txtCaseNo(Index)
    Select Case Index
        Case 2, 3, 4
            CloseIme
    End Select
End Sub

Private Sub txtCaseNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
