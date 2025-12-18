VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010602_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "核准函輸入"
   ClientHeight    =   5748
   ClientLeft      =   -1524
   ClientTop       =   1056
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9360
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   780
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "1"
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7152
      TabIndex        =   11
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   6324
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8376
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm06010602_2.frx":0000
      Left            =   1140
      List            =   "frm06010602_2.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4380
      TabIndex        =   5
      Top             =   690
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   3
      Top             =   690
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   2
      Top             =   690
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   1
      Top             =   690
      Width           =   255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   0
      Top             =   690
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3732
      Left            =   180
      TabIndex        =   15
      Top             =   1560
      Width           =   9012
      _ExtentX        =   15896
      _ExtentY        =   6583
      _Version        =   393216
      Cols            =   12
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label LblFM2 
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Top             =   1110
      Width           =   7455
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "13150;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "訴願，行政訴訟，上訴的核准請改至  一般來函輸 1502撤銷原處分！"
      ForeColor       =   &H000000FF&
      Height          =   360
      Index           =   2
      Left            =   360
      TabIndex        =   16
      Top             =   120
      Width           =   3000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   180
      X2              =   9180
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   180
      X2              =   9180
      Y1              =   1464
      Y2              =   1464
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1:核准, 2:改變原處分)"
      Height          =   180
      Left            =   1140
      TabIndex        =   14
      Top             =   5400
      Width           =   1740
   End
   Begin VB.Label Label3 
      Caption         =   "結果:"
      Height          =   252
      Left            =   180
      TabIndex        =   13
      Top             =   5400
      Width           =   492
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   1080
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Index           =   0
      Left            =   3420
      TabIndex        =   6
      Top             =   690
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   690
      Width           =   765
   End
End
Attribute VB_Name = "frm06010602_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/4/23 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strTemp As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim intWhere As Integer
Dim intLastRow As Integer, intCols As Integer

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         FormConfirm
      Case 1
         frm06010602_1.Show
         Unload Me
      Case 2
         Unload frm06010602_1
         Unload Me
   End Select
End Sub

' 確認鈕
Private Sub FormConfirm()
Dim bolChk As Boolean, i As Integer, j As Integer, strTmp(1 To 2) As String, strCP10 As String
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            bolChk = True
            Me.Tag = .TextMatrix(i, 1)
            strExc(5) = .TextMatrix(i, 5)
            strCP10 = .TextMatrix(i, 9) 'Add by Morgan 2008/3/19
            Exit For
         End If
      Next
   End With
   
   If bolChk = False Then
      MsgBox "請選擇資料 !", vbInformation
      Exit Sub
   End If
   
   'Added by Morgan 2013/1/14
   '檢查移轉或讓與的受讓人(5個)與基本檔是否相同
   If InStr("701,702,703,708", strCP10) > 0 Then
      If PUB_ChkAsignCaseCustNo(Me.Tag) = False Then
         Exit Sub
      End If
   End If
   'end 2013/1/14
   
   'Add Morgan 2008/3/19
   If strCP10 = "928" Then
      strExc(0) = "select 1 from nextprogress where np01='" & Me.Tag & "' and np07='202' and (np06 is null or np06='N')"
      intI = 1
      Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "本案下一程序仍有重新委任的補文件期限，故重新委任不可輸核准！"
         Exit Sub
      End If
   End If
   'end 2008/3/19
   
   '2011/7/12 ADD BY SONIA
   '非新申請案或改請案檢查來函記錄檔期限
   'Modified by Morgan 2015/10/20 +再審
   If InStr(CaseMapIn & ",107", strCP10) = 0 And (strCP10 < "3" Or strCP10 >= "4") Then
      If ClsLawChkMRec(TransDate(frm06010602_1.Text5, 2), Text2 & Text3 & Text4 & Text5, strTmp(1), strTmp(2)) Then
         If strTmp(1) <> "" Then
            If MsgBox("與櫃台之來函收文記錄期限 ( " & TransDate(strTmp(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
         End If
      'Modified by Morgan 2017/5/10 電子公文
      'Else
      ElseIf frm06010602_1.m_DocNo = "" Then
      'end 2017/5/10
         If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
      End If
   '新申請案或改請案只檢查是否有來函記錄
   Else
      If ClsLawChkMRec(TransDate(frm06010602_1.Text5, 2), Text2 & Text3 & Text4 & Text5, strTmp(1), strTmp(2)) Then
      'Modified by Morgan 2017/5/10 電子公文
      'Else
      ElseIf frm06010602_1.m_DocNo = "" Then
      'end 2017/5/10
         If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbYesNo) = vbNo Then Exit Sub
      End If
   End If
   '2011/7/12 END
   
   'Added by Morgan 2017/5/10 電子公文
   frm06010602_3.m_DocWord = frm06010602_1.m_DocWord
   frm06010602_3.m_DocNo = frm06010602_1.m_DocNo
   frm06010602_3.m_DocDate = frm06010602_1.m_DocDate
   frm06010602_3.m_AppNo = frm06010602_1.m_AppNo
   frm06010602_3.m_DeadLine = frm06010602_1.m_DeadLine
   'end 2017/5/10
   frm06010602_3.Show
   Me.Hide
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         LblFM2 = pa(5)
      Case "英"
         LblFM2 = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         LblFM2 = pa(7)
   End Select
End Sub

Private Sub Form_Activate()
   ReadPatent 1
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()


   MoveFormToCenter Me
   intWhere = 國外_FC
   pa(1) = strExc(1)
   pa(2) = strExc(2)
   pa(3) = strExc(3)
   pa(4) = strExc(4)
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
End Sub

Private Sub ReadPatent(ByVal iSitu As Integer)
 Dim Lbl As Label, txt As TextBox, i As Integer
 Dim strTmp(0 To 5) As String
   LblFM2 = ""
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      LblFM2 = pa(5)
      Text1 = pa(11)
   End If
   '核准
   If iSitu = 1 Then
      'Modify By Cheng 2002/01/25
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is null and " & _
'         "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B' or " & _
'         "(substr(cp09,1,1)='C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))"
      'Modify By Cheng 2002/04/12
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is null and " & _
'         "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B' or " & _
'         "(substr(cp09,1,1)='C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))" & _
'         " And ((cp10>='101' and cp10<='105') or cp10 ='107' or (cp10>='301' and cp10<='307') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804')) "
        'Modify By Cheng 2003/07/28
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is null and " & _
'         "( cp09<'C' or " & _
'         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))" & _
'         " And ((cp10>='101' and cp10<='105') or cp10 ='107' or (cp10>='301' and cp10<='307') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10>='602' and cp10<='604') or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804')) "

      'Modify by Morgan 2005/8/26 加922(檢還樣品證據)
      'Modify by Morgan 2006/1/12 加421(技術報告)
      '2007/5/3 MODIFY BY SONIA 加406(申請英文證明)
      '2007/6/27 MODIFY BY SONIA 加928(重新委任)
      'Modify by Morgan 2007/8/31 加807第三人申請技術報告
      '2007/11/5 modify by sonia 加417提早公開
      '2007/12/3 modify by sonia 加429放棄專利權
      'Modify by Morgan 2009/10/8 +908退費
      '2010/11/12 MODIFY BY SONIA 訴願或行政訴訟或上訴的核准請改輸  一般來函的撤銷原處分
      '2010/11/17 modify by sonia C類來函被異議理由1801,被舉發理由1802,通知參加訴願1504,通知參加訴訟1505只抓下一程序未續辦者,改在下面另外抓
      'strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is null and CP10<>'501' AND CP10<>'503' AND CP10<>'507' AND " & _
         "(( cp09<'C' And ((cp10>='101' and cp10<='105') or cp10 ='107' or (cp10>='301' and cp10<='307') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10='421') or (cp10='807') or (cp10='406') or (cp10='417') or (cp10='429') or (cp10>='602' and cp10<='604') or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804') or (cp10='922') or (cp10='928') or (cp10='908'))) Or " & _
         "( cp09>'C' And (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "' Or CP10='1504' Or CP10='1505' Or CP10='1211' Or CP10='1210' ))) "
      'Modified by Morgan 2012/12/24 +衍生設計125,改請衍生設計308
      'Modified by Morgan 2013/3/25 +608加註專利權延長
      'Modified by Morgan 2013/4/1 +433誤譯訂正--江如玉 Ex.FCP-43321
      'modify by sonia 2014/4/24 +405申請優先權證明書 FCP-049265
      'modify by sonia 2014/11/26 +432回復原狀 FCP-055075
      'modify by sonia 2015/3/6 +124回復優先權主張 FCP-051344
      'modify by sonia 2018/11/6 +439專利權部分拋棄
      'modify by sonia 2019/8/22 +440申請權部分拋棄
      'modify by sonia 2020/2/24 +412延緩公告
      'modify by sonia 2023/12/5 +935案件轉至本所
      'modify by sonia 2024/3/15 +508行政上訴答辯
      'Modified by Morgan 2024/4/18 +443申請證書副本--Winfrey
      'Modified by Lydia 2025/02/12 +245延緩審查
      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is null and CP10<>'501' AND CP10<>'503' AND CP10<>'507' AND " & _
         "(( cp09<'C' And ((cp10>='101' and cp10<='105') or cp10 ='107' or cp10 ='124' or cp10 ='125' or (cp10>='301' and cp10<='308') or (cp10>='401' and cp10<='403') or cp10='405' or cp10='432' or (cp10>='413' and cp10<='415') or cp10='421' or cp10='807' or cp10='406' or cp10='412' or cp10='417' or cp10='429' or cp10='439' or cp10='440' or cp10='433' or (cp10>='602' and cp10<='604') or cp10='608' or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804')" & _
         " or cp10='922' or cp10='928' or cp10='908' or cp10='935' or cp10='508' or cp10='443' or cp10='245')) Or " & _
         "( cp09>'C' And (CP10='1211' Or CP10='1210' ))) "
         
'   '改變原處份
   Else
      'Modify By Cheng 2002/01/25
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is not null and " & _
'         "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B' or " & _
'         "(substr(cp09,1,1)='C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))"
      'Modify By Cheng 2002/04/12
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is not null and " & _
'         "(substr(cp09,1,1)='A' or substr(cp09,1,1)='B' or " & _
'         "(substr(cp09,1,1)='C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))" & _
'         " And ((cp10>='101' and cp10<='105') or cp10 ='107' or (cp10>='301' and cp10<='307') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804')) "
        'Modify By Cheng 2003/07/28
'      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'         " and cp27 is not null and cp24 is not null and " & _
'         "( cp09<'C' or " & _
'         "( cp09>'C' and (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "')))" & _
'         " And ((cp10>='101' and cp10<='105') or cp10 ='107' or (cp10>='301' and cp10<='307') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10>='602' and cp10<='604') or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804')) "
      'Modify by Morgan 2006/1/12 加421(技術報告)
      '2007/5/3 MODIFY BY SONIA 加406(申請英文證明)
      '2010/11/12 MODIFY BY SONIA 訴願或行政訴訟或上訴的核准請改輸  一般來函的撤銷原處分
      '2010/11/17 modify by sonia C類來函被異議理由1801,被舉發理由1802,通知參加訴願1504,通知參加訴訟1505只抓下一程序未續辦者,改在下面另外抓
      'strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is not null and CP10<>'501' AND CP10<>'503' AND CP10<>'507' AND " & _
         "(( cp09<'C' And ((cp10>='101' and cp10<='105') or cp10 ='107' or (cp10>='301' and cp10<='307') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10='421') or (cp10='406') or (cp10>='602' and cp10<='604') or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804'))) Or " & _
         "( cp09>'C' And (cp10='" & 被異議理由 & "' or cp10='" & 被舉發理由 & "' Or CP10='1504' Or CP10='1505' Or CP10='1211' Or CP10='1210' ))) "
      'Modified by Morgan 2012/12/24 +衍生設計125,改請衍生設計308
      strExc(1) = ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " and cp27 is not null and cp24 is not null and CP10<>'501' AND CP10<>'503' AND CP10<>'507' AND " & _
         "(( cp09<'C' And ((cp10>='101' and cp10<='105') or cp10 ='107' or cp10 ='125' or (cp10>='301' and cp10<='308') or (cp10>='401' and cp10<='403') or (cp10>='413' and cp10<='415') or (cp10='421') or (cp10='406') or (cp10>='602' and cp10<='604') or (cp10>='701' and cp10<='707') or (cp10>='801' and cp10<='804'))) Or " & _
         "( cp09>'C' AND (CP10='1211' Or CP10='1210' ))) "
   End If
   ' 91.09.13 modify by louis (排序)
   'strExc(2) = "'',CP09,CPM03," & _
   '   "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
   '   SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64 " & _
   '   "from caseprogress,casepropertymap,CUSTOMER"
   strExc(2) = "'',CP09,CPM03," & _
      "DECODE(CP10,'704',CP50,'705',CP50,'706',CP50,'701',NVL(CU04,NVL(CU05,CU06)),CP40)," & _
      SQLDate("CP06") & "," & SQLDate("CP07") & "," & SQLDate("CP27") & ",decode(CP24,'1','准,勝','2','駁,敗',''),CP64,CP10 " & _
      ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " & _
      "from caseprogress,casepropertymap,CUSTOMER"
   '2010/11/15 add by sonia 已有專用期間者不帶出新申請案件性質
   If pa(25) <> "" Then
      strExc(3) = " and instr('" & NewCasePtyList & "',cp10)=0 "
   Else
      strExc(3) = ""
   End If
   '2010/11/15 end
   'strExc(0) = "select " & strExc(2) & " where substr(cp10,1,1)='1' and " & strExc(1) & " and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
   '   " and (cp01,cp02,cp03,cp04) not in " & _
   '   "(select cp01,cp02,cp03,cp04 from caseprogress where substr(cp10,1,1)='3' and " & _
   '   strExc(1) & ") union " & _
   '   "select " & strExc(2) & " where substr(cp10,1,1)<>'1' and " & strExc(1) & " and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)"
   '94.2.2 modify by sonia
   'strExc(0) = "select " & strExc(2) & " where substr(cp10,1,1)='1' and " & strExc(1) & " and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
   '   " and (cp01,cp02,cp03,cp04) not in " & _
   '   "(select cp01,cp02,cp03,cp04 from caseprogress where substr(cp10,1,1)='3' and " & _
   '   strExc(1) & ") union " & _
   '   "select " & strExc(2) & " where substr(cp10,1,1)<>'1' and " & strExc(1) & " and " & _
   '   "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) " & _
   '   "ORDER BY SORTFIELD DESC "
   strExc(0) = "select " & strExc(2) & " where substr(cp10,1,1)='1' and " & strExc(1) & strExc(3) & " and " & _
      "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+)" & _
      " and (cp01,cp02,cp03,cp04) not in " & _
      "(select cp01,cp02,cp03,cp04 from caseprogress where substr(cp10,1,1)='3' and " & _
      strExc(1) & ") union " & _
      "select " & strExc(2) & " where (substr(cp10,1,1)<>'1' or cp10='107') and " & strExc(1) & strExc(3) & " and " & _
      "cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) "
   '2010/11/17 add by sonia C類來函被異議理由1801,被舉發理由1802,通知參加訴願1504,通知參加訴訟1505只抓下一程序未續辦者
   strExc(0) = strExc(0) & " union " & _
      "select " & strExc(2) & ",nextprogress where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      "and cp09>'C' And (CP10='1801' Or CP10='1802' Or CP10='1504' Or CP10='1505') " & _
      "and cp09=np01(+) and (np06 is null or np06='N') " & _
      "and cp01=cpm01(+) and cp10=cpm02(+) AND SUBSTR(CP56,1,8)=CU01(+) AND SUBSTR(CP56,9,1)=CU02(+) "
   '核准
   If iSitu = 1 Then
      strExc(0) = strExc(0) & " and cp24 is null "
   Else
      strExc(0) = strExc(0) & " and cp24 is not null "
   End If
   '2010/11/17 end
   strExc(0) = strExc(0) & "ORDER BY SORTFIELD DESC "
   '94.2.2
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   Combo1.ListIndex = 0
   'Add by Morgan 2006/3/30 輸入2(改變原處分)時,預設最大發文日的"再審"
   If iSitu = 2 Then
      With RsTemp
      If .RecordCount > 0 Then
         '因為本來排序就由發文日大到小,所以找第一筆再審就好
         .Find "CP10='107'"
         If Not .EOF Then
            MSHFlexGrid1.row = .AbsolutePosition
            GridClick MSHFlexGrid1, intLastRow, 0
         End If
      End If
      End With
   End If
   '2006/3/30 end
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm06010602_2 = Nothing
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1500: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1500: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1500: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1500: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1500: .Text = "發文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 7: .ColWidth(7) = 1500: .Text = "結果"
      .CellAlignment = flexAlignCenterCenter
      .col = 8: .ColWidth(8) = 1500: .Text = "進度備註"
      .CellAlignment = flexAlignCenterCenter
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
   'Add by Morgan 2003/11/25
   If GridDataCheck() = False Then Exit Sub
   '---End
   GridClick MSHFlexGrid1, intLastRow, 0
   cmdOK(0).SetFocus
End Sub
'Add by Morgan 2003/11/25
Private Function GridDataCheck() As Boolean
   
   Dim strSql As String, strTemp As String, bolRtn As Boolean
   Dim rsCheck As New ADODB.Recordset
   
   bolRtn = False
   If (MSHFlexGrid1.row = 0) Then
      bolRtn = True
   ElseIf (pa(9) <> "000") Then
      bolRtn = True
   Else
      RsTemp.Move MSHFlexGrid1.row - 1, 1
      strTemp = RsTemp.Fields("CP10")
      If (Len(strTemp) = 3 And strTemp >= "101" And strTemp <= "105") Then
         strTemp = pa(11)
         'Add by Morgan 2007/3/9 台灣案的基本檔申請號第一碼沒有存 '0'
         If Left(strTemp, 1) <> "0" Then
            strTemp = "0" & strTemp
         End If
         'end 2007/3/9
         If (Trim(strTemp) = Empty) Then
            bolRtn = True
         Else
            strSql = "SELECT PD01||'-'||PD02||PD03||PD04 FROM PRIDATE WHERE PD06='" & strTemp & "' AND PD07='000'"
            rsCheck.CursorLocation = adUseClient
            rsCheck.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If Not (rsCheck.BOF And rsCheck.EOF) Then
               MsgBox "此案已被 " & rsCheck.Fields(0).Value & " 主張國內優先權！", vbCritical
            Else
               bolRtn = True
            End If
            rsCheck.Close
         End If
      Else
         bolRtn = True
      End If
   End If
   GridDataCheck = bolRtn
   
End Function

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   Else
      If KeyAscii = 49 Then
         ReadPatent 1
      Else
         ReadPatent 2
      End If
   End If
End Sub
