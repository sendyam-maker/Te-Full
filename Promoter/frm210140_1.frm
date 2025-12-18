VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm210140_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "開拓資料本所客戶檢查"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   7140
   Begin VB.TextBox txtFileName 
      Height          =   264
      Left            =   450
      TabIndex        =   3
      Top             =   1530
      Width           =   6045
   End
   Begin VB.CommandButton CmdOpenFile 
      Caption         =   "<="
      Height          =   345
      Left            =   6540
      TabIndex        =   2
      Top             =   1500
      Width           =   345
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Default         =   -1  'True
      Height          =   405
      Left            =   6060
      TabIndex        =   1
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton CmdTrans 
      Caption         =   "檢查"
      Height          =   405
      Left            =   5040
      TabIndex        =   0
      Top             =   60
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Caption         =   "程式執行前請確認「公司名稱」不可為非中文名稱，且名稱後面不可有空白。"
      ForeColor       =   &H00FF0000&
      Height          =   370
      Left            =   120
      TabIndex        =   10
      Top             =   330
      Width           =   4995
   End
   Begin VB.Label lblCount 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0 / 0 )"
      Height          =   165
      Left            =   2715
      TabIndex        =   9
      Top             =   2250
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "程式執行過程中不可使用Word軟體，一筆資料需查5~10秒。"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   $"frm210140_1.frx":0000
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1005
      Width           =   6735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "注意事項："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   960
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "開拓資料檔案："
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1290
      Width           =   1320
   End
End
Attribute VB_Name = "frm210140_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/27 Form2.0已檢查 (無需修改的物件)
'Create By Amy 2013/05/31
Option Explicit

Dim strDocTransFileName As String, strTmp As String
Dim i As Integer, intWordRow As Integer
Dim strErrMsg As String, m_strFileName1 As String
Dim bolFirstSA As Boolean 'Modify by Amy 2014/03/05 是否為業務助理第一筆資料
Dim intCol As Integer, CmpName As String, bolSetColor As Boolean 'Modify by Amy 2017/08/30


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub CmdOpenFile_Click()
   Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "" 'Modify by Amy 2014/03/13 原*.doc
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      'Modify by Amy 2014/09/18 經理:大部分的人都用doc 所以限制只能以doc做檢查先不開放docx
      .Filter = "Word檔案 (*.doc 或 *.docx)|*.doc;*.docx" 'Modify by Amy 2014/03/13 +docx
      .Filter = "Word檔案 (*.doc )|*.doc"
      .InitDir = PUB_Getdesktop
      '.MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         txtFileName.Text = .FileName
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub CmdTrans_Click()
 
On Error GoTo ErrHnd

 If txtFileName = "" Then
    MsgBox "檔案不可空白！"
    txtFileName.SetFocus
    Exit Sub
 End If
 
 strDocTransFileName = txtFileName.Text
 
 Dim strSql As String
 Dim insRow As Integer, intCount As Integer, j As Integer
 'Add by Amy 2017/08/30
 Dim strTp As String
 Dim strCheckWay As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String 'Add by Amy 2021/08/13
 
 Screen.MousePointer = vbHourglass
 CmdExit.Enabled = False
 If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
 g_WordAp.Visible = False
 g_WordAp.Documents.Open FileName:=strDocTransFileName
 
 'Add by Amy 2015/03/23 +第一行非表格檢查
 If g_WordAp.Selection.Tables.Count = 0 Then
    MsgBox "請確認檔案中第一行是否為表格！"
    GoTo ChkEnd
 End If

 intWordRow = g_WordAp.Selection.Tables(1).Rows.Count - 1
 'Call ReplaceWord 'Mark by Amy 2014/04/02 怕名稱中間有空白
 ProgressBar1.Min = 0
 ProgressBar1.max = intWordRow
 ProgressBar1.Value = 0
 ProgressBar1.Visible = True
 lblCount.Visible = True
 DoEvents

 'Add by Amy 2017/08/30 排序
 g_WordAp.Selection.Sort ExcludeHeader:=True, FieldNumber:="欄位1", sortFieldType:=wdSortFieldStroke, SortOrder:=wdSortOrderAscending, _
        CaseSensitive:=False, LanguageID:=wdTraditionalChinese
 'end 2017/08/30

 For i = 1 To intWordRow
    strSql = "": CmpName = "": bolFirstSA = False: bolSetColor = False

    '#抓取word資料
      Call RunWordFind(i + insRow) '執行尋找
      CmpName = Trim(RunWordReadData("公司名稱", i + insRow))
      'Add by Amy 2018/08/29 (股)公司取代 為 股份有限公司,取代造字(公報)
      CmpName = Replace(Replace(CmpName, "(股)公司", "股份有限公司"), "（股）公司", "股份有限公司")
      CmpName = Replace(Replace(CmpName, "(股)有限公司", "股份有限公司"), "（股）有限公司", "股份有限公司")
      If ChkBSpecWord(CmpName, strTp) = True Then
        CmpName = strTp
      End If
      'end 2017/08/30
      
      'Add by Amy 2021/08/13 +檢查對造(其他相關人也要出現)
      strSQL1 = " AND CP01 IN (" & SQLGrpStr(GetGroupKindByTwo, 2) & ") "
      strSQL2 = " AND CP01 IN (" & SQLGrpStr("", 1) & ") "
      StrSQL3 = " AND CP01 IN (" & SQLGrpStr("", 3) & ") "
      StrSQL4 = " AND CP01 IN (" & SQLGrpStr("", 4) & ") "
      strSQL5 = " AND CP01 IN (" & SQLGrpStr("", 5) & ") "
      strCheckWay = "="
      Call Pub_ProcR100102_1(strUserNum & "@" & Me.Name, strSQL1, strSQL2, StrSQL3, StrSQL4, strSQL5, ChgSQL(CmpName), strCheckWay)
      'end 2021/08/13
        
      'Modify by Amy 2013/09/18 +因DB部分資料有前後空白故全部 Trim,若有造字查不出
      'Modify by Amy 2014/04/02 智權人員查出來可能為空值,導致轉資料又會轉入,空值時顯示編號或文字 ex:Y5282000(創盟科技股分有限公司)
      'Modify by Amy 2015/03/03 CP50~52原:顯示對造 /+案件性質 for 所有商標案(cp01有T)且案件性質為1202(核駁前先行通知)顯示為 其他相關人
      'Modify by Amy 2015/03/23 +POC14狀態
      'Modify by Amy 2017/08/30 +是否寄電子報
      strSql = "Select CU04 ,NVL(ST02,CU01) ,NVL(CU20,''),NVL(CU30||CU31,NVL(CU112||CU23,'')),NVL(CU16,NVL(CU17,'')),CU01 as CusNo,0 as StartDate,'' as CP10,'' as POC14,Nvl(cu132,'') as SMail From Customer,STAFF WHERE rtrim(ltrim(CU04))='" & CmpName & "' And CU13=ST01(+) "
      strSql = strSql & "Union Select PCU08,NVL(ST02,PCU01),NVL(PCU18,''),NVL(PCU27,''),NVL(PCU13,NVL(PCU14,'')),PCU01 as CusNo,0 as StartDate,'' as CP10,'' as POC14,Nvl(PCU35,'') as SMail From PotCustomer,staff Where rtrim(ltrim(PCU08))='" & CmpName & "' And substr(LTrim(PCU38),1,5)=ST01(+) "
      strSql = strSql & "Union Select POC03,NVL(ST02,POC01),NVL(POC09,''),NVL(POC10,''),NVL(POC05,NVL(POC06,'')),POC01 as CusNo,POC12 as StartDate,'' as CP10,Nvl(POC14,'') as POC14,Nvl(POC11,'') as SMail From PotCustomer1,staff Where rtrim(ltrim(POC03))='" & CmpName & "' And poc13=ST01(+) "
      strSql = strSql & "Union Select FA04,NVL(ST02,FA01||FA02),NVL(FA16,''),NVL(FA17,''),NVL(FA12,NVL(FA13,'')),FA01 as CusNo,0 as StartDate,'' as CP10,'' as POC14,Nvl(FA97,'') as SMail From FAGENT,STAFF Where  rtrim(ltrim(FA04))='" & CmpName & "' And substr(LTrim(FA94),1,5)=ST01(+) "
      strSql = strSql & "Union Select NT02,NVL(ST02,NT01),'',NVL(NT09,''),'',NT01 as CusNo,0 as StartDate,'' as CP10,'' as POC14,'' as SMail From Notagent,STAFF Where rtrim(ltrim(NT02))='" & CmpName & "' And NT18=ST01(+) "
      '聯絡人
      strSql = strSql & "Union Select PCC05,NVL(ST02,PCC01),NVL(PCC08,''),NVL(PCC21||PCC22,''),'',PCC01 as CusNo,0 as StartDate,'' as CP10,'' as POC14,Nvl(PCC10,'') as SMail From (Select * From potcustcont Where rtrim(ltrim(PCC05))='" & CmpName & "') A,CUSTOMER,STAFF Where CU13=ST01(+) AND CU01(+)=PCC01 AND CU02='0' "
      strSql = strSql & "Union Select PCC05,NVL(ST02,PCC01),NVL(PCC08,''),NVL(PCC21||PCC22,''),'',PCC01 as CusNo,0 as StartDate,'' as CP10,'' as POC14,Nvl(PCC10,'') as SMail From (Select * From potcustcont Where rtrim(ltrim(PCC05))='" & CmpName & "' ) A,PotCustomer,STAFF Where PCU01(+)=PCC01 AND PCU02='0' and substr(LTrim(PCU38),1,5)=ST01(+) "
      strSql = strSql & "Union Select PCC05,PCC01,NVL(PCC08,''),NVL(PCC21||PCC22,''),'',PCC01 as CusNo,0 as StartDate,'' as CP10,'' as POC14,Nvl(PCC10,'') as SMail From (Select * From potcustcont Where rtrim(ltrim(PCC05))='" & CmpName & "' ) A,FAGENT Where FA01(+)=PCC01 AND FA02='0' "
      strSql = strSql & "Union Select PCC05,NVL(ST02,PCC01),NVL(PCC08,''),NVL(PCC21||PCC22,''),'',PCC01 as CusNo,0 as StartDate,'' as CP10,'' as POC14,Nvl(PCC10,'') as SMail From (Select * From potcustcont Where rtrim(ltrim(PCC05))='" & CmpName & "') A,PotCustomer1,STAFF Where POC01(+)=PCC01 AND POC02='0' and poc13=ST01(+) "
      'Add by Amy 2013/11/14 +開拓客戶
      strSql = strSql & "Union Select ecd03||' '||ecd04,'外法開拓',Nvl(ECD13,''),ECD05||' '||ECD06||' '||ECD07||' '||ECD08||' '||ECD09,'',ecd02||'-'||LPAD(ecd01,6,'0') as CusNo,0 as StartDate,'' as CP10,'' as POC14,Nvl(ECD14,'') as SMail From ExPandCusDetail Where (ltrim(rtrim(ECD03))='" & CmpName & "' Or ltrim(rtrim(ECD04))='" & CmpName & "') "
      strSql = strSql & "Union Select ecd11||' '||ecd12,'外法開拓',Nvl(ECD13,''),ECD05||' '||ECD06||' '||ECD07||' '||ECD08||' '||ECD09,'',ecd02||'-'||LPAD(ecd01,6,'0') as CusNo,0 as StartDate,'' as CP10,'' as POC14,Nvl(ECD14,'') as SMail From ExPandCusDetail Where (ltrim(rtrim(ECD11))='" & CmpName & "' Or ltrim(rtrim(ECD12))='" & CmpName & "') "
     'end 2014/04/02
     'Modify by Amy 2021/08/13 改成共用function
'     'Add by Amy 2013/09/18 +對造 CP40 CP50
'     'Modify by Amy 2015/12/03 增加案件性質為202/303/404時,案件性質顯示為* (for 原對造顯示為 其他相關人)
'      strSql = strSql & "Union Select CP40,'對造','案號：'||CP01||'-'||CP02||'-'||CP03||'-'||CP04,'','','Z'||CP01||CP02||CP03||CP04 as CusNo,0 as StartDate,Decode(CP10,'1202','*',Decode(CP10,'202','*',Decode(CP10,'303','*',Decode(CP10,'404','P*','')))) as CP10,'' as POC14,'' as SMail From (Select * From CaseProgress Where CP40>' ') Where rtrim(ltrim(CP40))='" & CmpName & "' "
'      strSql = strSql & "Union Select CP50,'其他相關人','案號：'||CP01||'-'||CP02||'-'||CP03||'-'||CP04,'','','Z'||CP01||CP02||CP03||CP04 as CusNo,0 as StartDate,Decode(CP10,'1202','*',Decode(CP10,'202','*',Decode(CP10,'303','*',Decode(CP10,'404','P*','')))) as CP10,'' as POC14,'' as SMail From (Select * From CaseProgress Where CP50>' ') Where rtrim(ltrim(CP50))='" & CmpName & "' "
'      'Add by Amy 2014/03/05 +對造英日(同案號只出現一筆）
'      strSql = strSql & "Union Select CP41,'對造','案號：'||CP01||'-'||CP02||'-'||CP03||'-'||CP04,'','','Z'||CP01||CP02||CP03||CP04 as CusNo,0 as StartDate,Decode(CP10,'1202','*',Decode(CP10,'202','*',Decode(CP10,'303','*',Decode(CP10,'404','P*','')))) as CP10,'' as POC14,'' as SMail From (Select * From CaseProgress Where CP41>' ') Where rtrim(ltrim(CP41))='" & CmpName & "' "
'      strSql = strSql & "Union Select CP51,'其他相關人','案號：'||CP01||'-'||CP02||'-'||CP03||'-'||CP04,'','','Z'||CP01||CP02||CP03||CP04 as CusNo,0 as StartDate,Decode(CP10,'1202','*',Decode(CP10,'202','*',Decode(CP10,'303','*',Decode(CP10,'404','P*','')))) as CP10,'' as POC14,'' as SMail From (Select * From CaseProgress Where CP51>' ') Where rtrim(ltrim(CP51))='" & CmpName & "' "
'      strSql = strSql & "Union Select CP42,'對造','案號：'||CP01||'-'||CP02||'-'||CP03||'-'||CP04,'','','Z'||CP01||CP02||CP03||CP04 as CusNo,0 as StartDate,Decode(CP10,'1202','*',Decode(CP10,'202','*',Decode(CP10,'303','*',Decode(CP10,'404','P*','')))) as CP10,'' as POC14,'' as SMail From (Select * From CaseProgress Where CP42>' ') Where rtrim(ltrim(CP42))='" & CmpName & "' "
'      strSql = strSql & "Union Select CP52,'其他相關人','案號：'||CP01||'-'||CP02||'-'||CP03||'-'||CP04,'','','Z'||CP01||CP02||CP03||CP04 as CusNo,0 as StartDate,Decode(CP10,'1202','*',Decode(CP10,'202','*',Decode(CP10,'303','*',Decode(CP10,'404','P*','')))) as CP10,'' as POC14,'' as SMail From (Select * From CaseProgress Where CP52>' ') Where rtrim(ltrim(CP52))='" & CmpName & "' "
'      'end 2015/12/03
'      'end 2015/03/03
      strSql = strSql & "Union Select R021002,Decode(R021004,1,'對造','其他相關人'),'案號：'||R021001,'','','Z'||R021002 as CusNo,0 as StartDate,Decode(R021018,'1202','*',Decode(R021018,'202','*',Decode(R021018,'303','*',Decode(R021018,'404','P*','')))) as CP10,'' as POC14,'' as SMail From R100102_1 Where ID='" & strUserNum & "@" & Me.Name & "' And R021004<3 "
      'end 2021/08/13
      'Add by Amy 2015/03/26 +國內開拓函特定公司不列印者
      strSql = strSql & "Union Select TBNP01,'***','國內開拓函特定公司不寄','','','' as CusNo,0 as StartDate,'' as CP10,'' as POC14,'' as SMail From TMBulletinnp Where rtrim(ltrim(TBNP01))='" & CmpName & "' "
      'end 2017/08/30
      strSql = "Select * From (" & strSql & ") Order by CusNo,StartDate Desc"

      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
        If RsTemp.RecordCount > 1 Then bolSetColor = True 'Add by Amy 2017/08/30
        For j = 0 To RsTemp.RecordCount - 1 'Modify by Amy 2015/03/03
            If j = 0 Then
                If bolSetColor = True Then
                    g_WordAp.Selection.Cells(1).Select
                    g_WordAp.Selection.Range.HighlightColorIndex = wdYellow
                End If
                For intCol = 1 To RsTemp.Fields.Count - 6 'Modify by Amy 2017/08/30 不顯示不寄電子報 '2015/03/25 不顯示案件性質/POC14 '2014/03/05 原-1 改因不抓編號及POC12
                    g_WordAp.Selection.MoveRight Unit:=wdCell

                    Call BackFillWord 'Modify by Amy 2017/08/30 原程式般至BackFillWord

                Next intCol
                If RsTemp.RecordCount > 1 Then '多筆
                    g_WordAp.Selection.MoveRight Unit:=wdCharacter, Count:=1
                    g_WordAp.Selection.InsertRows RsTemp.RecordCount - 1
                End If
            Else
                insRow = insRow + 1 '新增列數
                g_WordAp.Selection.Tables(1).Rows(i + insRow + 1).Select
                g_WordAp.Selection.Cells(1).Select
                For intCol = 0 To RsTemp.Fields.Count - 6 'Modify by Amy 2017/08/30 不顯示不寄電子報 '2015/03/25 不顯示案件性質/POC14 原-1 改因不抓編號及POC12
                    If intCol = 0 Then
                        g_WordAp.Selection.MoveLeft Unit:=wdCharacter, Count:=1
                    Else
                        g_WordAp.Selection.MoveRight Unit:=wdCell
                    End If
                    Call BackFillWord  'Modify by Amy 2017/08/30 原程式般至BackFillWord
                Next intCol
            End If
            If Not RsTemp.EOF Then
                RsTemp.MoveNext
            End If
        Next j

      End If
      ProgressBar1.Value = ProgressBar1.Value + 1
      lblCount.Caption = "( " & ProgressBar1.Value & " / " & ProgressBar1.max & " )"
      DoEvents
 Next i
  
 Screen.MousePointer = vbDefault
 MsgBox ("檢查已完成！")
 CmdExit.Enabled = True
 ProgressBar1.Visible = False
 lblCount.Visible = False
 g_WordAp.Selection.WholeStory
 g_WordAp.Selection.HomeKey
 g_WordAp.Visible = True
 g_WordAp.Activate
 g_WordAp.WindowState = wdWindowStateMaximize
 Set g_WordAp = Nothing
 
 Exit Sub
   
ErrHnd:
   MsgBox "匯入失敗！" & vbCrLf & Err.Description
ChkEnd:
   g_WordAp.ActiveDocument.Close
   g_WordAp.Quit
   Set g_WordAp = Nothing
   
   Screen.MousePointer = vbDefault
   CmdExit.Enabled = True
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
End Sub

'執行取代空白 2014/04/02 不使用 怕名稱中間有空白
Private Sub ReplaceWord()
    g_WordAp.Selection.Find.ClearFormatting
    g_WordAp.Selection.Find.Replacement.ClearFormatting
    With g_WordAp.Selection.Find
        .Text = " "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchByte = True
    End With
    g_WordAp.Selection.Find.Execute Replace:=wdReplaceAll
   
End Sub

'執行Find
Private Sub RunWordFind(intRow As Integer)
Dim strKey As String
   strKey = "ROW" & String(Len(CStr(intWordRow + intRow)) - Len(CStr(intRow)), "0") & intRow
   With g_WordAp
      .Selection.HomeKey Unit:=wdStory
      .Selection.Find.ClearFormatting
      .Selection.Find.Text = strKey
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
      .Selection.SelectRow
   End With
End Sub

'讀取欄位內容
Private Function RunWordReadData(strTitle As String, intRow As Integer) As String
Dim intCol As Integer
   
   RunWordReadData = ""
   Select Case strTitle
      Case "公司名稱"
         intCol = 1
   End Select
   
   With g_WordAp
      .Selection.Tables(1).Rows(intRow + 1).Select
      .Selection.Cells(intCol).Select
      RunWordReadData = Replace(PUB_StringFilter(Trim(.Selection.Text)), Chr(7), "")
   End With
End Function

'限定字串長度
'Remove by Lydia 2018/08/24 與basQuery重複
'Private Function convForm(ByVal p_InStr As String, ByVal p_Num As Integer, Optional ByVal p_Char As String = " ") As String
'   convForm = StrConv(LeftB(StrConv(p_InStr & String(p_Num, p_Char), vbFromUnicode), p_Num), vbUnicode)
'End Function

'Add by Amy 2017/08/30 寫成共用並修改
Private Sub BackFillWord()
    Dim strTmp As String
    
    If bolSetColor = True Then g_WordAp.Selection.Range.HighlightColorIndex = wdYellow
    If intCol = 1 Then
        'Modify by Amy 2015/03/03 將所有商標案InStr(RsTemp.Fields(2),'T')且案件性質為1202(核駁前先行通知)顯示為 其他相關人 ex:東陽實業廠股份有限公司
        'Modify by Amy 2015/12/03 增加 商標CFC/S-案件性質為202(申請意見書)及303(延期) 及 所有專利-案件性質404(延期) 顯示為 其他相關人
        If RsTemp.Fields(1) = "對造" And (((InStr(RsTemp.Fields(2), "T") > 0 Or Left(Replace("" & RsTemp.Fields(2), "案號：", ""), 4) = "CFC-" Or Left(Replace("" & RsTemp.Fields(2), "案號：", ""), 2) = "S-") _
            And "" & RsTemp.Fields("CP10") = "*") Or ((InStr(RsTemp.Fields(2), "P") > 0 Or Left(Replace("" & RsTemp.Fields(2), "案號：", ""), 3) = "FG-" > 0) And "" & RsTemp.Fields("CP10") = "P*")) Then
            strTmp = "其他相關人"
        'Modify by Amy 2014/03/05 +若為業務助理第一筆資料智權人員欄位顯示開拓日
        ElseIf RsTemp.Fields(1) = "業務助理" And bolFirstSA = False Then
            strTmp = ChangeWStringToTDateString(RsTemp.Fields("StartDate"))
            bolFirstSA = True
        Else
            strTmp = "" & RsTemp.Fields(intCol)
        End If
        g_WordAp.Selection.TypeText Text:=strTmp
    'Add by Amy 2015/03/23 +開拓不寄及國內開拓函特定公司不寄
    'Modify by Amy 2017/08/30 +不寄電子報
    ElseIf intCol = 2 And ("" & RsTemp.Fields("POC14") = "開拓不寄" Or "" & RsTemp.Fields(2) = "國內開拓函特定公司不寄" Or "" & RsTemp.Fields("SMail") = "N") Then
        If RsTemp.Fields("POC14") = "開拓不寄" Then
            g_WordAp.Selection.Font.ColorIndex = wdDarkRed
            strTmp = "" & RsTemp.Fields("POC14")
        Else
            If "" & RsTemp.Fields(2) = "國內開拓函特定公司不寄" Then
                strTmp = "" & RsTemp.Fields(intCol)
            Else
                strTmp = "不寄電子報"
            End If
            g_WordAp.Selection.Font.ColorIndex = wdBlue
        End If
        g_WordAp.Selection.TypeText Text:=strTmp
    ElseIf intCol = 3 Then
        g_WordAp.Selection.TypeText Text:=Replace("" & RsTemp.Fields(intCol), "　", "") '去除郵遞區號多全型空白
    Else
        g_WordAp.Selection.TypeText Text:="" & RsTemp.Fields(intCol)
    End If
    
End Sub


'確認是否公報特殊對照檔有相同公司名稱
Private Function ChkBSpecWord(ByVal stCmpN As String, ByRef stName As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    
    strQ = "Select * From BulletinSpecWord Where BS02='" & stCmpN & "'"
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        ChkBSpecWord = True
        stName = "" & RsQ.Fields("BS03")
    End If
    RsQ.Close
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frm210140_1 = Nothing
End Sub
