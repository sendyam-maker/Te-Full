VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm1105 
   BorderStyle     =   1  '單線固定
   Caption         =   "定稿資料維護"
   ClientHeight    =   5784
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   9000
   Begin VB.CheckBox Check2 
      Caption         =   "不讀原始檔"
      Height          =   225
      Left            =   4650
      TabIndex        =   26
      Top             =   1110
      Width           =   1425
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   5
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1050
      Width           =   2290
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   4
      Left            =   1410
      MaxLength       =   30
      TabIndex        =   5
      Top             =   750
      Width           =   2290
   End
   Begin VB.OptionButton OptKind 
      Caption         =   "審定號數/證書號數："
      Height          =   220
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2055
   End
   Begin VB.OptionButton OptKind 
      Caption         =   "申請案號："
      Height          =   220
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   780
      Width           =   1215
   End
   Begin VB.OptionButton OptKind 
      Caption         =   "本所案號："
      Height          =   220
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含同部門其他人定稿"
      Height          =   225
      Left            =   1680
      TabIndex        =   23
      Top             =   78
      Width           =   1995
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "發 FC 郵件"
      Height          =   320
      Index           =   7
      Left            =   6210
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   300
      Left            =   4605
      Style           =   2  '單純下拉式
      TabIndex        =   21
      Top             =   5355
      Width           =   4305
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "報價轉定稿(&T)"
      CausesValidation=   0   'False
      Height          =   320
      Index           =   6
      Left            =   45
      TabIndex        =   18
      Top             =   30
      Width           =   1290
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "不印(&N)"
      Height          =   320
      Index           =   5
      Left            =   4590
      TabIndex        =   9
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "修改(&E)"
      Height          =   320
      Index           =   0
      Left            =   5385
      TabIndex        =   10
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "刪除(&D)"
      CausesValidation=   0   'False
      Height          =   320
      Index           =   1
      Left            =   6210
      TabIndex        =   11
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "列印(&P)"
      CausesValidation=   0   'False
      Height          =   320
      Index           =   2
      Left            =   7020
      TabIndex        =   12
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   320
      Index           =   3
      Left            =   7824
      TabIndex        =   13
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   320
      Index           =   4
      Left            =   3795
      TabIndex        =   8
      Top             =   24
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   3
      Left            =   3795
      MaxLength       =   2
      TabIndex        =   3
      Top             =   450
      Width           =   420
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   2
      Left            =   3435
      MaxLength       =   1
      TabIndex        =   2
      Top             =   450
      Width           =   276
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   1
      Left            =   2115
      MaxLength       =   6
      TabIndex        =   1
      Top             =   450
      Width           =   1236
   End
   Begin VB.TextBox Text1 
      Height          =   280
      Index           =   0
      Left            =   1410
      MaxLength       =   3
      TabIndex        =   0
      Top             =   450
      Width           =   612
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2775
      Left            =   30
      TabIndex        =   14
      Top             =   2535
      Width           =   8895
      _ExtentX        =   15685
      _ExtentY        =   4890
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   13
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   435
      Left            =   7680
      TabIndex        =   22
      Top             =   660
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2096
      _ExtentY        =   762
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frm1105.frx":0000
   End
   Begin MSForms.Label LblFM2 
      Height          =   240
      Index           =   4
      Left            =   3780
      TabIndex        =   31
      Top             =   5370
      Width           =   795
      Caption         =   "印表機："
      Size            =   "1402;423"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   3
      Left            =   900
      TabIndex        =   30
      Top             =   2190
      Width           =   7965
      Size            =   "14049;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   29
      Top             =   1930
      Width           =   7545
      Size            =   "13309;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   28
      Top             =   1700
      Width           =   7545
      Size            =   "13309;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   27
      Top             =   1470
      Width           =   7545
      Size            =   "13309;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSendMailDt 
      AutoSize        =   -1  'True
      Caption         =   "寄件日期:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   6210
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   3840
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　(日)："
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   20
      Top             =   1930
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "　　　　(英)："
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   17
      Top             =   1700
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "申請人："
      Height          =   225
      Index           =   2
      Left            =   90
      TabIndex        =   16
      Top             =   2190
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱(中)："
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   15
      Top             =   1470
      Width           =   1230
   End
End
Attribute VB_Name = "frm1105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/04/14 改成Form2.0 (LblFM2) ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim strPrinter As String
'Add By Sindy 2012/8/7
Dim strSendDate As String
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_CP09 As String
Dim m_CP10 As String
Dim m_EditMode As String
Dim intRow As Integer
Dim strAppendix As String
'2012/8/7 End
Dim m_CP28 As String 'Add By Sindy 2012/11/08
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
Dim strTemplatePath As String, strTempFolder As String 'Add By Sindy 2014/9/17
Public m_strNationNo As String 'Modify By Sindy 2020/10/30


Private Sub cmdMove_Click(Index As Integer)
Dim ii As Integer
'Add By Cheng 2003/03/27
Dim blnV As Boolean '是否有勾選資料
Dim strContent As String
Dim strFilePathName As String
Dim PrinterIndex As Integer
Dim i As Integer
Dim rsA As New ADODB.Recordset, bolDownFile As Boolean 'Add By Sindy 2015/11/2
'Add By Sindy 2016/7/13
Dim strEfile As String
Dim kk As Integer
Dim stContent As String
'2016/7/13 END
Dim EMailType As String 'Add By Sindy 2016/8/4
Dim strAttach As String 'Add By Sindy 2018/7/17
Dim bolAsked As Boolean 'Added by Morgan 2018/10/15
Dim bolECustLetter As Boolean, strECustNo As String, bolELtrAddr As Boolean 'Added by Morgan 2018/11/1
Dim stPDFfileName As String 'Add By Sindy 2019/3/18
Dim strCmd As String
Dim process_id As Long
Dim process_handle As Long
Dim strMergeFN As String, strMergeName As String
Dim fs
'Add By Sindy 2013/1/4
Dim strCP09 As String
Dim m_bolInsCP As Boolean
Dim strCP10 As String
Dim rsC As New ADODB.Recordset
Dim strNewCP09 As String
Dim strLD11 As String '處理狀況:延展/使用宣誓
'2013/1/4 END
Dim strNP22 As String 'Add By Sindy 2015/4/30
Dim strLP01 As String, stDoc As String, bolOK As Boolean 'Add By Sindy 2019/12/4
Dim bolMergeWord As Boolean 'Add By Sindy 2022/3/29
   
   'Add By Sindy 2011/2/11
   If OptKind(0).Value = True Then 'Added by Lydia 2017/12/19
       If Text1(2) = "" Then Text1(2) = "0"
       If Text1(3) = "" Then Text1(3) = "00"
   End If 'end 2017/12/19
   
   'Add By Cheng 2003/03/27
   blnV = False
   If Me.grdDataList.Rows > 1 Then
       For ii = 1 To Me.grdDataList.Rows - 1
           If Me.grdDataList.TextMatrix(ii, 0) = "V" Then
               blnV = True
               Exit For
           End If
       Next ii
   End If
   
   m_EditMode = 0
   Select Case Index
   Case 0, 7 '0.修改 7.發FC郵件
      m_AutoStampNameInWord = "" 'Added by Morgan 2022/3/22
      m_EditMode = 2
      pub_OsPrinter = PUB_GetOsDefaultPrinter
      PUB_SetOsDefaultPrinter cboPrinter
      'PUB_SetWordActivePrinter
      
      For ii = 1 To Me.grdDataList.Rows - 1
         If Me.grdDataList.TextMatrix(ii, 0) = "V" Then
            Screen.MousePointer = vbHourglass
            
            'Add by Morgan 2008/3/13 電腦中心人員執行時暫時將strUserNum設定為定稿的使用者編號,這樣才抓得到例外欄位資料,定稿產生後再設回來
            'Modify by Morgan 2009/11/10 改可查詢就可編輯
            'If Pub_StrUserSt03 = "M51" Then
               strUserNum = Me.grdDataList.TextMatrix(ii, 8)
           ' End If
           'end 2009/11/10
            'end 2008/3/13
            
            Call GetNowTMNo(ii) 'Added by Lydia 2017/12/20 點選的本所案號
            
            'Add By Sindy 2019/12/4
            '檢查原始檔是否有客戶函/指示信DOC
            'Modified by Morgan 2020/9/18
            'If Index = 0 Then '修改
            If Index = 0 And Check2.Value = vbUnchecked Then '修改
            'end 2020/9/18
               If "" & Me.grdDataList.TextMatrix(ii, 19) <> "" Then
                  strLP01 = Me.grdDataList.TextMatrix(ii, 19)
               ElseIf "" & Me.grdDataList.TextMatrix(ii, 20) <> "" Then
                  strLP01 = Me.grdDataList.TextMatrix(ii, 20)
               End If
               If strLP01 <> "" And Me.grdDataList.TextMatrix(ii, 24) <> "" Then
                  strSql = "select cpf13,cp01,cp02,cp03,cp04,cp10,cpf02 from CasePaperFile,caseprogress" & _
                           " where cpf01='" & strLP01 & "' and cp09(+)=cpf01"
                  '副檔名
                  If Me.grdDataList.TextMatrix(ii, 24) = UCase("CUS") Then
                     strSql = strSql & " and substr(upper(cpf02),-8)='.CUS.DOC'"
                  ElseIf Me.grdDataList.TextMatrix(ii, 24) = UCase("DATA") Then
                     strSql = strSql & " and substr(upper(cpf02),-9)='.DATA.DOC'"
                  ElseIf Me.grdDataList.TextMatrix(ii, 24) = UCase("BLANK") Then
                     strSql = strSql & " and substr(upper(cpf02),-10)='.BLANK.DOC'"
                  End If
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     stDoc = App.path & "\$TEMP"
                     With rsA
                     If PUB_GetFtpFile(.Fields("cpf13"), stDoc, "CASEPAPERFILE", True) Then
                        If PUB_OpenWord(stDoc) = True Then
                           bolOK = True
                        End If
                     End If
                     End With
                  End If
               End If
            End If
            '2019/12/4 END
            
            'Add By Sindy 2014/9/18
            If Index = 7 Then
               '暫存資料夾
               'Modified by Lydia 2020/03/12 改放在App.path
               'strTempFolder = "C:\$$LetterTemp"
               strTempFolder = App.path & "\$$LetterTemp"

               If Dir(strTempFolder, vbDirectory) = "" Then
                  MkDir strTempFolder
               Else
                  If Dir(strTempFolder & "\.") <> "" Then
                     Kill strTempFolder & "\*.*"
                  End If
               End If
               
'               If grdDataList.TextMatrix(ii, 11) <> "04" And _
'                  grdDataList.TextMatrix(ii, 11) <> "08" And _
'                  grdDataList.TextMatrix(ii, 11) <> "14" Then
'                  MsgBox "此定稿不可發FC郵件", vbExclamation + vbOKOnly
'                  Screen.MousePointer = vbDefault
'                  Exit Sub
'               End If
               'Modified by Lydia 2017/12/20 點選的本所案號
               'If Left(Text1(0), 1) <> "T" Then 'Add By Sindy 2016/8/3 非內商案件
               If Left(m_TM01, 1) <> "T" Then
                  'Add By Sindy 2015/11/2 08.催年費催實審 : 改直接從卷宗區讀取客戶函
                  bolDownFile = False
                  'Modify By Sindy 2016/11/18 Mark 江如玉:繳年費通知函, 請改定稿信日期(遇E化案件)-定稿資料維護, 發FC郵件, 附件的PDF信上的日期請自動改為當天的日期, (而非原大批定稿產生的日期)
'                  If Me.grdDataList.TextMatrix(ii, 11) = "08" Then
'                     strFilePathName = Text1(0) & Text1(1) & IIf(Text1(2) & Text1(3) <> "000", Text1(2) & Text1(3), "") & "_Letter"
'                     strSql = "Select cp09,cp10,cp43,cp30,np07,cp66,cp67,cpp14" & _
'                              " From nextProgress, Caseprogress, Casepaperpdf" & _
'                              " Where CP01='" & Text1(0) & "' and CP02='" & Text1(1) & "' and CP03='" & Text1(2) & "' and CP04='" & Text1(3) & "' and CP10='1913'" & _
'                              " and cp43=np01(+) and cp30=np22(+) and np07 in('416','605')" & _
'                              " and cp09=cpp01(+) and cpp02 is not null and instr(upper(cpp02),upper('.CUS.PDF'))>0" & _
'                              " order by cp66 desc,cp67 desc"
'                     intI = 1
'                     Set rsA = ClsLawReadRstMsg(intI, strSql)
'                     If intI = 1 Then
'                        rsA.MoveFirst
'                        bolDownFile = PUB_GetFtpFile(rsA.Fields("cpp14"), strTempFolder & "\" & strFilePathName & ".pdf", "Casepaperpdf")
'                     End If
'                  End If
                  '2015/11/2 END
                  
                  'Add By Sindy 2016/7/13
                  Dim strChk04 As String
                  If Me.grdDataList.TextMatrix(ii, 11) = "07" Then
                     strChk04 = Me.grdDataList.TextMatrix(ii, 4) '定稿別
                     '檢查(定稿,譯文,年費表)合併檔是否存在
                     If Pub_StrUserSt03 = "M51" Then
                        'Modified by Lydia 2017/12/20 點選的本所案號
                        'strEfile = PUB_Getdesktop & "\" & Text1(0) & Text1(1) & "Letter(Patent Certificate).pdf"
                        strEfile = PUB_Getdesktop & "\" & m_TM01 & m_TM02 & "Letter(Patent Certificate).pdf"
                     Else
                        'Modified by Lydia 2017/12/20 點選的本所案號
                        'strEfile = "\\typing2\fcp_workflow\patent certificate\" & Text1(0) & Text1(1) & "Letter(Patent Certificate).pdf"
                        'Modified by Lydia 2024/07/22 改用變數
                        'strEfile = "\\typing2\fcp_workflow\patent certificate\" & m_TM01 & m_TM02 & "Letter(Patent Certificate).pdf"
                        strEfile = "\\" & strTyping2Path & "\fcp_workflow\patent certificate\" & m_TM01 & m_TM02 & "Letter(Patent Certificate).pdf"
                     End If
                     If Dir(strEfile) = "" Then
                        strUserLevel = "發FC郵件" '這電子檔是要E給客戶的,因此不要加蓋Confirmation的章
                        '將多筆定稿內容全部匯入至一個Word檔
'                        stContent = ""
'                        For kk = 1 To (Me.grdDataList.Rows - 1) - 1
'                           'Modify by Amy 2018/07/27 +Me.Name
'                           NowPrint Me.grdDataList.TextMatrix(kk, 1), Me.grdDataList.TextMatrix(kk, 11), Me.grdDataList.TextMatrix(kk, 7), False, strUserNum, , stContent, True, stContent, , , , , , , , , , , , , Me.Name
'                           DoEvents
'                        Next kk
'                        'Modify by Amy 2018/07/27 +Me.Name
'                        NowPrint Me.grdDataList.TextMatrix(kk, 1), Me.grdDataList.TextMatrix(kk, 11), Me.grdDataList.TextMatrix(kk, 7), True, strUserNum, , stContent, , , , , True, , , , , , , , True, , Me.Name
'                        strUserLevel = "" '取消
                        
                        '逐一定稿產生為PDF檔,最後再合併
                        For kk = 1 To (Me.grdDataList.Rows - 1) '- 1
                           'Add by Sindy 2021/3/22 同定稿別才合併
                           'Modify by Sindy 2021/4/16 排除傳真封面: And InStr(Me.grdDataList.TextMatrix(kk, 6), "傳真封面") = 0 ex: FAX 證書 Our Ref: FCP-056612 [PROC.1603]
                           If strChk04 = Me.grdDataList.TextMatrix(kk, 4) And InStr(Me.grdDataList.TextMatrix(kk, 6), "傳真封面") = 0 Then
                           '2021/3/22 END
                              grdDataList.row = kk
                              grdDataList.col = 0
                              grdDataList.Text = ""
                              For i = 0 To grdDataList.Cols - 1
                                 grdDataList.col = i
                                 grdDataList.CellBackColor = QBColor(15)
                              Next i
                              stPDFfileName = "$$Temp" & Format(kk, "000")
                              'Modify by Amy 2018/07/27 +Me.Name
                              NowPrint Me.grdDataList.TextMatrix(kk, 1), Me.grdDataList.TextMatrix(kk, 11), Me.grdDataList.TextMatrix(kk, 7), True, strUserNum, , , , , , , True, , False, , , , , , True, , Me.Name
                              DoEvents
                              
   '                           '用Word轉PDF
   '                           'g_WordAp.ActiveDocument.SaveAs strTempFolder & "\" & stPDFfileName, FileFormat:=0
   '                           g_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strTempFolder & "\" & stPDFfileName, ExportFormat:=17, OpenAfterExport:=False
   '                           g_WordAp.Quit wdDoNotSaveChanges
   '                           'Set g_WordAp = Nothing
   '                           DoEvents
                              
                              'Modify By Sindy 2018/9/20 用Word轉Pdf功能
                              'Removed by Morgan 2025/3/28 已都改用Word另存方式產生PDF
                              'frmPDF.Show
                              'If pub_Word2Pdf Then
                              'end 2025/3/28
                              
                                 g_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strTempFolder & "\" & stPDFfileName & ".pdf", ExportFormat:=17, OpenAfterExport:=False
                                 
                              'Removed by Morgan 2025/3/28 已都改用Word另存方式產生PDF
                              'Else
                              '   '轉PDF
                              '   'frmPDF.Show
                              '   frmPDF.StartProcess strTempFolder, stPDFfileName
                              '   '切換印表機
                              '   If PUB_PdfCreatorNameInWord = "" Then PUB_PdfCreatorNameInWord = PUB_GetCreatorNameInWord
                              '   g_WordAp.ActivePrinter = PUB_PdfCreatorNameInWord
                              '   g_WordAp.ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
                              '   frmPDF.EndtProcess
                              '   'Unload frmPDF
                              'End If
                              'end 2025/3/28
                              
                              g_WordAp.Quit wdDoNotSaveChanges
                              Set g_WordAp = Nothing
                              
                              'Unload frmPDF 'Removed by Morgan 2025/3/28 已都改用Word另存方式產生PDF
                              
                              '記錄檔案位置
                              strMergeFN = strMergeFN & IIf(strMergeFN <> "", " ", "") & ".\" & stPDFfileName & ".pdf"
                           End If
                        Next kk
                        strUserLevel = "" '取消
                        
                        '合併
                        If strTempFolder <> "." Then ChDir strTempFolder '切換至來源目錄
                        strMergeName = "merge" & ServerTime & ".pdf"
                        strCmd = pub_PdftkEXE & " " & strMergeFN & " cat output .\" & strMergeName
                        process_id = SHELL(strCmd, vbHide)
                        process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
                        If process_handle <> 0 Then
                           For intI = 1 To 10
                              If PUB_CheckIsRunning(pub_PdftkName) = True Then
                                 Sleep 1000
                              Else
                                 Exit For
                              End If
                           Next
                           If intI > 10 Then
                              TerminateProcess process_handle, 0&
                              CloseHandle process_handle
                              MsgBox "合併PDF失敗！"
                              Exit Sub
                           Else
                              CloseHandle process_handle
                           End If
                        Else
                           MsgBox "合併PDF失敗！"
                           Exit Sub
                        End If
                        '複製檔案至 Server
                        Set fs = CreateObject("Scripting.FileSystemObject")
                        fs.CopyFile strTempFolder & "\" & strMergeName, strEfile 'Memo by Lydia 2020/03/12 產生的PDF檔回存到\\typing2\fcp_workflow\patent certificate\FCP0xxxxxLetter(Patent Certificate).pdf
                        DoEvents
                     End If
                  End If
                  '2016/7/13 END
               End If
               
               'Modify By Sindy 2015/7/16 定稿別(07.證書函)不夾帶此定稿電子檔
               'Modify By Sindy 2015/9/11 定稿別(05.公告通知函)不夾帶此定稿電子檔
               'Modify By Sindy 2015/11/2 + Me.grdDataList.TextMatrix(ii, 11) <> "08" 催年費催實審
               'Modify By Sindy 2016/8/3 + 內商案件(Left(Text1(0), 1) = "T")
               'Modified by Lydia 2017/12/20 點選的本所案號
               'If Left(Text1(0), 1) = "T" Or
               If Left(m_TM01, 1) = "T" Or _
                  ((Me.grdDataList.TextMatrix(ii, 11) <> "07" And _
                   Me.grdDataList.TextMatrix(ii, 11) <> "05" And _
                   Me.grdDataList.TextMatrix(ii, 11) <> "08") Or _
                   (Me.grdDataList.TextMatrix(ii, 11) = "08" And bolDownFile = False)) Then
               '2015/7/16 END
               
                  '檢查是否有安裝PDFCreator
                  'Removed by Morgan 2025/3/28 已都改用Word另存方式產生PDF
                  'PrinterIndex = -1
                  'For i = 0 To Printers.Count - 1
                  ' If UCase(Printers(i).DeviceName) = UCase$("PDFCreator") Then
                  '  PrinterIndex = i
                  '  Exit For
                  ' End If
                  'Next i
                  'If PrinterIndex < 0 Then
                  '   MsgBox "請通知電腦中心安裝PDFCreator !!!"
                  '   Exit Sub
                  'End If
                  '
                  'pub_OsPrinter = PUB_GetOsDefaultPrinter '取得作業系統預設印表機
                  'PUB_SetOsDefaultPrinter Printers(PrinterIndex).DeviceName 'Printer.DeviceName '作業系統預設印表機指到PDFCreator
                  'PUB_SetWordActivePrinter
                  'end 2025/3/28
                  
                  '產生定稿
                  'NowPrint Me.grdDataList.TextMatrix(ii, 1), Me.grdDataList.TextMatrix(ii, 11), Me.grdDataList.TextMatrix(ii, 7), False, strUserNum, , , , , , , , True, False, , , , grdDataList.TextMatrix(ii, 13), True
                  strUserLevel = "發FC郵件" 'Add By Sindy 2015/7/9
                  'Modify by Amy 2018/07/27 +Me.Name
                  NowPrint Me.grdDataList.TextMatrix(ii, 1), Me.grdDataList.TextMatrix(ii, 11), Me.grdDataList.TextMatrix(ii, 7), True, strUserNum, , , , , , , True, , False, , , , grdDataList.TextMatrix(ii, 13), True, , , Me.Name
                  strUserLevel = "" 'Add By Sindy 2015/7/9
                  'Modified by Lydia 2017/12/20 點選的本所案號
                  'strFilePathName = Text1(0) & Text1(1) & IIf(Text1(2) & Text1(3) <> "000", Text1(2) & Text1(3), "") & "_Letter"
                  strFilePathName = m_TM01 & m_TM02 & IIf(m_TM03 & m_TM04 <> "000", m_TM03 & m_TM04, "") & "_Letter"
                  
                  '轉PDF
                  'strFileName = "$$" & Text1(0) & Text1(1) & Text1(2) & Text1(3) & "_Letter"
                  'Removed by Morgan 2025/3/28 已都改用Word另存方式產生PDF
                  'frmPDF.Show
                  'Modify By Sindy 2020/4/24 用Word轉Pdf功能
                  'If pub_Word2Pdf Then
                  'end 2025/3/28
                  
                     g_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strTempFolder & "\" & strFilePathName & ".pdf", ExportFormat:=17, OpenAfterExport:=False
                     
                  'Removed by Morgan 2025/3/28 已都改用Word另存方式產生PDF
                  'Else
                  '   frmPDF.StartProcess strTempFolder, strFilePathName
                  '   '切換印表機
                  '   'g_WordAp.Visible = True
                  '   If PUB_PdfCreatorNameInWord = "" Then PUB_PdfCreatorNameInWord = PUB_GetCreatorNameInWord
                  '   g_WordAp.ActivePrinter = PUB_PdfCreatorNameInWord
                  '   g_WordAp.ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
                  '   frmPDF.EndtProcess
                  'End If
                  'end 2025/3/28
                  
                  g_WordAp.Quit wdDoNotSaveChanges
                  Set g_WordAp = Nothing
                  
                  'Unload frmPDF 'Removed by Morgan 2025/3/28 已都改用Word另存方式產生PDF
                  
                  PUB_SetOsDefaultPrinter pub_OsPrinter '復原作業系統預設印表機
                  Screen.MousePointer = vbDefault 'Add By Sindy 2020/4/29
                  
                  'Add By Sindy 2018/7/17
                  If Left(m_TM01, 1) = "T" Then
                     strAttach = IIf(strFilePathName = "", "", strTempFolder & "\" & strFilePathName & ".pdf")
                     '夾帶電子公文
                     'Modify By Sindy 2019/12/13 ex:T-218295 1611.對方延期
                     If Left(Me.grdDataList.TextMatrix(ii, 1), 1) = "C" Then
                        strSql = "select cp09,cp43,cp10,cp05,cp27,cp82" & _
                                 " From caseprogress" & _
                                 " where cp09='" & Me.grdDataList.TextMatrix(ii, 1) & "'" & _
                                 " order by cp27||substr('000000'||cp82,-6) desc"
                     Else
                     '2019/12/13 END
                        strSql = "select cp09,cp43,cp10,cp05,cp27,cp82" & _
                                 " From caseprogress" & _
                                 " where cp43='" & Me.grdDataList.TextMatrix(ii, 1) & "'" & _
                                 " order by cp27||substr('000000'||cp82,-6) desc"
                     End If
                     intI = 1
                     Set rsA = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        strFilePathName = Text1(0) & Text1(1) & IIf(Text1(2) & Text1(3) <> "000", Text1(2) & Text1(3), "") & "_Letter" & IIf(strAttach <> "", "2", "")
                        rsA.MoveFirst
                        'Modified by Morgan 2025/3/27 +CPP19
                        strSql = "select cpp14,cpp19" & _
                                 " From casepaperpdf" & _
                                 " where cpp01='" & rsA.Fields("cp09") & "' and cpp02 is not null" & _
                                 " and instr(upper(cpp02),upper('." & rsA.Fields("cp10") & ".PDF'))>0"
                        intI = 1
                        Set rsA = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           If PUB_GetFtpFile(rsA.Fields("cpp14"), strTempFolder & "\" & strFilePathName & ".pdf", "Casepaperpdf", , , "" & rsA.Fields("cpp19") <> "") = True Then
                              If strAttach <> "" Then strAttach = strAttach & ";"
                              strAttach = strAttach & strTempFolder & "\" & strFilePathName & ".pdf"
                           End If
                        End If
                     End If
                  End If
                  '2018/7/17 END
               End If
               '2015/7/16 END
               
               'Modify By Sindy 2016/8/3 + 內商案件
               'Modified by Lydia 2017/12/20 點選的本所案號
               'If Left(Text1(0), 1) = "T" Then 'T開頭即為內商
               If Left(m_TM01, 1) = "T" Then
                  '取得郵件範本檔名
                  'Modify By Sindy 2018/5/14 Mark : 欲改call Form
'                  strTemplatePath = PUB_DownloadOftPath("P20", "", EMailType, False)
                  'Modified by Lydia 2017/12/20 點選的本所案號
                  'Call PUB_SettingTeMail(strTemplatePath, _
                                         Me.Text1(0), Me.Text1(1), Me.Text1(2), Me.Text1(3), _
                                         IIf(strFilePathName = "", "", strTempFolder & "\" & strFilePathName & ".pdf"), _
                                         Me.grdDataList.TextMatrix(ii, 21), _
                                         Me.grdDataList.TextMatrix(ii, 11))
                  lblSendMailDt.Visible = True 'Add By Sindy 2018/5/14
                  'Add By Sindy 2020/4/29 讀取要歸寄件備份的總收文號 ex:T-227182
                  strCP09 = ""
                  If InStr(Me.grdDataList.TextMatrix(ii, 1), "&") > 0 Then
                     strCP10 = Mid(Me.grdDataList.TextMatrix(ii, 1), InStr(Me.grdDataList.TextMatrix(ii, 1), "&") + 1)
                     strSql = "SELECT cp09,cp10 FROM caseprogress" & _
                              " WHERE cp43 IN(SELECT cp09 FROM caseprogress WHERE cp01='" & Me.grdDataList.TextMatrix(ii, 14) & "' AND cp02='" & Me.grdDataList.TextMatrix(ii, 15) & "' AND cp03='" & Me.grdDataList.TextMatrix(ii, 16) & "' AND cp04='" & Me.grdDataList.TextMatrix(ii, 17) & "' AND cp10='" & strCP10 & "')" & _
                              " order by cp27 desc"
                     intI = 1
                     Set rsA = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        strCP09 = rsA.Fields("cp09")
                        strCP10 = rsA.Fields("cp10")
                     End If
                  Else
                     strCP09 = Me.grdDataList.TextMatrix(ii, 1)
                     strCP10 = Me.grdDataList.TextMatrix(ii, 21)
                  End If
                  'Me.grdDataList.TextMatrix(ii, 1) => strCP09
                  'Me.grdDataList.TextMatrix(ii, 21) => strCP10
                  Call PUB_SettingTeMail(Me, strTemplatePath, m_TM01, m_TM02, m_TM03, m_TM04, strAttach, _
                                         strCP10, strCP09, _
                                         Me.grdDataList.TextMatrix(ii, 11), , , lblSendMailDt)
               Else
               '2016/8/3 END
                  '呼叫發FC郵件
                  frm090401.bolFrom1105Callme = True
                  'Add By Sindy 2015/1/8 傳入案件性質
                  'Add By Sindy 2015/5/29 傳入定稿別和處理狀況
                  'Modify By Sindy 2018/5/18
                  '此定稿的案件性質:主要是年費申請人會有差別
                  If Me.grdDataList.TextMatrix(ii, 18) <> "" And Me.grdDataList.TextMatrix(ii, 18) <> "000" Then
                     frm090401.OutCallCP10 = Me.grdDataList.TextMatrix(ii, 18)
                  ElseIf Me.grdDataList.TextMatrix(ii, 21) <> "" And Me.grdDataList.TextMatrix(ii, 21) <> "000" Then
                     frm090401.OutCallCP10 = Me.grdDataList.TextMatrix(ii, 21)
                  ElseIf InStr(Me.grdDataList.TextMatrix(ii, 1), "&") > 0 Then
                     frm090401.OutCallCP10 = Mid(Me.grdDataList.TextMatrix(ii, 1), InStr(Me.grdDataList.TextMatrix(ii, 1), "&") + 1)
                  End If
                  If frm090401.OutCallCP10 = "000" Then frm090401.OutCallCP10 = ""
                  '要帶入主旨使用的案件性質,分信歸檔使用
                  If m_TM01 = "FCP" And Me.grdDataList.TextMatrix(ii, 11) = "08" Then '08-繳年費通知函 (實審通知函.1913)
                     frm090401.OutCallProcCP10 = "1913" '[PROC.1913]=通知期限
                  'Modify By Sindy 2018/5/21
                  ElseIf m_TM01 = "FCP" And Me.grdDataList.TextMatrix(ii, 11) = "04" Then '04-核准函
                     frm090401.OutCallProcCP10 = "1917" '[PROC.1917]=通知告淮
                  'Modify By Sindy 2018/5/22
                  ElseIf m_TM01 = "FCP" And Me.grdDataList.TextMatrix(ii, 11) = "14" Then '14-公開通知函
                     frm090401.OutCallProcCP10 = "1229" '[PROC.1229]=公開公報
                  ElseIf m_TM01 = "FCP" And Me.grdDataList.TextMatrix(ii, 11) = "05" Then '05-公告通知函
                     frm090401.OutCallProcCP10 = "1228" '[PROC.1228]=公告公報
                  '2018/5/22 END
                  End If
                  'Add By Sindy 2019/5/23
                  If frm090401.OutCallProcCP10 = "" Then
                     frm090401.OutCallProcCP10 = frm090401.OutCallCP10
                  End If
                  '2019/5/23 END
                  '2018/5/18 END
                  '2015/1/8 END
                  'Add By Sindy 2015/5/29
                  frm090401.m_ET01 = Me.grdDataList.TextMatrix(ii, 11) '定稿別
                  frm090401.m_ET03 = Me.grdDataList.TextMatrix(ii, 7) '處理狀況
                  '2015/5/29 END
                  frm090401.m_ET99 = Me.grdDataList.TextMatrix(ii, 5) '份數 Add By Sindy 2017/10/6
                  frm090401.OutCallCP09 = Me.grdDataList.TextMatrix(ii, 1) 'Add By Sindy 2015/6/18
                  frm090401.Hide
                  frm090401.EMailType = "HTML"
                  frm090401.strAttach = IIf(strFilePathName = "", "", strTempFolder & "\" & strFilePathName & ".pdf")
                  'Modified by Lydia 2017/12/20 點選的本所案號
'                  frm090401.Text1 = Me.Text1(0)
'                  frm090401.Text2 = Me.Text1(1)
'                  frm090401.Text3 = Me.Text1(2)
'                  frm090401.Text4 = Me.Text1(3)
                  frm090401.Text1 = m_TM01
                  frm090401.Text2 = m_TM02
                  frm090401.Text3 = m_TM03
                  frm090401.Text4 = m_TM04
                  'end 2017/12/20
                  Call frm090401.Read
                  Call frm090401.cmdFCMail_Click(1)
                  Unload frm090401
               End If
'               If m_strFilePath = "" Then
'                  MsgBox "無讀到定稿的檔案路徑，請重新操作！", vbExclamation + vbOKOnly
'                  Screen.MousePointer = vbDefault
'                  Exit Sub
'               Else
'                  If PUB_DocToHtml(m_strFilePath, Text1(0) & Text1(1) & Text1(2) & Text1(3), strFilePathName) = False Then
'                     MsgBox "定稿DOC轉HTML語法有誤，請重新操作！", vbExclamation + vbOKOnly
'                     Screen.MousePointer = vbDefault
'                     Exit Sub
'                  End If
'                  'HTML內文
'                  RichTextBox1.LoadFile (strFilePathName)
'                  strContent = RichTextBox1.Text
'                  'Modify By Sindy 2014/9/26 Mark 此段程式在Word 2007會錯
''                  strContent = Mid(RichTextBox1.Text, InStr(UCase(RichTextBox1.Text), "<BODY>") + 6)
''                  If InStr(UCase(strContent), UCase("Best regards,")) > 0 Then
''                     strContent = Left(strContent, InStr(UCase(strContent), UCase("Best regards,")) - 1)
''                  Else
''                     strContent = Left(strContent, InStr(UCase(strContent), "</BODY>") - 1)
''                  End If
'                  '2014/9/26 END
'                  strContent = Replace(strContent, vbCrLf, "")
'                  '呼叫發FC郵件
'                  Call PUB_SettingFCeMail(m_StrUserSt03, strTemplatePath, "HTML", _
'                                          Text1(0), Text1(1), Text1(2), Text1(3), _
'                                          strContent)
'               End If
            Else
            '2014/9/18 END
               'Added by Lydia 2017/06/27 更新FCT延展、移轉、變更(102,501,301)案之核准定稿日期
               'Modified by Lydia 2017/09/18 +LD03定稿時間
               UpdateFCTld02et07 Me.grdDataList.TextMatrix(ii, 8), Me.grdDataList.TextMatrix(ii, 9), Me.grdDataList.TextMatrix(ii, 1), PUB_MGridGetValue(ii, "LD05", Me.grdDataList), Me.grdDataList.TextMatrix(ii, 11), Me.grdDataList.TextMatrix(ii, 7), Me.grdDataList.TextMatrix(ii, 10)
               
               'Modify By Sindy 2019/12/4
               '修改
               If bolOK = False Then  '無原始檔
               '2019/12/4 END
                  'Modify by Morgan 2010/3/4 FMP定稿要加傳真封面及信頭(譯文除外)
                  If strUserNum = strFMPNum And InStr(grdDataList.TextMatrix(ii, 6), "譯文") = 0 Then
                     'Modified by Morgan 2016/11/21
                     'If Left(grdDataList.TextMatrix(ii, 1), 3) = "FCP" Then
                     strExc(9) = ""
                     'Modify by Amy 2018/07/27 +Me.Name
                     If grdDataList.TextMatrix(ii, 14) = "FCP" Then
                     'end 2016/11/21
                        NowPrint Me.grdDataList.TextMatrix(ii, 1), "04", "97", False, strUserNum, , , True, strExc(9), , , True, , False, , , , grdDataList.TextMatrix(ii, 13), , , , Me.Name
                     Else
                        NowPrint Me.grdDataList.TextMatrix(ii, 1), "01", "98", False, strUserNum, , , True, strExc(9), , , True, , False, , , , grdDataList.TextMatrix(ii, 13), , , , Me.Name
                     End If
                     DoEvents
                     NowPrint Me.grdDataList.TextMatrix(ii, 1), Me.grdDataList.TextMatrix(ii, 11), Me.grdDataList.TextMatrix(ii, 7), True, strUserNum, , strExc(9), , , , , True, , False, , 2, , grdDataList.TextMatrix(ii, 13), , , , Me.Name
                  Else
                     'Added by Morgan 2018/11/11
                     '商標ｅ化客戶案件也要能上傳及EMail
                     'Modified by Morgan 2021/8/18 內商也已電子化不必
                     If strSrvDate(1) >= e化客戶啟用日 And (grdDataList.TextMatrix(ii, 14) = "CFT" Or grdDataList.TextMatrix(ii, 14) = "CFC") Then
                        bolECustLetter = ChkECustLetter(grdDataList.TextMatrix(ii, 8), grdDataList.TextMatrix(ii, 9), grdDataList.TextMatrix(ii, 10), strECustNo, bolELtrAddr)
                        g_bolELtrAddr = bolELtrAddr
                     End If
                     'end 2018/11/11
                     
                     'Modify By Sindy 2022/3/29 有信函文號的定稿,就檢查是否為多個函,若是就合併在一個Word檔案裡
                     'grdDataList.TextMatrix(ii, 24): 信函副檔名 LD27 in('CUS','BLANK')
                     bolMergeWord = False: strExc(9) = ""
                     If grdDataList.TextMatrix(ii, 13) <> "" Then
                        strSql = "select LP01,cp01,cp02,cp03,cp04,cp10,LP41,cp27,LETTERDEMAND.* from letterprogress,caseprogress,LETTERDEMAND" & _
                                 " where cp09(+)=lP01 AND LD18(+)=CP09 and ld01 is not null and ld18='" & grdDataList.TextMatrix(ii, 13) & "'" & _
                                 " AND LD27='" & grdDataList.TextMatrix(ii, 24) & "'" & _
                                 " ORDER BY LD02 ASC,LD03 ASC"
                        intI = 1
                        Set rsA = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 And rsA.RecordCount > 1 Then
                           bolMergeWord = True
                           rsA.MoveFirst
                           For i = 1 To rsA.RecordCount - 1
                              NowPrint rsA.Fields("ld04"), rsA.Fields("ld10"), rsA.Fields("ld11"), True, strUserNum, , , True, strExc(9), , , , , False, , , , rsA.Fields("lp01"), , , IIf(rsA.Fields("cp10") = "990", True, False), Me.Name
                              rsA.MoveNext
                           Next i
                           NowPrint rsA.Fields("ld04"), rsA.Fields("ld10"), rsA.Fields("ld11"), True, strUserNum, , strExc(9), , , , , , , False, , , , rsA.Fields("lp01"), , , IIf(rsA.Fields("cp10") = "990", True, False), Me.Name
                        End If
                     End If
                     If bolMergeWord = False Then
                     '2022/3/29 END
                        NowPrint Me.grdDataList.TextMatrix(ii, 1), Me.grdDataList.TextMatrix(ii, 11), Me.grdDataList.TextMatrix(ii, 7), True, strUserNum, , , , , , , , , False, , , , grdDataList.TextMatrix(ii, 13), , , IIf(grdDataList.TextMatrix(ii, 18) = "990", True, False), Me.Name
                     End If
                     
                     g_bolELtrAddr = False 'Added by Morgan 2018/11/21
                  End If
                  'end 2018/07/27
               End If
            End If
            DoEvents
            strUserNum = strUser1Num 'Add by Morgan 2008/3/13
            Screen.MousePointer = vbDefault

            'Added by Morgan 2014/3/28
            With grdDataList
            '有信函進度的進修改畫面
            'Modified by Morgan 2017/12/25 +CFP
            'Modified by Morgan 2018/11/1 +bolECustLetter(商標ｅ化客戶案件)
            'Modify By Sindy 2019/11/5 不鎖系統別
            'If bolECustLetter Or ((.TextMatrix(ii, 14) = "P" Or .TextMatrix(ii, 14) = "CFP") And (.TextMatrix(ii, 19) <> "" Or .TextMatrix(ii, 20) <> "")) Then
            'Modify By Sindy 2019/11/19 + And Index = 0 => 0.修改
            If (bolECustLetter Or .TextMatrix(ii, 19) <> "" Or .TextMatrix(ii, 20) <> "") And _
               Index = 0 Then
            '2019/11/5 END
            
               'Added by Morgan 2025/8/11 若定稿為案件回覆單且已收文時會沒有內容，繼續往下執行會出錯 Ex:CFP-034181
               If g_WordAp Is Nothing Then
                  MsgBox "定稿載入失敗！" & vbCrLf & vbCrLf & "注意:案件回覆單若已收文將無定稿可載入。", vbCritical
                  Exit For
               End If
               'end 2025/8/11
            
               'Added by Morgan 2018/10/15
               '檢查若已定稿維護畫面已開啟時確認是否只開Word並提醒無法直接上傳
               If PUB_CheckFormExist("frm1105_1") Then
                  If Not bolAsked Then
                     MsgBox "定稿維護畫面已開啟，此次修改將以 Word 開啟且無法直接上傳！", vbExclamation
                     bolAsked = True
                  End If
                  
               'Added by Morgan 2018/11/1
               '商標ｅ化客戶案件
               ElseIf bolECustLetter Then
                  frm1105_1.m_eCustNo = strECustNo
                  frm1105_1.m_RecNo = .TextMatrix(ii, 1)
                  frm1105_1.Show
                  frm1105_1.cboPrinter.Text = cboPrinter.Text
               'end 2018/11/1
               
               Else
               'end 2018/10/15
                  'Modified by Morgan 2015/11/4 +14解除期限指示信
                  'Modified by Morgan 2016/3/30
                  'If .TextMatrix(ii, 20) <> "" And (.TextMatrix(ii, 11) = "01" Or .TextMatrix(ii, 11) = "16" Or .TextMatrix(ii, 11) = "14") Then
                  'Modified by Morgan 2016/5/23 要判斷是申請書/指示信還是客戶函
                  'If .TextMatrix(ii, 20) <> "" Then
                  'Modified by Morgan 2020/4/14 改抓 LD27
                  'If .TextMatrix(ii, 20) <> "" And (.TextMatrix(ii, 12) = "5" Or .TextMatrix(ii, 12) = "6") Then
                  If .TextMatrix(ii, 20) <> "" And .TextMatrix(ii, 24) = "DATA" Then
                  'end 2016/3/30
                     frm1105_1.m_RecNo = .TextMatrix(ii, 20)
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(.TextMatrix(ii, 14), .TextMatrix(ii, 15), .TextMatrix(ii, 16), .TextMatrix(ii, 17)) & "." & .TextMatrix(ii, 18) & ".DATA.PDF"
                     frm1105_1.Show
                     frm1105_1.cboPrinter.Text = cboPrinter.Text
                  ElseIf .TextMatrix(ii, 19) <> "" Then
                     frm1105_1.m_RecNo = .TextMatrix(ii, 19)
                     If .TextMatrix(ii, 23) = "Y" Then '是否為回覆單
                        frm1105_1.m_PdfName = PUB_CaseNo2FileName(.TextMatrix(ii, 14), .TextMatrix(ii, 15), .TextMatrix(ii, 16), .TextMatrix(ii, 17)) & "." & .TextMatrix(ii, 18) & ".BLANK.PDF"
                     Else
                        frm1105_1.m_PdfName = PUB_CaseNo2FileName(.TextMatrix(ii, 14), .TextMatrix(ii, 15), .TextMatrix(ii, 16), .TextMatrix(ii, 17)) & "." & .TextMatrix(ii, 18) & ".CUS.PDF"
                     End If
                     frm1105_1.Show
                     frm1105_1.cboPrinter.Text = cboPrinter.Text
                  End If
                  Exit For
               End If 'Added by Morgan 2018/10/15
            End If
            End With
            'end 2014/3/28
            
            'Modify By Cheng 2002/11/14
            '讓Word直接跳至畫面
'            'Add By Cheng 2002/10/28
'            MsgBox "修改完成!!!", vbExclamation + vbOKOnly
            'Modify By Cheng 2003/03/27
'            Exit For
         End If
      Next ii
'      If ii = Me.grdDataList.Rows Then
      If blnV = False Then
         MsgBox "請點選欲修改的定稿!!!", vbExclamation + vbOKOnly
      End If
        
      PUB_SetOsDefaultPrinter pub_OsPrinter
      
   Case 1 '刪除
      m_EditMode = 3
      'Added by Morgan 2011/12/13 問一次以免誤點
      If MsgBox("您是否要刪除定稿資料???", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
      
      For ii = 1 To Me.grdDataList.Rows - 1
         If Me.grdDataList.TextMatrix(ii, 0) = "V" Then
'            If MsgBox("您是否要刪除定稿資料???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
               DeleteData Me.grdDataList.TextMatrix(ii, 8), Me.grdDataList.TextMatrix(ii, 9), Me.grdDataList.TextMatrix(ii, 10)
'                'Add By Cheng 2002/10/28
'                MsgBox "刪除完成!!!", vbExclamation + vbOKOnly
'            End If
'            Exit For
         End If
      Next ii
'      If ii = Me.grdDataList.Rows Then
        If blnV = False Then
            MsgBox "請點選欲刪除的定稿!!!", vbExclamation + vbOKOnly
        Else
            QueryData
            MsgBox "刪除完成!!!", vbExclamation + vbOKOnly
        End If
   Case 2 '列印
      pub_OsPrinter = PUB_GetOsDefaultPrinter
      PUB_SetOsDefaultPrinter cboPrinter
      PUB_SetWordActivePrinter
      
      For ii = 1 To Me.grdDataList.Rows - 1
         If Me.grdDataList.TextMatrix(ii, 0) = "V" Then
            Screen.MousePointer = vbHourglass
            'Add by Morgan 2008/3/13 電腦中心人員執行時暫時將strUserNum設定為定稿的使用者編號,這樣才抓得到例外欄位資料,定稿產生後再設回來
            'Modify by Morgan 2009/11/10 改可查詢就可列印
            'If Pub_StrUserSt03 = "M51" Then
               strUserNum = Me.grdDataList.TextMatrix(ii, 8)
            'End If
            'end 2009/11/10
            'end 2008/3/13
            
            'Add By Sindy 2013/1/4 要產生來函通知進度
            strCP09 = Me.grdDataList.TextMatrix(ii, 1)
            'Modified by Lydia 2017/12/20 改成模組
'            m_TM01 = SystemNumber(Me.grdDataList.TextMatrix(ii, 2), 1)
'            m_TM02 = SystemNumber(Me.grdDataList.TextMatrix(ii, 2), 2)
'            m_TM03 = SystemNumber(Me.grdDataList.TextMatrix(ii, 2), 3)
'            m_TM04 = SystemNumber(Me.grdDataList.TextMatrix(ii, 2), 4)
            Call GetNowTMNo(ii)
            'end 2017/12/20
            strLD11 = Trim(Me.grdDataList.TextMatrix(ii, 7)) '處理狀況
            If strLD11 = "01" Or strLD11 = "02" Or strLD11 = "03" Or strLD11 = "07" Or strLD11 = "09" Or strLD11 = "11" Then
               strLD11 = "延展"
            ElseIf strLD11 = "04" Or strLD11 = "05" Or strLD11 = "08" Or strLD11 = "10" Or strLD11 = "12" Then
               strLD11 = "使用宣誓"
            End If
            If m_TM01 = "CFT" And _
               Trim(Me.grdDataList.TextMatrix(ii, 4)) = "業務員期限管制表" And _
               (strLD11 = "延展" Or strLD11 = "使用宣誓") Then
               m_bolInsCP = True
               'Modify By Sindy 2015/4/30 要抓取NP22值
               'strSql = "select cp10 from caseprogress where cp09='" & strCP09 & "'"
               strSql = "select cp10,np07,np22 from caseprogress,nextprogress where cp09='" & strCP09 & "' and cp09=np01(+)"
               '2015/4/30 END
               rsC.CursorLocation = adUseClient
               rsC.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsC.RecordCount > 0 Then
                  strNP22 = "" & rsC.Fields("np22") 'Add By Sindy 2015/4/30
                  strCP10 = "" & rsC.Fields("cp10")
               End If
               If rsC.State <> adStateClosed Then rsC.Close
               '若為使用宣誓時,收文的案件性質為1711.通知使用宣誓時不產生進度
               If strLD11 = "使用宣誓" Then
                  If strCP10 = "1711" Then
                     m_bolInsCP = False
                  End If
               End If
               If m_bolInsCP = True Then
                  strSql = "select cp05 from caseprogress " & _
                           "where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "'"
                  If strLD11 = "延展" Then
                     strSql = strSql & " and cp10='1717'"
                     strCP10 = "1717"
                  ElseIf strLD11 = "使用宣誓" Then
                     strSql = strSql & " and cp10='1723'"
                     strCP10 = "1723"
                  End If
                  strSql = strSql & " order by cp05 desc"
                  rsC.CursorLocation = adUseClient
                  rsC.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsC.RecordCount > 0 Then
                     rsC.MoveFirst
                     'Modified by Lydia 2016/11/25 因為在報價轉定稿會將報價點寫到進度備註,所以當天重印不再產生定稿 (ex. CFT-001411)
                     If rsC.Fields(0) = strSrvDate(1) Then
                         m_bolInsCP = False
                     Else
                         '檢查30天內重新以此方式作業,以詢問方式確認是否要產生通知進度
                         If CompWorkDay(30, rsC.Fields(0), 0) >= strSrvDate(1) Then
                            'Modified by Lydia 2016/11/25 次日之後　(ex. CFT-007663)
                            'If MsgBox("是否產生來函通知進度？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                            If MsgBox("30天內已有來函通知進度，是否重新產生來函通知進度？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                               m_bolInsCP = False
                            End If
                         End If
                     End If
                     'end 2016/11/25
                  End If
                  If rsC.State <> adStateClosed Then rsC.Close
                  Set rsC = Nothing
                  If m_bolInsCP = True Then
                     'Modified by Lydia 2016/12/22 本所管控C類進度自2017/01/01起改用D類收文
                     'strNewCP09 = AutoNo("C", 6)
                     If strSrvDate(1) >= 本所D類收文啟用日 Then
                         strNewCP09 = AutoNo("D", 6)
                     Else
                         strNewCP09 = AutoNo("C", 6)
                     End If
                    'end 2016/12/22
                     
                     'Modify By Sindy 2015/4/30 +CP30
                     strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP27,CP30) " & _
                     "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & _
                             "'" & strNewCP09 & "','" & strCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & strUserNum & "'," & _
                             "'N','N','N'," & _
                             "'" & strCP09 & "'," & strSrvDate(1) & "," & CNULL(strNP22) & ")"
                     cnnConnection.Execute strSql
                  End If
               End If
            End If
            '2013/1/4 End
            'Modify by Amy 2018/07/30 +showColorFirst=>是否先抓彩色代表圖
            PrinterLetterDB strUserNum, Me.grdDataList.TextMatrix(ii, 12), Me.grdDataList.TextMatrix(ii, 9), Me.grdDataList.TextMatrix(ii, 10), , , , , Me.Name
            'Added by Lydia 2017/06/27 更新FCT延展、移轉、變更(102,501,301)案之核准定稿日期
            'Modified by Lydia 2017/09/18 +LD03定稿時間
            UpdateFCTld02et07 Me.grdDataList.TextMatrix(ii, 8), Me.grdDataList.TextMatrix(ii, 9), Me.grdDataList.TextMatrix(ii, 1), PUB_MGridGetValue(ii, "LD05", Me.grdDataList), Me.grdDataList.TextMatrix(ii, 11), Me.grdDataList.TextMatrix(ii, 7), Me.grdDataList.TextMatrix(ii, 10)
            
            strUserNum = strUser1Num 'Add by Morgan 2008/3/13
            Screen.MousePointer = vbDefault
         End If
      Next ii
      If blnV = False Then
         MsgBox "請點選欲列印的定稿!!!", vbExclamation + vbOKOnly
      Else
         MsgBox "定稿列印完成!!!", vbExclamation + vbOKOnly
      End If
      PUB_SetOsDefaultPrinter pub_OsPrinter
      
   Case 3 '離開
      Unload Me
   Case 4 ' 查詢
      m_EditMode = 4
      QueryData
    'Add By Cheng 2003/03/13
   Case 5 '不印(上列印註記)
      For ii = 1 To Me.grdDataList.Rows - 1
         If Me.grdDataList.TextMatrix(ii, 0) = "V" Then
               MarkData Me.grdDataList.TextMatrix(ii, 8), Me.grdDataList.TextMatrix(ii, 9), Me.grdDataList.TextMatrix(ii, 10)
'                MsgBox "定稿註記不印完成!!!", vbExclamation + vbOKOnly
                Screen.MousePointer = vbDefault
'            Exit For
         End If
      Next ii
'      If ii = Me.grdDataList.Rows Then
        If blnV = False Then
           MsgBox "請點選欲不印 (上列印註記) 的定稿!!!", vbExclamation + vbOKOnly
        Else
            MsgBox "定稿註記不印完成!!!", vbExclamation + vbOKOnly
        End If
   'Add by Morgan 2009/9/15
   Case 6 '已確認報價轉定稿
      'Added by Lydia 2017/12/19 用本所案號查詢
      If OptKind(0).Value = False Then
          MsgBox "請選擇本所案號!!!", vbExclamation + vbOKOnly
           Screen.MousePointer = vbDefault
           Exit Sub
      End If
      If Trim(Me.Text1(0).Text) = "" Or Trim(Me.Text1(0).Text) = "" Then
        If Me.Text1(0).Text = "" Then
           MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
           Me.Text1(0).SetFocus
           Text1_GotFocus 0
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
        If Me.Text1(1).Text = "" Then
           MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
           Me.Text1(1).SetFocus
           Text1_GotFocus 1
           Screen.MousePointer = vbDefault
           Exit Sub
        End If
      End If
      'end 2017/12/19
      Screen.MousePointer = vbHourglass 'Added by Lydia 2016/11/25
      Transfer
      Screen.MousePointer = vbDefault 'Added by Lydia 2016/11/25
   End Select
   
   Set rsA = Nothing 'Add By Sindy 2015/11/2
End Sub

Private Sub DeleteData(strDL01 As String, strDL02 As String, strDL03 As String)
   strSql = "Delete From LETTERDEMAND WHERE LD01='" & strDL01 & "' AND LD02=" & Val(strDL02) & " AND LD03=" & Val(strDL03) & " "
   cnnConnection.Execute strSql
   'Added by Lydia 2024/05/03 刪除例外欄位; T-246593的變更(AB2046388)核准，分別在112/12/14的湘芸和112/12/28的桂英輸入核准，雖然12/14的D類收文和定稿刪除，但是保留exceptcondition造成12/28的收款寄証發文時產生Duplicate Key錯誤
   strSql = "delete from exceptcondition where et04='" & strDL01 & "' and et07=" & Val(strDL02) & " and et08=" & Val(Left(String(6 - Len(strDL03), "0") & strDL03, 4))
   cnnConnection.Execute strSql
End Sub

'Add By Cheng 2003/03/13
'上列印註記
Private Sub MarkData(strDL01 As String, strDL02 As String, strDL03 As String)
   strSql = "Update LETTERDEMAND Set LD16='*' WHERE LD01='" & strDL01 & "' AND LD02=" & Val(strDL02) & " AND LD03=" & Val(strDL03) & " "
   cnnConnection.Execute strSql
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   'Add by Morgan 2010/2/1 +可選擇印表機
   PUB_SetPrinter Me.Name, cboPrinter, strPrinter
   
   'Modify By Sindy 2014/9/18
   'Download郵件範本
   'Modify By Sindy 2016/8/3 + 開放內商人員
   If Pub_StrUserSt03 = "F22" Then
      strTemplatePath = PUB_DownloadOftPath("F23", "")
      cmdMove(7).Visible = True
   'Modify By Sindy 2021/6/23 電子化已穩定上線,Mark
'   ElseIf Pub_StrUserSt03 = "P20" Or Pub_StrUserSt03 = "P21" Or Pub_StrUserSt03 = "P22" Then
'      strTemplatePath = PUB_DownloadOftPath(Pub_StrUserSt03, "")
'      cmdMove(7).Visible = True
   ElseIf Pub_StrUserSt03 = "M51" Then
      strTemplatePath = PUB_DownloadOftPath("F23", "")
      strTemplatePath = PUB_DownloadOftPath("P20", "")
      cmdMove(7).Visible = True
   '2016/8/3 END
   Else
      cmdMove(7).Visible = False
   End If
   
   'Added by Morgan 2017/2/16
   If Pub_StrUserSt03 = "F12" Or Pub_StrUserSt03 = "F22" Or Pub_StrUserSt03 = "P12" Or Pub_StrUserSt03 = "P22" Then
      Check1.Visible = True
   Else
      Check1.Visible = False
   End If
   'end 2017/2/16
   
   'Added by Morgan 2020/9/18 測試定稿修改是否正確用
   If Pub_StrUserSt03 = "M51" Then
      Check2.Visible = True
   Else
      Check2.Visible = False
   End If
   'end 2020/9/18
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strUserNum = strUser1Num
   If cboPrinter.Text <> cboPrinter.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cboPrinter.Name, "0", "0", Me.cboPrinter.Text
   End If
   Set frm1105 = Nothing
End Sub

Private Sub grdDataList_SelChange()
Dim i As Integer
Dim j As Integer

   grdDataList.Visible = False
   grdDataList.row = grdDataList.MouseRow
   grdDataList.col = 0
   If Me.grdDataList.TextMatrix(Me.grdDataList.row, 8) = "" Then
      Me.grdDataList.Visible = True
      Exit Sub
   End If
   If grdDataList.row <> 0 Then
      If grdDataList.Text = "V" Then
         grdDataList.Text = ""
         For i = 0 To grdDataList.Cols - 1
                grdDataList.col = i
                grdDataList.CellBackColor = QBColor(15)
         Next i
      Else
         grdDataList.Text = "V"
         For i = 0 To grdDataList.Cols - 1
             grdDataList.col = i
             grdDataList.CellBackColor = &HFFC0C0
         Next i
      End If
'      For j = 1 To Me.grdDataList.Rows - 1
'         If j <> Me.grdDataList.MouseRow Then
'            Me.grdDataList.Row = j
'            Me.grdDataList.Col = 0
'            Me.grdDataList.Text = ""
'            For i = 0 To grdDataList.Cols - 1
'                   grdDataList.Col = i
'                   grdDataList.CellBackColor = QBColor(15)
'            Next i
'         End If
'      Next j
   End If
   grdDataList.Visible = True
End Sub

'Add By Cheng 2002/10/25
Private Sub SelectFirstRow()
Dim i As Integer
Dim j As Integer

   grdDataList.Visible = False
   grdDataList.row = 1
   grdDataList.col = 0
   If Me.grdDataList.TextMatrix(Me.grdDataList.row, 8) = "" Then
      Me.grdDataList.Visible = True
      Exit Sub
   End If
   If grdDataList.row <> 0 Then
        grdDataList.Text = "V"
        For i = 0 To grdDataList.Cols - 1
            grdDataList.col = i
            grdDataList.CellBackColor = &HFFC0C0
        Next i
   End If
   grdDataList.Visible = True
End Sub


Private Sub SetDataListWidth()
   Dim ii As Integer
   With Me.grdDataList
      .ClearStructure
      .Rows = 2
      .FixedRows = 1
      .row = 0
      .col = 0: .Text = "V"
      .ColWidth(0) = 200
      .CellAlignment = flexAlignLeftCenter
      .col = 1: .Text = "總收文號"
      .ColWidth(1) = 1200
      .CellAlignment = flexAlignLeftCenter
      .col = 2: .Text = "本所案號"
      .ColWidth(2) = 1600
      .CellAlignment = flexAlignLeftCenter
      .col = 3: .Text = "案件性質"
      .ColWidth(3) = 1000
      .CellAlignment = flexAlignLeftCenter
      .col = 4: .Text = "定稿別"
      .ColWidth(4) = 1200
      .CellAlignment = flexAlignLeftCenter
      .col = 5: .Text = "份數"
      .ColWidth(5) = 400
      .CellAlignment = flexAlignLeftCenter
      .col = 6: .Text = "定稿說明"
      .ColWidth(6) = 2000
      .CellAlignment = flexAlignLeftCenter
      .col = 7: .Text = "處理狀況"
      .ColWidth(7) = 800
      .CellAlignment = flexAlignLeftCenter
      .col = 8: .Text = "使用者代號"
      .ColWidth(8) = 0
      .CellAlignment = flexAlignLeftCenter
      .col = 9: .Text = "日期"
      .ColWidth(9) = 0
      .CellAlignment = flexAlignLeftCenter
      .col = 10: .Text = "時間"
      .ColWidth(10) = 0
      .CellAlignment = flexAlignLeftCenter
      'Modified by Morgan 2022/5/18 後面欄位非電腦中心不顯示
      If Pub_StrUserSt03 <> "M51" Then
         For ii = 11 To .Cols - 1
            .col = ii
            .ColWidth(ii) = 0
         Next
      End If
      'end 2022/5/18
   End With
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
   Case 3
      'Modify By Sindy 2011/2/11
      'QueryData False
      'Remove by Lydia 2017/12/19 改成Option選項,由Enter或查詢啟動
      'Call cmdMove_Click(4)
      'Mark by Lydia 2017/12/19 保留
'      If OptKind(0).Value = True Then Call cmdMove_Click(4)
'   'Added by Lydia 2017/12/19
'   Case 4 '申請案號
'      If OptKind(1).Value = True Then Call cmdMove_Click(4)
'   Case 5 '審定號數/證書號數
'      If OptKind(2).Value = True Then Call cmdMove_Click(4)
'   'end 2017/12/19
   End Select
End Sub

Private Function QueryData(Optional bQuery As Boolean = True) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim adoRst As ADODB.Recordset, strCP10 As String 'Add By Sindy 2018/11/26
Dim i As Integer
'Added by Lydia 2017/12/19
Dim strCon(1 To 5) As String
Dim strNoList As String
'end 2017/12/19
Dim ii As Integer
   
   lblSendMailDt.Visible = False 'Add By Sindy 2018/5/14
   Screen.MousePointer = vbHourglass
   QueryData = False
   If OptKind(0).Value = True Then 'Added by Lydia 2017/12/19 改成Option選項
        If Me.Text1(0).Text = "" Then
           MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
           Me.Text1(0).SetFocus
           Text1_GotFocus 0
           Screen.MousePointer = vbDefault
           Exit Function
        End If
        If Me.Text1(1).Text = "" Then
           MsgBox "請輸入本所案號!!!", vbExclamation + vbOKOnly
           Me.Text1(1).SetFocus
           Text1_GotFocus 1
           Screen.MousePointer = vbDefault
           Exit Function
        End If
   'Added by Lydia 2017/12/19
   Else
        Me.Text1(0).Text = ""
        Me.Text1(1).Text = ""
        Me.Text1(2).Text = ""
        Me.Text1(3).Text = ""
        If OptKind(1).Value = True Then
            If Me.Text1(4).Text = "" Then
               MsgBox "請輸入申請案號!!!", vbExclamation + vbOKOnly
               Me.Text1(4).SetFocus
               Text1_GotFocus 0
               Screen.MousePointer = vbDefault
               Exit Function
            End If
        ElseIf OptKind(2).Value = True Then
            If Me.Text1(5).Text = "" Then
               MsgBox "請輸入審定號數/證書號數!!!", vbExclamation + vbOKOnly
               Me.Text1(5).SetFocus
               Text1_GotFocus 0
               Screen.MousePointer = vbDefault
               Exit Function
            End If
        End If
   End If 'end 2017/12/19
   
   'Added by Lydia 2017/12/19 設查詢條件
   strNoList = ""
   If OptKind(0).Value = True Then '本所案號
        strCon(1) = ChgPatent(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
        strCon(2) = ChgTradeMark(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
        strCon(3) = ChgLawcase(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
        strCon(4) = ChgHirecase(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
        strCon(5) = ChgService(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
   ElseIf OptKind(1).Value = True Then  '申請案號
        strCon(1) = " PA11='" & Me.Text1(4).Text & "' "
        strCon(2) = " TM12='" & Me.Text1(4).Text & "' "
        strCon(3) = "":         strCon(4) = ""
        strCon(5) = " SP11='" & Me.Text1(4).Text & "' "
   ElseIf OptKind(2).Value = True Then  '審定號數/證書號數
        strCon(1) = " PA22='" & Me.Text1(5).Text & "' "
        strCon(2) = " TM15='" & Me.Text1(5).Text & "' "
        strCon(3) = "":         strCon(4) = ""
        strCon(5) = " (SP14='" & Me.Text1(5).Text & "' OR SP32='" & Me.Text1(5).Text & "') "
   End If
   'end 2017/12/19
   
   ClearContorls
    'Modify By Cheng 2003/01/10
'   strSQLA = "SELECT PA05,PA06,PA07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA09 FROM PATENT,CUSTOMER WHERE SUBSTR(PA26,1,8)=CU01 AND SUBSTR(PA26,9,1)=CU02 AND " & ChgPatent(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
'   strSQLA = strSQLA & " UNION SELECT TM05,TM06,TM07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM10 FROM TRADEMARK,CUSTOMER WHERE SUBSTR(TM23,1,8)=CU01 AND SUBSTR(TM23,9,1)=CU02 AND " & ChgTradeMark(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
'   strSQLA = strSQLA & " UNION SELECT LC05,LC06,LC07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),LC15 FROM LAWCASE,CUSTOMER WHERE SUBSTR(LC11,1,8)=CU01 AND SUBSTR(LC11,9,1)=CU02 AND " & ChgLawcase(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
'   strSQLA = strSQLA & " UNION SELECT HC06,'','',NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),'000' FROM HIRECASE,CUSTOMER WHERE SUBSTR(HC05,1,8)=CU01 AND SUBSTR(HC05,9,1)=CU02 AND " & ChgHirecase(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
'   strSQLA = strSQLA & " UNION SELECT SP05,SP06,SP07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP09 FROM SERVICEPRACTICE,CUSTOMER WHERE SUBSTR(SP08,1,8)=CU01 AND SUBSTR(SP08,9,1)=CU02 AND " & ChgService(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
   'Modified by Lydia 2017/12/19 改成Option選項
'   StrSQLa = "SELECT PA05,PA06,PA07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),PA09 FROM PATENT,CUSTOMER WHERE SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND " & ChgPatent(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
'   StrSQLa = StrSQLa & " UNION SELECT TM05,TM06,TM07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),TM10 FROM TRADEMARK,CUSTOMER WHERE SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND " & ChgTradeMark(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
'   StrSQLa = StrSQLa & " UNION SELECT LC05,LC06,LC07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),LC15 FROM LAWCASE,CUSTOMER WHERE SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND " & ChgLawcase(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
'   StrSQLa = StrSQLa & " UNION SELECT HC06,'','',NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),'000' FROM HIRECASE,CUSTOMER WHERE SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) AND " & ChgHirecase(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
'   StrSQLa = StrSQLa & " UNION SELECT SP05,SP06,SP07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),SP09 FROM SERVICEPRACTICE,CUSTOMER WHERE SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND " & ChgService(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text)
   StrSQLa = "SELECT PA05,PA06,PA07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CNAME,PA09,PA01,PA02,PA03,PA04 FROM PATENT,CUSTOMER WHERE SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND " & strCon(1)
   StrSQLa = StrSQLa & " UNION SELECT TM05,TM06,TM07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CNAME,TM10,TM01,TM02,TM03,TM04 FROM TRADEMARK,CUSTOMER WHERE SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND " & strCon(2)
   If OptKind(0).Value = True Then '法務和顧問案只能用本所案號
       StrSQLa = StrSQLa & " UNION SELECT LC05,LC06,LC07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CNAME,LC15,LC01,LC02,LC03,LC04 FROM LAWCASE,CUSTOMER WHERE SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND " & strCon(3)
       StrSQLa = StrSQLa & " UNION SELECT HC06,'','',NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CNAME,'000',HC01,HC02,HC03,HC04 FROM HIRECASE,CUSTOMER WHERE SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) AND " & strCon(4)
   End If
   StrSQLa = StrSQLa & " UNION SELECT SP05,SP06,SP07,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CNAME,SP09,SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE,CUSTOMER WHERE SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND " & strCon(5)
    'end 2017/12/19
     
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   
   If rsA.RecordCount > 0 Then
      'Added by Lydia 2017/12/20 複數筆彈訊息
      If rsA.RecordCount > 1 Then
          MsgBox "共" & rsA.RecordCount & "筆案件，請改用本所案號查詢!!!", vbExclamation + vbOKOnly
          rsA.MoveFirst
          Do While Not rsA.EOF
               strNoList = strNoList & rsA.Fields(5).Value & rsA.Fields(6).Value & rsA.Fields(7).Value & rsA.Fields(8).Value & ","
                rsA.MoveNext
          Loop
      Else
          strNoList = "" & rsA.Fields(5).Value & rsA.Fields(6).Value & rsA.Fields(7).Value & rsA.Fields(8).Value
          Me.Text1(0).Text = "" & rsA.Fields(5).Value
          Me.Text1(1).Text = "" & rsA.Fields(6).Value
          Me.Text1(2).Text = "" & rsA.Fields(7).Value
          Me.Text1(3).Text = "" & rsA.Fields(8).Value
      'end 2017/12/20
          Me.lblFM2(0).Caption = "" & rsA.Fields(0).Value
          Me.lblFM2(1).Caption = "" & rsA.Fields(1).Value
          Me.lblFM2(2).Caption = "" & rsA.Fields(2).Value
          Me.lblFM2(3).Caption = "" & rsA.Fields(3).Value
          m_strNationNo = "" & rsA.Fields(4).Value
      End If 'end 2017/12/20
     
      'Add by Morgan 2009/9/15
      If bQuery = False Then
         QueryData = True
         Screen.MousePointer = vbDefault
         Exit Function
      End If
      
   Else
      If OptKind(0).Value = True Then 'Added by Lydia 2017/12/19
          MsgBox "無此本所案號資料!!!", vbExclamation + vbOKOnly
          Me.Text1(0).SetFocus
      'Added by Lydia 2017/12/19
      ElseIf OptKind(1).Value = True Then
          MsgBox "無此申請案號資料!!!", vbExclamation + vbOKOnly
          Me.Text1(4).SetFocus
      ElseIf OptKind(2).Value = True Then
          MsgBox "無此審定號數/證書號數資料!!!", vbExclamation + vbOKOnly
          Me.Text1(5).SetFocus
      End If
      'end 2017/12/19
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      Screen.MousePointer = vbDefault
      Exit Function
   End If
   
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   'Added by Lydia 2017/12/20 定稿的案件條件
   If InStr(strNoList, ",") = 0 Then
         strExc(0) = " AND LD05='" & Me.Text1(0).Text & "' AND LD06='" & Left(Me.Text1(1).Text & "000000", 6) & "' AND LD07='" & Left(Me.Text1(2).Text & "0", 1) & "' AND LD08='" & Left(Me.Text1(3).Text & "00", 2) & "'"
   Else
         strExc(0) = " AND LD05||LD06||LD07||LD08 IN (" & GetAddStr(strNoList) & ") "
   End If
   'end 2017/12/20
   
    'Modify By Cheng 2003/01/21
    '加列印方向
'   strSQLA = "SELECT ' ' AS V,DECODE(LD04,'',LD05||LD06||LD07||LD08||'&'||LD09,LD04) AS 總收文號,LD05||'-'||LD06||'-'||LD07||'-'||LD08 AS 本所案號,DECODE('" & m_strNationNo & "','020',CPM04,CPM03) AS 案件性質,TYP03 AS 定稿別,NVL(LD13,0) 份數,FTM06 AS 定稿說明,LD11 AS 處理狀況,LD01 AS 使用者代號, LD02 AS 日期,LD03 AS 時間, LD10 " & _
'            " FROM LETTERDEMAND,CASEPROGRESS,CASEPROPERTYMAP,TEXTTYPE,FINALTEXTMAP " & _
'            " WHERE LD04=CP09(+) AND LD05=CPM01(+) AND LD09=CPM02(+) AND LD10=TYP01(+) AND LD05=TYP02(+) AND LD05=FTM01(+) AND LD10=FTM02(+) AND LD09=FTM03(+) AND LD11=FTM04(+) " & _
'            " AND LD05='" & Me.Text1(0).Text & "' AND LD06='" & Left(Me.Text1(1).Text & "000000", 6) & "' AND LD07='" & Left(Me.Text1(2).Text & "0", 1) & "' AND LD08='" & Left(Me.Text1(3).Text & "00", 2) & "' AND LD01='" & strUserNum & "' "
   'Modify by Morgan 2010/3/31 定稿說明要判斷沒有相同案件性質的抓 000 定稿
   'Modified by Morgan 2014/5/2 +LD18
   'Modified by Morgan 2014/5/7 +LD05,LD06,LD07,LP08,CP10;CP改用LD18串(原來串了沒用)
   'Modified by Morgan 2014/8/13 +LP01,AF01
   'Modified by Sindy 2016/8/4 +LD09
   'Modified by Lydia 2017/12/20 可能有複數筆案件
   'StrSQLa = "SELECT ' ' AS V,DECODE(LD04,'',LD05||LD06||LD07||LD08||'&'||LD09,LD04) AS 總收文號,LD05||'-'||LD06||'-'||LD07||'-'||LD08 AS 本所案號,DECODE('" & m_strNationNo & "','000',CPM03,CPM04)||decode(cp10,'990','-副本') AS 案件性質,TYP03 AS 定稿別,NVL(LD13,0) 份數,FTM06 AS 定稿說明,LD11 AS 處理狀況,LD01 AS 使用者代號, LD02 AS 日期,LD03 AS 時間, LD10, LD12, LD18,LD05,LD06,LD07,LD08,CP10,LP01,AF01,LD09,LD01" & _
            " FROM LETTERDEMAND,CASEPROGRESS,CASEPROPERTYMAP,TEXTTYPE,FINALTEXTMAP A,LetterProgress,AppForm" & _
            " WHERE LD18=CP09(+) AND LD05=CPM01(+) AND LD09=CPM02(+) AND LD10=TYP01(+) AND LD05=TYP02(+) AND LD05=FTM01(+) AND LD10=FTM02(+) AND LD11=FTM04(+) and lp01(+)=LD18 and af01(+)=LD18" & _
            " AND LD05='" & Me.Text1(0).Text & "' AND LD06='" & Left(Me.Text1(1).Text & "000000", 6) & "' AND LD07='" & Left(Me.Text1(2).Text & "0", 1) & "' AND LD08='" & Left(Me.Text1(3).Text & "00", 2) & "'" & _
            " AND FTM03=(SELECT NVL(MAX(B.FTM03),'000') FROM FINALTEXTMAP B WHERE B.FTM01=LD05 AND B.FTM02=LD10" & _
            " AND B.FTM03=LD09 AND B.FTM04=LD11)"
   'Modified by Morgan 2018/7/3 +LP40
   'Modify by Sindy 2020/2/6 +LD16
   '                 0        1                                                                  2                                                 3                                                                                       4               5                6                 7                8                   9           10            11    12    13   14   15   16   17   18   19   20   21   22   23   24   25
   StrSQLa = "SELECT ' ' AS V,DECODE(LD04,'',LD05||LD06||LD07||LD08||'&'||LD09,LD04) AS 總收文號,LD05||'-'||LD06||'-'||LD07||'-'||LD08 AS 本所案號,DECODE('" & m_strNationNo & "','000',CPM03,CPM04)||decode(cp10,'990','-副本') AS 案件性質,TYP03 AS 定稿別,NVL(LD13,0) 份數,FTM06 AS 定稿說明,LD11 AS 處理狀況,LD01 AS 使用者代號,LD02 AS 日期,LD03 AS 時間, LD10, LD12, LD18,LD05,LD06,LD07,LD08,CP10,LP01,AF01,LD09,LD01,LP41,LD27,LD16" & _
            " FROM LETTERDEMAND,CASEPROGRESS,CASEPROPERTYMAP,TEXTTYPE,FINALTEXTMAP A,LetterProgress,AppForm" & _
            " WHERE LD18=CP09(+) AND LD05=CPM01(+) AND LD09=CPM02(+) AND LD10=TYP01(+) AND LD05=TYP02(+) AND LD05=FTM01(+) AND LD10=FTM02(+) AND LD11=FTM04(+) and lp01(+)=LD18 and af01(+)=LD18" & _
           strExc(0) & " AND FTM03=(SELECT NVL(MAX(B.FTM03),'000') FROM FINALTEXTMAP B WHERE B.FTM01=LD05 AND B.FTM02=LD10" & _
            " AND B.FTM03=LD09 AND B.FTM04=LD11)"
            
   'Modify by Morgan 2008/1/10 電腦中心人員執行時抓所有人的定稿
   'StrSQLa = StrSQLa & "' AND LD01='" & strUserNum & "' "
   If Pub_StrUserSt03 <> "M51" Then
      'Add by Morgan 2009/11/10 外商程序可看同部門定稿
      'Modified by Morgan 2017/2/10 改程序都可看同部門定稿但非本人定稿需彈訊息提醒--Robert
      'If Pub_StrUserSt03 = "F12" Then
      'Modified by Morgan 2017/2/15 先剔除F22,因為同一案號會有不同人員產生不同定稿(如公告,證書)
      'Modified by Morgan 2017/2/16 改用勾選判斷
      'If Pub_StrUserSt03 = "F12" Or Pub_StrUserSt03 = "P12" Or Pub_StrUserSt03 = "P22" Then
      If Check1.Value = vbChecked Then
         'modify by sonia 2018/9/13 外商程序可看外商承辦定稿(解除期限定稿)
         'StrSQLa = StrSQLa & " AND EXISTS(SELECT * FROM STAFF WHERE ST01=LD01 AND ST03='" & Pub_StrUserSt03 & "')"
         If Pub_StrUserSt03 = "F12" Then
            StrSQLa = StrSQLa & " AND EXISTS(SELECT * FROM STAFF WHERE ST01=LD01 AND ST03 LIKE 'F1%')"
         Else
            StrSQLa = StrSQLa & " AND EXISTS(SELECT * FROM STAFF WHERE ST01=LD01 AND ST03='" & Pub_StrUserSt03 & "')"
         End If
         'end 2018/9/13
      'Add by Morgan 2009/11/19 外專承辦可看FMP定稿
      ElseIf Pub_StrUserSt03 = "F23" Then
         StrSQLa = StrSQLa & " AND LD01='" & strFMPNum & "' "
         
'Removed by Morgan 2017/2/10 已開放程序同部門皆可看
'      'Add by Morgan 2010/6/15 79075 可以看該部門的定稿
'      'modify by sonia 2013/11/1 郭說開放余彥葶A2023也可以看全部
'      ElseIf strUserNum = "79075" Or strUserNum = "A2023" Then
'         StrSQLa = StrSQLa & " AND EXISTS(SELECT * FROM STAFF WHERE ST01=LD01 AND ST03='" & Pub_StrUserSt03 & "')"
'end 2017/2/10

      Else
         StrSQLa = StrSQLa & " AND LD01='" & strUserNum & "' "
      End If
   End If
   
'   'Add By Sindy 2014/9/18
'   cmdMove(7).Visible = False
'   '2014/9/18 END
   
   StrSQLa = StrSQLa & " order by ld01,ld02,ld03"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      'Added by Morgan 2017/2/10
      'Removed by Morgan 2017/2/16 改用勾選不必再提醒
      'If rsA.Fields("ld01") <> strUserNum And Not (Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "F23") Then
      '   If MsgBox("你目前維護的不是自己的定稿，確定要繼續嗎？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
      '      Screen.MousePointer = vbDefault
      '      Exit Function
      '   End If
      'End If
      'end 2017/2/16
      'end 2017/2/10
      
      Set Me.grdDataList.Recordset = rsA
      
      'Added by Sindy 2019/11/29 有信函收文號或指示信收文號,就不可在此列印定稿
      '應使用[修改]列印或到卷宗區調檔列印
      cmdMove(2).Enabled = True '列印
      For ii = 1 To Me.grdDataList.Rows - 1
         If grdDataList.TextMatrix(ii, 19) <> "" Or grdDataList.TextMatrix(ii, 20) <> "" Then
            cmdMove(2).Enabled = False
            Exit For
         End If
      Next ii
      '2019/11/29 END
      
   Else
      Me.Text1(0).SetFocus
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      Screen.MousePointer = vbDefault
      'Add By Sindy 2012/8/7
      If m_EditMode = 4 Then '查詢時
         'Modified by Lydia 2017/12/20 輸入本所案號
         'If Me.Text1(0).Text = "FCT" Then
         If OptKind(0).Value = True And Me.Text1(0).Text = "FCT" Then
            m_TM01 = Text1(0)
            m_TM02 = Text1(1)
            m_TM03 = Text1(2)
            m_TM04 = Text1(3)
            'Modify By Sindy 2018/11/26 + 註冊後一文多案的日文變更定稿
            strSql = "select cp09,cp10,tm15 from caseprogress,trademark" & _
                     " where cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "'" & _
                     " and cp03='" & Right("0" & Text1(2), 1) & "' and cp04='" & Right("00" & Text1(3), 2) & "'" & _
                     " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
                     " and cp148='Y' and cp10 in('501','301')" & _
                     " order by cp27 desc"
            intI = 1
            Set adoRst = ClsLawReadRstMsg(intI, strSql)
            strCP10 = "": strSendDate = ""
            If intI > 0 Then
               If GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3" Then '日文
                  adoRst.MoveFirst
                  Do While Not adoRst.EOF
                     If adoRst.Fields("cp10") = "301" And "" & adoRst.Fields("tm15") <> "" Then
                        strSendDate = InputBox("無任何定稿資料！" & vbCrLf & "是否要產生一文多案的（日文註冊變更）定稿？" & vbCrLf & "若否，請按〔取消〕鍵〔離開〕。" & vbCrLf & "若是，請一定要輸入發文日：")
                     ElseIf adoRst.Fields("cp10") = "501" Then
                        strSendDate = InputBox("無任何定稿資料！" & vbCrLf & "是否要產生一文多案的（日文移轉）定稿？" & vbCrLf & "若否，請按〔取消〕鍵〔離開〕。" & vbCrLf & "若是，請一定要輸入發文日：")
                     End If
                     If strSendDate <> "" Then
                        strCP10 = adoRst.Fields("cp10")
                        Exit Do
                     End If
                     adoRst.MoveNext
                  Loop
               ElseIf GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "2" Then '英文
                  adoRst.MoveFirst
                  Do While Not adoRst.EOF
                     If adoRst.Fields("cp10") = "501" Then
                        strSendDate = InputBox("無任何定稿資料！" & vbCrLf & "是否要產生一文多案的（英文移轉）定稿？" & vbCrLf & "若否，請按〔取消〕鍵〔離開〕。" & vbCrLf & "若是，請一定要輸入發文日：")
                     End If
                     If strSendDate <> "" Then
                        strCP10 = adoRst.Fields("cp10")
                        Exit Do
                     End If
                     adoRst.MoveNext
                  Loop
               Else
                  strSendDate = ""
               End If
            '2018/11/26 END
               If strSendDate <> "" Then
                  If CheckIsTaiwanDate(strSendDate) = True Then
                     Screen.MousePointer = vbHourglass
                     m_CP10 = strCP10
                     Call ReadCPData
                     Screen.MousePointer = vbDefault
                  End If
               Else
                  MsgBox "無任何定稿資料!!!", vbExclamation + vbOKOnly
               End If
            Else
               MsgBox "無任何定稿資料!!!", vbExclamation + vbOKOnly
            End If
         Else
            MsgBox "無任何定稿資料!!!", vbExclamation + vbOKOnly
         End If
      Else
         MsgBox "無任何定稿資料!!!", vbExclamation + vbOKOnly
      End If
      '2012/8/7 End
      Exit Function
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   If Me.grdDataList.Rows = 2 And Me.grdDataList.TextMatrix(1, 0) <> "" Then
      SelectFirstRow
   End If
   
'   'Add By Sindy 2014/9/18
'   If Text1(0) = "FCP" Then
'      For i = 1 To grdDataList.Rows - 1
'         If grdDataList.TextMatrix(i, 11) = "04" Or _
'            grdDataList.TextMatrix(i, 11) = "08" Or _
'            grdDataList.TextMatrix(i, 11) = "14" Then
'            cmdMove(7).Visible = True
'            Exit For
'         End If
'      Next i
'   End If
'   '2014/9/18 END
   
   QueryData = True
   Screen.MousePointer = vbDefault
End Function

Private Sub ClearContorls()
   Me.lblFM2(0).Caption = Empty
   Me.lblFM2(1).Caption = Empty
   Me.lblFM2(2).Caption = Empty
   Me.lblFM2(3).Caption = Empty
   SetDataListWidth
    'Added by Lydia 2017/12/19 清除其他項目
    If OptKind(0).Value = True Then
       Me.Text1(4).Text = ""
       Me.Text1(5).Text = ""
    Else
       Me.Text1(0).Text = Empty
       Me.Text1(1).Text = Empty
       Me.Text1(2).Text = Empty
       Me.Text1(3).Text = Empty
       If OptKind(1).Value = True Then
           Me.Text1(5).Text = ""
       ElseIf OptKind(2).Value = True Then
           Me.Text1(4).Text = ""
       End If
    End If
    'end 2017/12/19
End Sub

'Add by Morgan 2009/9/15
'報價轉定稿
Private Sub Transfer()
   Dim adoRst As ADODB.Recordset
   Dim stSQL As String, intR As Integer
   Dim bQuery As Boolean
   Dim stDate As String
   
   stDate = CompWorkDay(報價確認天數, strSrvDate(1), 1)
   
ReReadData: 'Add By Sindy 2013/4/24
   If QueryData(False) = True Then
      stSQL = "select lc01,lc02,lc07,lc10 from caseprogress,lettercache" & _
         " where cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "'" & _
         " and cp03='" & Right("0" & Text1(2), 1) & "' and cp04='" & Right("00" & Text1(3), 2) & "'" & _
         " and lc01(+)=cp09 and lc13 is null and lc01 is not null"
      If Pub_StrUserSt03 <> "M51" Then
         'Modified by Morgan 2020/6/3 +同部門可轉--桂英
         If Check1.Value = vbChecked Then
            stSQL = stSQL & " and exists(select * from staff where st01=lc10 and st03='" & Pub_StrUserSt03 & "')"
         Else
            stSQL = stSQL & " and lc10='" & strUserNum & "'"
         End If
         'end 2020/6/3
      End If
      
      intR = 1
      Set adoRst = ClsLawReadRstMsg(intR, stSQL)
      If intR = 0 Then
         'Modify By Sindy 2013/4/24 若輸入的本所案號後3碼不是000且無資料時,則再使用一次母案讀取資料
         If Text1(2) & Text1(3) <> "000" Then
            Text1(2) = "0"
            Text1(3) = "00"
            GoTo ReReadData
         End If
         '2013/4/24 End
         MsgBox "該案號無未列印報價定稿！"
      Else
         'Modify by Morgan 2009/9/18 FF案件不必智權人員確認
         'Modify by Morgan 2011/8/26
         'stSQL = stSQL & " and (lc07 is not null or (cp01='CFP' and cp12='F23'))"
          stSQL = stSQL & " and (lc07 is not null or (cp01='CFP' and cp12='F23') or NVL(LC06,lc11)<" & stDate & ")"
         intR = 1
         Set adoRst = ClsLawReadRstMsg(intR, stSQL)
         If intR = 0 Then
            MsgBox "智權人員尚未確認報價，不可轉定稿！"
         Else
            With adoRst
            If MsgBox("本案將有 " & .RecordCount & " 筆報價轉定稿，是否要繼續？ ", vbYesNo + vbDefaultButton2) = vbYes Then
               bQuery = False
               Do While Not .EOF
                  strUserNum = "" & .Fields("lc10")
                  If PUB_Cache2Letter(.Fields("lc01"), .Fields("lc02"), False, , , True) = True Then
                     bQuery = True
                  Else
                     MsgBox "作業失敗！"
                  End If
                  .MoveNext
               Loop
               strUserNum = strUser1Num
               If bQuery = True Then QueryData
            End If
            End With
         End If
      End If
   End If
   Set adoRst = Nothing
End Sub

'Add By Sindy 2012/8/7
Private Sub ReadCPData()
Dim adoRst As ADODB.Recordset
   
   strSql = "SELECT * FROM CaseProgress WHERE CP01='" & Text1(0) & "' AND CP02='" & Text1(1) & "' " & _
            "AND CP03='" & Text1(2) & "' AND CP04='" & Text1(3) & "' " & _
            "AND CP27=" & DBDATE(strSendDate) & " AND CP10='" & m_CP10 & "' AND CP148='Y'" & _
            " order by cp27 desc"
   intI = 1
   'Modify By Sindy 2015/5/20 ex.FCT-037437
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      adoRst.MoveFirst
      Do While Not adoRst.EOF
   '2015/5/20 END
         m_TM01 = "" & adoRst("CP01")
         m_TM02 = "" & adoRst("CP02")
         m_TM03 = "" & adoRst("CP03")
         m_TM04 = "" & adoRst("CP04")
         m_CP09 = "" & adoRst("CP09")
         m_CP10 = "" & adoRst("CP10")
         m_CP28 = "" & adoRst("CP28") 'Add By Sindy 2012/11/08
         '一文多案-母案
         strSql = "SELECT c1.* " & _
                  "FROM CaseProgress c1,(select * from caseprogress where cp09='" & m_CP09 & "') c2 " & _
                  "WHERE c1.CP01='FCT' AND c1.CP27=" & DBDATE(strSendDate) & " AND c1.CP10='" & m_CP10 & "' " & _
                  "AND c1.CP55||c1.CP93||c1.CP94||c1.CP95||c1.CP96||c1.CP56||c1.CP89||c1.CP90||c1.CP91||c1.CP92=c2.CP55||c2.CP93||c2.CP94||c2.CP95||c2.CP96||c2.CP56||c2.CP89||c2.CP90||c2.CP91||c2.CP92 " & _
                  "AND c1.CP123='Y' AND c1.CP148='Y'"
         RsTemp.Close
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            'Modify By Sindy 2013/5/20
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               If m_TM01 = "" & RsTemp("CP01") And m_TM02 = "" & RsTemp("CP02") And _
                  m_TM03 = "" & RsTemp("CP03") And m_TM04 = "" & RsTemp("CP04") Then
                  GoTo ReadEnd
               End If
               RsTemp.MoveNext
            Loop
            '比對不到母案則讀取第一筆
            RsTemp.MoveFirst
            '2013/5/20 End
            m_TM01 = "" & RsTemp("CP01")
            m_TM02 = "" & RsTemp("CP02")
            m_TM03 = "" & RsTemp("CP03")
            m_TM04 = "" & RsTemp("CP04")
            m_CP09 = "" & RsTemp("CP09")
            m_CP10 = "" & RsTemp("CP10")
            m_CP28 = "" & RsTemp("CP28") 'Add By Sindy 2012/11/08
         End If
ReadEnd:    'Add By Sindy 2013/5/20
   '      '一文多案
   '      'Modify By Sindy 2012/10/17
   '      strAppendix = PUB_GetFCTAppendix(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, DBDATE(strSendDate), "01", m_CP28, m_CP09, "05")
   '      strAppendix = Empty
   '      intRow = 0
   '      strSql = "SELECT tm15,tm09,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm05,c1.cp123,tm45 " & _
   '               "FROM CaseProgress c1,trademark,(select * from caseprogress where cp09='" & m_CP09 & "') c2 " & _
   '               "WHERE c1.CP01='FCT' AND c1.CP27=" & DBDATE(strSendDate) & " AND c1.CP10='" & m_CP10 & "' " & _
   '               "AND c1.CP01=TM01(+) AND c1.CP02=TM02(+) AND c1.CP03=TM03(+) AND c1.CP04=TM04(+) " & _
   '               "AND c1.CP55||c1.CP93||c1.CP94||c1.CP95||c1.CP96||c1.CP56||c1.CP89||c1.CP90||c1.CP91||c1.CP92=c2.CP55||c2.CP93||c2.CP94||c2.CP95||c2.CP96||c2.CP56||c2.CP89||c2.CP90||c2.CP91||c2.CP92 " & _
   '               "ORDER BY c1.CP05 asc"
   '      RsTemp.Close
   '      intI = 1
   '      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   '      If intI = 1 Then
   '         RsTemp.MoveFirst
   '         Do While Not RsTemp.EOF
   '            intRow = intRow + 1
   '            If Len(Trim("" & RsTemp.Fields(1))) >= 11 Then
   '               strAppendix = strAppendix & "(" & intRow & ") Reg. No.：" & Trim("" & RsTemp.Fields(0)) & " "
   '               strAppendix = strAppendix & "Class：" & Trim("" & RsTemp.Fields(1)) & " "
   '            Else
   '               strAppendix = strAppendix & "(" & intRow & ") Reg. No.：" & Left(("" & RsTemp.Fields(0)) & "            ", 12)
   '               strAppendix = strAppendix & "Class：" & Left(("" & RsTemp.Fields(1)) & "           ", 11)
   '            End If
   '            strAppendix = strAppendix & "   Trademark：" & RsTemp.Fields(3) & vbCrLf
   '            strAppendix = strAppendix & "   Our Ref：" & IIf(Right(RsTemp.Fields(2), 5) = "-0-00", Left(RsTemp.Fields(2), Len(RsTemp.Fields(2)) - 5), RsTemp.Fields(2)) & "     Your Ref：" & RsTemp.Fields("tm45") & vbCrLf & vbCrLf
   '            RsTemp.MoveNext
   '         Loop
   '      End If
'         If GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "2" Then
            Call PrintLetter
'         Else
'            MsgBox "此案定稿語文不是2.英文，無定稿！"
'         End If
         adoRst.MoveNext
      Loop
   Else
      MsgBox "此案" & ChangeWStringToTDateString(DBDATE(strSendDate)) & "無" & IIf(m_CP10 = "501", "移轉", "變更") & "資料！"
      Exit Sub
   End If
End Sub

'Add By Sindy 2012/8/7
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   Dim ET01 As String, ET03 As String, ET03_1 As String, stContent As String
   
   Select Case Me.Text1(0).Text
      Case "FCT"
         If m_CP10 = "501" Then '移轉
            'Add By Sindy 2018/11/26
            If GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3" Then '日文
               ET01 = "01"
               ET03 = "06"
            '2018/11/26 END
            Else
               ET01 = "01"
               ET03 = "05"
            End If
         'Add By Sindy 2018/11/26
         ElseIf m_CP10 = "301" Then '註冊變更
            If GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3" Then '日文
               ET01 = "01"
               ET03 = "05" '註冊後變更申請書譯文(一文多案)
            End If
         '2018/11/26 END
         End If
   End Select
   
   'Add By Sindy 2012/11/26
   bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper) '檢查是否以E-Mail通知
   '2012/11/26 End
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   Call InsExpField(ET01, ET03)
   
   If ET03 <> "" Then
      'Add by Morgan 2008/6/12
'      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      'Modify by Amy 2018/07/27 +Me.Name
      If bolEmail Then
         'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'end 2009/10/20
         If ET03_1 <> "" Then
            NowPrint m_CP09, "01", ET03, False, strUserNum, , , , , iCopy, , , , , , , , , , , , Me.Name
            NowPrint m_CP09, "01", ET03_1, False, strUserNum, , , , , iCopy, , , , , , , , , , , , Me.Name
            NowPrint m_CP09, "01", ET03, False, strUserNum, , , True, stContent, , , , True, , , , , , , , , Me.Name
            NowPrint m_CP09, "01", ET03_1, False, strUserNum, , stContent, , , , , True, True, , , , , , , , , Me.Name
         Else
            NowPrint m_CP09, "01", ET03, False, strUserNum, , , , , iCopy, , True, True, , , , , , , , , Me.Name
         End If
         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
      Else
      'end 2008/6/12
         NowPrint m_CP09, "01", ET03, False, strUserNum, , , , , , , , , , , , , , , , , Me.Name
         If ET03_1 <> "" Then
            NowPrint m_CP09, "01", ET03_1, False, strUserNum, , , , , , , , , , , , , , , , , Me.Name
         End If
      End If
      'end 2018/07/27
      
      m_EditMode = 1
      Me.Text1(0).Text = m_TM01
      Me.Text1(1).Text = m_TM02
      Me.Text1(2).Text = m_TM03
      Me.Text1(3).Text = m_TM04
      Call QueryData
   End If
End Sub

'Add By Sindy 2012/8/7
'列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField(ET01 As String, ET03 As String)
Dim nCount As Integer
Dim nIndex As Integer
Dim strTemp As String
Dim strTemp1 As String
Dim strTextAdd As String
Dim strDebitNote As String 'Add By Sindy 2017/4/13
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   'Modify By Sindy 2017/4/13【FCT 01 501  00 英->函知已提移轉】
   m_MySt(1) = m_TM01: m_MySt(2) = m_TM02: m_MySt(3) = m_TM03: m_MySt(4) = m_TM04: m_Rule = m_CP09
   strDebitNote = ExceptFieldData2("FCT特殊請款文字對照")
   '2017/4/13 END
   
   Select Case Me.Text1(0).Text
      Case "FCT"
         If m_CP10 = "501" Then '移轉
            Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
               '英文
               Case "2":
                  '清除定稿例外欄位檔原有資料
                  EndLetter ET01, m_CP09, ET03, strUserNum
                  '是否補件
                  strTemp = Empty
                  strTextAdd = InputBox("是否補件(可複選)：(1:受讓人委任狀 2:移轉契約書 3:受讓人法人證明 4:註冊證)")
                  Screen.MousePointer = vbHourglass
                  nCount = GetSubStringCount(strTextAdd)
                  For nIndex = 1 To nCount
                     strTemp1 = GetSubString(strTextAdd, nIndex)
                     Select Case strTemp1
                        Case "1": '受讓人委任狀
                           If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                           strTemp = strTemp & "    * Power of Attorney of the Assignee."
                        Case "2": '移轉契約書
                           If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                           strTemp = strTemp & "    * Deed of Assignment respectively signed by the Assignee and Assignor."
                        Case "3": '受讓人法人證明
                           If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                           strTemp = strTemp & "    * A notarized Certificate of Corporation of the Assignee."
                        Case "4": '證冊證
                           If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                           strTemp = strTemp & "    * Original Certificates of Registration."
                     End Select
                  Next nIndex
                  If strTemp <> Empty Then: strTemp = vbCrLf & "    The remaining document(s) for the referenced assignment application follow : " & Chr(13) & Chr(10) & strTemp
                  If IsEmptyText(strTemp) = False Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
                              "','是否補件','" & strTemp & "')"
                     cnnConnection.Execute strSql
                  End If
                  'Modify By Sindy 2013/5/2 程式移到PUB_GetFCTAppendix
'                  '一案多件清單
'                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                           "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
'                           "','一案多件清單','" & ChgSQL(Trim(strAppendix)) & "')"
'                  cnnConnection.Execute strSql
                  'Modify By Sindy 2012/10/17 一文多案
                  Call PUB_GetFCTAppendix(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, DBDATE(strSendDate), "01", m_CP28, m_CP09, "05")
                  'Add By Sindy 2012/11/26 eMail Only定稿 : 以電子郵件通知，並且不寄紙本
                  If bolEmail = True And bolPlusPaper = False Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
                              "','例外內文','Enclosed herewith are scanned copies of the assignment application and the filing receipt for your reference. " & IIf(strDebitNote = "", "Our debit note is also enclosed herewith for your kind settlement.", strDebitNote) & "')"
                     cnnConnection.Execute strSql
                  Else '郵件
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
                              "','例外內文','" & IIf(strDebitNote = "", "Enclosed please find our debit note for services rendered for your kind settlement.", strDebitNote) & " Copies of the assignment application and the filing receipt will be mailed to you with the confirmation copy of this letter for your records.')"
                     cnnConnection.Execute strSql
                  End If
                  '2012/11/26 End
               'Add By Sindy 2018/11/26
               '日文
               Case "3":
                  '清除定稿例外欄位檔原有資料
                  EndLetter ET01, m_CP09, ET03, strUserNum
                  Call PUB_GetFCTAppendix_JP(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, DBDATE(strSendDate), ET01, m_CP28, m_CP09, ET03)
               '2018/11/26 END
            End Select
            
         'Add By Sindy 2018/11/26
         ElseIf m_CP10 = "301" Then '變更
            Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
               '日文
               Case "3":
                  '清除定稿例外欄位檔原有資料
                  EndLetter ET01, m_CP09, ET03, strUserNum
                  Call PUB_GetFCTAppendix_JP(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, DBDATE(strSendDate), ET01, m_CP28, m_CP09, ET03)
                  
                  '讀取變更檔
                  StrSQLa = "Select ce01,ce04,ce05,ce06,ce07,ce08,ce09" & _
                            ",ce23,ce24,ce25,ce26,ce27,ce28,ce29,ce30,ce31,ce32,ce33,ce34,ce35,ce36,ce37,ce38" & _
                            ",ce10,ce11,ce12,ce13,ce14,ce15,ce16" & _
                            ",ce68,ce69,ce70,ce71,ce72,ce73,ce74,ce75,ce76,ce77,ce78,ce79,ce80,ce81,ce82,ce83,ce84,ce85,ce86,ce87,ce88,ce89,ce90,ce91" & _
                            ",ce55,ce56 From changeevent Where ce01='" & m_CP09 & "'"
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.Fields(0).Value > 0 Then
                     If "" & rsA.Fields("ce04") <> "" Or _
                        "" & rsA.Fields("ce05") <> "" Or _
                        "" & rsA.Fields("ce06") <> "" Or _
                        "" & rsA.Fields("ce07") <> "" Or _
                        "" & rsA.Fields("ce08") <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
                                 "','變更申請人名稱','♀')"
                        cnnConnection.Execute strSql
                     End If
                     If "" & rsA.Fields("ce23") <> "" Or _
                        "" & rsA.Fields("ce24") <> "" Or _
                        "" & rsA.Fields("ce25") <> "" Or _
                        "" & rsA.Fields("ce26") <> "" Or _
                        "" & rsA.Fields("ce27") <> "" Or _
                        "" & rsA.Fields("ce28") <> "" Or _
                        "" & rsA.Fields("ce29") <> "" Or _
                        "" & rsA.Fields("ce30") <> "" Or _
                        "" & rsA.Fields("ce31") <> "" Or _
                        "" & rsA.Fields("ce32") <> "" Or _
                        "" & rsA.Fields("ce33") <> "" Or _
                        "" & rsA.Fields("ce34") <> "" Or _
                        "" & rsA.Fields("ce35") <> "" Or _
                        "" & rsA.Fields("ce36") <> "" Or _
                        "" & rsA.Fields("ce37") <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
                                 "','變更申請人住所','♀')"
                        cnnConnection.Execute strSql
                     End If
                     If "" & rsA.Fields("ce10") <> "" Or "" & rsA.Fields("ce11") <> "" Or "" & rsA.Fields("ce12") <> "" Or _
                        "" & rsA.Fields("ce13") <> "" Or "" & rsA.Fields("ce14") <> "" Or "" & rsA.Fields("ce15") <> "" Or _
                        "" & rsA.Fields("ce68") <> "" Or "" & rsA.Fields("ce69") <> "" Or "" & rsA.Fields("ce70") <> "" Or _
                        "" & rsA.Fields("ce71") <> "" Or "" & rsA.Fields("ce72") <> "" Or "" & rsA.Fields("ce73") <> "" Or _
                        "" & rsA.Fields("ce74") <> "" Or "" & rsA.Fields("ce75") <> "" Or "" & rsA.Fields("ce76") <> "" Or _
                        "" & rsA.Fields("ce77") <> "" Or "" & rsA.Fields("ce78") <> "" Or "" & rsA.Fields("ce79") <> "" Or _
                        "" & rsA.Fields("ce80") <> "" Or "" & rsA.Fields("ce81") <> "" Or "" & rsA.Fields("ce82") <> "" Or _
                        "" & rsA.Fields("ce83") <> "" Or "" & rsA.Fields("ce84") <> "" Or "" & rsA.Fields("ce85") <> "" Or _
                        "" & rsA.Fields("ce86") <> "" Or "" & rsA.Fields("ce87") <> "" Or "" & rsA.Fields("ce88") <> "" Or _
                        "" & rsA.Fields("ce89") <> "" Or "" & rsA.Fields("ce90") <> "" Or "" & rsA.Fields("ce91") <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
                                 "','變更代表人','♀')"
                        cnnConnection.Execute strSql
                     End If
                     If "" & rsA.Fields("ce55") <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & ET01 & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & _
                                 "','變更出名代理人','♀')"
                        cnnConnection.Execute strSql
                     End If
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  '2019/3/26 END
            End Select
         End If
   End Select
End Sub

'Added by Lydia 2017/06/28 更新FCT延展、移轉、變更(102,501,301)案之核准定稿日期
'Modified by Lydia 2017/09/18 +LD03定稿時間
Private Sub UpdateFCTld02et07(ByVal pLD01 As String, ByVal pLD02 As String, ByVal pLD04 As String, pLD05 As String, pLD10 As String, pLD11 As String, pLD03 As String)
Dim intA As Integer
'Added by Lydia 2017/09/18
Dim strCP10 As String
Dim rsAD As New ADODB.Recordset
Dim strTime As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String 'Added by Lydia 2019/04/01

On Error GoTo ErrHandle
    
Exit Sub 'Added by Lydia 2023/08/01 因為現在FCT所有定稿(催延展除外)在產生於定稿作業維護同時，會另將定稿儲存於FCT-workflow，所以程序人員都在FCT -workflow做修改或列印的動作, 不會每件都從定稿作業維護列印定稿了
    
    If pLD05 <> "FCT" Or pLD02 <> "99999999" Or pLD10 <> "03" Then Exit Sub
    
    'Added by Lydia 2017/09/18 判斷一申請書多件時,一併更新定日期
'    cnnConnection.BeginTrans
'        strSql = "update letterdemand set ld02=" & strSrvDate(1) & ",ld01='" & strUserNum & "' where ld04='" & pLD04 & "' and ld01='" & pLD01 & "' and ld02=" & pLD02 & " and ld10='" & pLD10 & "' and ld11='" & pLD11 & "' "
'        cnnConnection.Execute strSql, IntA
'        strSql = "update exceptcondition set et07=" & strSrvDate(1) & ",et04='" & strUserNum & "' where et02='" & pLD04 & "' and et04='" & pLD01 & "' and et07=" & pLD02 & " and et01='" & pLD10 & "' and et03='" & pLD11 & "' "
'        cnnConnection.Execute strSql, IntA
'    cnnConnection.CommitTrans
    'Modified by Lydia 2019/04/01 +cp01~cp04
    strSql = "select cp09,cp10,cp27,cp28,cp148,cp01,cp02,cp03,cp04 from caseprogress where cp09='" & pLD04 & "' "
    intA = 1
    Set rsAD = ClsLawReadRstMsg(intA, strSql)
    If intA = 1 Then
       strCP10 = rsAD.Fields("CP10")
       'Added by Lydia 2019/04/01
       strCP01 = rsAD.Fields("CP01")
       strCP02 = rsAD.Fields("CP02")
       strCP03 = rsAD.Fields("CP03")
       strCP04 = rsAD.Fields("CP04")
       'end 2019/04/041
       'If "" & rsAD.Fields("cp148") = "Y" Then 'Remove by Lydia 2022/11/09
           'Modified by Lydia 2017/12/20 點選的本所案號
           'strSql = UCase(PUB_GetOneAppMuchCaseSql(Text1(0), Text1(1), Text1(2), Text1(3), rsAD.Fields("cp10"), rsAD.Fields("cp27"), rsAD.Fields("cp28")))
           If "" & rsAD.Fields("cp148") = "Y" Then 'Added by Lydia 2022/11/09
                strSql = UCase(PUB_GetOneAppMuchCaseSql(m_TM01, m_TM02, m_TM03, m_TM04, rsAD.Fields("cp10"), "" & rsAD.Fields("cp27"), "" & rsAD.Fields("cp28")))
               'Modified by Lydia 2022/11/09  列印過定稿或譯文即算有列印; ex.FCT-049141 承辦自己寫定稿,只要印譯文
               'strSql = "select ld01,ld02,ld03,ld04,ld05,ld06,ld07,ld08,ld09,ld10,ld11,ld16 from letterdemand ,(" & Mid(strSql, 1, InStr(strSql, "ORDER") - 1) & _
                        ") X1 where ld02=" & pLD02 & " and LD11='" & pLD11 & "' and ld04=cp09 order by LD04"
               strSql = "select ld01,ld02,ld03,ld04,ld05,ld06,ld07,ld08,ld09,ld10,ld11,ld16 from letterdemand ,(" & Mid(strSql, 1, InStr(strSql, "ORDER") - 1) & _
                        ") X1 where ld02=" & pLD02 & " and LD10='" & pLD10 & "' and ld04=cp09 order by LD04"
           'Added by Lydia 2022/11/09 非一文多案
           Else
                strSql = "select ld01,ld02,ld03,ld04,ld05,ld06,ld07,ld08,ld09,ld10,ld11,ld16 from letterdemand " & _
                            "where ld01='" & pLD01 & "' and ld02=" & pLD02 & " and LD04='" & pLD04 & "' and LD10='" & pLD10 & "' order by LD04"
           End If
           'end 2022/11/09
           intA = 1
           Set rsAD = ClsLawReadRstMsg(intA, strSql)
           If intA = 1 Then
              cnnConnection.BeginTrans
              rsAD.MoveFirst
              Do While Not rsAD.EOF
                 'Modified by Lydia 2018/10/01 改成目前操作者(ex.母案FCT-42446已列印,子案FCP-42447的定稿未能一起更新,猜測是換人列印)
                 'strTime = PUB_GetUniqeLD03(pLD01, strSrvDate(1), "" & rsAD.Fields("ld03"))
                 strTime = PUB_GetUniqeLD03(strUserNum, strSrvDate(1), "" & rsAD.Fields("ld03"))
                 'Modified by Lydia 2019/04/01 where條件+LD01,LD03 (ex.FCT-38643因為是註冊後變更所以沒有核准定稿，但是譯文產生2筆定稿，造成更新後成為重覆主鍵)
                 'strSql = "update letterdemand set ld02=" & strSrvDate(1) & ",ld01='" & strUserNum & "',ld03='" & strTime & "' where ld04 ='" & rsAD.Fields("ld04") & "' and ld02=" & pLD02 & " and ld09='" & strCP10 & "' and ld10='" & pLD10 & "' and LD11='" & pLD11 & "'"
                 'Modified by Lydia 2022/11/09  列印過定稿或譯文即算有列印
                 'strSql = "update letterdemand set ld02=" & strSrvDate(1) & ",ld01='" & strUserNum & "',ld03='" & strTime & "' where ld04 ='" & rsAD.Fields("ld04") & "' and ld02=" & pLD02 & " and ld09='" & strCP10 & "' and ld10='" & pLD10 & "' and LD11='" & pLD11 & "' and ld01='" & rsAD.Fields("ld01") & "' and LD03='" & rsAD.Fields("ld03") & "' "
                 strSql = "update letterdemand set ld02=" & strSrvDate(1) & ",ld01='" & strUserNum & "',ld03='" & strTime & "' where ld04 ='" & rsAD.Fields("ld04") & "' and ld02=" & pLD02 & " and ld09='" & strCP10 & "' and ld10='" & pLD10 & "' and LD11='" & rsAD.Fields("ld11") & "' and ld01='" & rsAD.Fields("ld01") & "' and LD03='" & rsAD.Fields("ld03") & "' "
                 cnnConnection.Execute strSql, intA
                 'Modified by Lydia 2019/04/01 where條件+ET04
                 'strSql = "update exceptcondition set et07=" & strSrvDate(1) & ",et04='" & strUserNum & "' where et02 ='" & rsAD.Fields("ld04") & "' and et07=" & pLD02 & " and et01='" & pLD10 & "' and et03='" & pLD11 & "'"
                 'Modified by Lydia 2022/11/09  列印過定稿或譯文即算有列印
                 'strSql = "update exceptcondition set et07=" & strSrvDate(1) & ",et04='" & strUserNum & "' where et02 ='" & rsAD.Fields("ld04") & "' and et07=" & pLD02 & " and et01='" & pLD10 & "' and et03='" & pLD11 & "' and et04='" & rsAD.Fields("ld01") & "'"
                 strSql = "update exceptcondition set et07=" & strSrvDate(1) & ",et04='" & strUserNum & "' where et02 ='" & rsAD.Fields("ld04") & "' and et07=" & pLD02 & " and et01='" & pLD10 & "' and et03='" & rsAD.Fields("ld11") & "'  and et04='" & rsAD.Fields("ld01") & "'"
                 cnnConnection.Execute strSql, intA
                 rsAD.MoveNext
              Loop
              cnnConnection.CommitTrans
           End If
       'Remove by Lydia 2022/11/09 列印過定稿或譯文即算有列印
       'Else
'            'Modified by Lydia 2018/10/01 改成目前操作者 'Memo by Lydia 2022/09/01 因為「暫時將strUserNum設定為定稿的使用者編號」，所以仍是以定稿LD01執行
'            'strTime = PUB_GetUniqeLD03(pLD01, strSrvDate(1), pLD03)
'            strTime = PUB_GetUniqeLD03(strUserNum, strSrvDate(1), pLD03)
'            cnnConnection.BeginTrans
'                'Modified by Lydia 2019/05/06 +LD03
'                'strSql = "update letterdemand set ld02=" & strSrvDate(1) & ",ld01='" & strUserNum & "',ld03='" & strTime & "' where ld04='" & pLD04 & "' and ld01='" & pLD01 & "' and ld02=" & pLD02 & " and ld10='" & pLD10 & "' and ld11='" & pLD11 & "' "
'                strSql = "update letterdemand set ld02=" & strSrvDate(1) & ",ld01='" & strUserNum & "',ld03='" & strTime & "' where ld04='" & pLD04 & "' and ld01='" & pLD01 & "' and ld02=" & pLD02 & " and ld03=" & pLD03 & "  and ld10='" & pLD10 & "' and ld11='" & pLD11 & "' "
'                cnnConnection.Execute strSql, intA
'                strSql = "update exceptcondition set et07=" & strSrvDate(1) & ",et04='" & strUserNum & "' where et02='" & pLD04 & "' and et04='" & pLD01 & "' and et07=" & pLD02 & " and et01='" & pLD10 & "' and et03='" & pLD11 & "' "
'                cnnConnection.Execute strSql, intA
'            cnnConnection.CommitTrans
       'End If
       'end 2022/11/09
    End If
    'end 2017/09/18
    
    Exit Sub
ErrHandle:
'Added by Lydia 2019/04/01 更新出錯通知電腦中心
If Err.Number <> 0 Then
    cnnConnection.RollbackTrans
    PUB_SendMail strUserNum, "A3034", "", Me.Caption & "-定稿日期更新失敗：" & strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04 & "(LD04=" & pLD04 & ")", vbCrLf & Err.Description
End If
End Sub

'Added by Lydia 2017/12/20 點選的本所案號
Private Sub GetNowTMNo(ByVal iRow As Integer)
        m_TM01 = SystemNumber(Me.grdDataList.TextMatrix(iRow, 2), 1)
        m_TM02 = SystemNumber(Me.grdDataList.TextMatrix(iRow, 2), 2)
        m_TM03 = SystemNumber(Me.grdDataList.TextMatrix(iRow, 2), 3)
        m_TM04 = SystemNumber(Me.grdDataList.TextMatrix(iRow, 2), 4)
End Sub
'Added by Morgan 2018/11/1
'檢查是否ｅ化客戶通知函
Private Function ChkECustLetter(pLD01 As String, pLD02 As String, pLD03 As String, Optional ByRef pRcvrNo As String, Optional ByRef pPrtAddr As Boolean) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stCustNo As String
   
   pRcvrNo = ""
   'Modified by Morgan 2021/12/1 +ld10
   stSQL = "select nvl(a.ftm05,b.ftm05) ftm05,c1.cp55,c1.cp56,c1.cp72,c2.cp72 cp72r,ld05||ld06||ld07||ld08 CaseNo,ld05,ld10,ld11" & _
      " from letterdemand,finaltextmap a,finaltextmap b,caseprogress c1,caseprogress c2" & _
      " where ld01='" & pLD01 & "' and ld02=" & pLD02 & " and ld03=" & pLD03 & _
      " and a.ftm01(+)=ld05 and a.ftm02(+)=ld10 and a.ftm03(+)=ld09 and a.ftm04(+)=ld11" & _
      " and b.ftm01(+)=ld05 and b.ftm02(+)=ld10 and b.ftm03(+)='000' and b.ftm04(+)=ld11" & _
      " and c1.cp09(+)=ld04 and c2.cp09(+)=c1.cp43"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      If InStr("" & .Fields("ftm05"), "<中申開窗郵號") > 0 Then
         If InStr(.Fields("ftm05"), "<中申開窗郵號/被授權>") > 0 Then
            stCustNo = "" & .Fields("cp72")
         ElseIf InStr(.Fields("ftm05"), "<中申開窗郵號/相關被授權>") > 0 Then
            stCustNo = "" & .Fields("cp72r")
         ElseIf InStr(.Fields("ftm05"), "<中申開窗郵號/移轉人>") > 0 Then
            stCustNo = "" & .Fields("cp55")
         ElseIf InStr(.Fields("ftm05"), "<中申開窗郵號/移轉>") > 0 Then
            stCustNo = "" & .Fields("cp56")
         Else
            stCustNo = PUB_GetCustNo(.Fields("CaseNo"))
         End If
         If stCustNo <> "" Then
            'Modified by Morgan 2022/2/21 全E化客戶才要
            'If PUB_ChkECust(stCustNo, .Fields("ld05")) Then
            If PUB_ChkECust(stCustNo, .Fields("ld05"), , intI) Then
               If intI = 1 Then
                  ChkECustLetter = True
                  pRcvrNo = stCustNo
                  'If InStr(.Fields("ftm05"), "<中申開窗郵號/掛>") = 0 Then 'Removed by Morgan 2018/11/19 掛號直寄E化也不要寄紙本--文雄
                     'Added by Morgan 2021/12/1 CFT改判斷定稿別
                     If .Fields("ld05") = "CFT" Then
                        If .Fields("ld10") = "05" Then
                           pPrtAddr = True
                        End If
                     Else
                     'end 2021/12/1
                        If MsgBox("本信函是否需寄送【紙本】？" & vbCrLf & vbCrLf & "※有實體文件(例如收據、證書…)請選【是】" & vbCrLf & "※選【否】將不印【郵遞區號及地址】", vbYesNo + vbDefaultButton2 + vbQuestion, "ｅ化客戶通知函提醒") = vbYes Then
                           pPrtAddr = True
                        End If
                     End If
                  'End If
               End If
            End If
         End If
      End If
      End With
   End If
   
   Set rsQuery = Nothing
End Function
