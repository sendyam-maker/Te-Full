VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060104_j 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳年費整批發文"
   ClientHeight    =   5916
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8508
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5916
   ScaleWidth      =   8508
   Begin VB.TextBox txtPayToday 
      Height          =   264
      Left            =   7515
      MaxLength       =   1
      TabIndex        =   11
      Top             =   5160
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4740
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   5520
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印清單(&O)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   0
      Left            =   4590
      TabIndex        =   6
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "開啟..."
      Height          =   315
      Left            =   7245
      TabIndex        =   3
      Top             =   660
      Width           =   825
   End
   Begin VB.TextBox txtSample 
      Height          =   315
      Left            =   1305
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   660
      Width           =   5910
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "發文(&O)"
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   400
      Index           =   1
      Left            =   6435
      TabIndex        =   1
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7260
      TabIndex        =   0
      Top             =   120
      Width           =   800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3150
      Top             =   5430
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   3945
      Left            =   270
      TabIndex        =   5
      Top             =   1080
      Width           =   7950
      _ExtentX        =   14034
      _ExtentY        =   6964
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
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
      _Band(0).Cols   =   13
   End
   Begin VB.Label lblPayToday 
      AutoSize        =   -1  'True
      Caption         =   "電子送件是否當日扣款:         (Y/N)"
      Height          =   180
      Left            =   5580
      TabIndex        =   12
      Top             =   5190
      Width           =   2655
   End
   Begin VB.Label lblCount 
      Height          =   180
      Left            =   1035
      TabIndex        =   10
      Top             =   5490
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件數："
      Height          =   180
      Index           =   1
      Left            =   270
      TabIndex        =   9
      Top             =   5490
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   "印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3780
      TabIndex        =   8
      Top             =   5550
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CSV檔案："
      Height          =   180
      Index           =   0
      Left            =   405
      TabIndex        =   4
      Top             =   690
      Width           =   870
   End
End
Attribute VB_Name = "frm060104_j"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/18 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Created by Morgan 2012/8/28
Option Explicit

Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim bolBarShow As Boolean
Dim strPrinter As String
Dim m_iCols As Integer
'Added by Lydia 2019/10/01
Dim colCP148 As Integer, colCP60 As Integer '特殊請款單,請款單號的欄位
Dim strCP09List As String '已發文的收文號


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 0
      If grdDataList.Rows > 1 Then
         If grdDataList.TextMatrix(1, 1) <> "" Then
            PUB_RestorePrinter Combo1
            DoPrint
            PUB_RestorePrinter strPrinter
         End If
      End If
   Case 1
      Screen.MousePointer = vbHourglass
      strCP09List = "" 'Added by Lydia 2019/10/01
      doBatch
      Screen.MousePointer = vbDefault
     
      'Added by Lydia 2019/10/01 領證/年費發文直接產生:承辦單+請款定稿+帳單(請款單)
      If strCP09List <> "" Then
            MsgBox "開始產生帳單！"
            Screen.MousePointer = vbHourglass
            Call doBatchAddAcc1k0(strCP09List)
            Screen.MousePointer = vbDefault
      End If
      'end 2019/10/01
   Case 2
      Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean
   FormSave = True
End Function

Private Sub cmdOpen_Click()
   Dim strPath As String
   strPath = GetSaveName(txtSample)
   If strPath <> "" Then
      txtSample = strPath
      If Dir(txtSample) <> "" Then
         grdDataList.Visible = False
         ReadData
         grdDataList.Visible = True
      Else
         MsgBox "無法讀取檔案！"
      End If
   End If
End Sub

Private Function GetSaveName(ByVal pFileName As String) As String
   Dim strPath As String, strFileName As String
   
   If InStrRev(pFileName, "\") > 0 Then
      strPath = Left(pFileName, InStrRev(pFileName, "\") - 1)
      strFileName = Mid(pFileName, InStrRev(pFileName, "\") + 1)
   Else
      'strPath = PUB_Getdesktop
      'Modified by Morgan 2015/8/13 CSV檔只留最後一次的
      'strPath = EFilePath & "\CSV\" & Format(Now, "YYYYMMDD")
      'If Dir(strPath, vbDirectory) = "" Then
      '   strPath = EFilePath & "\CSV"
      'End If
      strPath = EFilePath & "\CSV"
      strFileName = pFileName
   End If
   
On Error GoTo ErrHnd

   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "CSV 檔 (*.CSV)|*.CSV"
      .InitDir = strPath
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         GetSaveName = .FileName
      End If
   End With
   
   Exit Function
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   cmdok(0).Enabled = IsUserHasRightOfFunction(Me.Name, strPrint, False)
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   
   'Added by Morgan 2013/5/15
   If Val(ServerTime) <= 153000 Then
      txtPayToday = "Y"
   End If
   'end 2013/5/15
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2020/08/17
   
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   Set frm060104_j = Nothing
End Sub

'Modified by Morgan 2023/6/20 e網通新網頁改用UTF-8編碼的CSV(以LF符號斷行)
Private Sub ReadData()
   Dim ii As Integer, jj As Integer
   Dim ff As Integer, strData As String
   Dim arrData
   Dim strDueDate As String
   Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean 'Added by Lydia 2016/04/28
   Dim kk As Integer, arrRow() As String 'Added by Morgan 2023/6/20
   
   ii = 1
   'Modified by Morgan 2023/6/20
   'If ff > 0 Then Close #ff
   'ff = FreeFile
   'Open txtSample For Input Access Read As #ff
   'Do While Not EOF(ff)
   '   Line Input #ff, strData
   '   arrData = Split(strData, ",")
   strData = PUB_ReadTextFile(txtSample, "UTF-8")
   If strData = "" Then Exit Sub
   arrRow = Split(strData, vbLf)
   For kk = 0 To UBound(arrRow)
   If arrRow(kk) <> "" Then
      arrData = Split(arrRow(kk), ",")
   ' end 2023/6/20
      ii = ii + 1
      grdDataList.Rows = ii
      For jj = 0 To UBound(arrData)
         'Modified by Morgan 2021/4/21 第6欄(jj=5)收據抬頭(固定為A 專利權人)要略過
         If jj < 5 Then
            grdDataList.TextMatrix(ii - 1, jj) = arrData(jj)
         ElseIf jj > 5 And jj < 8 Then 'Modified by Morgan 2023/6/20 + And jj < 8
            grdDataList.TextMatrix(ii - 1, jj - 1) = arrData(jj)
         End If
         'end 2021/4/21
      Next
      'Modified by Lydia 2016/04/28 +特殊請款單CP148
      'Modified by Lydia 2019/10/01 +CP60
      'Modified by Morgan 2020/12/3 +order by nvl(cp27,cp05) desc,有可能退費後重新繳納 Ex:FCP-056910
      strExc(0) = "select pa01||'-'||pa02,sqldatet(cp27),cp09,pa01,pa02,pa03,pa04,cp14,st02,pa14,cp148,CP60 " & _
                        "from patent,caseprogress,staff where pa11='" & arrData(0) & "' and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10(+)='605' and cp53(+)='" & arrData(2) & "' and st01(+)=cp14 order by nvl(cp27,cp05) desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         For jj = 0 To 8
            grdDataList.TextMatrix(ii - 1, jj + 7) = "" & RsTemp(jj)
         Next
         strDueDate = CompDate(0, Val(grdDataList.TextMatrix(ii - 1, 2)) - 1, RsTemp(9))
         strDueDate = CompDate(2, -1, strDueDate)
         strDueDate = PUB_GetWorkDay1(strDueDate, False)
         If RsTemp(1) <> "" Then
            strExc(1) = DBDATE(RsTemp(1))
         Else
            strExc(1) = strSrvDate(1)
         End If
         'Added by Lydia 2016/04/28 判斷是否產生電子檔
         jj = 16
         m_bolEmail = PUB_GetEMailFlag(RsTemp.Fields("pa01") & RsTemp.Fields("pa02") & RsTemp.Fields("pa03") & RsTemp.Fields("pa04"), True, , m_bolPlusPaper)
         strExc(2) = IIf(m_bolEmail, IIf(m_bolPlusPaper, "Ｅ", "ｅ"), "")
         If strExc(2) <> "" Then
            grdDataList.TextMatrix(ii - 1, jj) = strExc(2)
         End If
         jj = jj + 1
         If colCP148 = 0 Then colCP148 = jj 'Added by Lydia 2019/10/01
         'Added by Lydia 2016/04/28 +特殊請款單CP148
         If Mid(RsTemp.Fields("cp09"), 1, 1) < "C" And Not IsNull(RsTemp.Fields("cp148")) Then
            grdDataList.TextMatrix(ii - 1, jj) = "Y"
         End If
         jj = jj + 1
         'Memo by Lydia 2017/09/06 是否逾期補繳
         If Val(strExc(1)) > Val(strDueDate) Then
            'Modified by Lydia 2016/04/28
            'grdDataList.TextMatrix(ii - 1, 16) = "Y"
            grdDataList.TextMatrix(ii - 1, jj) = "Y"
         End If
         'end 2016/04/28
         'Added by Lydia 2019/10/01 請款單號
         jj = jj + 1
         If colCP60 = 0 Then colCP60 = jj
         grdDataList.TextMatrix(ii - 1, jj) = "" & RsTemp.Fields("cp60")
         'end 2019/10/01
      End If
   
   'Modified by Morgan 2023/6/20
   'Loop
   'Close #ff
   End If
   Next kk
   'end 2023/6/20
   If ii > 1 Then cmdok(1).Enabled = True
   lblCount = ii - 1
End Sub

Private Sub SetDataListWidth()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iCol As Integer
   'Added by Lydia 2016/04/28 +特殊請款單CP148,Ｅｅ化
   'Modified by Lydia 2019/10/01 +CP60
   arrGridHeadText = Array("申請案號", "證書號", "起年", "迄年", "金額", "減收資格", "減收資格類型", "本所案號", "發文日", "收文號", "PA01", "PA02", "PA03", "PA04", "CP14", "ST02", "ｅ", "特殊請款", "補繳", "CP60")
   arrGridHeadWidth = Array(1000, 850, 450, 450, 650, 820, 1200, 1100, 850, 0, 0, 0, 0, 0, 0, 0, 0, 0, 450, 0)
   grdDataList.Cols = UBound(arrGridHeadText) + 1
   For iCol = 0 To grdDataList.Cols - 1
      grdDataList.row = 0
      grdDataList.col = iCol
      grdDataList.Text = arrGridHeadText(iCol)
      grdDataList.ColWidth(iCol) = arrGridHeadWidth(iCol)
      grdDataList.CellAlignment = flexAlignCenterCenter
      Select Case iCol
         Case 4
            grdDataList.ColAlignment(iCol) = flexAlignRightCenter
         'Modified by Lydia 2016/04/28 e化(16),+特殊請款單CP148 (17)
         'Case 2, 3, 5, 6, 16
         Case 2, 3, 5, 6, 16, 17, 18
            grdDataList.ColAlignment(iCol) = flexAlignCenterCenter
         Case Else
            grdDataList.ColAlignment(iCol) = flexAlignLeftCenter
      End Select
   Next iCol
   grdDataList.BackColor = &HFFC0C0
End Sub

Private Function doBatch() As Boolean
   Dim iRow As Integer
   Dim bSuccess As Boolean
   Dim pa(4) As String, strCP09 As String
   Dim stCP152 As String 'Added by Morgan 2013/5/15
   
   'Added by Morgan 2013/5/15
   If txtPayToday = "" Then
      MsgBox "請輸入是否當日扣扣款(Y/N)！", vbExclamation
      txtPayToday.SetFocus
      Exit Function
   Else
      'Modified by Lydia 2018/09/11 改成模組
      'If txtPayToday = "Y" Then
      '   stCP152 = CompWorkDay(2, strSrvDate(1))
      'Else
      '   stCP152 = CompWorkDay(3, strSrvDate(1))
      'End If
      stCP152 = Pub_FcpSetPayToday("2", strSrvDate(1), txtPayToday)
   End If
   'end 2013/5/15
   
   doBatch = True
   With grdDataList
   For iRow = 1 To .Rows - 1
      If .TextMatrix(iRow, 8) = "" Then
         bSuccess = False
         If .TextMatrix(iRow, 9) = "" Then
            MsgBox .TextMatrix(iRow, 7) & " 資料錯誤，請確認！", vbExclamation
         Else
            pa(1) = .TextMatrix(iRow, 10)
            pa(2) = .TextMatrix(iRow, 11)
            pa(3) = .TextMatrix(iRow, 12)
            pa(4) = .TextMatrix(iRow, 13)
            strCP09 = .TextMatrix(iRow, 9)
            If PUB_ChkNP605(pa(1) & pa(2) & pa(3) & pa(4)) Then
               MsgBox pa(1) & pa(2) & pa(3) & pa(4) & " 下一程序有<年費>期限不可發文!!!", vbExclamation + vbOKOnly
            ElseIf PUB_ApproveCheck(strCP09) Then
               With frm060104_a
                  .m_bolBeCalled = True
                  .m_CP01 = pa(1)
                  .m_CP02 = pa(2)
                  .m_CP03 = pa(3)
                  .m_CP04 = pa(4)
                  .m_CP09 = strCP09
                  .Text5(4) = strSrvDate(2)
                  .Text5(0) = grdDataList.TextMatrix(iRow, 2)
                  .Text5(1) = grdDataList.TextMatrix(iRow, 3)
                  .txtCP118 = "A" '自動扣款
                  'Modified by Lydia 2016/04/28
                  '.Text5(2) = grdDataList.TextMatrix(iRow, 16)
                  'Modified by Lydia 2017/09/06 debug
                  '.Text5(2) = GrdDataList.TextMatrix(iRow, 17)
                  .Text5(2) = grdDataList.TextMatrix(iRow, 18) '是否逾期補繳
                  .m_CP84 = Val(grdDataList.TextMatrix(iRow, 4))
                  .m_CP152 = stCP152 'Added by Morgan 2013/5/15
                  bSuccess = .Process()
               End With
               DoEvents
               Unload frm060104_a
            End If
         End If
         
         
         If bSuccess Then
            .TextMatrix(iRow, 8) = ChangeTStringToTDateString(strSrvDate(2))
            'Added by Lydia 2019/10/01 記錄已發文的收文號
            If Trim("" & .TextMatrix(iRow, colCP60)) = "" Then
                strCP09List = strCP09List & strCP09 & ","
            End If
            'end 2019/10/01
         Else
            doBatch = False
            Exit For
         End If
      End If
   Next
   End With
   
   NoBatchCaseInform 'Added by Morgan 2025/7/24
   
   If doBatch = True Then
      MsgBox "全部發文完成！"
   Else
      MsgBox "發文失敗，第" & iRow & "筆(" & grdDataList.TextMatrix(iRow, 7) & ")！"
   End If
End Function

Private Sub DoPrint()

   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp() As String, lngSum As Long

   iOrientation = Printer.Orientation
   Printer.Orientation = 1
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With grdDataList
      GetPleft
      ReDim strTemp(1 To m_iCols)
      iPage = 1
      PrintPageHeader
      PrintPageHeader1
      For iRow = 1 To .Rows - 1
         strTemp(1) = grdDataList.TextMatrix(iRow, 15)
         strTemp(2) = grdDataList.TextMatrix(iRow, 7)
         'Modified by Lydia 2016/04/28
'         strTemp(3) = GrdDataList.TextMatrix(iRow, 0)
'         strTemp(4) = Format(GrdDataList.TextMatrix(iRow, 4), "#,###")
'         strTemp(5) = GrdDataList.TextMatrix(iRow, 2) & IIf(GrdDataList.TextMatrix(iRow, 2) = GrdDataList.TextMatrix(iRow, 3), "", "-" & GrdDataList.TextMatrix(iRow, 3))
'         strTemp(6) = GrdDataList.TextMatrix(iRow, 8)
         strTemp(3) = grdDataList.TextMatrix(iRow, 16)
         strTemp(4) = grdDataList.TextMatrix(iRow, 17)
         strTemp(5) = grdDataList.TextMatrix(iRow, 0)
         strTemp(6) = Format(grdDataList.TextMatrix(iRow, 4), "#,###")
         strTemp(7) = grdDataList.TextMatrix(iRow, 2) & IIf(grdDataList.TextMatrix(iRow, 2) = grdDataList.TextMatrix(iRow, 3), "", "-" & grdDataList.TextMatrix(iRow, 3))
         strTemp(8) = grdDataList.TextMatrix(iRow, 8)
         'end 2016/04/28
         lngSum = lngSum + grdDataList.TextMatrix(iRow, 4)
         PrintDetail strTemp
      Next
      Call PrintReportFooter(.Rows - 1, lngSum)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
   
End Sub


Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   'Modified by Lydia 2016/04/28
   'm_iCols = 6
   m_iCols = 8
   ReDim PLeft(1 To m_iCols + 1)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(3, "　")) + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(7, "　")) + ciColGap
   'Modified by Lydia 2016/04/28 + e化,特殊請款單
'   PLeft(4) = PLeft(3) + Printer.TextWidth(String(5, "　")) + ciColGap
'   PLeft(5) = PLeft(4) + Printer.TextWidth(String(5, "　")) + ciColGap
'   PLeft(6) = PLeft(5) + Printer.TextWidth(String(5, "　")) + ciColGap
'   PLeft(7) = PLeft(6) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(1, "　")) + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(2, "　")) + ciColGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(9) = PLeft(8) + Printer.TextWidth(String(5, "　")) + ciColGap
   'end 2016/04/28
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      PrintLine
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
    
End Sub

Private Sub PrintLine()
   Dim iNo As Integer
   iNo = (Printer.ScaleWidth - Printer.CurrentX - 500) \ Printer.TextWidth("-")
   Printer.Print String(iNo, "-")
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
      Select Case iCol
         'Modified by Lydia 2016/04/28 + 2
         'Case 4
         Case 6
            Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - ciColGap
            Printer.CurrentY = iPrint
            Printer.Print strData(iCol)
         Case Else
            Printer.CurrentX = PLeft(iCol)
            Printer.CurrentY = iPrint
            Printer.Print strData(iCol)
      End Select
    Next
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = "繳年費整批發文明細表"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "CSV檔案：" & txtSample
   PrintNewLine
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   PrintLine
End Sub

Sub PrintPageHeader1()

   Call PrintNewLine(False, 1)
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   
   'Added by Lydia 2016/04/28
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "ｅ"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "特殊"
   
   'Modified by Lydia 2016/04/28 +2
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "申請號"
    
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "規費"
   
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "繳費年度"
   
   Printer.CurrentX = PLeft(8)
   'end 2016/04/28
   Printer.CurrentY = iPrint
   Printer.Print "發文日"
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   PrintLine
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0, Optional ByVal lngSum As Long)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    PrintLine
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & Format(lngSum, "#,###") & " 元 ( " & iRecCount & " 筆 )"
    Printer.EndDoc
End Sub

Private Sub txtPayToday_GotFocus()
   TextInverse txtPayToday
   CloseIme
End Sub

Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2019/10/01 整批:領證/年費發文直接產生:承辦單+請款定稿+帳單(請款單), 並且列印整批清單
Private Sub doBatchAddAcc1k0(ByVal pList As String)
Dim nFrm As Form
    If pList <> "" Then
       '檢查表單是否已開啟，若是，則關閉
        For Each nFrm In Forms
           If StrComp(nFrm.Name, "frm060307", vbTextCompare) = 0 Then
              Unload frm060307
              Exit For
           End If
        Next
        frm060307.m_KeyCP09 = pList
        frm060307.m_KeyCP10 = "605"

        Call frm060307.SetData(0, "2", True) 'Added by Lydia 2020/08/17 外部呼叫,預設類別

        frm060307.Show
        Call frm060307.cmdok_Click(0)
        Unload frm060307
    End If
End Sub

'Added by Morgan 2025/7/24
'有設定整批年費暫不繳案件通知
Private Sub NoBatchCaseInform()
   strExc(0) = "select * from caseprogress where cp05>to_char(sysdate-365,'yyyymmdd') and cp01='FCP' and cp10='605' and cp158=0 and cp159=0 and cp141='4'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         '正本
         strExc(1) = "" & .Fields("cp14") '年費承辦人
         '副本
         strExc(2) = PUB_GetFCPHandler(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), "605") '程序管制人
         If strExc(2) = strExc(1) Then
            strExc(2) = ""
         End If
         '主旨
         strExc(3) = "FCP-" & .Fields("cp02") & IIf(.Fields("cp03") & .Fields("cp04") = "000", "", "-" & .Fields("cp03") & "-" & .Fields("cp04")) & "年費設定暫不繳納，請確認是否可進行年費繳納！"
         '內文
         strExc(4) = "如旨。" & vbCrLf & "(若為誤設暫不繳，請盡速完成繳納)"
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            "select * from (select '" & strUserNum & "' c01,'" & strExc(1) & "' c02,to_char(sysdate,'yyyymmdd') c03,to_char(sysdate,'hh24miss') c04" & _
            ",'" & strExc(3) & "' c07,'" & strExc(4) & "' c08,'" & strExc(2) & "' c09,'" & .Fields("cp09") & "' c13" & _
            " from dual) X  where not exists(select * from mailcache where mc02=c02 and mc07=c07 and mc13=c13)"
         cnnConnection.Execute strSql, intI
         .MoveNext
      Loop
      End With
   End If
   PUB_SendMailCache
End Sub
