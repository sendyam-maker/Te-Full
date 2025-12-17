VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14m0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "收款扣繳改年度清單"
   ClientHeight    =   2592
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5712
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2592
   ScaleWidth      =   5712
   Begin VB.CommandButton CmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   510
      Style           =   1  '圖片外觀
      TabIndex        =   13
      Top             =   2000
      Width           =   4692
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
      Left            =   1530
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   3000
      Width           =   3450
   End
   Begin VB.OptionButton Option1 
      Caption         =   "第二次以上"
      Height          =   315
      Index           =   1
      Left            =   3660
      TabIndex        =   3
      Top             =   840
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "第一次列印"
      Height          =   315
      Index           =   0
      Left            =   2130
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   1
      Top             =   420
      Width           =   705
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   0
      Top             =   36
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   510
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   2580
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   312
      Left            =   2136
      TabIndex        =   4
      Top             =   1200
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "資料請存檔備查並自行記錄上次印日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Index           =   2
      Left            =   960
      TabIndex        =   14
      Top             =   1600
      Width           =   4500
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   510
      TabIndex        =   12
      Top             =   3030
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "列印範圍："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   960
      TabIndex        =   11
      Top             =   840
      Width           =   1248
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "收款年度："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   960
      TabIndex        =   10
      Top             =   60
      Width           =   1248
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   3240
      TabIndex        =   9
      Top             =   1080
      Width           =   252
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "上次列印日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   516
      TabIndex        =   8
      Top             =   1236
      Width           =   1608
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   960
      TabIndex        =   7
      Top             =   456
      Width           =   1248
   End
End
Attribute VB_Name = "Frmacc14m0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Created by Sindy 2012/4/27
Option Explicit

Dim PLeft() As Integer, intY As Integer, iPageNo As Integer, strPrinter As String
Dim strF() As String, arrWidth, strSum(1) As String, intField As Integer  'Add by Amy 2024/06/03

'Add by Amy 2024/06/03
Private Sub cmdExcel_Click()
   Dim rsQ1 As ADODB.Recordset, RsQ2 As ADODB.Recordset, strQ As String, strCon As String
   Dim intQ1 As Integer, intQ2 As Integer, strMsg As String
   Dim hLocalFile As Long 'Add by Amy 2024/06/05
   
   If FormCheck = False Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   strCon = "": strSum(0) = "": strSum(1) = ""
   '修改扣繳年度的修改日期條件
   '第一次
   If Option1(0).Value = True Then
      strCon = strCon & " And A0K15<=" & FCDate(MaskEdBox1)
   End If
   '第二次以上
   If Option1(1).Value = True Then
      'Modify by Amy 2025/03/21 因之前2025/01/06 休長年假,畫面列印日期預帶前一工作天,因帶錯
      '                              故將strDate變數Mark , 未想到第2次印的迄日需用到 ex:20250223 前一工作天帶錯,應帶1140124
      strDate = Val(CompWorkDay(2, strSrvDate(1), 1)) - 19110000 '前1個工作天
      strCon = strCon & " And a0k15>=" & FCDate(MaskEdBox1) & " And a0k15<=" & strDate & " "
   End If
   
'*** 收款轉前一年列表 ***
   strSql = GetSql(0, True, strCon) '服務費合計
   intQ1 = 1
   Set rsQ1 = ClsLawReadRstMsg(intQ1, strSql)
   If intQ1 = 1 Then
      If Not IsNull(rsQ1.Fields(0)) Then strSum(0) = rsQ1.Fields(0)
   End If
   '列表
   strSql = GetSql(0, False, strCon)
   intQ1 = 1
   Set rsQ1 = ClsLawReadRstMsg(intQ1, strSql)
   If intQ1 = 1 Then
      rsQ1.MoveFirst
      If SaveExcel(0, rsQ1) = False Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   End If
'*** End 收款轉前一年列表 ***

'*** 扣繳年度改回收款年度清單 ***
   strSql = GetSql(1, True, strCon) '服務費合計
   intQ2 = 1
   Set RsQ2 = ClsLawReadRstMsg(intQ2, strSql)
   If intQ2 = 1 Then
      If Not IsNull(RsQ2.Fields(0)) Then strSum(1) = RsQ2.Fields(0)
   End If
   
   strSql = GetSql(1, False, strCon)
   intQ2 = 1
   Set RsQ2 = ClsLawReadRstMsg(intQ2, strSql)
   If intQ2 = 1 Then
      RsQ2.MoveFirst
      If SaveExcel(1, RsQ2) = False Then
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
   End If
'*** End 扣繳年度改回收款年度清單 ***
   If intQ1 = 1 Or intQ2 = 1 Then
      If intQ1 = 1 Then strMsg = "收款轉前一年列表" & vbCrLf
      If intQ2 = 1 Then
         If strMsg <> MsgText(601) Then
            strMsg = strMsg & "　　　　及" & vbCrLf
         End If
         strMsg = strMsg & "扣繳年度改回收款年度清單" & vbCrLf
      End If
      strMsg = strMsg & "已產生!" & vbCrLf
      'Add by Amy 2024/06/05
      If MsgBox(strMsg & "是否開啟資料夾？", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
         ShellExecute hLocalFile, "explore", strExcelPath, vbNullString, vbNullString, 1
      End If
   Else
      MsgBox "無資料可供列印！"
   End If
   Screen.MousePointer = vbDefault
   Set rsQ1 = Nothing
   Set RsQ2 = Nothing
End Sub

Private Sub Command1_Click()
   'Mark by Amy 已不使用
'   Dim strCon As String
'   Dim rsReport As ADODB.Recordset
'   'Dim iItemCount As Integer 'Modify By Amy 2013/07/04
'   'Dim i As Integer 'Modify By Amy 2013/07/04
'   Dim dblSum As Double
'   Dim rsReport2  As ADODB.Recordset, dblSum2 As Double, intR As Integer 'Add by Amy 2013/07/04
'
'   If FormCheck = False Then Exit Sub
'
'   Screen.MousePointer = vbHourglass
'   strCon = ""
'
'   'Modify by Amy 2013/06/20 搬至Form_Load
'   '系統日前一天
'   'strDate = Val(CompDate(2, -1, strSrvDate(1))) - 19110000 '日曆天
'   'strDate = Val(CompWorkDay(1, strSrvDate(1), 1)) - 19110000 '工作天
'   'End 2013/06/20
'
'   '修改扣繳年度的修改日期條件
'   '第一次
'   If Option1(0).Value = True Then
'      'Modify by Amy 2013/06/20 改抓A0K15<=畫面輸入之上次列印日期
'      'strCon = strCon & "AND (A0K27 IS NULL OR A0K27<=" & strDate & ") "
'      strCon = strCon & "And A0K15<=" & FCDate(MaskEdBox1)
'   End If
'   '第二次以上
'   If Option1(1).Value = True Then
'      'Modify by Amy 2013/06/20 改抓A0K15<=畫面輸入之上次列印日期
'      'strCon = strCon & "AND a0k27>=" & FCDate(MaskEdBox1.Text) & " and a0k27<=" & strDate & " "
'      strCon = strCon & "AND a0k15>=" & FCDate(MaskEdBox1.Text) & " and a0k15<=" & strDate & " "
'
'   End If
'
'   'Modify by Amy 2024/06/03 原程式改至共用函數
'   strExc(0) = GetSql(0, True, strCon)
'   intI = 1
'   Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
'   dblSum = 0
'   If intI = 1 Then
'      If Not IsNull(rsReport.Fields(0)) Then dblSum = rsReport.Fields(0)
'   End If
'
'   'Modify by Amy 2024/06/03 原程式改至共用函數
'   strExc(0) = GetSql(0, False, strCon)
'   intI = 1
'   iPageNo = 0 ': iItemCount = 1'Modify by Amy 2013/07/04
'   Set rsReport = ClsLawReadRstMsg(intI, strExc(0))
'
''Add & Modify By Amy 20130/07/04 增加列印扣繳年度改回收款年度清單
''   If intI = 1 Then
''      PUB_RestorePrinter Combo1
''      With rsReport
'''         Printer.EndDoc
''         Printer.Orientation = 1 '1.直印 2.橫印
''         Printer.Font.Name = "細明體"
''         GetPleft
''         PrintHead
''         Do While Not .EOF
''            If iItemCount > 46 Then
''               Printer.NewPage
''               PrintHead
''               iItemCount = 1
''            End If
''
''            For i = 0 To 9
''               If i = 8 Then
''                  Printer.CurrentX = PLeft(i) - Printer.TextWidth(Format(.Fields(i), DDollar2))
''               Else
''                  Printer.CurrentX = PLeft(i)
''               End If
''               Printer.CurrentY = intY
''               If i = 1 Then
''                  Printer.Print ChangeTStringToTDateString(.Fields(i))
''               ElseIf i = 8 Then
''                  Printer.Print Format(.Fields(i), DDollar2)
''               Else
''                  Printer.Print "" & .Fields(i)
''               End If
''            Next i
''            intY = intY + 300
''            iItemCount = iItemCount + 1
''            .MoveNext
''         Loop
''         Printer.CurrentX = PLeft(0)
''         Printer.CurrentY = intY
''         Printer.Print String(110, "-")
''         intY = intY + 300
''         Printer.CurrentX = PLeft(7)
''         Printer.CurrentY = intY
''         Printer.Print "服務費合計"
''         Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(dblSum, DDollar2))
''         Printer.CurrentY = intY
''         Printer.Print Format(dblSum, DDollar2)
''         Printer.EndDoc
''         Screen.MousePointer = vbDefault
''         Call ShowPrintOk
''         FormClear
''      End With
''
''      PUB_RestorePrinter strPrinter
''
''   Else
''      MsgBox "無資料可供列印！"
''      Screen.MousePointer = vbDefault
''   End If
'
'   'Modify by Amy 2024/06/03 原程式改至共用函數
'   strExc(0) = GetSql(1, True, strCon)
'   intR = 1
'   Set rsReport2 = ClsLawReadRstMsg(intR, strExc(0))
'   dblSum2 = 0
'   If intR = 1 Then
'      If Not IsNull(rsReport2.Fields(0)) Then dblSum2 = rsReport2.Fields(0)
'   End If
'
'   'Modify by Amy 2024/06/03 原程式改至共用函數
'   strExc(0) = GetSql(1, False, strCon)
'   intR = 1
'   Set rsReport2 = ClsLawReadRstMsg(intR, strExc(0))
'
'   If intI = 1 Or intR = 1 Then
'        PUB_RestorePrinter Combo1
'        Printer.Orientation = 1 '1.直印 2.橫印
'        Printer.Font.Name = "細明體"
'        GetPleft
'        If intI = 1 Then
'            PrintHead 1
'            PrintData rsReport, 1, dblSum
'
'        End If
'        If intR = 1 Then
'           PrintHead 2
'           PrintData rsReport2, 2, dblSum2
'        End If
'
'        Screen.MousePointer = vbDefault
'        Call ShowPrintOk
'        FormClear
'   Else
'        MsgBox "無資料可供列印！"
'        Screen.MousePointer = vbDefault
'   End If
'   Set rsReport = Nothing
'   Set rsReport2 = Nothing
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1(0).Text = ""
   Text1(1).Text = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   Text1(0).SetFocus
   Option1(1).Value = True 'Add by Amy 2013/06/20
End Sub

Private Sub Form_Load()
   PUB_InitForm Me, Me.Width, Me.Height
   Me.Width = 5805
   'Modify by Amy 2024/06/03 原:3825
   Me.Height = 3012
   
   'Add by Amy 2013/06/20 由Command1_Click搬過來
   'Modify by Amy 2025/01/06 不預帶-秀玲
   '系統日前一天
   'strDate = Val(CompDate(2, -1, strSrvDate(1))) - 19110000 '日曆天
   'strDate = Val(CompWorkDay(1, strSrvDate(1), 1)) - 19110000 '工作天
   'End 2013/06/20
   
   'MaskEdBox1.Text = CFDate(strSrvDate(2))
   MaskEdBox1.Mask = DFormat
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add by Morgan 2011/7/8
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(151)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   Set Frmacc14m0 = Nothing
End Sub

'Add by Amy 2013/06/20
Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            'Modify by Amy 2025/01/06 改不預帶 原:選擇第一次列印，上次列印日期預帶系統日前一日曆天(因放年假,預帶錯)
            'MaskEdBox1.Text = ChangeTStringToTDateString(strDate)
            MaskEdBox1.Mask = ""
            'Modify by Amy 2025/02/25 瑞婷選第1次未輸上次列印日,會錯(因以為一定會輸入期)-秀玲:選第一次列印時,預帶前一個工作天
            '                                                     若當天修改當天報表也要出現,User 要自行修改日期
            'MaskEdBox1.Text = ""
            MaskEdBox1.Text = ChangeTStringToTDateString(Val(CompWorkDay(2, strSrvDate(1), 1)) - 19110000) '前1個工作天
            'end 2025/02/25
            'end 2025/01/06
            MaskEdBox1.Mask = DFormat
        Case 1
            MaskEdBox1.Mask = ""
            MaskEdBox1.Text = ""
            MaskEdBox1.Mask = DFormat
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'*************************************************
'  畫面輸入檢查
'*************************************************
Private Function FormCheck() As Boolean
   If Trim(Text1(0)) = "" Then
      MsgBox "收款年度不可空白！", vbCritical
      FormCheck = False
      Text1(0).SetFocus
      Exit Function
   End If
   If Trim(Text1(1)) = "" Then
      MsgBox "扣繳年度不可空白！", vbCritical
      FormCheck = False
      Text1(1).SetFocus
      Exit Function
   End If
   'Modify by Amy 2025/02/25 原第一次印會預帶日期,改為不預帶時未輸上次列印日期 程式會Error,故改為日期必輸
   'If Option1(1).Value = True And MaskEdBox1.Text = MsgText(29) Then
   If MaskEdBox1.Text = MsgText(29) Then
      MsgBox "列印範圍為第二次以上時，上次列印日期不可空白！", vbCritical
      FormCheck = False
      Exit Function
   End If
   FormCheck = True
End Function

Private Sub GetPleft()
   ReDim PLeft(9) As Integer
   
   '公司別
   PLeft(0) = 300
   '收款日
   PLeft(1) = 600
   '傳票號碼
   PLeft(2) = 1700
   '本所案號
   PLeft(3) = 2900
   '智權人員
   PLeft(4) = 4300
   '收據抬頭
   PLeft(5) = 5100
   '申請國家
   PLeft(6) = 6900
   '案件性質
   PLeft(7) = 7900
   '服務費
   PLeft(8) = 10000
   '收據號碼
   PLeft(9) = 10200
End Sub

Private Sub PrintHead(ByVal h As Integer) 'Modify By Amy 2013/07/04 +H-1:收款扣款改年度清單/2:扣繳年度改回清單
'Mark by Amy 不使用
'   iPageNo = iPageNo + 1
'
'   intY = 300
'
'   If h = 1 Then 'Add by Amy 2013/07/04
'     If Val(Text1(0)) < Val(Text1(1)) Then
'        strExc(1) = Text1(0) & "年收款轉" & Text1(1) & "年扣繳（9101及9102）"
'     Else
'        strExc(1) = Text1(0) & "年收款轉" & Text1(1) & "年扣繳（9201及9202）"
'     End If
'   Else
'    strExc(1) = "扣繳年度改回收款年度清單"
'   End If
'   Printer.FontSize = 14
'   Printer.FontBold = True
'   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strExc(1)) / 2)
'   Printer.CurrentY = intY
'   Printer.Print strExc(1)
'
'   intY = intY + 400
'   'Add by Amy 2013/07/23 扣繳年度改回清單 顯示文字修改
'   If h = 1 Then
'        If Val(Text1(0)) < Val(Text1(1)) Then
'            strExc(1) = "轉次年：注意若已產生規費借方傳票，請刪除該案號於隔日傳票的借方規費項次；" & vbCrLf & _
'                             "　　　　　　另請自行執行複委託及翻譯費之收款資料檢查並修改摘要的收款日期"
'        Else
'            strExc(1) = "轉前一年：注意要改收款年度第一個工作天的規費借方金額＝前一年度最後一天的規費貸方金額；" & vbCrLf & _
'                          "　　　　　　　另請自行執行複委託及翻譯費之收款資料檢查並修改摘要的收款日期"
'        End If
'   Else
'        strExc(1) = "注意若已產生規費借方傳票，請檢查該案號的借方規費；" & vbCrLf & _
'                         "   另請自行執行複委託及翻譯費之收款資料檢查並修改摘要的收款日期"
'   End If
'   'end 2013/07/23
'
'   Printer.FontSize = 10
'   Printer.FontBold = False
'   Printer.CurrentX = PLeft(0)
'   Printer.CurrentY = intY
'   Printer.Print strExc(1)
'   If Option1(1).Value = True Then
'      Printer.CurrentX = 9000
'      Printer.CurrentY = intY
'      Printer.Print MaskEdBox1.Text & " 至 " & ChangeTStringToTDateString(strDate)
'   Else
'      Printer.CurrentX = PLeft(9)
'      Printer.CurrentY = intY
'      Printer.Print ChangeTStringToTDateString(strDate)
'   End If
'
'   intY = intY + 600
'
'   Printer.CurrentX = PLeft(0)
'   Printer.CurrentY = intY
'   Printer.Print ""
'
'   Printer.CurrentX = PLeft(1)
'   Printer.CurrentY = intY
'   Printer.Print "收款日"
'
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = intY
'   Printer.Print "傳票號碼"
'
'   Printer.CurrentX = PLeft(3)
'   Printer.CurrentY = intY
'   Printer.Print "本所案號"
'
'   Printer.CurrentX = PLeft(4)
'   Printer.CurrentY = intY
'   Printer.Print "智權人員"
'
'   Printer.CurrentX = PLeft(5)
'   Printer.CurrentY = intY
'   Printer.Print "收據抬頭"
'
'   Printer.CurrentX = PLeft(6)
'   Printer.CurrentY = intY
'   Printer.Print "申請國家"
'
'   Printer.CurrentX = PLeft(7)
'   Printer.CurrentY = intY
'   Printer.Print "案件性質"
'
'   Printer.CurrentX = PLeft(8) - Printer.TextWidth("服務費")
'   Printer.CurrentY = intY
'   Printer.Print "服務費"
'
'   Printer.CurrentX = PLeft(9)
'   Printer.CurrentY = intY
'   Printer.Print "收據號碼"
'
'   intY = intY + 300
'   Printer.CurrentX = PLeft(0)
'   Printer.CurrentY = intY
'   Printer.Print String(110, "-")
'   intY = intY + 300
End Sub

'Add by Amy 2013/07/04
Public Sub PrintData(RsTemp As ADODB.Recordset, ByVal h As Integer, ByVal dblSum As Double)
'Mark by Amy 不使用
'Dim iItemCount As Integer, i As Integer
'iItemCount = 1
'    With RsTemp
'         Do While Not .EOF
'            If iItemCount > 46 Then
'               Printer.NewPage
'               PrintHead h
'               iItemCount = 1
'            End If
'
'            For i = 0 To 9
'               If i = 8 Then
'                  Printer.CurrentX = PLeft(i) - Printer.TextWidth(Format(.Fields(i), DDollar2))
'               Else
'                  Printer.CurrentX = PLeft(i)
'               End If
'               Printer.CurrentY = intY
'               If i = 1 Then
'                  Printer.Print ChangeTStringToTDateString(.Fields(i))
'               ElseIf i = 8 Then
'                  Printer.Print Format(.Fields(i), DDollar2)
'               Else
'                  Printer.Print "" & .Fields(i)
'               End If
'            Next i
'            intY = intY + 300
'            iItemCount = iItemCount + 1
'            .MoveNext
'         Loop
'         Printer.CurrentX = PLeft(0)
'         Printer.CurrentY = intY
'         Printer.Print String(110, "-")
'         intY = intY + 300
'         Printer.CurrentX = PLeft(7)
'         Printer.CurrentY = intY
'         Printer.Print "服務費合計"
'         Printer.CurrentX = PLeft(8) - Printer.TextWidth(Format(dblSum, DDollar2))
'         Printer.CurrentY = intY
'         Printer.Print Format(dblSum, DDollar2)
'         Printer.EndDoc
'
'    End With
End Sub

'Add by Amy 2024/06/03 取得資料語法
'intChoose:0-收款轉前一年列表 / 1-扣繳年度改回收款年度清單
Private Function GetSql(intChoose As Integer, bolSum As Boolean, stCon As String) As String
   '收款轉前一年列表
   If intChoose = 0 Then
      If bolSum = True Then
         '服務費合計
         'Modify by Amy 2016/03/10 因收款年度104 扣繳年度105 上次列印105/02/05 F10413475會抓到a1p01='J'的資料,J公司不需扣繳,故排除
         'Modify by Amy 2024/04/10 原:And a1p01='1',a0k11為L時,傳票傳號要抓a1p01='L'的a1p22
         'Modify by Amy 2024/06/03 原:SUBSTR(A0K04,1,8) 收據抬頭 ,改Excel可不用限制字數
         GetSql = "Select SUM(服務費) From ( " & _
                     "Select A0K11 公司別, A0L02 收款日, A1P22 傳票號碼, A0J02 本所案號, ST02 智權人員, A0K04 收據抬頭, " & _
                     "SUBSTR(na03,1,4) 申請國家, SUBSTR(cpm03,1,4) 案件性質, DECODE(A0J07,'Y',A1U04+A1U05,A1U04) 服務費, A0M02 收據號碼 " & _
                     "From (Select DISTINCT A0K11, A0L02, A1P22, A0J02, ST02, A0K04,na03, decode(a0j04,'000',cpm03,cpm04) cpm03, A0M02,A0M01,A0J13,A0J01,A0J07 " & _
                     "From ACC0K0, ACC0L0, Acc1p0, STAFF, ACC0J0, ACC0M0, Nation, casepropertymap, caseprogress " & _
                     "Where substr((A0L02+19110000),1,4)-1911=" & Text1(0) & " And A0M01=A0L01(+) And A0M07=" & Text1(1) & " And A0M02=A0K01(+) And A0M01=A1P04(+) " & _
                     "And A0M10=ST01(+) And A0M02=A0J13(+) And DECODE(A0J07,'Y',A0J09+A0J10,A0J09)>0 " & _
                     "And a0j04=na01(+) And a0j01=cp09(+) And cp01=cpm01(+) And cp10=cpm02(+) And a1p01=Decode(a0k11, 'L', 'L', '1')" & stCon & _
                     ") B,ACC1U0 Where B.A0M01=A1U01(+) And B.A0J13=A1U02(+) And B.A0J01=A1U03(+) And A1U03<>A1U01) "
      Else
         'And A1P05>='4' And A1P05<='5' 不可限制收入科目否則只收規費的傳票會抓不到
         'Modify by Amy 2016/03/10 因收款年度104 扣繳年度105 上次列印105/02/05 F10413475會抓到a1p01='J'的資料,J公司不需扣繳,故排除
         'Modify by Amy 2024/04/10 原:And a1p01='1',a0k11為L時,傳票傳號要抓a1p01='L'的a1p22
         'Modify by Amy 2024/06/03 原:SUBSTR(A0K04,1,8) 收據抬頭 ,改Excel可不用限制字數
         GetSql = "Select A0K11 公司別, A0L02 收款日, A1P22 傳票號碼, A0J02 本所案號, ST02 智權人員, A0K04 收據抬頭, " & _
                     "SUBSTR(na03,1,4) 申請國家, SUBSTR(cpm03,1,4) 案件性質, DECODE(A0J07,'Y',A1U04+A1U05,A1U04) 服務費, A0M02 收據號碼 " & _
                     "From (Select DISTINCT A0K11, A0L02, A1P22, A0J02, ST02, A0K04,na03, decode(a0j04,'000',cpm03,cpm04) cpm03, A0M02,A0M01,A0J13,A0J01,A0J07 " & _
                     "From ACC0K0, ACC0L0, Acc1p0, STAFF, ACC0J0, ACC0M0, Nation, casepropertymap, caseprogress " & _
                     "Where substr((A0L02+19110000),1,4)-1911=" & Text1(0) & " And A0M01=A0L01(+) And A0M07=" & Text1(1) & " And A0M02=A0K01(+) And A0M01=A1P04(+) " & _
                     "And A0M10=ST01(+) And A0M02=A0J13(+) And DECODE(A0J07,'Y',A0J09+A0J10,A0J09)>0 " & _
                     "And a0j04=na01(+) And a0j01=cp09(+) And cp01=cpm01(+) And cp10=cpm02(+) And a1p01=Decode(a0k11, 'L', 'L', '1')" & stCon & _
                     ") B,ACC1U0 Where B.A0M01=A1U01(+) And B.A0J13=A1U02(+) And B.A0J01=A1U03(+) And A1U03<>A1U01 "
      End If
   '扣繳年度改回收款年度清單
   Else
      If bolSum = True Then
         '扣繳年度改回收款年度 服務費合計
         'Modify by Amy 2016/03/10 因收款年度104 扣繳年度105 上次列印105/02/05 F10413475會抓到a1p01='J'的資料,J公司不需扣繳,故排除
         'Modify by Amy 2024/04/10 原:And a1p01='1',a0k11為L時,傳票傳號要抓a1p01='L'的a1p22
         'Modify by Amy 2024/06/03 原:SUBSTR(A0K04,1,8) 收據抬頭 ,改Excel可不用限制字數
         GetSql = "Select SUM(服務費) From ( " & _
                        "Select A0K11 公司別, A0L02 收款日, A1P22 傳票號碼, A0J02 本所案號, ST02 智權人員, A0K04 收據抬頭, " & _
                        "SUBSTR(na03,1,4) 申請國家, SUBSTR(cpm03,1,4) 案件性質, DECODE(A0J07,'Y',A1U04+A1U05,A1U04) 服務費, A0M02 收據號碼 " & _
                        "From (Select DISTINCT A0K11, A0L02, A1P22, A0J02, ST02, A0K04,na03, decode(a0j04,'000',cpm03,cpm04) cpm03, A0M02,A0M01,A0J13,A0J01,A0J07 " & _
                        "From ACC0K0, ACC0L0, Acc1p0, STAFF, ACC0J0, ACC0M0, Nation, casepropertymap, caseprogress " & _
                        "Where substr(A0L02,1,3)='" & Text1(0) & "' And substr((A0L02+19110000),1,4)-1911=A0K16 And A0M01=A0L01(+) And A0M02=A0K01(+) And A0M01=A1P04(+) " & _
                        "And A0M10=ST01(+) And A0M02=A0J13(+) And DECODE(A0J07,'Y',A0J09+A0J10,A0J09)>0 " & _
                        "And a0j04=na01(+) And a0j01=cp09(+) And cp01=cpm01(+) And cp10=cpm02(+) And a1p01=Decode(a0k11, 'L', 'L', '1')" & stCon & _
                        ") B,ACC1U0 Where B.A0M01=A1U01(+) And B.A0J13=A1U02(+) And B.A0J01=A1U03(+) And A1U03<>A1U01) "
      Else
         'Modify by Amy 2016/03/10 因收款年度104 扣繳年度105 上次列印105/02/05 F10413475會抓到a1p01='J'的資料,J公司不需扣繳,故排除
         'Modify by Amy 2024/04/10 原:And a1p01='1',a0k11為L時,傳票傳號要抓a1p01='L'的a1p22
         'Modify by Amy 2024/06/03 原:SUBSTR(A0K04,1,8) 收據抬頭 ,改Excel可不用限制字數
         GetSql = "Select A0K11 公司別, A0L02 收款日, A1P22 傳票號碼, A0J02 本所案號, ST02 智權人員, A0K04 收據抬頭, " & _
                     "SUBSTR(na03,1,4) 申請國家, SUBSTR(cpm03,1,4) 案件性質, DECODE(A0J07,'Y',A1U04+A1U05,A1U04) 服務費, A0M02 收據號碼 " & _
                     "From (Select DISTINCT A0K11, A0L02, A1P22, A0J02, ST02, A0K04,na03, decode(a0j04,'000',cpm03,cpm04) cpm03, A0M02,A0M01,A0J13,A0J01,A0J07 " & _
                     "From ACC0K0, ACC0L0, Acc1p0, STAFF, ACC0J0, ACC0M0, Nation, casepropertymap, caseprogress " & _
                     "Where substr(A0L02,1,3)='" & Text1(0) & "' And substr((A0L02+19110000),1,4)-1911=A0K16 And A0M01=A0L01(+) And A0M02=A0K01(+) And A0M01=A1P04(+) " & _
                     "And A0M10=ST01(+) And A0M02=A0J13(+) And DECODE(A0J07,'Y',A0J09+A0J10,A0J09)>0 " & _
                     "And a0j04=na01(+) And a0j01=cp09(+) And cp01=cpm01(+) And cp10=cpm02(+) And a1p01=Decode(a0k11, 'L', 'L', '1')" & stCon & _
                     ") B,ACC1U0 Where B.A0M01=A1U01(+) And B.A0J13=A1U02(+) And B.A0J01=A1U03(+) And A1U03<>A1U01 "
      End If
   End If
End Function

Private Function SaveExcel(intChoose As Integer, RsQ As ADODB.Recordset) As Boolean
   Dim xlsAp As New Excel.Application, wksrpt As New Worksheet, i As Integer, intR As Integer, intTitleR As Integer, bolOpenXls As Boolean
   Dim strAllF As String, strAllW As String, strFileN As String, strFormat As String, strTmp As String

On Error GoTo ErrHnd
   '收款轉前一年列表
   If intChoose = 0 Then
      strFileN = Text1(0) & "年收款轉" & Text1(1) & "年扣繳列表"
   '扣繳年度改回收款年度清單
   Else
      strFileN = "扣繳年度改回收款年度清單"
   End If
   
   If Dir(strExcelPath & strFileN & ACDate(ServerDate) & ServerTime & ".xlsx") = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strExcelPath & strFileN & ACDate(ServerDate) & ServerTime & ".xlsx"
   End If
   
   strAllF = "公司別,收款日,傳票號碼,本所案號,智權人員,收據抬頭,申請國家,案件性質,服務費,收據號碼"
   strAllW = "1.5, 8, 10, 12.5, 8.5, 11.56, 8.56, 10.56, 9, 9.22"
   strF = Split(strAllF, ",")
   arrWidth = Split(strAllW, ",")
   
   intField = 65: intR = 1
   xlsAp.SheetsInNewWorkbook = 3
   xlsAp.Workbooks.add
   bolOpenXls = True
   Set wksrpt = xlsAp.Worksheets(1)
   Call SetTitle(intChoose, xlsAp, wksrpt, intR)
   intTitleR = intR
   intR = intR + 1
   
   Do While RsQ.EOF = False
      For i = LBound(strF) To UBound(strF)
         strFormat = ""
         strTmp = RsQ.Fields(strF(i))
         
         Select Case strF(i)
            Case "公司別"
               strFormat = "@"
            Case "收款日"
               strTmp = Format(strTmp, "###/##/##")
            Case "服務費"
               strFormat = "#,##0"
         End Select
        
         wksrpt.Range(Chr(i + intField) & intR).Value = strTmp
         If strFormat <> MsgText(601) Then
            wksrpt.Range(Chr(i + intField) & intR).NumberFormatLocal = strFormat
         End If
      Next i
      intR = intR + 1
      RsQ.MoveNext
   Loop
   '畫線
   wksrpt.Range(Chr(intField) & intR - 1 & ":" & Chr(intField + UBound(strF)) & intR - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
   '公司別
   wksrpt.Range(Chr(GetColVal(strF, "公司別", LBound(strF)) + intField) & intTitleR).Value = ""
   '合計
   wksrpt.Range(Chr(GetColVal(strF, "案件性質", LBound(strF)) + intField) & intR).Value = "服務費合計"
   wksrpt.Range(Chr(GetColVal(strF, "案件性質", LBound(strF)) + intField) & intR).Font.Bold = True
   wksrpt.Range(Chr(GetColVal(strF, "案件性質", LBound(strF)) + intField) & intR).HorizontalAlignment = xlRight
   strTmp = "=Sum(" & Chr(GetColVal(strF, "服務費", LBound(strF)) + intField) & intTitleR + 1 & ":" & Chr(GetColVal(strF, "服務費", LBound(strF)) + intField) & intR - 1 & ")"
   wksrpt.Range(Chr(GetColVal(strF, "服務費", LBound(strF)) + intField) & intR).Value = strTmp
   strTmp = wksrpt.Range(Chr(GetColVal(strF, "服務費", LBound(strF)) + intField) & intR).Value
   If Val(strSum(intChoose)) <> Val(strTmp) Then
      MsgBox strFileN & " 合計有誤,請洽電腦中心！"
   End If
   '內容字大小
   wksrpt.Range(Chr(intField) & intTitleR & ":" & Chr(intField + UBound(strF)) & intR).Font.Size = 11
   '版面設定
   wksrpt.PageSetup.Orientation = xlPortrait '直印
   wksrpt.PageSetup.Zoom = 100 '縮放比例為100%,列印頁面水平置中
   wksrpt.PageSetup.HeaderMargin = xlsAp.Application.InchesToPoints(0) '頁首
   wksrpt.PageSetup.FooterMargin = xlsAp.Application.InchesToPoints(0) '頁尾
   wksrpt.PageSetup.TopMargin = xlsAp.InchesToPoints(0.5) '上
   wksrpt.PageSetup.BottomMargin = xlsAp.InchesToPoints(0.5) '下
   wksrpt.PageSetup.LeftMargin = xlsAp.InchesToPoints(0.3) '左邊界
   wksrpt.PageSetup.RightMargin = xlsAp.InchesToPoints(0.3) '右邊界
   wksrpt.PageSetup.PrintTitleRows = "$1:$" & intTitleR '表頭保留列
   wksrpt.PageSetup.CenterHorizontally = True '水平置中(版面設定->邊界->水平置中)
   
   If Val(xlsAp.Version) < 12 Then
      xlsAp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
   Else
      xlsAp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN & ACDate(ServerDate) & ServerTime & ".xlsx", FileFormat:=51
   End If
   xlsAp.Workbooks.Close
   xlsAp.Quit
   Set xlsAp = Nothing
    
   SaveExcel = True
   Exit Function
    
ErrHnd:
   If bolOpenXls = True Then
      If Val(xlsAp.Version) < 12 Then
         xlsAp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN & ACDate(ServerDate) & ServerTime & MsgText(43), FileFormat:=-4143
      Else
         xlsAp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN & ACDate(ServerDate) & ServerTime & ".xlsx", FileFormat:=51
      End If
       xlsAp.Workbooks.Close
       xlsAp.Quit
       Set xlsAp = Nothing
    End If
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function
    
Private Sub SetTitle(intChoose As Integer, xlsApp As Excel.Application, Wks As Worksheet, ByRef intRow As Integer)
   Dim ii As Integer, stTitleN As String, stTxt As String, stDateTxt As String
   
   '收款轉前一年列表
   If intChoose = 0 Then
      If Val(Text1(0)) < Val(Text1(1)) Then
         stTitleN = "（9101及9102）"
         stTxt = "轉次年：注意若已產生規費借方傳票，請刪除該案號於隔日傳票的借方規費項次；"
      Else
         stTitleN = "（9201及9202）"
         stTxt = "轉前一年：注意要改收款年度第一個工作天的規費借方金額＝前一年度最後一天的規費貸方金額；"
      End If
      stTitleN = Text1(0) & "年收款轉" & Text1(1) & "年扣繳" & stTitleN
      stTxt = stTxt & vbCrLf & _
                  "另請自行執行複委託及翻譯費之收款資料檢查並修改摘要的收款日期"
   '扣繳年度改回收款年度清單
   Else
      stTitleN = "扣繳年度改回收款年度清單"
      stTxt = "注意若已產生規費借方傳票，請檢查該案號的借方規費；" & vbCrLf & _
                     "另請自行執行複委託及翻譯費之收款資料檢查並修改摘要的收款日期"
   End If
   If Option1(1).Value = True Then
      stDateTxt = MaskEdBox1.Text & " 至 " & ChangeTStringToTDateString(strDate)
   Else
      stDateTxt = ChangeTStringToTDateString(strDate)
   End If
   Wks.Range(Chr(intField) & intRow).Value = stTitleN
   Wks.Range(Chr(intField) & intRow).Font.Size = 18
   Wks.Range(Chr(intField) & intRow).Font.Bold = True
   Wks.Range(Chr(intField) & intRow & ":" & Chr(UBound(strF) + intField) & intRow).Select
  
   With xlsApp.Selection
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlBottom
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .ShrinkToFit = False
      .MergeCells = True
   End With
   intRow = intRow + 1
   
   '說明文字
   Wks.Range(Chr(intField) & intRow).Value = stTxt
   Wks.Range(Chr(intField) & intRow).RowHeight = 28
   Wks.Range(Chr(intField) & intRow & ":" & Chr(GetColVal(strF, "案件性質", LBound(strF)) + intField) & intRow).MergeCells = True
   '日期
   Wks.Range(Chr(GetColVal(strF, "服務費", LBound(strF)) + intField) & intRow).Value = stDateTxt
   Wks.Range(Chr(GetColVal(strF, "服務費", LBound(strF)) + intField) & intRow).VerticalAlignment = xlTop '靠上
   Wks.Range(Chr(GetColVal(strF, "服務費", LBound(strF)) + intField) & intRow & ":" & Chr(UBound(strF) + intField) & intRow).MergeCells = True
   
   Wks.Range(Chr(intField) & intRow & ":" & Chr(UBound(strF) + intField) & intRow).Font.Name = "微軟正黑體"
   Wks.Range(Chr(intField) & intRow & ":" & Chr(UBound(strF) + intField) & intRow).Font.Size = 9
   Wks.Range(Chr(intField) & intRow & ":" & Chr(UBound(strF) + intField) & intRow).Font.Color = vbBlue
   Wks.Range(Chr(intField) & intRow & ":" & Chr(UBound(strF) + intField) & intRow).Font.Bold = False
   intRow = intRow + 1
   
   For ii = LBound(strF) To UBound(strF)
      Wks.Range(Chr(intField + ii) & intRow).Value = strF(ii)
      Wks.Range(Chr(intField + ii) & intRow).Font.Bold = True
      Wks.Range(Chr(intField + ii) & intRow).ColumnWidth = Val(arrWidth(ii))
      Wks.Range(Chr(intField + ii) & intRow).HorizontalAlignment = xlCenter
   Next ii
   '畫線
   Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strF)) & intRow).Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

