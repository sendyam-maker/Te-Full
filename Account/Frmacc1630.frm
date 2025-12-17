VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc1630 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "銷貨退回折讓單列印"
   ClientHeight    =   5325
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢(&Q)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   5700
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   810
      Width           =   1365
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   3870
      MaxLength       =   12
      TabIndex        =   5
      Top             =   900
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1635
      Style           =   2  '單純下拉式
      TabIndex        =   9
      Top             =   4980
      Width           =   3450
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3885
      MaxLength       =   15
      TabIndex        =   4
      Top             =   540
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "單筆"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   570
      Width           =   1125
   End
   Begin VB.OptionButton Option1 
      Caption         =   "整批"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   150
      Value           =   -1  'True
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   7170
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   810
      Width           =   1365
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   3885
      TabIndex        =   1
      Top             =   90
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   315
      Left            =   5805
      TabIndex        =   2
      Top             =   90
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3615
      Left            =   60
      TabIndex        =   8
      Top             =   1290
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|銷退單號|銷退日期|發票編號|折讓單列印日期"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請至盟立平台"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   1
      Left            =   690
      TabIndex        =   16
      Top             =   840
      Width           =   1700
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "注意：列印時不要使用Word！"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   5160
      TabIndex        =   15
      Top             =   5040
      Width           =   2925
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   480
      X2              =   8520
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "本 所 案 號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   14
      Top             =   930
      Width           =   1380
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   765
      TabIndex        =   13
      Top             =   5040
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "銷 退 單 號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   12
      Top             =   570
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   135
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5565
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "銷退／轉開日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2190
      TabIndex        =   10
      Top             =   120
      Width           =   1710
   End
End
Attribute VB_Name = "Frmacc1630"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2014/4/7
Option Explicit

Public adoacc0k0 As New ADODB.Recordset
Dim strPrinter As String
Dim m_TempFileName As String, m_FileName As String
Dim m_A0S01 As String
Dim m_A0S03 As String
Dim m_A4302 As String
Dim m_Number As String
Dim m_A4602 As String
Dim m_Money As String
Dim m_TaxFee As String
Dim m_A0K04 As String
Dim m_A0K04ID As String
Dim m_Fillopen As Boolean


Private Sub cmdok_Click(Index As Integer)
Dim bolHaveData As Boolean

   Select Case Index
      Case 0 '列印
         Call PrintData
      Case 1 '查詢
         Call QueryData
   End Select
End Sub

Private Sub CallWordPrint()
Dim i As Integer, jj As Integer
Dim strName As String
Dim strText As String
Dim intRunTime As Integer
   
On Error GoTo ErrHand
   
   '判斷word是否已開啟
   If g_WordAp Is Nothing Then
RestarWord:
      Set g_WordAp = New Word.Application
      g_WordAp.Visible = False
   End If
   m_TempFileName = "$$銷退折讓證明單_temp.doc"
   If Dir(App.path & "\" & m_TempFileName) <> "" Then
      Kill App.path & "\" & m_TempFileName
   End If
   g_WordAp.Documents.Open App.path & "\" & m_FileName
   g_WordAp.ActiveDocument.SaveAs App.path & "\" & m_TempFileName
   g_WordAp.ActiveDocument.Close
   g_WordAp.Documents.Open App.path & "\" & m_TempFileName
   With g_WordAp
      For i = 1 To 11
         strName = ""
         strText = ""
         intRunTime = 0
         If i = 1 Then
            strName = "銷退日期"
            strText = Left(m_A0S03, 3) & "年" & Mid(m_A0S03, 5, 2) & "月" & Mid(m_A0S03, 8, 2) & "日" & IIf(m_Fillopen = True, " ***  補開  ***", "")
            intRunTime = 4
         ElseIf i = 2 Then
            strName = "Yr"
            strText = Left(m_A4302, 3)
            intRunTime = 4
         ElseIf i = 3 Then
            strName = "M"
            strText = Mid(m_A4302, 5, 2)
            intRunTime = 4
         ElseIf i = 4 Then
            strName = "Dy"
            strText = Mid(m_A4302, 8, 2)
            intRunTime = 4
         ElseIf i = 5 Then
            strName = "ZZ"
            strText = Left(m_Number, 2)
            intRunTime = 4
         ElseIf i = 6 Then
            strName = "號碼"
            strText = Mid(m_Number, 3)
            intRunTime = 4
         ElseIf i = 7 Then
            strName = "品名"
            strText = StrToStr(CheckStr(m_A4602), 12)
            intRunTime = 4
         ElseIf i = 8 Then
            strName = "金額"
            strText = Format(m_Money, DDollar)
            intRunTime = 8
         ElseIf i = 9 Then
            strName = "稅"
            strText = Format(m_TaxFee, DDollar)
            intRunTime = 8
         ElseIf i = 10 Then
            strName = "客戶名稱"
            'strText = StrToStr(CheckStr(m_A0K04), 12)
            strText = CheckStr(m_A0K04)
            intRunTime = 4
         ElseIf i = 11 Then
            strName = "客戶統一編號"
            strText = m_A0K04ID
            intRunTime = 4
         End If
         If Trim(strName) <> "" Then
            For jj = 1 To intRunTime
               .Selection.WholeStory
               .Selection.Copy
               .Selection.Find.ClearFormatting
               If i = 2 Or i = 3 Or i = 4 Or i = 5 Then
                  .Selection.Find.Text = strName
               Else
                  .Selection.Find.Text = "|#" & strName & "#|"
               End If
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
               .Selection.Delete
               .Selection.Font.ColorIndex = wdBlack
               .Selection.TypeText strText
            Next jj
         End If
      Next i
   End With
   '列印整份 (Range:=wdPrintRangeOfPages, Pages:="1")
   g_WordAp.PrintOut FileName:="", Range:=wdPrintAllDocument, Item:=wdPrintDocumentContent, _
                     Copies:=1, Pages:="", PageType:=wdPrintAllPages, _
                     ManualDuplexPrint:=False, Collate:=True, Background:=False, PrintToFile:=False
   g_WordAp.ActiveDocument.Close
   If m_Fillopen = False Then
      If m_A0S01 <> "" Then
         adoTaie.Execute "update acc0s0 set a0s25=sysdate where a0s01='" & m_A0S01 & "'"
      ElseIf m_Number <> "" Then
         adoTaie.Execute "update acc430 set a4318=sysdate where a4301='" & m_Number & "'"
      End If
   End If
   
   g_WordAp.Quit
   Set g_WordAp = Nothing
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   ElseIf Err.Number <> 0 Then
      MsgBox (Err.Description)
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   '改單線固定(調整大小不用再設定)
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
      For intY = 0 To Int(ScaleHeight / sglHeight)
         PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
      Next
   Next
   
   'Add by Amy 2019/12/17
   Label1(1).Caption = "請至盟立平台" & vbCrLf & _
                                    "列印折讓單"
   
   '預設上一個工作日至當天
   MaskEdBox1.Text = CFDate(TransDate(PUB_GetWorkDay1(strSrvDate(1) - 1, 1), 1))
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Text = CFDate(strSrvDate(2))
   MaskEdBox2.Mask = DFormat
   If Option1(0).Value = True Then Call Option1_Click(0)
   If Option1(1).Value = True Then Call Option1_Click(1)
   
   SetGrd
   Frmacc0000.StatusBar1.Panels(1).Text = ""
   
   m_FileName = "$$銷退折讓證明單.doc"
   If Dir(App.path & "\" & m_FileName) <> "" Then
      Kill App.path & "\" & m_FileName
   End If
   Call PUB_GetSampleFile(m_FileName, "M31-000002-0-00")
   PUB_SetPrinter Me.Name, Combo1, strPrinter
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHand
   
   If Not g_WordAp Is Nothing Then
      g_WordAp.Quit
CloseWord:
      Set g_WordAp = Nothing
   End If
   
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc1630 = Nothing
   
   Exit Sub
   
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo CloseWord
   ElseIf Err.Number <> 0 Then
      MsgBox (Err.Description)
   End If
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Option1(0).Value = True
   Option1(1).Value = False
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   Text1(0).Text = ""
'   Text1(1).Text = ""
   Text1(2).Text = ""
End Sub

Private Sub grd1_SelChange()
Dim i As Integer

grd1.Visible = False
If grd1.row <> 0 Then
   grd1.col = 0
'   GRD1.row = GRD1.MouseRow
   If grd1.TextMatrix(grd1.row, 3) <> "" Then
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
End If
grd1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow grd1, x, y, nCol, nRow
grd1.col = nCol
grd1.row = nRow
End Sub

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
   If MaskEdBox2.Text = "" Then
      MaskEdBox2.Mask = MsgText(601)
      MaskEdBox2.Text = MaskEdBox1.Text
      MaskEdBox2.Mask = DFormat
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'列印
Private Sub PrintData()
Dim strSql As String
Dim ii As Integer
Dim bolSelect As Boolean
   
   '整批
   If Option1(0).Value = True Then
      If MaskEdBox1.Text = MsgText(29) Then
         MsgBox "請輸入起始銷退／轉開日期！"
         MaskEdBox1.SetFocus
         Exit Sub
      End If
      If MaskEdBox2.Text = MsgText(29) Then
         MsgBox "請輸入迄止銷退／轉開日期！"
         MaskEdBox2.SetFocus
         Exit Sub
      End If
   '單筆
   Else
      If grd1.Rows <= 2 And grd1.TextMatrix(1, 3) = "" Then
         MsgBox "無資料，請查詢！"
         Exit Sub
      End If
      bolSelect = False
      For ii = 1 To grd1.Rows - 1
         If grd1.TextMatrix(ii, 0) = "V" Then
            bolSelect = True
         End If
      Next ii
      If bolSelect = False Then
         MsgBox "請至少點選一筆欲列印的資料！"
         Exit Sub
      End If
   End If
   
On Error GoTo Checking:
   
   Screen.MousePointer = vbHourglass
   
   '整批
   If Option1(0).Value = True Then
      strSql = "select ' ' as V,a0s01 as 銷退單號,nvl(sqldatet(a0s03),'') as 銷退日期,a4301 as 發票編號,nvl(sqldatet(to_char(a0s25,'yyyymmdd')),'') as 折讓單列印日期,nvl(sqldatet(a4302),''),a4602,a4604,a4605,a0k04,a4303" & _
               " From acc0s0,acc460,acc431,acc430,acc0k0" & _
               " where a0s27='N' and a0s25 is null" & _
               " and a0s03>=" & Val(FCDate(MaskEdBox1.Text)) & " and a0s03<=" & Val(FCDate(MaskEdBox2.Text)) & _
               " and a0s01=a4601(+)" & _
               " and a0s02=axc02(+) and axc01=a4301(+) and axc02=a0k01(+)"
       'Modify by Amy 2019/12/16 原:axc02=a0k01(+) 會有「轉」字會抓不到acc0k0資料
       strSql = strSql & " Union" & _
               " select ' ' as V,'',nvl(sqldatet(a4310),''),a4301,nvl(sqldatet(to_char(a4318,'yyyymmdd')),''),nvl(sqldatet(a4302),''),a4602,a4604,a4605,a0k04,a4303" & _
               " From acc430,acc460,acc431,acc0k0" & _
               " where a4310 is not null and a4318 is null" & _
               " and a4310>=" & Val(FCDate(MaskEdBox1.Text)) & " and a4310<=" & Val(FCDate(MaskEdBox2.Text)) & _
               " and a4301=a4601(+) and a4301=axc01(+) and substr(axc02,1,9)=a0k01(+)"
      adoacc0k0.CursorLocation = adUseClient
      adoacc0k0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc0k0.RecordCount = 0 Then
         adoacc0k0.Close
         Screen.MousePointer = vbDefault
         MsgBox MsgText(28), , MsgText(5)
         Exit Sub
      End If
   End If
   
   'PUB_RestorePrinter Combo1
   PUB_SetOsDefaultPrinter Combo1
   '整批
   If Option1(0).Value = True Then
      adoacc0k0.MoveFirst
      Do While Not adoacc0k0.EOF
         m_A0S01 = "" & adoacc0k0.Fields(1)
         m_A0S03 = "" & adoacc0k0.Fields(2)
         m_Number = "" & adoacc0k0.Fields(3)
         m_A4302 = "" & adoacc0k0.Fields(5)
         m_A4602 = "" & adoacc0k0.Fields(6)
         m_Money = "" & adoacc0k0.Fields(7)
         m_TaxFee = "" & adoacc0k0.Fields(8)
         m_A0K04 = "" & adoacc0k0.Fields(9)
         m_A0K04ID = "" & adoacc0k0.Fields(10)
         m_Fillopen = IIf("" & adoacc0k0.Fields(4) = "", False, True)
         CallWordPrint
         
         adoacc0k0.MoveNext
      Loop
      adoacc0k0.Close
      MsgBox "列印完畢！"
   '單筆
   Else
      For ii = 1 To grd1.Rows - 1
         If grd1.TextMatrix(ii, 0) = "V" Then
            m_A0S01 = grd1.TextMatrix(ii, 1)
            m_A0S03 = grd1.TextMatrix(ii, 2)
            m_Number = grd1.TextMatrix(ii, 3)
            m_A4302 = grd1.TextMatrix(ii, 5)
            m_A4602 = grd1.TextMatrix(ii, 6)
            m_Money = grd1.TextMatrix(ii, 7)
            m_TaxFee = grd1.TextMatrix(ii, 8)
            m_A0K04 = grd1.TextMatrix(ii, 9)
            m_A0K04ID = grd1.TextMatrix(ii, 10)
            m_Fillopen = IIf(grd1.TextMatrix(ii, 4) = "", False, True)
            CallWordPrint
         End If
      Next ii
      MsgBox "列印完畢！"
   End If
   
   'PUB_RestorePrinter strPrinter
   PUB_SetOsDefaultPrinter strPrinter
   
   Screen.MousePointer = vbDefault
   Set adoacc0k0 = Nothing
   Exit Sub
   
Checking:
   Set adoacc0k0 = Nothing
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
End Sub

'查詢
Private Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
   
   '整批
   If Option1(0).Value = True Then
      'Memo by Amy 2019/12/16 由Acc0s0抓的資料 axc02不會有「轉」字,不需改語法
      strSql = "select ' ' as V,a0s01 as 銷退單號,nvl(sqldatet(a0s03),'') as 銷退日期,a4301 as 發票編號,nvl(sqldatet(to_char(a0s25,'yyyymmdd')),'') as 折讓單列印日期,nvl(sqldatet(a4302),''),a4602,a4604||'' as 金額,a4605||'' as 營業稅額,a0k04,a4303" & _
               " From acc0s0,acc460,acc431,acc430,acc0k0" & _
               " where a0s27='N' and a0s25 is null" & _
               " and a0s03>=" & Val(FCDate(MaskEdBox1.Text)) & " and a0s03<=" & Val(FCDate(MaskEdBox2.Text)) & _
               " and a0s01=a4601(+)" & _
               " and a0s02=axc02(+) and axc01=a4301(+) and axc02=a0k01(+)"
        'Modify by Amy 2019/12/16 原:axc02=a0k01(+)會有「轉」字會抓不到acc0k0資料
        strSql = strSql & " Union" & _
               " select ' ' as V,'',nvl(sqldatet(a4310),''),a4301,nvl(sqldatet(to_char(a4318,'yyyymmdd')),''),nvl(sqldatet(a4302),''),a4602,a4604||'',a4605||'',a0k04,a4303" & _
               " From acc430,acc460,acc431,acc0k0" & _
               " where a4310 is not null and a4318 is null" & _
               " and a4310>=" & Val(FCDate(MaskEdBox1.Text)) & " and a4310<=" & Val(FCDate(MaskEdBox2.Text)) & _
               " and a4301=a4601(+) and a4301=axc01(+) and substr(axc02,1,9)=a0k01(+)"
   Else
      'If Text1(0).Text = "" And Text1(1).Text = "" And Text1(2).Text = "" Then
      If Text1(0).Text = "" And Text1(2).Text = "" Then
         MsgBox "請至少輸入一項查詢條件！"
         Text1(0).SetFocus
         Exit Sub
      End If
      
      If Text1(0).Text <> "" Then '銷退單號
         If strSql <> "" Then strSql = strSql & " union "
         strSql = strSql & "select ' ' as V,a0s01 as 銷退單號,nvl(sqldatet(a0s03),'') as 銷退日期,a4301 as 發票編號,nvl(sqldatet(to_char(a0s25,'yyyymmdd')),'') as 折讓單列印日期,nvl(sqldatet(a4302),''),a4602,a4604||'' as 金額,a4605||'' as 營業稅額,a0k04,a4303" & _
                           " From acc0s0,acc460,acc431,acc430,acc0k0" & _
                           " where a0s01='" & Text1(0) & "' and a0s27='N'" & _
                           " and a0s01=a4601(+)" & _
                           " and a0s02=axc02(+) and axc01=a4301(+) and axc02=a0k01(+)"
      End If
'      If Text1(1).Text <> "" Then '發票號碼
'         If strSql <> "" Then strSql = strSql & " union "
'         strSql = strSql & "select ' ' as V,'' as 銷退單號,nvl(sqldatet(a4310),'') as 銷退日期,a4301 as 發票編號,nvl(sqldatet(to_char(a4318,'yyyymmdd')),'') as 折讓單列印日期,nvl(sqldatet(a4302),''),a4602,a4604||'' as 金額,a4605||'' as 營業稅額,a0k04,a4303" & _
'                           " From acc430,acc460,acc431,acc0k0" & _
'                           " where a4301='" & Text1(1) & "' and a4310 is not null" & _
'                           " and a4301=a4601(+) and a4301=axc01(+) and axc02=a0k01(+)"
'      End If
      If Text1(2).Text <> "" Then '本所案號
         If strSql <> "" Then strSql = strSql & " union "
         'Modify by Amy 2049/12/17 查 P108188000 出現型態不同的錯誤 原:a4604,a4605;原:a0k01=axc02會有「轉」字會抓不到acc0k0資料
         strSql = strSql & "select ' ' as V,a0s01 as 銷退單號,nvl(sqldatet(a0s03),'') as 銷退日期,a4301 as 發票編號,nvl(sqldatet(to_char(a0s25,'yyyymmdd')),'') as 折讓單列印日期,nvl(sqldatet(a4302),''),a4602,a4604||'' as 金額,a4605||'' as 營業稅額,a0k04,a4303" & _
                           " From acc0j0,acc0s0,acc460,acc431,acc430,acc0k0" & _
                           " where a0j02='" & Text1(2) & "' and a0j13=a0s02(+)" & _
                           " and a0s27='N'" & _
                           " and a0s01=a4601(+)" & _
                           " and a0s02=axc02(+) and axc01=a4301(+) and axc02=a0k01(+)" & _
                           " Union" & _
                           " select ' ' as V,'',nvl(sqldatet(a4310),''),a4301,nvl(sqldatet(to_char(a4318,'yyyymmdd')),''),nvl(sqldatet(a4302),''),a4602,a4604||'',a4605||'',a0k04,a4303" & _
                           " From acc0j0,acc460,acc431,acc430,acc0k0" & _
                           " where a0j02='" & Text1(2) & "' and a0j13=a0k01(+) and a0k01=substr(axc02(+),1,9) and axc01=a4301(+)" & _
                           " and a4310 is not null" & _
                           " and a4301=a4601(+)"
      End If
   End If
   
   grd1.Clear
   SetGrd
   
   Screen.MousePointer = vbHourglass
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd1.Recordset = rsTmp
      cmdok(0).Enabled = True
   Else
      cmdok(0).Enabled = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      ShowNoData
      Exit Sub
   End If
   
   '若有資料游標停在第一筆
   grd1.Visible = False
   grd1.col = 0
   grd1.row = 1
'   dblPrevRow = GRD1.row
   If rsTmp.RecordCount > 0 Then
      grd1.Text = "V"
      For i = 0 To grd1.Cols - 1
         grd1.col = i
         grd1.CellBackColor = &HFFC0C0
      Next i
   End If
   grd1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("V", "銷退單號", "銷退日期", "發票編號", "折讓單列印日期", "a4302", "a4602", "金額", "營業稅額", "a0k04", "a4303")
   '整批
   If Option1(0).Value = True Then
      arrGridHeadWidth = Array(0, 1500, 1500, 1500, 1600, 0, 0, 1000, 1000, 0, 0)
   Else
      arrGridHeadWidth = Array(300, 1500, 1500, 1500, 1600, 0, 0, 1000, 1000, 0, 0)
   End If
   grd1.Visible = False
   grd1.Cols = UBound(arrGridHeadText) + 1
   grd1.Rows = 2
   For iRow = 0 To grd1.Cols - 1
      grd1.row = 0
      grd1.col = iRow
      grd1.Text = arrGridHeadText(iRow)
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1.CellAlignment = flexAlignCenterCenter
   Next
   grd1.Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
   '整批
   If Index = 0 Then
      MaskEdBox1.Enabled = True
      MaskEdBox2.Enabled = True
      Text1(0).Enabled = False
'      Text1(1).Enabled = False
      Text1(2).Enabled = False
      cmdok(1).Enabled = True 'False Modify By Sindy 2018/10/24 整批開放可以查詢
      cmdok(0).Enabled = True
   '單筆
   Else
      MaskEdBox1.Enabled = False
      MaskEdBox2.Enabled = False
      Text1(0).Enabled = True
'      Text1(1).Enabled = True
      Text1(2).Enabled = True
      cmdok(1).Enabled = True
   End If
End Sub
