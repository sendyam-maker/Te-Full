VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc12e0 
   AutoRedraw      =   -1  'True
   Caption         =   "法律與智慧所案件對照表"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9075
   Begin VB.TextBox txtcp01 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2100
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1110
      Width           =   550
   End
   Begin VB.TextBox txtcp02 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2700
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1110
      Width           =   855
   End
   Begin VB.TextBox txtcp03 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3615
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1110
      Width           =   255
   End
   Begin VB.TextBox txtcp04 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3930
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1110
      Width           =   375
   End
   Begin VB.TextBox txtSalesArea1 
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
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   3
      Top             =   390
      Width           =   705
   End
   Begin VB.TextBox txtSalesArea 
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
      Left            =   2100
      MaxLength       =   6
      TabIndex        =   2
      Top             =   390
      Width           =   705
   End
   Begin VB.TextBox txtSales 
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
      Left            =   2100
      MaxLength       =   6
      TabIndex        =   4
      Top             =   750
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5130
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   630
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel(&E)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6960
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   630
      Width           =   1755
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   2100
      TabIndex        =   0
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   3570
      TabIndex        =   1
      Top             =   30
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "Frmacc12e0.frx":0000
      Height          =   4185
      Left            =   30
      TabIndex        =   15
      Top             =   1500
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7382
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "公司別|收據日期|本所案號|智權人員|總收文號|申請人名稱|收據抬頭|收據號碼|收據金額|收款日|客戶編號|案件性質"
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
      _Band(0).Cols   =   15
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "本所案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   930
      TabIndex        =   16
      Top             =   1155
      Width           =   1185
   End
   Begin MSForms.Label LblSalesName 
      Height          =   285
      Left            =   3060
      TabIndex        =   14
      Top             =   810
      Width           =   1575
      VariousPropertyBits=   19
      Caption         =   "LblFM2"
      Size            =   "2778;503"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin VB.Line Line2 
      X1              =   2730
      X2              =   2940
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "業務區："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   13
      Top             =   390
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "介紹人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      TabIndex        =   12
      Top             =   780
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "法律所收據日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   630
      TabIndex        =   11
      Top             =   60
      Width           =   1485
   End
   Begin VB.Line Line1 
      X1              =   3300
      X2              =   3510
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1110
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc12e0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/17 Form2.0已修改
'Create By Sindy 2020/9/21
Option Explicit

Dim adoquery As New ADODB.Recordset
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intCounter As Integer
Dim lngPageNo As Long '頁數
Dim strPrinter As String
Dim dblPrevRow As Double


'Excel
Private Sub Command1_Click()
Dim ii As Integer
   
On Error GoTo ErrHnd
   
   If GRD1.Rows > 1 Then
      If GRD1.TextMatrix(1, 1) = "" Then
         MsgBox "無資料！"
         Exit Sub
      End If
   Else
      MsgBox "無資料！"
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   Set xlsAnnuity = New Excel.Application
   Call SetExcelWorksheets
   PrintHead_Excel intCounter '頁首
   With wksAnnuity
   For ii = 1 To GRD1.Rows - 1
'      '第2頁切頁有誤 +  And intCounter <> 48 判斷
'      If (lngPageNo = 1 And intCounter Mod 32 = 0) Or _
'         (lngPageNo <> 1 And intCounter Mod 32 = 0 And intCounter <> 32) Then
'         '換頁
'         intCounter = intCounter + 1
'         .Range("A" & intCounter).Select
'         .HPageBreaks.add Before:=.Application.ActiveCell
'         PrintHead_Excel intCounter '頁首
'      End If
      
      '明細資料
      intCounter = intCounter + 1
      .Range("A" & intCounter).Value = GRD1.TextMatrix(ii, 1) '案源單號
      .Range("B" & intCounter).Value = GRD1.TextMatrix(ii, 4) '公司別
      .Range("C" & intCounter).Value = GRD1.TextMatrix(ii, 5) '收據日期
      .Range("D" & intCounter).Value = GRD1.TextMatrix(ii, 6) '本所案號
      .Range("E" & intCounter).Value = GRD1.TextMatrix(ii, 7) '智權人員
      .Range("F" & intCounter).Value = GRD1.TextMatrix(ii, 8) '總收文號
      .Range("G" & intCounter).Value = GRD1.TextMatrix(ii, 9) '申請人名稱
      .Range("H" & intCounter).Value = GRD1.TextMatrix(ii, 10) '收據抬頭
      .Range("I" & intCounter).Value = GRD1.TextMatrix(ii, 11) '收據號碼
      .Range("J" & intCounter).Value = GRD1.TextMatrix(ii, 12) '服務費
      .Range("K" & intCounter).Value = GRD1.TextMatrix(ii, 13) '規費
      .Range("L" & intCounter).Value = GRD1.TextMatrix(ii, 14) '收款日
      .Range("M" & intCounter).Value = GRD1.TextMatrix(ii, 15) '客戶編號
      .Range("N" & intCounter).Value = GRD1.TextMatrix(ii, 16) '案件性質
   Next ii
   End With
   
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize

   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   
   Screen.MousePointer = vbDefault

'   MaskEdBox1.Mask = ""
'   MaskEdBox2.Mask = ""
'   MaskEdBox1.Text = ""
'   MaskEdBox2.Text = ""
'   MaskEdBox1.Mask = DFormat
'   MaskEdBox2.Mask = DFormat
   
   Exit Sub

ErrHnd:
   Screen.MousePointer = vbDefault
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing

   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

Private Sub SetExcelWorksheets()
   xlsAnnuity.Visible = True
   xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   'wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   '設定各欄位長度
   wksAnnuity.Columns("A:A").ColumnWidth = 10
   wksAnnuity.Columns("B:B").ColumnWidth = 9
   wksAnnuity.Columns("C:C").ColumnWidth = 9
   wksAnnuity.Columns("D:D").ColumnWidth = 10
   wksAnnuity.Columns("E:E").ColumnWidth = 10
   wksAnnuity.Columns("F:F").ColumnWidth = 10
   wksAnnuity.Columns("G:G").ColumnWidth = 10
   wksAnnuity.Columns("H:H").ColumnWidth = 10
   wksAnnuity.Columns("I:I").ColumnWidth = 10
   wksAnnuity.Columns("J:J").ColumnWidth = 10
   wksAnnuity.Columns("K:K").ColumnWidth = 10
   wksAnnuity.Columns("L:L").ColumnWidth = 10
   wksAnnuity.Columns("M:M").ColumnWidth = 10
   wksAnnuity.Columns("N:N").ColumnWidth = 10
   
   wksAnnuity.Range("B:B").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("C:C").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("F:F").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("G:G").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   
   intCounter = 1
End Sub

'表頭
Private Sub PrintHead_Excel(ByRef iRow As Integer)
Dim i As Integer, strTemp As String

   lngPageNo = lngPageNo + 1
   With wksAnnuity
      .Range("E" & iRow).Value = "法律與智慧所案件對照表"
      '選取,儲存格合併,置中,粗體字
      strTemp = "A" & iRow & ":H" & iRow
      .Range(strTemp).Select
      With .Application.Selection
          .HorizontalAlignment = xlGeneral
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      With .Application.Selection
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      .Application.Selection.Font.Bold = True

      iRow = iRow + 1
      .Range("A" & iRow).Value = "列印人：" & strUserName
      .Range("D" & iRow).Value = "法律所收據日：" & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
      .Range("G" & iRow).Value = "列印日期："
      .Range("H" & iRow).Value = Format(strSrvDate(2), "###/##/##")
      iRow = iRow + 1
'      .Range("G" & iRow).Value = "頁數："
'      .Range("H" & iRow).Value = lngPageNo
      strTemp = "D" & iRow - 1 & ":D" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter '置中
      End With
      strTemp = "F" & iRow - 1 & ":F" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlRight '靠右
      End With
      strTemp = "G" & iRow & ":H" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlLeft '靠左
      End With
      
      iRow = iRow + 1
      .Range("A" & iRow).Value = "案源單號"
      .Range("B" & iRow).Value = "公司別"
      .Range("C" & iRow).Value = "收據日期"
      .Range("D" & iRow).Value = "本所案號"
      .Range("E" & iRow).Value = "智權人員"
      .Range("F" & iRow).Value = "總收文號"
      .Range("G" & iRow).Value = "申請人名稱"
      .Range("H" & iRow).Value = "收據抬頭"
      .Range("I" & iRow).Value = "收據號碼"
      .Range("J" & iRow).Value = "服務費"
      .Range("K" & iRow).Value = "規費"
      .Range("L" & iRow).Value = "收款日"
      .Range("M" & iRow).Value = "客戶編號"
      .Range("N" & iRow).Value = "案件性質"
      strTemp = "A" & iRow & ":N" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter '置中
      End With
'      With .Application.Selection.Borders(xlEdgeLeft)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeTop)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
'      With .Application.Selection.Borders(xlEdgeRight)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlInsideVertical)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
   End With
End Sub

'Private Sub PrintData_Excel(p_Rst As ADODB.Recordset, ByRef iRow As Integer)
'Dim strTemp As String
'
'   iRow = iRow + 1
'   With wksAnnuity
'      If m_strT01 = "" & p_Rst.Fields("T01") And _
'         m_strT15 = "" & p_Rst.Fields("T15") And _
'         m_strT19 = "" & p_Rst.Fields("T19") Then
'         .Range("A" & iRow).Value = "---"
'         .Range("B" & iRow).Value = "---"
'         .Range("C" & iRow).Value = "---"
'         .Range("F" & iRow).Value = "---"
'         .Range("G" & iRow).Value = "---"
'      Else
'         'Modify By Sindy 2019/12/18
'         '.Range("A" & iRow).Value = "" & p_Rst.Fields("T01") '客戶編號
'         If p_Rst.Fields("T24") = "1" Then '代填方式
'            .Range("A" & iRow).Value = "每筆代繳"
'         ElseIf p_Rst.Fields("T24") = "2" Then
'            .Range("A" & iRow).Value = "單筆收據稅額超過2000元"
'         Else
'            .Range("A" & iRow).Value = "" & p_Rst.Fields("T24")
'         End If
'         '2019/12/18 END
'         .Range("B" & iRow).Value = "" & p_Rst.Fields("T15") '收據抬頭
'         .Range("C" & iRow).Value = "" & p_Rst.Fields("T19") '統一編號
'         .Range("F" & iRow).Value = "" & p_Rst.Fields("T20") '繳款書地址
'         .Range("G" & iRow).Value = "" & p_Rst.Fields("T16") '收件人
'      End If
'      .Range("D" & iRow).Value = IIf("" & p_Rst.Fields("T07") = "X", "", "" & p_Rst.Fields("T07")) '票號
'      .Range("E" & iRow).Value = ChangeTStringToTDateString("" & p_Rst.Fields("T08")) '票據到期日
'      .Range("H" & iRow).Value = "" & p_Rst.Fields("T26") '會計備註
'      .Range("A" & iRow & ":G" & iRow).Select
'      .Application.Selection.VerticalAlignment = xlTop '靠上
'      .Range("H" & iRow).Select
'      .Application.Selection.WrapText = True '自動換行
'   End With
'   m_strT01 = "" & p_Rst.Fields("T01")
'   m_strT15 = "" & p_Rst.Fields("T15")
'   m_strT19 = "" & p_Rst.Fields("T19")
'End Sub

'查詢
Private Sub Command2_Click()
   Call QueryData
End Sub

Private Sub Form_Activate()
   strFormName = Name
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single

   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   'Modify by Amy 2023/10/06 原W:9045/H:6120
   Me.Width = 9200
   Me.Height = 6390
   '改單線固定(調整大小不用再設定)
   'Modify by Amy 2023/10/06 原(lngWidth - Me.Width) / 2,切畫面不需再左移-瑞婷
   Me.Move 0, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
      For intY = 0 To Int(ScaleHeight / sglHeight)
         PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
      Next
   Next

   '起日預設為上月1日
   strDate = CompDate(1, -1, strSrvDate(1))
   strDate = Left(strDate, 6) & "01"
   MaskEdBox1.Text = CFDate(ACDate(strDate))
   MaskEdBox1.Mask = DFormat
   '止日預設為上月底
   'strDate = GetMonthStdDay(Left(strDate, 6), 1, True)
   MaskEdBox2.Text = CFDate(ACDate(strSrvDate(1))) '系統日
   MaskEdBox2.Mask = DFormat
   
   lblSalesName.Caption = "" 'Add By Sindy 2022/2/17
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc12e0 = Nothing
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   '日期檢查
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox "收據起始日期格式錯誤！", vbExclamation
      FormCheck = False
      MaskEdBox1.SetFocus
      Exit Function
   End If

   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox "收據迄止日期格式錯誤！", vbExclamation
      FormCheck = False
      MaskEdBox2.SetFocus
      Exit Function
   End If
   FormCheck = True
End Function

Private Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
Dim strDate_S As String, strDate_E As String
Dim strCon As String, strConCP As String
   
   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
'   m_blnColOrderAsc = True
   GRD1.Clear
   SetGrd
   
   strDate_S = DBDATE(Me.MaskEdBox1) - 19110000
   strDate_E = DBDATE(Me.MaskEdBox2) - 19110000
   If txtSales <> "" Then
      strCon = strCon & " and substr(los04,1,5)='" & txtSales & "'"
   End If
   If txtSalesArea <> "" Then
      strCon = strCon & " and s2.st15>='" & txtSalesArea & "'"
   End If
   If txtSalesArea1 <> "" Then
      strCon = strCon & " and s2.st15<='" & txtSalesArea1 & "'"
   End If
   
   'Add By Sindy 2021/2/4
   If txtcp01 <> "" And txtcp02 <> "" Then
      If txtcp03 = "" Then txtcp03 = "0"
      If txtcp04 = "" Then txtcp04 = "00"
      strConCP = " AND los15 in (select cp162 from caseprogress where cp01='" & txtcp01 & "' AND cp02='" & txtcp02 & "' AND cp03='" & txtcp03 & "' AND cp04='" & txtcp04 & "')"
   End If
   '2021/2/4 END
   
   Screen.MousePointer = vbHourglass
   strSql = "select V,key1,los02,dtype,a0k11,a0k02,CaseNo,s1.st02,cp09,NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),a0k04,a0k01,TO_CHAR(nvl(a0k06,cp16),'9,999,999'),TO_CHAR(nvl(a0k07,cp17),'9,999,999'),sqldatet(a0L02),a0k03,cpm03,cpm04,cp162,Los15" & _
            " from ("
   strSql = strSql & _
               "select '' V,key1,los02,'1' dtype,a0k11,sqldatet(a0k02) a0k02,cp01||cp02||cp03||cp04 CaseNo,cp13,cp09,nvl(pa26,nvl(tm23,nvl(sp08,''))) app1,a0k04,a0k01,a0k06,a0k07,'',nvl(a0k03,nvl(pa26,nvl(tm23,nvl(sp08,'')))) a0k03,cp10,cp01,cp16,cp17,nvl(pa09,nvl(tm10,nvl(sp09,''))) NaID,los04,cp162,Los15" & _
               " FROM (SELECT c.cp162 key1 FROM acc0k0,acc0j0 j,caseprogress c" & _
               " Where a0k02>=" & strDate_S & " And a0k02<=" & strDate_E & " And a0k09=0" & _
               " AND j.a0j13(+)=a0k01 AND c.cp09(+)=j.a0j01 AND c.cp162 IS NOT NULL) x" & _
               ",LawOfficeSource,acc0j0,acc0k0,caseprogress,patent,Trademark,servicepractice" & _
               " WHERE los15(+)=key1 AND a0j01(+)=los01 AND a0k01(+)=a0j13 AND los01=cp09" & _
               " AND cp01=pa01(+) AND cp02=pa02(+) AND cp03=pa03(+) AND cp04=pa04(+)" & _
               " AND cp01=tm01(+) AND cp02=tm02(+) AND cp03=tm03(+) AND cp04=tm04(+)" & _
               " AND cp01=sp01(+) AND cp02=sp02(+) AND cp03=sp03(+) AND cp04=sp04(+)"
   strSql = strSql & " Union " & _
               "select '' V,key1,los02,'1' dtype,a0k11,sqldatet(a0k02) a0k02,cp01||cp02||cp03||cp04 CaseNo,cp13,cp09,nvl(pa26,nvl(tm23,nvl(sp08,''))) app1,a0k04,a0k01,a0k06,a0k07,'',nvl(a0k03,nvl(pa26,nvl(tm23,nvl(sp08,'')))) a0k03,cp10,cp01,cp16,cp17,nvl(pa09,nvl(tm10,nvl(sp09,''))) NaID,los04,cp162,Los15" & _
               " FROM (SELECT c.cp162 key1 FROM acc0k0,acc0j0 j,caseprogress c" & _
               " Where a0k02>=" & strDate_S & " And a0k02<=" & strDate_E & " And a0k09=0" & _
               " AND j.a0j13(+)=a0k01 AND c.cp09(+)=j.a0j01 AND c.cp162 IS NOT NULL) x" & _
               ",LawOfficeSource,acc0j0,acc0k0,caseprogress,patent,Trademark,servicepractice" & _
               " WHERE los15(+)=key1 AND a0j01(+)=los10 AND a0k01(+)=a0j13 AND los10=cp09" & _
               " AND cp01=pa01(+) AND cp02=pa02(+) AND cp03=pa03(+) AND cp04=pa04(+)" & _
               " AND cp01=tm01(+) AND cp02=tm02(+) AND cp03=tm03(+) AND cp04=tm04(+)" & _
               " AND cp01=sp01(+) AND cp02=sp02(+) AND cp03=sp03(+) AND cp04=sp04(+)"
   strSql = strSql & " Union " & _
               "SELECT '' V,cp162 key1,los02,'2' dtype,a0k11,sqldatet(a0k02) a0k02,cp01||cp02||cp03||cp04 CaseNo,cp13,cp09,nvl(hc05,nvl(lc11,'')) app1,a0k04,a0k01,a0k06,a0k07,'',nvl(a0k03,nvl(hc05,nvl(lc11,''))) a0k03,cp10,cp01,cp16,cp17,nvl(lc15,'000') NaID,los04,cp162,Los15" & _
               " FROM acc0k0,acc0j0,LawOfficeSource,caseprogress,lawcase,hirecase" & _
               " Where a0k02>=" & strDate_S & " And a0k02<=" & strDate_E & " And a0k09=0" & _
               " AND a0j13(+)=a0k01 AND cp09(+)=a0j01 AND cp162 IS NOT NULL and los15(+)=cp162" & _
               " AND cp01=lc01(+) AND cp02=lc02(+) AND cp03=lc03(+) AND cp04=lc04(+)" & _
               " AND cp01=hc01(+) AND cp02=hc02(+) AND cp03=hc03(+) AND cp04=hc04(+)" & _
            "),staff s1,casepropertymap,customer,staff s2,acc0m0,acc0L0" & _
            " Where cp13=s1.st01(+) and cp01=cpm01(+) and cp10=cpm02(+)" & _
            " and substr(app1,1,8)=cu01(+) and substr(app1,9,1)=cu02(+) and substr(los04,1,5)=s2.st01(+)" & strCon & _
            " and a0k01=a0m02(+) and a0m01=a0L01(+)" & strConCP & _
            " order by 2,3,4"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GRD1.FixedCols = 0
      Set GRD1.Recordset = rsTmp.Clone 'rsTmp
      GRD1.FixedCols = 4
   Else
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      MsgBox "查無資料！"
      Exit Sub
   End If
   
   Call SetGrd(True)
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   dblPrevRow = GRD1.row
   If rsTmp.RecordCount > 0 Then
      GRD1.Text = "V"
      For i = 4 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &HFFC0C0
      Next i
   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub SetGrd(Optional bolReplaceHeader As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "案源單號", "los02", "dtype", "公司別", "收據日期", "本所案號", "智權人員", "總收文號", "申請人名稱", "收據抬頭", "收據號碼", "服務費", "規費", "收款日", "客戶編號", "案件性質")
   arrGridHeadWidth = Array(0, 800, 0, 0, 300, 900, 900, 800, 1000, 1300, 1000, 1000, 1000, 1000, 900, 1000, 1000)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   If bolReplaceHeader = False Then
      GRD1.Rows = 2
   End If
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub grd1_SelChange()
Dim i As Integer

GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 0
      GRD1.row = dblPrevRow
      GRD1.Text = ""
      For i = 4 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   dblPrevRow = GRD1.row
'   If grd1.Text = "V" Then
'      grd1.Text = ""
'      For i = 0 To grd1.Cols - 1
'         grd1.col = i
'         grd1.CellBackColor = QBColor(15)
'      Next i
'   Else
      If GRD1.TextMatrix(GRD1.row, 1) <> "" Then
         GRD1.Text = "V"
         For i = 4 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
'   End If
End If
GRD1.Visible = True
End Sub

Private Sub txtcp01_GotFocus()
   TextInverse txtcp01
   CloseIme
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
  Dim strTit As String
  Dim strMsg As String
  
  txtcp01.Text = UCase(txtcp01.Text)
  If IsEmptyText(txtcp01) = False Then
      '2011/5/20 MODIFY BY SONIA
      'If Not IsCorrectSysKindLaw(txtcp01) Then
      'Modify By Sindy 2021/2/4 Mark
'      If CheckSys(txtcp01) <> "3" And CheckSys(txtcp01) <> "4" Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "系統類別不正確"
'         MsgBox strMsg, vbOKOnly, strTit
'         txtcp01_GotFocus
'         Exit Sub
'      End If
      
'      ' 檢查使用者是否有使用該系統類別的權限
'      If IsUserHasRightOfSystem(strUserNum, txtcp01) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "您沒有使有此系統別的權限"
'         MsgBox strMsg, vbOKOnly, strTit
'         txtcp01_GotFocus
'         Exit Sub
'      End If
   End If

'   If txtcp01 <> "" Then
'      txtcp01 = UCase(txtcp01)
'      If txtcp01 = "L" Or txtcp01 = "LA" Or txtcp01 = "FCL" Then
'         blnCom1 = True
'      Else
'         DataErrorMessage 1, "系統類別"
'         blnCom1 = False
'         Cancel = True
'      End If
'   End If
'   ChkCmd
   If Cancel Then TextInverse txtcp01
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
   CloseIme
End Sub

Private Sub txtcp03_GotFocus()
   TextInverse txtcp03
   CloseIme
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
   CloseIme
End Sub

Private Sub txtSalesArea_GotFocus()
   TextInverse txtSalesArea
   CloseIme
End Sub

Private Sub txtSalesArea_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSalesArea1_GotFocus()
   TextInverse txtSalesArea1
   CloseIme
End Sub

Private Sub txtSalesArea1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'智權人員
Private Sub txtSales_Change()
   If txtSales = MsgText(601) Then
      Exit Sub
   End If
   lblSalesName = StaffQuery(txtSales)
End Sub

Private Sub txtSales_GotFocus()
   TextInverse txtSales
   CloseIme
End Sub

Private Sub txtSales_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
