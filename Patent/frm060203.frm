VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060203 
   BorderStyle     =   1  '單線固定
   Caption         =   "翻譯完稿案件查詢/列印"
   ClientHeight    =   5352
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5352
   ScaleWidth      =   8952
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   6435
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1395
      Width           =   1080
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5175
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1395
      Width           =   1080
   End
   Begin VB.TextBox txtCP14 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2295
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1395
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1125
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   990
      Width           =   7725
   End
   Begin VB.CheckBox Check1 
      Caption         =   "只抓未付翻譯費案件"
      Height          =   225
      Left            =   180
      TabIndex        =   7
      Top             =   1800
      Width           =   1995
   End
   Begin VB.TextBox txtCP14 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1035
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1395
      Width           =   975
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
      Left            =   5145
      TabIndex        =   8
      Top             =   1740
      Width           =   3675
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   1
      Left            =   3510
      MaxLength       =   7
      TabIndex        =   1
      Top             =   600
      Width           =   1080
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   7065
      TabIndex        =   10
      Top             =   60
      Width           =   1020
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2205
      MaxLength       =   7
      TabIndex        =   0
      Top             =   600
      Width           =   1080
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6030
      TabIndex        =   9
      Top             =   60
      Width           =   1020
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   8100
      TabIndex        =   11
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2628
      Left            =   132
      TabIndex        =   13
      Top             =   2376
      Width           =   8688
      _ExtentX        =   15325
      _ExtentY        =   4636
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "完稿日　|發文日　|承辦人　　　|本所案號　　　|案件名稱　　|中/原文字數|中/原文字數(財)|數學式字數|相似折扣|瑕疵折扣|加成比率"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
   End
   Begin VB.Label lblRemark2 
      AutoSize        =   -1  'True
      Caption         =   "約定薪資的翻譯人員比照外翻人員的規則"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   192
      Left            =   168
      TabIndex        =   23
      Top             =   2112
      Width           =   3552
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H8000000E&
      Height          =   180
      Index           =   0
      Left            =   2460
      TabIndex        =   21
      Top             =   270
      Width           =   900
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BackColor       =   &H8000000E&
      Height          =   180
      Index           =   1
      Left            =   3540
      TabIndex        =   20
      Top             =   270
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "完稿日：                         －"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4320
      TabIndex        =   19
      Top             =   1440
      Width           =   2100
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "承辦人：                        －"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   18
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblRemark 
      AutoSize        =   -1  'True
      Caption         =   "( 所內同仁(含巨京F5595寰華F5639)下班翻譯案件 98/5/1 以後完稿的則過濾條件為發文日或閉卷日 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   390
      Left            =   4680
      TabIndex        =   17
      Top             =   540
      Width           =   4155
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "印表機："
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
      Left            =   4320
      TabIndex        =   16
      Top             =   1770
      Width           =   975
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "翻譯該程序若有延期，本所案號前作 * 標記，若發文日後有 # 號則表示該日期為取消收文日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   15
      Top             =   5070
      Width           =   7770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件來源："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   14
      Top             =   1050
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "完稿/發文/取消收文日：                        －"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   12
      Top             =   630
      Width           =   3315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "前次完稿/發文/取消收文日：                     －                     ( 按列印後更新)"
      Height          =   180
      Left            =   180
      TabIndex        =   22
      Top             =   270
      Width           =   5565
   End
End
Attribute VB_Name = "frm060203"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改 (Printer列印未改)
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
'Add by Morgan 2007/6/8
Option Explicit

Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500, ciColGap = 250
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim bolBarShow As Boolean
Dim strPrinter As String
Dim m_iCols As Integer


Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub DoPrint()

   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 2
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
         For iCol = LBound(strTemp) To UBound(strTemp)
            '案件名稱抓8個字
            If iCol = 5 Then
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 8)
            Else
               strTemp(iCol) = .TextMatrix(iRow, iCol - 1)
            End If
         Next
         PrintDetail strTemp
      Next
      Call PrintReportFooter(.Rows - 1)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
   
   
End Sub

Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   intI = m_iCols + 1
   ReDim PLeft(1 To intI)
   PLeft(1) = ciStartX
   For intI = 2 To intI
      '案件名稱可印8個中文
      If intI = 6 Then
         PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(String(8, "　")) + ciColGap
      Else
         PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(grdDataList.TextMatrix(0, intI - 2)) + ciColGap
      End If
   Next
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.Print String(130, "-")
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
    
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
      'Modify by Morgan 2007/8/13 +加成比率
      'If iCol > 4 And iCol < 7 Then
      'Modify by Morgan 2007/10/24 +日文字數,數學式字數
      'If iCol > 4 And iCol < 8 Then
      If iCol > 5 Then
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - ciColGap
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
      Else
        Printer.CurrentX = PLeft(iCol)
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
      End If
    Next
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = "FCP翻譯完稿案件明細表"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   If txtNo(0) <> "" Or txtNo(1) <> "" Then
      strPTmp = "完稿/發文日："
      PrintNewLine
      Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(9, "　"))
      Printer.CurrentY = iPrint
      Printer.Print strPTmp & CFDate(txtNo(0)) & " － " & IIf(txtNo(1) <> "", CFDate(txtNo(1)), "")
   End If
   
   strPTmp = "案件來源：" & Combo2.Text
   PrintNewLine
   'Modified by Morgan 2019/9/3
   'Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(9, "　"))
   Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(strPTmp) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   If txtCP14(0) & txtCP14(1) <> "" Then
      strPTmp = "承辦人：" & txtCP14(0) & " ～ " & txtCP14(1)
      PrintNewLine
      Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(9, "　"))
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
   End If
   
   'Added by Morgan 2016/5/11
   If txtNo(2) <> "" Or txtNo(3) <> "" Then
      strPTmp = "完稿日："
      PrintNewLine
      Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(9, "　"))
      Printer.CurrentY = iPrint
      Printer.Print strPTmp & CFDate(txtNo(2)) & " － " & IIf(txtNo(3) <> "", CFDate(txtNo(3)), "")
   End If
   
   'end 2016/5/11
   If Me.Check1.Value = 1 Then
      strPTmp = "（ 未付翻譯費案件 ）"
      PrintNewLine
      Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(strPTmp) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
   End If
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   strPTmp = lblMemo.Caption
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print String(130, "-")
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    For intI = 1 To m_iCols
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print grdDataList.TextMatrix(0, intI - 1)
    Next
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(130, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(130, "-")
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

Private Sub cmdPrint_Click()
   If Not grdDataList.Recordset Is Nothing Then
      If grdDataList.Recordset.RecordCount > 0 Then
         PUB_RestorePrinter Combo1
         DoPrint
         PUB_RestorePrinter strPrinter
         'Modify by Amy 2014/07/14
         'SaveSetting "TAIE", "FCP", Me.Name & "#DATE01", txtNo(0).Text
         'Modified by Morgan 2016/3/29 會有一個以上的人使用,加員工號
         'PUB_SaveLastDate Me.Name, "txtNo(0)", txtNo(0)
         PUB_SaveLastDate Me.Name, "txtNo(0)" & "-" & strUserNum, txtNo(0)
         'end 2016/3/29
         Label4(0) = txtNo(0)
         'SaveSetting "TAIE", "FCP", Me.Name & "#DATE02", txtNo(1).Text
         'Modified by Morgan 2016/3/29 會有一個以上的人使用,加員工號
         'PUB_SaveLastDate Me.Name, "txtNo(1)", txtNo(1)
         PUB_SaveLastDate Me.Name, "txtNo(1)" & "-" & strUserNum, txtNo(1)
         'end 2016/3/29
         Label4(1) = txtNo(1)
         'end 2014/07/14
      End If
   End If
End Sub

'Modify by Morgan 2009/4/20
'日期條件:外翻人員一律用完稿日,其他人員則判斷98/5/1 以後完稿的改抓發文日
Private Sub doQuery()
Dim stCon As String, stConEP09 As String, stConCP27 As String
Dim stSQLInner As String
Dim stConCFP As String
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
   
   'Modify by Morgan 2011/8/1 +北京寰華 F5639
   'Modified by Morgan 2024/6/18 約定薪資的翻譯人員比照外翻人員的規則
   '內翻以發文日/取消收文日管控
   stSQLInner = " ( exists(select * from staff_idmap a1,staff a2" & _
      " where a1.sim02=c1.cp14 and a2.st01(+)=a1.sim01 and a2.st04='1' and not substr(sim01,-2)>='9A') or c1.cp14='F5595' or c1.cp14='F5639') "
   
On Error GoTo flgErr
   
   stCon = "": stConEP09 = "": stConCP27 = ""
   If txtNo(0) <> "" Then
     stConEP09 = stConEP09 & " and EP09>=" & DBDATE(txtNo(0))
     'Modify by Morgan 2009/8/21 +閉卷日(取消收文)
     'stConCP27 = stConCP27 & " and c1.cp27>=" & DBDATE(txtNo(0))
     stConCP27 = stConCP27 & " and (c1.cp27>=" & DBDATE(txtNo(0)) & " or (c1.cp27 is null and c1.cp57>=" & DBDATE(txtNo(0)) & ") )"
   End If
   If txtNo(1) <> "" Then
     stConEP09 = stConEP09 & " and EP09<=" & DBDATE(txtNo(1))
      'Modify by Morgan 2009/8/21 +閉卷日(取消收文)
     'stConCP27 = stConCP27 & " and c1.cp27<=" & DBDATE(txtNo(1))
     stConCP27 = stConCP27 & " and (c1.cp27<=" & DBDATE(txtNo(1)) & " or (c1.cp27 is null and c1.cp57<=" & DBDATE(txtNo(1)) & ") )"
   End If
   If txtNo(0) <> "" Or txtNo(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1, 12) & txtNo(0) & "-" & txtNo(1) 'Add By Sindy 2010/12/7
   End If
   
   If txtNo(2) <> "" Then
      stCon = stCon & " and EP09>=" & DBDATE(txtNo(2))
   End If
   If txtNo(3) <> "" Then
      stCon = stCon & " and EP09<=" & DBDATE(txtNo(3))
   End If
   If txtNo(2) <> "" Or txtNo(3) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label6, 4) & txtNo(2) & "-" & txtNo(3) 'Add By Sindy 2010/12/7
   End If
    
   'Modify by Morgan 2009/4/20 抓外譯編號時要排除 F5588舜禹 & F5595巨京
   Select Case Combo2.ItemData(Combo2.ListIndex)
      Case 2 '所有承辦人為F編號案件( 排除 F5588舜禹 & F5595巨京 )
         'Modify by Morgan 2011/8/1 +北京寰華 F5639
         'Modified by Morgan 2012/8/30 +F5614通用 & F5653捷恩凱
         'Modified by Morgan 2017/8/30 +F5698 迅達翻譯 --婧瑄
         'Modified by Lydia 2018/01/04
         'stCon = stCon & " and SUBSTR(c1.CP14,1,1)='F' and c1.cp14<>'F5588' and c1.cp14<>'F5595' and c1.cp14<>'F5639' and c1.cp14<>'F5614' and c1.cp14<>'F5653' and c1.cp14<>'F5698' "
         'Modified by Lydia 2025/03/13 改用模組取得
         'stCon = stCon & " and SUBSTR(c1.CP14,1,1)='F' and c1.cp14<>'" & 外翻_舜禹 & "' and c1.cp14<>'" & 外翻_捷恩凱 & "' and c1.cp14<>'" & 外翻_迅達 & "' and c1.cp14<>'F5595' and c1.cp14<>'F5639' and c1.cp14<>'F5614'  "
         stCon = stCon & " and SUBSTR(c1.CP14,1,1)='F' and instr('" & Pub_SetF51Order("F", "") & "',c1.cp14)=0 and c1.cp14<>'F5595' and c1.cp14<>'F5639' and c1.cp14<>'F5614'  "
      Case 3 '所內同仁下班翻譯案件( 含 F5595巨京 & F5639寰華 )
         stCon = stCon & " and SUBSTR(c1.CP14,1,1)='F' and " & stSQLInner
   End Select
   If Combo2.Text <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label2(1) & Combo2.Text 'Add By Sindy 2010/12/7
   End If
   
   If txtCP14(0) <> "" Then
     stCon = stCon & " and c1.cp14>='" & txtCP14(0) & "'"
   End If
   If txtCP14(1) <> "" Then
     stCon = stCon & " and c1.cp14<='" & txtCP14(1) & "'"
   End If
   If txtCP14(0) <> "" Or txtCP14(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label5, 4) & txtCP14(0) & "-" & txtCP14(1) 'Add By Sindy 2010/12/7
   End If
   
   If Check1.Value = 1 Then
     stCon = stCon & " and tf07 is null"
     pub_QL05 = pub_QL05 & ";" & Check1.Caption 'Add By Sindy 2010/12/7
   End If
   
   stConCFP = stCon 'Added by Morgan 2025/11/12
   
   'Modify by Morgan 2008/1/10 改控制承辦人部門為F21,F51,F52的
   'stCon = " and c1.cp01='FCP'"
   '2008/4/8 MODIFY BY SONIA 加F81
   stCon = stCon & " and ST03 in ('F21','F51','F52','F81')"
   'Added by Morgan 2011/2/24 +過濾國外部收文案件
   stCon = stCon & " and substr(c1.cp12,1,1)='F'"
   
    
   'Modify by Morgan 2007/8/13 -法定期限,+加成比率
   'Modify by Morgan 2007/10/24 -本所期限,+日文字數,數學式字數
   'Modified by Morgan 2019/9/3 日文字數改原文字數(財)=財務處輸入的字數--婧瑄
   
   'Modify by Morgan 2010/8/13 百年蟲
   '98/5/1 以前完稿的案件
   strExc(0) = "select SQLDATET(EP09) C01,substrb(' '||sqldatet(nvl(c1.CP27,c1.CP57)),-9)||decode(c1.cp57,null,'','#') C12" & _
      ",st02||' '||c1.cp14 C02" & _
      ",DECODE(c2.CP01,NULL,' ','*')||pa01||'-'||pa02||decode(pa03||pa04,'000','',pa03||pa04) C03" & _
      ",pa05 C04,tf02 C05,tf21 C06,tf04 C07,tf05||decode(tf05,null,'','%') C08" & _
      ",tf06||decode(tf06,null,'','%') C09,tf18||decode(tf18,null,'','%') C10" & _
      ",ep09,pa01,pa02,pa03,pa04,c1.cp14 C11" & _
      " from engineerprogress,caseprogress c1,staff,patent,transfee,caseprogress c2" & _
      " where ep09<20090501" & stCon & stConEP09 & _
      " and c1.cp09(+)=ep02 and c1.cp10='201' and st01(+)=c1.cp14" & _
      " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04" & _
      " and tf01(+)=c1.cp09 AND c2.CP43(+)=c1.CP09 AND c2.CP10(+)='404'"
      
   '98/5/1 以後完稿且承辦人為外翻人員
   strExc(0) = strExc(0) & " union" & _
      " select SQLDATET(EP09) C01,substrb(' '||sqldatet(nvl(c1.CP27,c1.CP57)),-9)||decode(c1.cp57,null,'','#') C12" & _
      ",st02||' '||c1.cp14 C02" & _
      ",DECODE(c2.CP01,NULL,' ','*')||pa01||'-'||pa02||decode(pa03||pa04,'000','',pa03||pa04) C03" & _
      ",pa05 C04,tf23 C05,tf21 C06,tf04 C07,tf05||decode(tf05,null,'','%') C08" & _
      ",tf06||decode(tf06,null,'','%') C09,tf18||decode(tf18,null,'','%') C10" & _
      ",ep09,pa01,pa02,pa03,pa04,c1.cp14 C11" & _
      " from engineerprogress,caseprogress c1,staff,patent,transfee,caseprogress c2" & _
      " where ep09>=20090501" & stCon & stConEP09 & _
      " and c1.cp09(+)=ep02 and c1.cp10='201' and st01(+)=c1.cp14 and not " & stSQLInner & _
      " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04" & _
      " and tf01(+)=c1.cp09 AND c2.CP43(+)=c1.CP09 AND c2.CP10(+)='404'"
   
   '98/5/1 以後完稿且承辦人非外翻人員
   strExc(0) = strExc(0) & " union" & _
      " select SQLDATET(EP09) C01,substrb(' '||sqldatet(nvl(c1.CP27,c1.CP57)),-9)||decode(c1.cp57,null,'','#') C12" & _
      ",st02||' '||c1.cp14 C02" & _
      ",DECODE(c2.CP01,NULL,' ','*')||pa01||'-'||pa02||decode(pa03||pa04,'000','',pa03||pa04) C03" & _
      ",pa05 C04,tf23 C05,nvl(nvl(tf21,tf03),tf23) C06,tf04 C07,tf05||decode(tf05,null,'','%') C08" & _
      ",tf06||decode(tf06,null,'','%') C09,tf18||decode(tf18,null,'','%') C10" & _
      ",ep09,pa01,pa02,pa03,pa04,c1.cp14 C11" & _
      " from engineerprogress,caseprogress c1,staff,patent,transfee,caseprogress c2" & _
      " where ep09>=20090501" & stCon & stConCP27 & _
      " and c1.cp09(+)=ep02 and c1.cp10='201' and st01(+)=c1.cp14 and " & stSQLInner & _
      " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04" & _
      " and tf01(+)=c1.cp09 AND c2.CP43(+)=c1.CP09 AND c2.CP10(+)='404'"
      
   'Added by Morgan 2025/11/12
   'CFP案
   If Left(Pub_StrUserSt03, 1) = "M" Then
      strExc(0) = strExc(0) & " union" & _
         " select SQLDATET(EP09) C01,substrb(' '||sqldatet(nvl(c1.CP27,c1.CP57)),-9)||decode(c1.cp57,null,'','#') C12" & _
         ",st02||' '||c1.cp14 C02" & _
         ",pa01||'-'||pa02||decode(pa03||pa04,'000','',pa03||pa04) C03" & _
         ",pa05 C04,tf23 C05,tf21 C06,tf04 C07,tf05||decode(tf05,null,'','%') C08" & _
         ",tf06||decode(tf06,null,'','%') C09,tf18||decode(tf18,null,'','%') C10" & _
         ",ep09,pa01,pa02,pa03,pa04,c1.cp14 C11" & _
         " from engineerprogress,caseprogress c1,staff,patent,transfee" & _
         " where cp01='CFP' and cp14 like 'F%'" & stConCFP & stConEP09 & _
         " and c1.cp09(+)=ep02 and st01(+)=c1.cp14" & _
         " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04" & _
         " and tf01(+)=c1.cp09 and tf23>0"
   End If
      
   If Me.Combo2.ListIndex = 2 Then
      strExc(0) = strExc(0) & " order by C12,pa01,pa02,pa03,pa04,C11"
   Else
      strExc(0) = strExc(0) & " order by C01,pa01,pa02,pa03,pa04,C11"
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   m_iCols = 11
   SetGrid RsTemp
   RecordShow
   If RsTemp.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/7
      MsgBox "查無資料！", vbInformation
   Else
      InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2010/12/7
   End If
   
flgErr:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
    
End Sub

Private Sub SetGrid(p_Rst As ADODB.Recordset)
   Dim iCol As Integer
   With grdDataList
      .Visible = False
      Set .Recordset = p_Rst.Clone
      .FormatString = .FormatString
      For iCol = 0 To 1
         .ColAlignment(iCol) = flexAlignCenterCenter
      Next
      For iCol = 2 To 4
         .ColAlignment(iCol) = flexAlignLeftCenter
      Next
      For iCol = 5 To 10
         .ColAlignment(iCol) = flexAlignRightCenter
      Next
      For iCol = 11 To .Cols - 1
         .ColWidth(iCol) = 0
      Next
      .Visible = True
   End With
End Sub
Private Sub cmdQuery_Click()
    Screen.MousePointer = vbHourglass
    If TxtValidate Then doQuery
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Label4(0).BackColor = Label3.BackColor
   Label4(1).BackColor = Label3.BackColor
   'Modify by Amy 2014/07/14 第一次DB可能沒資料所以先抓client
   'Modified by Morgan 2016/3/29 會有一個以上的人使用,加員工號;取消抓client
   'If PUB_GetLastDate(Me.Name, "txtNo(0)") <> "" Then
   '     Label4(0) = PUB_GetLastDate(Me.Name, "txtNo(0)")
   'Else
   '     Label4(0) = GetSetting("TAIE", "FCP", Me.Name & "#DATE01", "")
   'End If
   'If PUB_GetLastDate(Me.Name, "txtNo(1)") <> "" Then
   '     Label4(0) = PUB_GetLastDate(Me.Name, "txtNo(1)")
   'Else
   '     Label4(1) = GetSetting("TAIE", "FCP", Me.Name & "#DATE02", "")
   'End If
   Label4(0) = PUB_GetLastDate(Me.Name, "txtNo(0)" & "-" & strUserNum)
   Label4(1) = PUB_GetLastDate(Me.Name, "txtNo(1)" & "-" & strUserNum)
   'end 2016/3/29
   'end 2014/07/14
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter, , , , , True
   
   Combo2.AddItem "3.所內同仁下班翻譯案件(含 F5595巨京,F5639寰華)", 0
   Combo2.ItemData(0) = 3
   'Modified by Morgan 2017/8/30 +F5698 迅達翻譯 --婧瑄
   'Modified by Lydia 2018/01/04
   'Combo2.AddItem "2.承辦人為F編號案件(排除 F5588舜禹,F5595巨京,F5639寰華,F5614通用,F5653捷恩凱,F5698迅達)", 0
   'Modified by Lydia 2025/03/13 改用模組取得
   'Combo2.AddItem "2.承辦人為F編號案件(排除 " & 外翻_舜禹 & "舜禹,F5595巨京,F5639寰華,F5614通用," & 外翻_捷恩凱 & "捷恩凱," & 外翻_迅達 & "迅達)", 0
   Combo2.AddItem "2.承辦人為F編號案件(排除 " & Pub_SetF51Order("F", "2") & ",F5595巨京,F5639寰華,F5614通用)", 0
   Combo2.ItemData(0) = 2
   Combo2.AddItem "1.全部", 0
   Combo2.ItemData(0) = 1
   
   '財務處預設1
   If Pub_StrUserSt03 = "M31" Then
     Combo2.ListIndex = 1
   Else
     Combo2.ListIndex = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   MenuEnabled
   Set frm060203 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub


Private Sub txtCP14_GotFocus(Index As Integer)
   If Index = 1 And txtCP14(1) = "" And txtCP14(0) <> "" Then
      txtCP14(1) = txtCP14(0)
   End If
   TextInverse txtCP14(Index)
End Sub

Private Sub txtCP14_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtNo_GotFocus(Index As Integer)
   If Index = 1 Then
      If txtNo(0) <> "" And txtNo(1) = "" Then
         txtNo(1) = txtNo(0)
      End If
   ElseIf Index = 3 Then
      If txtNo(2) <> "" And txtNo(3) = "" Then
         txtNo(3) = txtNo(2)
      End If
   End If
   TextInverse txtNo(Index)
End Sub

Private Sub txtNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtNo_Validate(Index As Integer, Cancel As Boolean)
   If txtNo(Index) <> "" Then
      If Not ChkDate(txtNo(Index)) Then
        Cancel = True
      End If
   End If
End Sub


Private Function TxtValidate() As Boolean
   
   Dim bolCancel As Boolean, ii As Integer
   
   If txtNo(0) = "" And txtNo(3) = "" Then
      MsgBox "完稿日條件不可空白！", vbExclamation
      txtNo(0).SetFocus
      Exit Function
   End If
   
   For ii = 0 To 3
      bolCancel = False
      Call txtNo_Validate(ii, bolCancel)
      If bolCancel Then
         Exit Function
      End If
   Next
   
   TxtValidate = True

End Function
