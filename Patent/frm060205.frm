VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060205 
   BorderStyle     =   1  '單線固定
   Caption         =   "翻譯案件承辦期限查詢/列印"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   8955
   Begin VB.TextBox txtCP14 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1215
      MaxLength       =   6
      TabIndex        =   3
      Top             =   840
      Width           =   1110
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   11
      Top             =   815
      Width           =   3495
   End
   Begin VB.TextBox txtDept 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1215
      MaxLength       =   1
      TabIndex        =   2
      Top             =   540
      Width           =   525
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   1
      Left            =   2475
      MaxLength       =   7
      TabIndex        =   1
      Top             =   235
      Width           =   1080
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6750
      TabIndex        =   5
      Top             =   90
      Width           =   1200
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   1215
      MaxLength       =   7
      TabIndex        =   0
      Top             =   235
      Width           =   1080
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5670
      TabIndex        =   4
      Top             =   90
      Width           =   1020
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   8010
      TabIndex        =   6
      Top             =   90
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3825
      Left            =   135
      TabIndex        =   8
      Top             =   1170
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   6747
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "承辦期限　|承辦人　　　|本所案號　　　|案件名稱　　|進度備註 "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSForms.Label lblName 
      Height          =   300
      Left            =   2370
      TabIndex        =   14
      Top             =   870
      Width           =   1635
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2884;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   13
      Top             =   875
      Width           =   780
   End
   Begin VB.Label Label2 
      Caption         =   "印表機"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   12
      Top             =   845
      Width           =   975
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   " 翻譯該程序若有延期，本所案號前作 * 標記"
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   5070
      Width           =   3465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "對象：    　               (1:外譯編號 2:全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   570
      Width           =   3240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "承辦期限：                         －"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   270
      Width           =   2295
   End
End
Attribute VB_Name = "frm060205"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改 (Printer列印未改)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
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
            '案件名稱抓10個字
            'If iCol = 4 Or iCol = 5 Then
            If iCol >= 4 Then
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 20)
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
      '案件名稱可印10個中文
      If intI = 5 Then
         PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(String(20, "　")) + ciColGap
      ElseIf intI = 3 Then
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
      Printer.Print String(125, "-")
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
      
      If iCol > 4 Then
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) + ciColGap
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
   strPTmp = "翻譯案件承辦期限查詢案件明細表"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   PrintNewLine
   strPTmp = "完稿日："
   Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(6, "　"))
   Printer.CurrentY = iPrint
   Printer.Print strPTmp & CFDate(txtNo(0)) & " － " & IIf(txtNo(1) <> "", CFDate(txtNo(1)), "")
    
   If txtDept = "1" Then
     strPTmp = "承辦人：外譯編號"
   Else
     strPTmp = "承辦人：全部"
   End If
   PrintNewLine
   Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(6, "　"))
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
       
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
   Printer.Print String(125, "-")
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    For intI = 1 To m_iCols
      If intI > 4 Then
         Printer.CurrentX = PLeft(intI) + ciColGap
         Printer.CurrentY = iPrint
         Printer.Print grdDataList.TextMatrix(0, intI - 1)
      Else
         Printer.CurrentX = PLeft(intI)
         Printer.CurrentY = iPrint
         Printer.Print grdDataList.TextMatrix(0, intI - 1)
      End If
    Next
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(125, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(125, "-")
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
      End If
   End If
End Sub

Private Sub doQuery()
Dim stCon As String
    
On Error GoTo flgErr
    
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
    
    'Modify by Morgan 2008/1/10 改控制承辦人部門為F21,F51,F52的
    '2008/4/8 MODIFY BY SONIA 加F81
    'stCon = " and c1.cp01='FCP'"
    stCon = " and ST15 in ('F21','F51','F52','F81')"
    
    '承辦期限
    If txtNo(0) <> "" Then
      stCon = stCon & " and c1.CP48>=" & DBDATE(txtNo(0))
    End If
    If txtNo(1) <> "" Then
      stCon = stCon & " and c1.CP48<=" & DBDATE(txtNo(1))
    End If
    If txtNo(0) <> "" Or txtNo(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1, 5) & txtNo(0) & "-" & txtNo(1) 'Add By Sindy 2010/12/7
    End If
    
    '對象 (1:外譯編號 2:全部)
    If txtDept = "1" Then
      stCon = stCon & " and SUBSTR(c1.CP14,1,1)='F'"
      pub_QL05 = pub_QL05 & ";" & Left(Label2(1), 3) & "1:外譯編號" 'Add By Sindy 2010/12/7
    Else
      pub_QL05 = pub_QL05 & ";" & Left(Label2(1), 3) & "2:全部" 'Add By Sindy 2010/12/7
    End If
    
    '承辦人
    If txtCP14 <> "" Then
      stCon = stCon & " and c1.cp14='" & txtCP14 & "'"
      pub_QL05 = pub_QL05 & ";" & Label2(2) & txtCP14 & lblName 'Add By Sindy 2010/12/7
    End If
    
   'Modify by Morgan 2010/8/13 百年蟲 substrb(' '||sqldatet(c1.CP48),-9)
   strExc(0) = "select distinct substrb(' '||sqldatet(c1.CP48),-9) C01" & _
   ",st02||' '||c1.cp14 C02" & _
   ",DECODE(c2.CP01,NULL,' ','*')||pa01||'-'||pa02||decode(pa03||pa04,'000','',pa03||pa04) C03" & _
   ",pa05 C4 ,c1.cp64 C5" & _
   " from engineerprogress,caseprogress c1,staff,patent,caseprogress c2" & _
   " where c1.cp09=ep02(+) and c1.cp10='201'" & stCon & _
   " and st01(+)=c1.cp14" & _
   " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04" & _
   " AND c2.CP43(+)=c1.CP09 AND c2.CP10(+)='404'" & _
   " and EP08 is NULL  and EP09 is Null " & _
   " order by C01,C02,C03"
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Morgan 2008/3/7 要固定10,否則報表會多印欄位
   'm_iCols = RsTemp.Fields.Count
   m_iCols = 5
   'end 2008/3/7
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
      .ColAlignment(0) = flexAlignLeftCenter
      .ColAlignment(4) = flexAlignLeftCenter
      

      .ColWidth(0) = 1000
      .ColWidth(1) = 1350
      .ColWidth(2) = 1200
      .ColWidth(3) = 2750
      .ColWidth(4) = 1800
      
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
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   
   '財務處預設1
   If Pub_StrUserSt03 = "M31" Then
     txtDept = "1"
   Else
     txtDept = "2"
   End If
   
   lblName.Caption = "" 'Add By Sindy 2021/12/16
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




Private Sub txtCP14_Change()
   Dim strTempName As String
   lblName = ""
   If Len(txtCP14) >= 5 Then
      If ClsPDGetStaff(txtCP14, strTempName) Then
         lblName = strTempName
      End If
   End If
End Sub

Private Sub txtCP14_GotFocus()
   TextInverse txtCP14
End Sub

Private Sub txtCP14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDept_GotFocus()
   TextInverse txtDept
End Sub

Private Sub txtDept_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtNo_GotFocus(Index As Integer)
   If Index = 1 Then
      If txtNo(0) <> "" And txtNo(1) = "" Then
         txtNo(1) = txtNo(0)
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
   
   Dim bolCancel As Boolean
   
   If txtNo(0) = "" Then
      MsgBox "承辦期限條件不可空白！", vbExclamation
      txtNo(0).SetFocus
      Exit Function
   End If
   
   bolCancel = False
   Call txtNo_Validate(0, bolCancel)
   If bolCancel Then
      Exit Function
   End If
   
   Call txtNo_Validate(1, bolCancel)
   If bolCancel Then
      Exit Function
   End If
   TxtValidate = True

End Function
