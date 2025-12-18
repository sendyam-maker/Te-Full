VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060207 
   BorderStyle     =   1  '單線固定
   Caption         =   "重新核稿案件明細查詢/列印"
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
      TabIndex        =   2
      Top             =   570
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
      TabIndex        =   8
      Top             =   540
      Width           =   3495
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
      Enabled         =   0   'False
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   5
      Top             =   90
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4305
      Left            =   135
      TabIndex        =   7
      Top             =   930
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   7594
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "發文日|本所案號|案件名稱|原核稿人|原核稿點數|新承辦人|新核稿點數"
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
      _Band(0).Cols   =   7
   End
   Begin MSForms.Label lblName 
      Height          =   300
      Left            =   2400
      TabIndex        =   11
      Top             =   600
      Width           =   1455
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
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
      TabIndex        =   10
      Top             =   600
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
      TabIndex        =   9
      Top             =   570
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發文日期：                         －"
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
      TabIndex        =   6
      Top             =   270
      Width           =   2295
   End
End
Attribute VB_Name = "frm060207"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改 (Printer列印未改)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Create by Morgan 2010/12/31
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
            '案件名稱
            If iCol = 3 Then
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 25)
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
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(3) = PLeft(2) + Printer.TextWidth(String(7, "　")) + ciColGap
   PLeft(4) = PLeft(3) + Printer.TextWidth(String(25, "　")) + ciColGap
   PLeft(5) = PLeft(4) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(6) = PLeft(5) + Printer.TextWidth(String(5, "　")) + ciColGap
   PLeft(7) = PLeft(6) + Printer.TextWidth(String(4, "　")) + ciColGap
   PLeft(8) = PLeft(7) + Printer.TextWidth(String(5, "　")) + ciColGap
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
         Case 1, 5, 7
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
   strPTmp = "重新核稿案件明細表"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   PrintNewLine
   strPTmp = "發文日："
   Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(8, "　"))
   Printer.CurrentY = iPrint
   Printer.Print strPTmp & CFDate(txtNo(0).Tag) & " － " & IIf(txtNo(1) <> "", CFDate(txtNo(1).Tag), "")
   
   If txtCP14.Tag <> "" Then
      PrintNewLine
      strPTmp = "承辦人："
      Printer.CurrentX = (lngPageWidth) / 2 - Printer.TextWidth(String(8, "　"))
      Printer.CurrentY = iPrint
      Printer.Print strPTmp & txtCP14.Tag & " " & lblName.Tag
   End If

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
    For intI = 1 To m_iCols
      Select Case intI
         Case 5, 7
            Printer.CurrentX = PLeft(intI + 1) - ciColGap - Printer.TextWidth(grdDataList.TextMatrix(0, intI - 1))
            Printer.CurrentY = iPrint
            Printer.Print grdDataList.TextMatrix(0, intI - 1)
         Case Else
            Printer.CurrentX = PLeft(intI)
            Printer.CurrentY = iPrint
            Printer.Print grdDataList.TextMatrix(0, intI - 1)
      End Select
    Next
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    PrintLine
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    PrintLine
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
    
    txtNo(0).Tag = txtNo(0).Text
    txtNo(1).Tag = txtNo(1).Text
    txtCP14.Tag = txtCP14.Text
    lblName.Tag = lblName.Caption
    
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
   
    '發文日
    If txtNo(0) <> "" Then
      stCon = stCon & " and c1.CP27>=" & DBDATE(txtNo(0))
    End If
    If txtNo(1) <> "" Then
      stCon = stCon & " and c1.CP27<=" & DBDATE(txtNo(1))
    End If
    If txtNo(0) <> "" Or txtNo(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1, 5) & txtNo(0) & "-" & txtNo(1) 'Add By Sindy 2010/12/7
    End If
    
    '承辦人
    If txtCP14 <> "" Then
      stCon = stCon & " and c1.cp14='" & txtCP14 & "'"
      pub_QL05 = pub_QL05 & ";" & Label2(2) & txtCP14 & lblName 'Add By Sindy 2010/12/7
    End If
   
   strExc(0) = "select substrb(' '||sqldatet(c1.CP27),-9) 發文日" & _
      ",c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) 本所案號" & _
      ",nvl(nvl(nvl(nvl(nvl(pa05,pa06),pa07),sp05),sp06),sp07) 案件名稱" & _
      ",s2.st02 原核稿人" & _
      ",round(decode(substr(c2.cp60,1,1),'E', c2.cp18*(0.3+0.7*(1-nvl(TF06,100)/100)),decode(a1k25,null,a1n05)),2) 原核稿點數" & _
      ",s1.st02 新承辦人" & _
      ",round(decode(substr(c2.cp60,1,1),'E',c2.cp18,(a1l05-nvl(a1l07,0))/1000)*0.3,2) 新核稿點數" & _
      " from caseprogress c1,caseprogress c2,engineerprogress,patent,servicepractice,staff s1,staff s2,acc1n0,acc1k0,acc1l0,transfee" & _
      " where c1.cp10='229' and c1.cp01 in ('FCP','P','PS','FG') and substr(c1.cp12,1,1)='F'" & stCon & _
      " and c2.cp01(+)=c1.cp01 and c2.cp02(+)=c1.cp02 and c2.cp03(+)=c1.cp03 and c2.cp04(+)=c1.cp04 and c2.cp10(+)='201'" & _
      " and ep02(+)=c2.cp09" & _
      " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04" & _
      " and sp01(+)=c1.cp01 and sp02(+)=c1.cp02 and sp03(+)=c1.cp03 and sp04(+)=c1.cp04" & _
      " and s1.st01(+)=c1.cp14 and s2.st01(+)=ep04" & _
      " and a1n02(+)='2' and a1n03(+)=ep02 and a1n04(+)=ep04 and a1n06(+)='Y'" & _
      " and a1k01(+)=c2.cp60 and a1l01(+)=c2.cp60 and a1l04(+)=c2.cp10" & _
      " and tf01(+)=c2.cp09 order by 1"
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   m_iCols = 7
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
      .ColAlignment(0) = flexAlignRightCenter
      .ColAlignment(4) = flexAlignRightCenter
      .ColAlignment(6) = flexAlignRightCenter
      
      .ColWidth(0) = 795
      .ColWidth(1) = 1320
      .ColWidth(2) = 2400
      .ColWidth(3) = 855
      .ColWidth(4) = 1050
      .ColWidth(5) = 870
      .ColWidth(6) = 1065
      
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
   cmdPrint.Enabled = IsUserHasRightOfFunction(Me.Name, strPrint, False)
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   
   lblName.Caption = "" 'Add By Sindy 2021/12/16
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   MenuEnabled
   Set frm060207 = Nothing
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

Private Sub txtNo_GotFocus(Index As Integer)
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
      MsgBox "發文日條件不可空白！", vbExclamation
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
