VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170105 
   BorderStyle     =   1  '單線固定
   Caption         =   "補充保費繳納作業"
   ClientHeight    =   5160
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8232
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8232
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1125
      MaxLength       =   1
      TabIndex        =   0
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3195
      MaxLength       =   1
      TabIndex        =   15
      Text            =   "1"
      Top             =   593
      Width           =   285
   End
   Begin VB.Frame Frame1 
      Caption         =   "印表機"
      Height          =   570
      Left            =   135
      TabIndex        =   12
      Top             =   1260
      Width           =   7905
      Begin VB.ComboBox cmbPrinter 
         Height          =   300
         Left            =   90
         Style           =   2  '單純下拉式
         TabIndex        =   13
         Top             =   180
         Width           =   7695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消繳費日期(&C)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4410
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   4650
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "紀錄繳費日期(&U)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   2430
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   4650
      Width           =   1875
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   6120
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7155
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5085
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox txtPayDate 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   5
      Top             =   4710
      Width           =   915
   End
   Begin VB.TextBox txtPayDate 
      Height          =   285
      Index           =   0
      Left            =   1125
      MaxLength       =   5
      TabIndex        =   2
      Top             =   240
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   2685
      Left            =   135
      TabIndex        =   11
      Top             =   1890
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   4741
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "補充保費繳費作業完成 ,請""立即填入繳費日期 "", 以確保資料不會被異動!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   150
      TabIndex        =   17
      Top             =   960
      Width           =   6150
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1530
      X2              =   1845
      Y1              =   728
      Y2              =   728
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公  司  別："
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   16
      Top             =   645
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "查詢別：          (1.保險對象 2.投保單位)"
      Height          =   180
      Index           =   2
      Left            =   2430
      TabIndex        =   14
      Top             =   645
      Width           =   3045
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "繳費日期："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   10
      Top             =   4770
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "給付年月："
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   9
      Top             =   292
      Width           =   900
   End
End
Attribute VB_Name = "frm170105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/27 Form2.0已修改 (Printer列印未改)
'Created by Morgan 2013/3/5
Option Explicit

Dim PLeft() As Integer, PColName() As String
Dim iPrint As Integer, iPage As Integer
Dim m_iStartX As Integer, m_iStartY As Integer, m_iColGap As Integer
Dim m_iPageHeight As Long, m_iLineHeight As Long, m_iMargin As Long
Dim stLstComp As String, stLstCompName As String
Dim m_DefaultPrinter As String


Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer
   With grdDataList
      .Visible = False
      If Text1 = "1" Then
         If p_bolHeaderOnly = False Then
            .Clear
            .Rows = 2: .Cols = 6: .FixedRows = 1: .FixedCols = 0
         End If
         .row = 0
         .col = 0: .ColWidth(.col) = 500: .Text = "公司"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 1: .ColWidth(.col) = 900: .Text = "統編"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 2: .ColWidth(.col) = 1650: .Text = "單位名稱"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 3: .ColWidth(.col) = 2250: .Text = "所得類別"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 4: .ColWidth(.col) = 1400: .Text = "應繳補充保費"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 5: .ColWidth(.col) = 900: .Text = "繳費日期"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      Else
         If p_bolHeaderOnly = False Then
            .Clear
            .Rows = 2: .Cols = 6: .FixedRows = 1: .FixedCols = 0
         End If
         .row = 0
         .col = 0: .ColWidth(.col) = 500: .Text = "公司"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 1: .ColWidth(.col) = 900: .Text = "統編"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 2: .ColWidth(.col) = 1800: .Text = "單位名稱"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 3: .ColWidth(.col) = 1400: .Text = "薪資所得總額"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 4: .ColWidth(.col) = 1600: .Text = "受僱者投保總額"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 5: .ColWidth(.col) = 1400: .Text = "應繳補充保費"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      End If
      .Refresh
      .Visible = True
   End With
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   If cmbPrinter <> Printer.DeviceName Then
      PUB_RestorePrinter cmbPrinter
   End If
   PrintSheet
   '若印表機變動, 則更新列印設定
   If cmbPrinter.Tag <> cmbPrinter Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, 0, 0, Me.cmbPrinter.Text
   End If
   If Printer.DeviceName <> m_DefaultPrinter Then
      PUB_RestorePrinter m_DefaultPrinter
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   If TxtValidate = True Then
      SetDataListWidth
      doQuery
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
   Screen.MousePointer = vbHourglass
   If TxtValidate(1) = True Then
      Me.Enabled = False
      If UpdateDate = True Then
         txtPayDate(1).Enabled = False
         Command1.Enabled = False
         Command2.Enabled = True
         cmdSearch.Value = True
      End If
      Me.Enabled = True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Function UpdateDate() As Boolean
   Dim stCon As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stCon = " and substr(nhi02,1,6)=" & Val(txtPayDate(0)) + 191100
   If txt1(0) <> "" Then
      stCon = stCon & " and nhi11>='" & txt1(0) & "'"
   End If
   
   If txt1(1) <> "" Then
      stCon = stCon & " and nhi11<='" & txt1(1) & "'"
   End If
   
   strSql = "update nhi2nd set nhi12=" & DBDATE(txtPayDate(1)) & " where nhi12 is null" & stCon
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
      
   cnnConnection.CommitTrans
   UpdateDate = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

EscPoint:
End Function

Private Function CancelDate() As Boolean
   Dim stCon As String
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   stCon = " and substr(nhi02,1,6)=" & Val(txtPayDate(0)) + 191100
   If txt1(0) <> "" Then
      stCon = stCon & " and nhi11>='" & txt1(0) & "'"
   End If
   
   If txt1(1) <> "" Then
      stCon = stCon & " and nhi11<='" & txt1(1) & "'"
   End If
  
   strSql = "update nhi2nd set nhi12=null where nhi12=" & DBDATE(txtPayDate(1)) & stCon
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql, intI
      
   cnnConnection.CommitTrans
   CancelDate = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox Err.Description, vbCritical

EscPoint:
End Function

Private Sub Command2_Click()
   If MsgBox("是否確定要取消？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   If CancelDate = True Then
      txtPayDate(1).Enabled = True
      Command1.Enabled = True
      Command2.Enabled = False
      cmdSearch.Value = True
   End If
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   PUB_SetPrinter Me.Name, cmbPrinter, m_DefaultPrinter
   
   If strUserNum = "86021" Then
      txt1(0) = "A": txt1(0).Enabled = False
      txt1(1) = "A": txt1(1).Enabled = False
   End If
   
   '預設最小未繳費月份
   'Modified by Morgan 2013/5/2 改15號以前預設上上月,以後預設上月
   'strExc(0) = "select min(nhi02) from nhi2nd where nhi12 is null"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'If intI = 1 Then
   '   strExc(1) = TransDate(RsTemp(0), 1) \ 100
   '   txtPayDate(0) = strExc(1)
   'End If
   If Val(Right(strSrvDate(1), 2)) > 15 Then
      strExc(1) = CompDate(1, -1, strSrvDate(1))
   Else
      strExc(1) = CompDate(1, -2, strSrvDate(1))
   End If
   txtPayDate(0) = (strExc(1) \ 100) - 191100
   'end 2013/5/2
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170105 = Nothing
End Sub

Private Function TxtValidate(Optional pMode As Integer = 0) As Boolean
   Dim bCancel As Boolean
   '查詢
   If pMode = 0 Then
      If txtPayDate(0) = "" Then
         MsgBox "請輸入給付年月!", vbInformation
         txtPayDate(0).SetFocus
         Exit Function
      End If
      If Text1 = "" Then
         MsgBox "請輸入查詢別!", vbInformation
         Text1.SetFocus
         Exit Function
      End If
      
      'Added by Morgan 2013/5/20
      If CheckUnPaidRecord = True Then
         Exit Function
      End If
      'end 2013/5/20
      
   '紀錄繳費日期
   Else
      If txtPayDate(1) = "" Then
         MsgBox "請輸入繳費日期!", vbInformation
         txtPayDate(1).SetFocus
         Exit Function
      End If
      txtPayDate_Validate 1, bCancel
      If bCancel = True Then
         txtPayDate(1).SetFocus
         Exit Function
      End If
   End If
   TxtValidate = True
End Function
   
Private Sub doQuery()
   Dim stVTB As String
   Dim stYM As String
   Dim stConNhi As String, stConSm As String
   Dim stNhiRate As String
   
   
   stYM = Val(txtPayDate(0)) + 191100
   stConSm = " and sm02=" & stYM
   stConNhi = " and nhi02>=" & stYM & "01 and nhi02<=" & stYM & "31"
   stNhiRate = PUB_GetNhiRate(stYM & "01") 'Added by Morgan 2016/2/23
   
   If txt1(0) <> "" Then
      stConSm = stConSm & " and sm37>='" & txt1(0) & "'"
      stConNhi = stConNhi & " and nhi11>='" & txt1(0) & "'"
   End If
   
   If txt1(1) <> "" Then
      stConSm = stConSm & " and sm37<='" & txt1(1) & "'"
      stConNhi = stConNhi & " and nhi11<='" & txt1(1) & "'"
   End If
   
   If Text1 = "1" Then
      stVTB = "select nhi11,x01,decode(x01,'1','四個月以上投保金額的獎金'" & _
         ",'2','兼職所得','3','執行業務收入','4','股利所得','5','利息所得','6','租金收入',x01) x02" & _
         ",sum(nhi06) x03,max(nhi12) x04 from ( select nhi11,decode(nhi03,'50',decode(sign(nhi05),1,'1','2')" & _
         ",'9A','3','9B','3','54','4','52','5','5A','5','5B','5','5C','5','51','6',nhi03) x01" & _
         ",nhi06,nhi12 from nhi2nd where 1=1 " & stConNhi & _
         ") x group by nhi11,x01"
   
      strExc(0) = "select nhi11,a0807,a0802,x02,x03,sqldatet(x04) x04" & _
         " from (" & stVTB & "),acc080 where a0801(+)=nhi11 order by nhi11,sign(x03) desc,x01"
   Else
      'Modified by Morgan 2013/3/29 排除未投保員工(當月離職者)及健保補助類別為低收者
      'Modified by Morgan 2014/2/26 73014,73017,73035 103/1轉換公司別,薪資所得-月薪資也要抓第3碼為A的員工
      'Modified by Morgan 2016/2/23 補充保費費率改抓設定
      'Modified by Morgan 2018/1/5 +sm04>0 也要 ,不管有無健保投保金額都要列出(Ex.退休退保)--辜
      'Modified by Morgan 2020/4/28 桂所長109.3月健保以非合夥人身分投保(算受雇人)--辜
      'Modify By Sindy 2020/6/25 + 證照津貼
      'Modified by Morgan 2022/8/30 +判斷有提撥的不管是否已設定為合夥人身分投保(未生效月份 Ex:蔣律師)--辜
      'Modified by Morgan 2024/2/27 投保身分可能會改，未投保改判斷沒有健保費及健保明細者 Ex:B2024 113/1月健保身分由低收改無
      stVTB = "select sm37,nvl(x2,0)+nvl(y2,0) C2,x3 c3,round((nvl(x2,0)+nvl(y2,0)-x3)*" & stNhiRate & "/100) c4" & _
         " from (select sm37,SUM((nvl(SM04,0)+nvl(SM05,0)+nvl(SM45,0)-nvl(SM21,0)+nvl(SM28,0))) x2" & _
         ",sum(decode(decode(sm01||sm02,'76012202003','N',decode(nvl(sm30,0),0,sd11)),'Y',0,decode(sign(sm15),1,1,decode(hm01,null,0,(decode(hm06,'12',0,1)))))*sm42) x3" & _
         " From salarymonth, salarydata,himonth where sm01<'F' and (sm04>0 or  sm42>0 ) " & stConSm & " and sd01(+)=sm01" & _
         " and hm03(+)=sm02 and hm01(+)=sm01 and hm02(+)=0 group by sm37) x,(select nhi11,sum(nhi07) y2" & _
         " from nhi2nd where nhi03='50'" & stConNhi & _
         " group by nhi11) y" & _
         " where nhi11(+)=sm37"
         
      strExc(0) = "select sm37,a0807,a0802,C2,c3,decode(sign(c4),1,c4,0) c4" & _
         " from (" & stVTB & "),acc080 where a0801(+)=sm37 order by sm37"
   End If
  
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set grdDataList.Recordset = RsTemp.Clone
   SetDataListWidth True
   If intI = 1 Then
      cmdPrint.Enabled = True
      
      If Text1 = "1" Then
         '檢查繳費日及申報日
         strExc(0) = "select nhi12,nhi09 from nhi2nd where nhi12>0" & stConNhi
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            txtPayDate(1) = TransDate(RsTemp(0), 1)
            '未申報
            If IsNull(RsTemp(1)) Then
               Command2.Enabled = True
            End If
         Else
            txtPayDate(1) = strSrvDate(2)
            txtPayDate(1).Enabled = True
            Me.Command1.Enabled = True
         End If
      End If
   Else
      cmdPrint.Enabled = False
   End If
End Sub

Private Sub Text1_Change()
   If cmdPrint.Enabled Then
      SetDataListWidth
      disableObj
   End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
End Sub

Private Sub txt1_Change(Index As Integer)
   If cmdPrint.Enabled Then
      SetDataListWidth
      disableObj
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPayDate_Change(Index As Integer)
   If Index = 0 Then
      If cmdPrint.Enabled Then
         SetDataListWidth
         disableObj
      End If
   End If
End Sub

Private Sub disableObj()
   cmdPrint.Enabled = False
   Command1.Enabled = False
   Command2.Enabled = False
   txtPayDate(1) = ""
   txtPayDate(1).Enabled = False
End Sub

Private Sub PrintSheet()
   
   Dim ii As Integer, iCompCol As Integer, stLstGroup As String, iLstPage As Integer
   
   GetPleft
   
   With grdDataList
   iPage = 0
   PrintPageHeader
   PrintPageHeader1
   For ii = 1 To .Rows - 1
      PrintNewLine
      PrintDetail ii
   Next
   Printer.EndDoc
   End With
   MsgBox "列印完畢！"
End Sub

Private Sub PrintDetail(iRow As Integer)
   Dim iCol As Integer
   Printer.Font.Size = 11
    
   With grdDataList
    
   If Text1 = "1" Then
      For iCol = 1 To UBound(PColName)
        Select Case iCol
        Case 5
           If .TextMatrix(iRow, iCol - 1) <> "" Then
              strExc(0) = Format(.TextMatrix(iRow, iCol - 1), DDollar2)
              Printer.CurrentX = PLeft(iCol + 1, 1) - Printer.TextWidth(strExc(0)) - m_iColGap
              Printer.CurrentY = iPrint
              Printer.Print strExc(0)
           End If
        Case Else
           Printer.CurrentX = PLeft(iCol, 1)
           Printer.CurrentY = iPrint
           Printer.Print .TextMatrix(iRow, iCol - 1)
        End Select
      Next
   Else
      For iCol = 1 To UBound(PColName)
        Select Case iCol
        Case 4, 5, 6
           If .TextMatrix(iRow, iCol - 1) <> "" Then
              strExc(0) = Format(.TextMatrix(iRow, iCol - 1), DDollar2)
              Printer.CurrentX = PLeft(iCol + 1, 1) - Printer.TextWidth(strExc(0)) - m_iColGap
              Printer.CurrentY = iPrint
              Printer.Print strExc(0)
           End If
        Case Else
           Printer.CurrentX = PLeft(iCol, 1)
           Printer.CurrentY = iPrint
           Printer.Print .TextMatrix(iRow, iCol - 1)
        End Select
      Next
   End If
   End With
End Sub

Private Sub GetPleft()
   Dim ii As Integer
   
   Printer.PaperSize = 9
   Printer.Orientation = 1
   Printer.FontSize = 12
   m_iStartX = 300
   m_iStartY = 500
   m_iPageHeight = Printer.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = (Printer.Height - Printer.ScaleHeight) / 2
   m_iColGap = 150 '欄位間隔
   
   With grdDataList
  
   If Text1 = "1" Then
      ReDim PLeft(7, 2)
      ReDim PColName(6)
      ii = 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"公司"
      PLeft(ii, 1) = m_iStartX
      PLeft(ii, 2) = 2
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"統編"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 4
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"單位名稱"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 13
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"所得類別"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 12
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"應繳補充保費"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 6
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"繳費日期"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 4
   Else
   
      ReDim PLeft(7, 2)
      ReDim PColName(6)
      ii = 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"公司"
      PLeft(ii, 1) = m_iStartX
      PLeft(ii, 2) = 2
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"統編"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 4
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"單位名稱"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 14
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"薪資所得總額"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 6
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"受僱者投保總額"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 7
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii - 1) '"應繳補充保費"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 6
   
   End If
   ii = ii + 1
   PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
   End With
End Sub


Private Sub PrintPageHeader()
   Dim strTmp As String
   
   If Text1 = "1" Then
      strTmp = "台一關係企業 保險對象補充保費統計表"
   Else
      strTmp = "台一關係企業 投保單位補充保費統計表"
   End If
   
   iPrint = m_iStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = 20
   Printer.Font.Bold = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   iPrint = iPrint + 600
   
   Printer.Font.Size = 11
   Printer.Font.Bold = False
   
  
   strTmp = "公司別：" & txt1(0) & " - " & txt1(1)
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(String(10, "　"))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   
   strTmp = "給付年月：" & txtPayDate(0)
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(String(10, "　"))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   iPrint = iPrint + m_iLineHeight
   
   iPage = iPage + 1
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
   
   PrintNewLine
   
End Sub


Private Sub PrintPageHeader1()
   
   iPrint = iPrint + m_iLineHeight
   Printer.Font.Size = 11
   Printer.FontBold = True
   
   For intI = 1 To UBound(PColName)
      Select Case intI
      Case 5
         Printer.CurrentX = PLeft(intI + 1, 1) - Printer.TextWidth(PColName(intI)) - m_iColGap
         Printer.CurrentY = iPrint
         Printer.Print PColName(intI)
      Case Else
         Printer.CurrentX = PLeft(intI, 1)
         Printer.CurrentY = iPrint
         Printer.Print PColName(intI)
      End Select
   Next
   iPrint = iPrint + m_iLineHeight
   DrawLine
End Sub

Private Sub DrawLine()
   Printer.DrawWidth = 5
   Printer.Line (PLeft(1, 1), iPrint)-(PLeft(UBound(PLeft, 1), 1), iPrint)
   iPrint = iPrint - m_iLineHeight / 2
End Sub

Private Sub PrintNewLine(Optional ByVal p_iExtraLines As Integer = 2)

   iPrint = iPrint + m_iLineHeight
   If iPrint >= (m_iPageHeight - m_iMargin - p_iExtraLines * m_iLineHeight) Then
      DrawLine
      Printer.NewPage
      PrintPageHeader
      PrintPageHeader1
      iPrint = iPrint + m_iLineHeight
   End If
   
End Sub

Private Sub txtPayDate_GotFocus(Index As Integer)
   TextInverse txtPayDate(Index)
   CloseIme
End Sub

Private Sub txtPayDate_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtPayDate_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then
      If txtPayDate(Index) <> "" Then
         If ChkDate(txtPayDate(Index)) = False Then
            Cancel = True
         End If
      End If
   End If
End Sub
'檢查是否有較早補充保費未上繳費日期
Private Function CheckUnPaidRecord() As Boolean
   Dim stCon As String
   stCon = " and nhi02<" & Val(txtPayDate(0) & "01") + 19110000
   If txt1(0) <> "" Then
      stCon = stCon & " and nhi11>='" & txt1(0) & "'"
   End If
   If txt1(1) <> "" Then
      stCon = stCon & " and nhi11<='" & txt1(1) & "'"
   End If
   strExc(0) = "select nvl(min(nhi02),0) from nhi2nd where nhi12 is null" & stCon
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp(0) > 0 Then
         CheckUnPaidRecord = True
         MsgBox Format((RsTemp(0) \ 100 - 191100), "###/##") & "補充保費尚未紀錄繳費日期，不可查詢！", vbExclamation
      End If
   End If
End Function
