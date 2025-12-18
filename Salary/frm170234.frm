VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170234 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工補充健保費查詢及列印"
   ClientHeight    =   5520
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9420
   Begin VB.TextBox txtPayDate 
      Height          =   285
      Index           =   2
      Left            =   4140
      MaxLength       =   7
      TabIndex        =   18
      Top             =   810
      Width           =   960
   End
   Begin VB.TextBox txtPayDate 
      Height          =   285
      Index           =   1
      Left            =   2835
      MaxLength       =   7
      TabIndex        =   17
      Top             =   810
      Width           =   960
   End
   Begin VB.OptionButton Option1 
      Caption         =   "明細"
      Height          =   285
      Index           =   1
      Left            =   1305
      TabIndex        =   15
      Top             =   810
      Width           =   780
   End
   Begin VB.OptionButton Option1 
      Caption         =   "統計"
      Height          =   285
      Index           =   0
      Left            =   1305
      TabIndex        =   14
      Top             =   502
      Value           =   -1  'True
      Width           =   780
   End
   Begin VB.Frame Frame1 
      Caption         =   "印表機"
      Height          =   570
      Left            =   4185
      TabIndex        =   13
      Top             =   4920
      Width           =   5070
      Begin VB.ComboBox cmbPrinter 
         Height          =   300
         Left            =   135
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   210
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   6120
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   90
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   1980
      MaxLength       =   1
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1305
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtPayDate 
      Height          =   285
      Index           =   0
      Left            =   2835
      MaxLength       =   5
      TabIndex        =   2
      Top             =   502
      Width           =   960
   End
   Begin VB.TextBox txtStaffNo 
      Height          =   285
      Left            =   1305
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1140
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8490
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   7590
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   90
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3405
      Left            =   90
      TabIndex        =   8
      Top             =   1470
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   6011
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   7
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "查  詢  別："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   19
      Top             =   502
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   3825
      X2              =   4140
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "日期："
      Height          =   180
      Index           =   1
      Left            =   2250
      TabIndex        =   16
      Top             =   855
      Width           =   540
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1710
      X2              =   2025
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公  司  別："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   12
      Top             =   180
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "年月："
      Height          =   180
      Index           =   0
      Left            =   2250
      TabIndex        =   11
      Top             =   554
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "所得人編號："
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   1185
      Width           =   1080
   End
   Begin VB.Label lblName 
      Height          =   180
      Left            =   2790
      TabIndex        =   9
      Top             =   1185
      Width           =   2250
   End
End
Attribute VB_Name = "frm170234"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Morgan 2013/2/25
Option Explicit
Dim PLeft() As Integer, PColName() As String
Dim iPrint As Integer, iPage As Integer
Dim m_iStartX As Integer, m_iStartY As Integer, m_iColGap As Integer
Dim m_iPageHeight As Long, m_iLineHeight As Long, m_iMargin As Long
Dim stLstComp As String, stLstCompName As String

Dim m_DefaultPrinter As String

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Screen.MousePointer = vbHourglass
   If TxtValidate = True Then
      SetDataListWidth
      doQuery
   End If
   Screen.MousePointer = vbDefault
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

Private Sub PrintSheet()
   
   Dim ii As Integer, iCompCol As Integer, stLstGroup As String, iLstPage As Integer
   
   GetPleft
   
   With grdDataList
   stLstComp = .TextMatrix(1, 0)
   stLstCompName = CompNameQuery(stLstComp)
   
   iPage = 0
   PrintPageHeader
   PrintPageHeader1
   
   If Option1(0).Tag = "1" Then
      iCompCol = 0
   Else
      'iCompCol = 8
      iCompCol = 0
   End If
   stLstGroup = .TextMatrix(1, iCompCol)
   For ii = 1 To .Rows - 1
      If stLstGroup <> .TextMatrix(ii, iCompCol) Then
         If stLstComp <> .TextMatrix(ii, 0) Then
            stLstComp = .TextMatrix(ii, 0)
            stLstCompName = CompNameQuery(stLstComp)
         End If
         
         Printer.NewPage
         PrintPageHeader
         PrintPageHeader1
         
         stLstGroup = .TextMatrix(ii, iCompCol)
      End If
      
      iLstPage = iPage
      PrintNewLine
      If .TextMatrix(ii, 1) = "" Then
         If iLstPage = iPage Then
            DrawLine
            PrintNewLine
         End If
         Printer.FontBold = True
      Else
         Printer.FontBold = False
      End If
      PrintDetail ii
      If .TextMatrix(ii, 1) = "" Then
         PrintNewLine
      End If
   Next
   Printer.EndDoc
   End With
   MsgBox "列印完畢！"
End Sub

Private Sub PrintDetail(iRow As Integer)
    Dim iCol As Integer
    Printer.Font.Size = 11
    
    With grdDataList
    If Option1(0).Tag = "1" Then
      For iCol = 1 To UBound(PColName)
        Select Case iCol
        
        Case 1, 2
           Printer.CurrentX = PLeft(iCol, 1)
           Printer.CurrentY = iPrint
           Printer.Print .TextMatrix(iRow, iCol)
        Case Else
           If .TextMatrix(iRow, iCol) <> "" Then
              strExc(0) = Format(.TextMatrix(iRow, iCol), DDollar2)
              Printer.CurrentX = PLeft(iCol + 1, 1) - Printer.TextWidth(strExc(0)) - m_iColGap
              Printer.CurrentY = iPrint
              Printer.Print strExc(0)
           End If
        
        End Select
      Next
   Else
      For iCol = 1 To UBound(PColName)
        Select Case iCol
        Case 3, 6, 7
           If .TextMatrix(iRow, iCol) <> "" Then
              strExc(0) = Format(.TextMatrix(iRow, iCol), DDollar2)
              Printer.CurrentX = PLeft(iCol + 1, 1) - Printer.TextWidth(strExc(0)) - m_iColGap
              Printer.CurrentY = iPrint
              Printer.Print strExc(0)
           End If
         Case 4
           Printer.CurrentX = PLeft(iCol, 1)
           Printer.CurrentY = iPrint
           Printer.Print PUB_StrToStr(.TextMatrix(iRow, iCol), 2 * PLeft(iCol, 2))
        Case Else
           Printer.CurrentX = PLeft(iCol, 1)
           Printer.CurrentY = iPrint
           Printer.Print .TextMatrix(iRow, iCol)
        End Select
      Next
   End If
   End With
End Sub

   
Private Sub doQuery()
   Dim stVTB As String
   Dim stYM As String, stY01 As String
   Dim stConNhi As String, stConSm As String, stCon As String
   
   '統計
   If Option1(0).Value = True Then
      stYM = Val(txtPayDate(0)) + 191100
      stY01 = (Val(stYM) \ 100) & "01"
      stConNhi = " and nhi02>=" & stY01 & "01 and nhi02<=" & stYM & "31"
      stConSm = " and sm02=" & stYM
   '明細
   Else
      stConNhi = " and nhi02>=" & (txtPayDate(1) + 19110000) & " and nhi02<=" & (txtPayDate(2) + 19110000)
      stConSm = " and sm02>=" & Left((txtPayDate(1) + 19110000), 6) & " and sm02<=" & Left((txtPayDate(2) + 19110000), 6)
   End If
   
   '公司別
   If txt1(0) <> "" Then
      stConNhi = stConNhi & " and nhi11>='" & txt1(0) & "'"
      stConSm = stConSm & " and sm37>='" & txt1(0) & "'"
   End If
   If txt1(1) <> "" Then
      stConNhi = stConNhi & " and nhi11<='" & txt1(1) & "'"
      stConSm = stConSm & " and sm37<='" & txt1(1) & "'"
   End If
   '員工編號
   If txtStaffNo <> "" Then
      stCon = stCon & " and x2='" & txtStaffNo & "'"
   End If
   
   '紀錄最後條件
   txt1(0).Tag = txt1(0)
   txt1(1).Tag = txt1(1)
   txtPayDate(0).Tag = txtPayDate(0)
   txtPayDate(1).Tag = txtPayDate(1)
   txtPayDate(2).Tag = txtPayDate(2)
   txtStaffNo.Tag = txtStaffNo
   lblName.Tag = lblName
   
   Option1(0).Tag = Abs(Option1(0).Value)
   
   If Option1(0).Value = True Then
      '以補充保費資料為主再加有月薪資但無補充保費者(因會有有補充保費但無月薪資者 Ex.76028 何尤玉 102/1月)
      'Modified by Morgan 2018/1/5 +sm04>0 也要 ,不管有無健保投保金額都要列出(Ex.退休退保)--辜
      stVTB = "select nhi11 x1,nhi01 x2,substr(max(nhi02||nhi05),9) x3" & _
         ",sum(decode(substr(nhi02,1,6)," & stYM & ",nhi07)) x4" & _
         ",sum(nhi07) x5,sum(decode(substr(nhi02,1,6)," & stYM & ",nhi08)) x6,sum(decode(substr(nhi02,1,6)," & stYM & ",nhi06)) x7" & _
         " from (select nhi11,nhi01,nhi02,nhi05,nhi06,nhi07,nhi08" & _
         " from nhi2nd where nhi03='50' and nhi04<>'4' and nhi05>0" & stConNhi & _
         " union all select nhi11 c1,nvl(x1,nhi01) nhi01,nhi02,nhi05,nhi06,nhi07,nhi08" & _
         " from nhi2nd,staff,(select sm01 x1,st26 x2" & _
         " from salarymonth,staff where sm01<'F' and (sm04>0 or sm42>0) " & stConSm & _
         " and st01(+)=sm01)" & _
         " Where nhi03='50' and nhi04='4' and nhi05>0" & stConNhi & _
         " and st01(+)=nhi01 and x2(+)=st26" & _
         " union all select sm37,sm01," & stYM & "01,sm42,0,0,0" & _
         " From salarymonth Where sm01<'F' and (sm04>0 or sm42>0) " & stConSm & _
         ") group by nhi11,nhi01"
         
      
      'Modified by Morgan 2013/7/29 排除當月無薪資及補充保費者,投保薪資
      'Modified by Morgan 2014/2/25 73014,73017,73035 103/1轉換公司別,故抓月薪資投保金額要加公司別
      'strExc(0) = "select x1 公司,x2 員工編號,st02 姓名,decode(sd11,'Y',0,decode(sign(sm15),1,1,decode(hm06,'12',0,null,0,1)))*sm42 投保金額" & _
         ",NVL(SM04,0)+NVL(SM05,0)+NVL(SM28,0)-NVL(SM21,0) 當月薪資 " & _
         ",x3*4 四倍投保金額,x4 當月獎金,x5 累計獎金,x6 費基,x7 補充保費" & _
         " from (" & stVTB & "),salarymonth,himonth,staff,salarydata where sm01(+)=x2 and sm02(+)=" & stYM & " and st01(+)=x2" & stCon & _
         " and hm03(+)=sm02 and hm01(+)=sm01 and hm02(+)=0 and sd01(+)=sm01 and (sm42>0 or x4>0) order by 1,st03,2"
      'Modify By Sindy 2020/6/25 + 證照津貼
      'Modified by Morgan 2022/8/30 +判斷有提撥的不管是否已設定為合夥人身分投保(未生效月份 Ex:蔣律師)--辜
      'Modified by Morgan 2023/12/27 st03-->sm03
      strExc(0) = "select x1 公司,x2 員工編號,st02 姓名,decode(decode(nvl(sm30,0),0,sd11),'Y',0,decode(sign(sm15),1,1,decode(hm06,'12',0,null,0,1)))*decode(sm37,x1,sm42,0) 投保金額" & _
         ",NVL(SM04,0)+NVL(SM05,0)+NVL(SM45,0)+NVL(SM28,0)-NVL(SM21,0) 當月薪資 " & _
         ",x3*4 四倍投保金額,x4 當月獎金,x5 累計獎金,x6 費基,x7 補充保費" & _
         " from (" & stVTB & ") x,(select decode(substr(sm01,3,1),'A',substr(sm01,1,2)||'0'||substr(sm01,4),sm01) sm01,sm02,sm03,sm04,sm05,sm45,sm15,sm21,sm28,sm37,sm42,sm30" & _
         " from salarymonth where 1=1" & stConSm & ") y,himonth,staff,salarydata where sm01(+)=x2 and sm37(+)=x1 and sm02(+)=" & stYM & " and st01(+)=x2" & stCon & _
         " and hm03(+)=sm02 and hm01(+)=sm01 and hm02(+)=0 and sd01(+)=sm01 and (sm04>0 or sm42>0 or x4>0) order by 1,sm03,2"
      'end 2014/2/25
   Else
   
      stVTB = "select nhi11 x1,nhi01 x8,nhi05*4 x3,'年終獎金' x4,nhi02 x5,nhi07 x6,nhi06 x7,nhi01 x2" & _
         " From nhi2nd where nhi04='1' and nhi03='50' and nhi05>0" & stConNhi & _
         " Union All select nhi11,nhi01, nhi05*4,mb14,nhi02,nhi07,nhi06,nhi01" & _
         " From nhi2nd, monthbonus where nhi04 in ('2','6') and nhi03='50' and nhi05>0" & stConNhi & _
         " and mb01(+)=nhi14 and mb02(+)=nhi01" & _
         " Union All select nhi11,nhi01, nhi05*4,'同仁其他',nhi02,nhi07,nhi06,nhi01" & _
         " From nhi2nd where nhi04='3' and nhi03='50' and nhi05>0" & stConNhi & _
         " Union All select nhi11,nhi01, nhi05*4,'翻譯費',nhi02,nhi07,nhi06,nvl(x1,nhi01)" & _
         " From (select * from nhi2nd,staff where nhi04='4' and nhi03='50' and nhi05>0" & stConNhi & _
         " and st01(+)=nhi01) y,(select sm01 x1,st26 x2,sm02 x3" & _
         " from salarymonth,staff where sm42>0" & stConSm & _
         " and st01(+)=sm01) x where x2(+)=st26 and x3(+)=substr(nhi02,1,6)" & _
         " Union All select nhi11,nhi01, nhi05*4,'其他所得',nhi02,nhi07,nhi06,nhi01" & _
         " From nhi2nd where nhi04='0' and nhi03='50' and nhi05>0" & stConNhi
      'Modified by Morgan 2023/12/27 st03-->sm03
      strExc(0) = "select x1 公司,x8 員工編號,st02 姓名,x3 四倍投保金額" & _
         ",x4 名目 " & _
         ",sqldatet(x5) 代扣日期,x6 獎金金額,x7 補充保費,x2 投保編號" & _
         " from (" & stVTB & "),salarymonth,staff where sm01(+)=x2 and sm02(+)=substr(x5,1,6) and st01(+)=x2" & stCon & _
         " order by x1,sm03,x2,x5"
   End If
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Set grdDataList.Recordset = RsTemp.Clone
   SetDataListWidth True
   If intI = 1 Then
      AddSubTotal
      cmdPrint.Enabled = True
   Else
      cmdPrint.Enabled = False
   End If
   
End Sub

Private Sub AddSubTotal()
   Dim ii As Integer, jj As Integer, stComp As String, stGroup As String, stGroupName As String, strAddItem As String, lngColor As Long, lngColor1 As Long
   Dim dblSub1(9) As Double, dblTot(9) As Double
   
   lngColor = &H90EE90
   lngColor1 = &H90EE40
   
   With grdDataList
      .Visible = False
      '統計
      If Option1(0).Value = True Then
         stGroup = .TextMatrix(1, 0)
         ii = 1
         Do While ii < .Rows
            If stGroup <> .TextMatrix(ii, 0) Then
               strAddItem = stGroup & vbTab & vbTab & "合計:"
               For jj = 3 To 9
                  strAddItem = strAddItem & vbTab & dblSub1(jj)
                  dblSub1(jj) = 0
               Next
               .AddItem strAddItem, ii
               
               .row = ii
               For jj = 0 To .Cols - 1
                  .col = jj: .CellBackColor = lngColor
               Next
      
               ii = ii + 1
               stGroup = .TextMatrix(ii, 0)
            End If
            For jj = 3 To 9
               dblSub1(jj) = dblSub1(jj) + Val(.TextMatrix(ii, jj))
            Next
            ii = ii + 1
         Loop
         strAddItem = stGroup & vbTab & vbTab & "合計:"
         For jj = 3 To 9
            strAddItem = strAddItem & vbTab & dblSub1(jj)
            dblSub1(jj) = 0
         Next
         .AddItem strAddItem, ii
         
         .row = ii
         For jj = 0 To .Cols - 1
            .col = jj: .CellBackColor = lngColor
         Next
      '明細
      Else
         stComp = .TextMatrix(1, 0)
         stGroup = .TextMatrix(1, 8)
         stGroupName = .TextMatrix(1, 2)
         ii = 1
         Do While ii < .Rows
            If stGroup <> .TextMatrix(ii, 8) Then
               strAddItem = stComp & vbTab & vbTab & stGroupName
               For jj = 3 To 7
                  Select Case jj
                  Case 4
                     strAddItem = strAddItem & vbTab & "小計:"
                  Case 6, 7
                     strAddItem = strAddItem & vbTab & dblSub1(jj)
                     dblTot(jj) = dblTot(jj) + dblSub1(jj)
                     dblSub1(jj) = 0
                  Case Else
                     strAddItem = strAddItem & vbTab
                  End Select
               Next
               strAddItem = strAddItem & vbTab & stGroup
               .AddItem strAddItem, ii
               .row = ii
               For jj = 0 To .Cols - 1
                  .col = jj: .CellBackColor = lngColor
               Next
               ii = ii + 1
               
               'Added by Morgan 2013/3/12 +合計
               If stComp <> .TextMatrix(ii, 0) Then
                  strAddItem = stComp & vbTab & vbTab
                  For jj = 3 To 7
                     Select Case jj
                     Case 4
                        strAddItem = strAddItem & vbTab & "合計:"
                     Case 6, 7
                        strAddItem = strAddItem & vbTab & dblTot(jj)
                        dblTot(jj) = 0
                     Case Else
                        strAddItem = strAddItem & vbTab
                     End Select
                  Next
                  strAddItem = strAddItem & vbTab
                  .AddItem strAddItem, ii
                  .row = ii
                  For jj = 0 To .Cols - 1
                     .col = jj: .CellBackColor = lngColor1
                  Next
                  ii = ii + 1
               End If
               'end 2013/3/12
               
               stComp = .TextMatrix(ii, 0)
               stGroup = .TextMatrix(ii, 8)
               stGroupName = .TextMatrix(ii, 2)
            End If
            
            dblSub1(6) = dblSub1(6) + Val(.TextMatrix(ii, 6))
            dblSub1(7) = dblSub1(7) + Val(.TextMatrix(ii, 7))
            
            ii = ii + 1
         Loop
         
         strAddItem = stComp & vbTab & vbTab & stGroupName
         For jj = 3 To 7
            Select Case jj
            Case 4
               strAddItem = strAddItem & vbTab & "小計:"
            Case 6, 7
               strAddItem = strAddItem & vbTab & dblSub1(jj)
               dblTot(jj) = dblTot(jj) + dblSub1(jj)
               dblSub1(jj) = 0
            Case Else
               strAddItem = strAddItem & vbTab
            End Select
         Next
         strAddItem = strAddItem & vbTab & stGroup
         
         .AddItem strAddItem, ii
         
         .row = ii
         For jj = 0 To .Cols - 1
            .col = jj: .CellBackColor = lngColor
         Next
         'Added by Morgan 2013/3/12
         ii = ii + 1
         strAddItem = stComp & vbTab & vbTab
         For jj = 3 To 7
            Select Case jj
            Case 4
               strAddItem = strAddItem & vbTab & "合計:"
            Case 6, 7
               strAddItem = strAddItem & vbTab & dblTot(jj)
               dblTot(jj) = 0
            Case Else
               strAddItem = strAddItem & vbTab
            End Select
         Next
         strAddItem = strAddItem & vbTab & stGroup
         .AddItem strAddItem, ii
         .row = ii
         For jj = 0 To .Cols - 1
            .col = jj: .CellBackColor = lngColor1
         Next
         'end 2013/3/12
      
      
      End If
      .Visible = True
   End With
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
   '   txtPayDate(1) = txtPayDate(0) & "01"
   '   txtPayDate(2) = TransDate(GetLastDay(txtPayDate(0) & "01"), 1)
   'End If
   If Val(Right(strSrvDate(1), 2)) > 15 Then
      strExc(1) = CompDate(1, -1, strSrvDate(1))
   Else
      strExc(1) = CompDate(1, -2, strSrvDate(1))
   End If
   txtPayDate(0) = (strExc(1) \ 100) - 191100
   txtPayDate(1) = txtPayDate(0) & "01"
   txtPayDate(2) = TransDate(GetLastDay(txtPayDate(1)), 1)
   'end 2013/5/2
   Option1_Click (0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170234 = Nothing
End Sub

Private Sub SetDataListWidth(Optional ByVal p_bolHeaderOnly As Boolean = False)
   Dim ii As Integer
   With grdDataList
      .Visible = False
      
      '統計
      If Option1(0).Value = True Then
         If p_bolHeaderOnly = False Then
            .Clear
            .Rows = 2: .Cols = 10: .FixedRows = 1: .FixedCols = 0
         End If
         
         .row = 0
         .col = 0: .ColWidth(.col) = 500: .Text = "公司"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 1: .ColWidth(.col) = 750: .Text = "員工號"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 2: .ColWidth(.col) = 750: .Text = "姓名"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 3: .ColWidth(.col) = 900: .Text = "投保金額"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 4: .ColWidth(.col) = 900: .Text = "當月薪資"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 5: .ColWidth(.col) = 1300: .Text = "4倍投保金額"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 6: .ColWidth(.col) = 900: .Text = "當月獎金"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 7: .ColWidth(.col) = 1000: .Text = "累計獎金"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 8: .ColWidth(.col) = 900: .Text = "費基"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 9: .ColWidth(.col) = 900: .Text = "補充保費"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
      '明細
      Else
         If p_bolHeaderOnly = False Then
            .Clear
            .Rows = 2: .Cols = 9: .FixedRows = 1: .FixedCols = 0
         End If
         
         .row = 0
         .col = 0: .ColWidth(.col) = 500: .Text = "公司"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 1: .ColWidth(.col) = 750: .Text = "員工號"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 2: .ColWidth(.col) = 750: .Text = "姓名"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 3: .ColWidth(.col) = 1300: .Text = "4倍投保金額"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 4: .ColWidth(.col) = 1100: .Text = "名目"
         .ColAlignment(.col) = flexAlignLeftCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 5: .ColWidth(.col) = 900: .Text = "代扣日期"
         .ColAlignment(.col) = flexAlignCenterCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 6: .ColWidth(.col) = 900: .Text = "獎金金額"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .col = 7: .ColWidth(.col) = 900: .Text = "補充保費"
         .ColAlignment(.col) = flexAlignRightCenter
         .CellAlignment = flexAlignCenterCenter: .CellFontBold = True
         .ColWidth(8) = 0 '投保編號
      End If
      .Refresh
      .Visible = True
   End With
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If Option1(0).Value = True Then
      If txtPayDate(0) = "" Then
         MsgBox "請輸入年月!"
         txtPayDate(0).SetFocus
         Exit Function
      End If
      
   Else
      If txtStaffNo = "" And (txtPayDate(1) = "" Or txtPayDate(2) = "") Then
         MsgBox "請輸入日期起訖或所得人編號!"
         txtPayDate(0).SetFocus
         Exit Function
      End If
   End If
   TxtValidate = True
End Function


Private Sub Option1_Click(Index As Integer)
   If Index = 0 Then
      txtPayDate(0).Enabled = True
      txtPayDate(1).Enabled = False
      txtPayDate(2).Enabled = False
   Else
      txtPayDate(0).Enabled = False
      txtPayDate(1).Enabled = True
      txtPayDate(2).Enabled = True
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
   Case 2
      If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
         KeyAscii = 0
      End If
   End Select
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
   If Index = 0 Then
      If IsDate(Format(txtPayDate(0) + 191100, "###/##/") & "01") = True Then
         txtPayDate(1) = txtPayDate(0) & "01"
         txtPayDate(2) = TransDate(GetLastDay(txtPayDate(0) & "01"), 1)
      End If
   End If
End Sub

Private Sub txtStaffNo_Change()
   lblName = ""
End Sub

Private Sub txtStaffNo_GotFocus()
   TextInverse txtStaffNo
   CloseIme
End Sub

Private Sub txtStaffNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtStaffNo_Validate(Cancel As Boolean)
   If txtStaffNo <> "" Then
      lblName = StaffQuery(txtStaffNo)
   End If
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
   '統計
   If Option1(0).Value = True Then
      ReDim PLeft(10, 2)
      ReDim PColName(9)
      ii = 1
      PColName(ii) = .TextMatrix(0, ii) '"員工號"
      PLeft(ii, 1) = m_iStartX
      PLeft(ii, 2) = 3
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"姓名"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 3
   
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"投保金額"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 4
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"當月薪資"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 4
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"4倍投保金額"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"當月獎金"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"累計獎金"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"費基"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"補充保費"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 4
      
   Else
      ReDim PLeft(8, 2)
      ReDim PColName(7)
      ii = 1
      PColName(ii) = .TextMatrix(0, ii) '"員工號"
      PLeft(ii, 1) = m_iStartX
      PLeft(ii, 2) = 3
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"姓名"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 3
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"4倍投保金額"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 6
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"名目"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 14
   
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"代扣日期"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 5
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"獎金金額"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 4
      
      ii = ii + 1
      PColName(ii) = .TextMatrix(0, ii) '"補充保費"
      PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
      PLeft(ii, 2) = 4
   End If
   ii = ii + 1
   PLeft(ii, 1) = PLeft(ii - 1, 1) + Printer.TextWidth(String(PLeft(ii - 1, 2), "　")) + m_iColGap
   End With
End Sub


Private Sub PrintPageHeader()
   Dim strTmp As String
   
   If Option1(0).Tag = "1" Then
      strTmp = "台一關係企業 補充保費統計表(員工)"
   Else
      strTmp = "台一關係企業 補充保費明細表(員工)"
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
   
   strTmp = "公  司  別：" & txt1(0).Tag & " - " & txt1(1).Tag
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(String(10, "　"))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   iPrint = iPrint + m_iLineHeight
   If Option1(0).Tag = "1" Then
      strTmp = "年　　月：" & txtPayDate(0).Tag
   Else
      strTmp = "日　　期：" & txtPayDate(1).Tag & " - " & txtPayDate(2).Tag
   End If
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(String(10, "　"))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strTmp
   
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   
   
   strTmp = "員工編號：" & txtStaffNo.Tag & " " & lblName.Tag
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
   
   iPrint = iPrint + m_iLineHeight
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "公司別：" & stLstComp & "　" & stLstCompName
   
   PrintNewLine
   
End Sub


Private Sub PrintPageHeader1()
   
   iPrint = iPrint + m_iLineHeight
   Printer.Font.Size = 11
   Printer.FontBold = True
   
   If Option1(0).Tag = "1" Then
      For intI = 1 To UBound(PColName)
         Select Case intI
         Case 3 To 9
            Printer.CurrentX = PLeft(intI + 1, 1) - Printer.TextWidth(PColName(intI)) - m_iColGap
            Printer.CurrentY = iPrint
            Printer.Print PColName(intI)
         Case Else
            Printer.CurrentX = PLeft(intI, 1)
            Printer.CurrentY = iPrint
            Printer.Print PColName(intI)
         End Select
      Next
   Else
      For intI = 1 To UBound(PColName)
         Select Case intI
         Case 3, 6, 7
            Printer.CurrentX = PLeft(intI + 1, 1) - Printer.TextWidth(PColName(intI)) - m_iColGap
            Printer.CurrentY = iPrint
            Printer.Print PColName(intI)
         Case Else
            Printer.CurrentX = PLeft(intI, 1)
            Printer.CurrentY = iPrint
            Printer.Print PColName(intI)
         End Select
      Next
   End If
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

