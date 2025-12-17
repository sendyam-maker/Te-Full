VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010030 
   BorderStyle     =   1  '單線固定
   Caption         =   "分所寄件統計"
   ClientHeight    =   2472
   ClientLeft      =   3780
   ClientTop       =   3696
   ClientWidth     =   8868
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2472
   ScaleWidth      =   8868
   Begin VB.CommandButton cmdOK 
      Caption         =   "明細(&D)"
      Height          =   345
      Index           =   1
      Left            =   4860
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   90
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2820
      MaxLength       =   7
      TabIndex        =   1
      Top             =   150
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   0
      Top             =   150
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1575
      Left            =   210
      TabIndex        =   4
      Top             =   780
      Width           =   7485
      _ExtentX        =   13208
      _ExtentY        =   2773
      _Version        =   393216
      Cols            =   3
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|所別|文件量"
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢(&S)"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   3915
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   90
      Width           =   900
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   5805
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "～"
      Height          =   180
      Index           =   1
      Left            =   2595
      TabIndex        =   7
      Top             =   195
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "共　0　件"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   5
      Left            =   225
      TabIndex        =   6
      Top             =   540
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "發文室發文日："
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   5
      Top             =   195
      Width           =   1260
   End
End
Attribute VB_Name = "frm010030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 (無需修改)
'Created by Morgan 2014/5/12
Option Explicit
Dim lPrevRow As Long '前次點選列
Dim m_stCon As String

Private Sub cmdExit_Click(Index As Integer)
   Unload Me
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   Select Case Index
   Case 0
      If TxtValidate = True Then
         QueryData
      End If
   Case 1
      If lPrevRow > 0 Then
         QueryDetail
      Else
         MsgBox "請點選一筆資料!!", vbInformation
      End If
   End Select
   Screen.MousePointer = vbDefault
End Sub

Private Sub QueryData()
   
   Dim ii As Integer
   Dim lCount1 As Long, lCount2 As Long, lCount3 As Long, lCount4 As Long, lCount5 As Long, lCount6 As Long
   Dim lTot As Long
   
   txt1(0).Tag = txt1(0).Text
   txt1(1).Tag = txt1(1).Text
   lPrevRow = 0

   m_stCon = ""
   
   If txt1(0) <> "" Then
      m_stCon = m_stCon & " and cp127>=" & DBDATE(txt1(0))
   End If
   
   If txt1(1) <> "" Then
      m_stCon = m_stCon & " and cp127<=" & DBDATE(txt1(1))
   End If
   
   SetGrid True
   
   'Modified by Morgan 2014/5/22 特殊設定 A7 的編號照北所的流程
   'Modified by Morgan 2014/6/12 +北所統計
   'Modified by Morgan 2015/6/29 排除發文人員為QPGMR(E化系統自動發文)
   'Modified by Morgan 2015/10/6 +臺灣,非臺灣個別統計--李佳寶
   'Modified by Morgan 2018/9/25 +判斷有信函進度的(lp03>0),(T案C類電子化但無LP)
   'Modified by Morgan 2018/9/28 +臺灣商標件數,非臺灣商標件數
   'Modified by Morgan 2018/10/4 +CFP
   'Modified by Morgan 2021/5/19 只抓北所發文室發文案件
   'Modified by Morgan 2022/1/3 +TC
   'Modified by Morgan 2024/6/7 P1004改走分所
   strExc(0) = "select '' V,decode(nvl(BNo,s1.st06),'2','中','3','南','4','高','北') 所別" & _
      ",count(*) 文件量,sum(decode(pa01||pa09,'P000',1,0)) P臺灣件數" & _
      ",sum(decode(pa01,'P',decode(pa09,'000',0,1))) P非臺灣件數" & _
      ",sum(decode(tm10,'000',1,0)) 臺灣商標件數" & _
      ",sum(decode(tm10,'000',0,decode(tm01,null,0,1))) 非臺灣商標件數" & _
      ",sum(decode(cp01,'CFP',1,0)) CFP件數" & _
      ",sum(decode(cp01,'TC',1,0)) TC件數" & _
      ",nvl(BNo,s1.st06) st06 From caseprogress, letterprogress, staff s1,staff s2,patent,trademark,servicepractice" & _
      ",(select st01 SNo,'1' BNo from setspecman,staff where ocode='A7' and instr(';'||replace(oMan,',',';')||';',';'||st01||';')>0 and st01<>'P1004')" & _
      " Where lp01(+)=cp09 and cp127 > 0 and CP154<>'QPGMR'" & m_stCon & _
      " and SNo(+)=cp13 and s1.st01(+)=cp13" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and s2.st01(+)=cp154 and s2.st06='1' group by nvl(BNo,s1.st06)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With MSHFlexGrid1
      .Visible = False
      .FixedCols = 0
      Set .Recordset = RsTemp
      SetGrid
      lTot = 0
      For ii = 1 To .Rows - 1
         lTot = lTot + Val(.TextMatrix(ii, 2))
         lCount1 = lCount1 + Val(.TextMatrix(ii, 3))
         lCount2 = lCount2 + Val(.TextMatrix(ii, 4))
         lCount3 = lCount3 + Val(.TextMatrix(ii, 5)) 'Added by Morgan 2018/9/28
         lCount4 = lCount4 + Val(.TextMatrix(ii, 6)) 'Added by Morgan 2018/9/28
         lCount5 = lCount5 + Val(.TextMatrix(ii, 7)) 'Added by Morgan 2018/10/4
         lCount6 = lCount6 + Val(.TextMatrix(ii, 8)) 'Added by Morgan 2022/1/3
      Next
      Label1(5).Caption = "共 " & lTot & " 件, P臺灣 " & lCount1 & " 件, P非臺灣 " & lCount2 & " 件, 臺灣商標 " & lCount3 & " 件, 非臺灣商標 " & lCount4 & " 件, CFP " & lCount5 & " 件, TC " & lCount6 & " 件"
      .Visible = True
      End With
      
   Else
      Label1(5).Caption = "共　0　件"
      ShowNoData
   End If
End Sub

Private Sub QueryDetail()
   'modify by sonia 2014/7/3 +LP11阿寶說只印非直寄的
   'Modified by Morgan 2015/7/24 排除發文人員為QPGMR(E化系統自動發文)
   'Modified by Morgan 2017/9/30 案件性質要判斷大陸案
   'Modified by Morgan 2018/9/25 +判斷有信函進度的(lp03>0)(T案C類電子化但無LP)
   'Modified by Morgan 2018/9/28 改回T案也要統計,但直寄用是否有期限來判斷
   'Modified by Morgan 2021/5/19 只抓北所發文室發文案件
   'Modified by Morgan 2024/6/7 P1004改走分所
   strExc(0) = "select cp01||'-'||cp02||'-'||CP03||'-'||CP04 本所案號,nvl(pa05,nvl(tm05,sp05)) 案件名稱,decode(nvl(pa09,nvl(tm10,sp09)),'000',cpm03,cpm04) 案件性質" & _
      ",sqldatet(cp127)||' '||sqltime(cp128) 發文室發文時間,decode(lp01,null,decode(cp07,null,'','直寄'),decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11)) 方式" & _
      " From  caseprogress, letterprogress, staff s1, staff s2, patent,trademark,servicepractice, casepropertymap" & _
      ",(select st01 SNo,'1' BNo from setspecman,staff where ocode='A7' and instr(';'||replace(oMan,',',';')||';',';'||st01||';')>0 and st01<>'P1004')" & _
      " Where lp01(+)=cp09 and cp127 > 0 and CP154<>'QPGMR'" & m_stCon & _
      " and s1.st01(+)=cp13 and SNo(+)=cp13 and nvl(BNo,s1.st06)='" & MSHFlexGrid1.TextMatrix(lPrevRow, 9) & "'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and s2.st01(+)=cp154 and s2.st06='1' order by cp127,cp128,1"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      frm010030_1.SetData RsTemp
      frm010030_1.lblZone = MSHFlexGrid1.TextMatrix(lPrevRow, 1)
      Screen.MousePointer = vbDefault
      frm010030_1.Show vbModal
   End If
   
End Sub

'Modified by Morgan 2022/1/3 +TC
Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   arrGridHeadWidth = Array(240, 500, 700, 950, 950, 950, 1000, 950, 950)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   .FixedCols = 1
   .FormatString = "V|所別|文件量|P臺灣|P非臺灣|臺灣商標|非臺灣商標|CFP|TC"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         If iCol >= 2 Then
            .ColAlignment(iCol) = flexAlignRightCenter
         Else
            .ColAlignment(iCol) = flexAlignLeftCenter
         End If
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txt1(0) = strSrvDate(2)
   txt1(1) = strSrvDate(2)
   SetGrid True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010030 = Nothing
End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nCol As Long, nRow As Long, lRow As Long
   If nCol < 0 Or nRow < 0 Then Exit Sub
   With MSHFlexGrid1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow > 0 Then
      If lPrevRow > 0 Then
         If lPrevRow <> nRow Then
            .row = lPrevRow
            ClickGrid MSHFlexGrid1
            .row = nRow
            ClickGrid MSHFlexGrid1
         End If
      Else
         .row = nRow
         ClickGrid MSHFlexGrid1
      End If
      lPrevRow = .row
   End If
   .Visible = True
   End With
End Sub

Private Sub ClickGrid(grdDataList As MSHFlexGrid)
   Dim iCol As Integer

   With grdDataList
   If .TextMatrix(grdDataList.row, 1) <> "" Then
      If .TextMatrix(.row, 0) = "V" Then
         .TextMatrix(.row, 0) = ""
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = .BackColor
          Next
      '已刪除資料標示為 X
      ElseIf .TextMatrix(.row, 0) = "" Then
         .TextMatrix(.row, 0) = "V"
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            .CellBackColor = &HFFC0C0
         Next
      End If
   End If
   End With
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If CheckIsTaiwanDate(txt1(Index), False) = False Then
      Cancel = True
      MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
      Call txt1_GotFocus(Index)
      Exit Sub
   End If
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim s As Integer

TxtValidate = False

'發文日期
If Len(Trim(txt1(0).Text)) = 0 Then
   s = MsgBox("發文起始日期不可空白", , "輸入條件錯誤")
   txt1(0).SetFocus
   Exit Function
End If
If Len(Trim(txt1(1).Text)) = 0 Then
   s = MsgBox("發文迄止日期不可空白", , "輸入條件錯誤")
   txt1(1).SetFocus
   Exit Function
End If

If Me.txt1(0).Enabled = True Then
   Cancel = False
   Call txt1_Validate(0, Cancel)
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.txt1(1).Enabled = True Then
   Cancel = False
   Call txt1_Validate(1, Cancel)
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

