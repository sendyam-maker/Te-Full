VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010032 
   BorderStyle     =   1  '單線固定
   Caption         =   "主管機關發文統計"
   ClientHeight    =   5745
   ClientLeft      =   3780
   ClientTop       =   3690
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   3000
      MaxLength       =   7
      TabIndex        =   1
      Top             =   120
      Width           =   945
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1740
      MaxLength       =   7
      TabIndex        =   0
      Top             =   120
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   4740
      Left            =   30
      TabIndex        =   4
      Top             =   990
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   8361
      _Version        =   393216
      Cols            =   5
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
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢(&S)"
      Height          =   345
      Index           =   0
      Left            =   6870
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   60
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   7830
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   60
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "～"
      Height          =   255
      Index           =   1
      Left            =   2730
      TabIndex        =   7
      Top             =   150
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "共　0　件"
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   5
      Left            =   6870
      TabIndex        =   6
      Top             =   690
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "發文日期："
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   150
      Width           =   1335
   End
End
Attribute VB_Name = "frm010032"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 GrdDataList
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer
Dim lngCounterI As Long
Dim m_bolPrintRight As Boolean
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_intRow As Integer, m_intCol As Integer
Dim bolV As Boolean
Public cmdState As Integer


Private Sub SetDataListWidth()
grdDataList.row = 0
grdDataList.col = 0: grdDataList.Text = "主管機關"
grdDataList.ColWidth(0) = 2000
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 1: grdDataList.Text = "內商"
grdDataList.ColWidth(1) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 2: grdDataList.Text = "外商"
grdDataList.ColWidth(2) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 3: grdDataList.Text = "內專"
grdDataList.ColWidth(3) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter
grdDataList.col = 4: grdDataList.Text = "外專"
grdDataList.ColWidth(4) = 1200
grdDataList.CellAlignment = flexAlignCenterCenter 'flexAlignRightCenter
End Sub

Private Sub cmdOK_Click(Index As Integer)
cmdState = Index
PubShowNextData
Exit Sub
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 0 '查詢
         Call SearchData
         
      Case 1 '結束
         'fnCloseAllFrm100
         Unload Me
         Set frm010032 = Nothing
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   SetDataListWidth
   
   '發文日期
   'txt1(0).Text = strSrvDate(2)
   'txt1(1).Text = strSrvDate(2)
   
   m_bolPrintRight = IsUserHasRightOfFunction("frm010032", strPrint, False)
   
   cmdOK(0).Default = True
   'Call SearchData
   cmdState = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm010032 = Nothing
End Sub

Private Sub SearchData()
Dim strSql As String
Dim dblCnt As Double

'重新檢查欄位有效性
If TxtValidate = False Then Exit Sub
'因考慮CFT內部收文之申請英文證明304,故CF02='000'而不抓案件之申請國家CFT-003049
'strSQL = "SELECT CF10,SUM(A1),SUM(A2),SUM(A3),SUM(A4) FROM ( " & _
'                  "SELECT CF10,1 AS A1,0 AS A2,0 AS A3,0 AS A4 FROM CaseProgress,CaseFee,Staff S1 " & _
'                  "Where CP124 >= " & ChangeTStringToWString(txt1(0)) & " And CP124 <= " & ChangeTStringToWString(txt1(1)) & " " & _
'                     "AND CP123 = 'Y' AND CP01 = CF01 AND CF02='000' AND CP10=CF03 " & _
'                     "AND CP83=S1.ST01(+) AND S1.ST03 like 'P2%' Union All " & _
'                  "SELECT CF10,0 AS A1,1 AS A2,0 AS A3,0 AS A4 FROM CaseProgress,CaseFee,Staff S1 " & _
'                  "Where CP124 >= " & ChangeTStringToWString(txt1(0)) & " And CP124 <= " & ChangeTStringToWString(txt1(1)) & " " & _
'                     "AND CP123 = 'Y' AND CP01 = CF01 AND CF02='000' AND CP10=CF03 " & _
'                     "AND CP83=S1.ST01(+) AND S1.ST03 like 'F1%' Union All " & _
'                  "SELECT CF10,0 AS A1,0 AS A2,1 AS A3,0 AS A4 FROM CaseProgress,CaseFee,Staff S1 " & _
'                  "Where CP124 >= " & ChangeTStringToWString(txt1(0)) & " And CP124 <= " & ChangeTStringToWString(txt1(1)) & " " & _
'                     "AND CP123 = 'Y' AND CP01 = CF01 AND CF02='000' AND CP10=CF03 " & _
'                     "AND CP83=S1.ST01(+) AND S1.ST03 like 'P1%' Union All " & _
'                  "SELECT CF10,0 AS A1,0 AS A2,0 AS A3,1 AS A4 FROM CaseProgress,CaseFee,Staff S1 " & _
'                  "Where CP124 >= " & ChangeTStringToWString(txt1(0)) & " And CP124 <= " & ChangeTStringToWString(txt1(1)) & " " & _
'                     "AND CP123 = 'Y' AND CP01 = CF01 AND CF02='000' AND CP10=CF03 " & _
'                     "AND CP83=S1.ST01(+) AND S1.ST03 like 'F2%' " & _
'                  ") GROUP BY CF10 ORDER BY CF10 "

strSql = "SELECT Org,SUM(A1),SUM(A2),SUM(A3),SUM(A4) FROM ( " & _
                  "SELECT decode(instr(cp130,','),0,CP130,substr(cp130,1,instr(cp130,',')-1)) as Org,1 AS A1,0 AS A2,0 AS A3,0 AS A4 FROM CaseProgress,Staff S1 " & _
                  "Where CP124 >= " & ChangeTStringToWString(txt1(0)) & " And CP124 <= " & ChangeTStringToWString(txt1(1)) & " " & _
                     "AND CP123 = 'Y' " & _
                     "AND CP83=S1.ST01(+) AND S1.ST03 like 'P2%' Union All " & _
                  "SELECT decode(instr(cp130,','),0,CP130,substr(cp130,1,instr(cp130,',')-1)) as Org,0 AS A1,1 AS A2,0 AS A3,0 AS A4 FROM CaseProgress,Staff S1 " & _
                  "Where CP124 >= " & ChangeTStringToWString(txt1(0)) & " And CP124 <= " & ChangeTStringToWString(txt1(1)) & " " & _
                     "AND CP123 = 'Y' " & _
                     "AND CP83=S1.ST01(+) AND S1.ST03 like 'F1%' Union All " & _
                  "SELECT decode(instr(cp130,','),0,CP130,substr(cp130,1,instr(cp130,',')-1)) as Org,0 AS A1,0 AS A2,1 AS A3,0 AS A4 FROM CaseProgress,Staff S1 " & _
                  "Where CP124 >= " & ChangeTStringToWString(txt1(0)) & " And CP124 <= " & ChangeTStringToWString(txt1(1)) & " " & _
                     "AND CP123 = 'Y' " & _
                     "AND CP83=S1.ST01(+) AND S1.ST03 like 'P1%' Union All " & _
                  "SELECT decode(instr(cp130,','),0,CP130,substr(cp130,1,instr(cp130,',')-1)) as Org,0 AS A1,0 AS A2,0 AS A3,1 AS A4 FROM CaseProgress,Staff S1 " & _
                  "Where CP124 >= " & ChangeTStringToWString(txt1(0)) & " And CP124 <= " & ChangeTStringToWString(txt1(1)) & " " & _
                     "AND CP123 = 'Y' " & _
                     "AND CP83=S1.ST01(+) AND S1.ST03 like 'F2%' " & _
                  ") GROUP BY Org ORDER BY Org "

Screen.MousePointer = vbHourglass
grdDataList.Clear
grdDataList.Rows = 2
SetDataListWidth
'GrdDataList.FixedCols = 0

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
dblCnt = 0
If adoRecordset.RecordCount <> 0 Then
   adoRecordset.MoveFirst
   Do While Not adoRecordset.EOF
      dblCnt = dblCnt + adoRecordset(1) + adoRecordset(2) + adoRecordset(3) + adoRecordset(4)
      adoRecordset.MoveNext
   Loop
   Label1(5).Caption = "共　" & dblCnt & "　件"
   Set grdDataList.Recordset = adoRecordset
Else
   Label1(5).Caption = "共　" & dblCnt & "　件"
   ShowNoData
   grdDataList.Clear
End If
SetDataListWidth
'GrdDataList.FixedCols = 3
CheckOC
Screen.MousePointer = vbDefault
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
