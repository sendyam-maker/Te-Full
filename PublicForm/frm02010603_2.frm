VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm02010603_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人其他來函輸入"
   ClientHeight    =   5745
   ClientLeft      =   210
   ClientTop       =   990
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form24"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9330
   Begin VB.Frame fraTF 
      BorderStyle     =   0  '沒有框線
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   2235
      Begin VB.Label lblTFCode 
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   23
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblTFCode 
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   22
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblTFCode 
         Height          =   252
         Index           =   1
         Left            =   840
         TabIndex        =   21
         Top             =   600
         Width           =   372
      End
      Begin VB.Label lblTFCode 
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   20
         Top             =   600
         Width           =   852
      End
   End
   Begin VB.Frame fraElse 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   15
      Top             =   540
      Width           =   2175
      Begin VB.Label lblCode 
         Height          =   252
         Index           =   2
         Left            =   1200
         TabIndex        =   18
         Top             =   0
         Width           =   492
      End
      Begin VB.Label lblCode 
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   17
         Top             =   0
         Width           =   252
      End
      Begin VB.Label lblCode 
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   972
      End
   End
   Begin VB.ComboBox cboCaseName 
      Height          =   300
      ItemData        =   "frm02010603_2.frx":0000
      Left            =   1080
      List            =   "frm02010603_2.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   0
      Top             =   1500
      Width           =   7275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      Height          =   405
      Index           =   1
      Left            =   7152
      TabIndex        =   3
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   6324
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   2
      Left            =   8376
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   2100
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   6165
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      RowSizingMode   =   1
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
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblSystem 
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblNumber1 
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   900
      Width           =   2415
   End
   Begin VB.Label lblAgent 
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label lblNumber2 
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label lblNumber 
      AutoSize        =   -1  'True
      Caption         =   "審定號數："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   900
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   900
   End
   Begin VB.Label lblCode 
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   5
      Top             =   900
      Width           =   255
   End
End
Attribute VB_Name = "frm02010603_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/10 改成Form2.0 (cboCaseName,lblAgent)
'Memo By Morgan 2012/12/17 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
Dim intCol As Integer
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim m_PrevForm As Form 'Add By Sindy 2016/10/11


'Add By Sindy 2016/10/11
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, bolRt As Boolean
   
   Select Case Index
      Case 0
         intLeaveKind = 0
         'Add By Sindy 2016/10/11
         If Not m_PrevForm Is Nothing Then
            Call frm02010603_3.SetParent(m_PrevForm)
         End If
         '2016/10/11 END
         'Add By Sindy 2016/10/7
         frm02010603_3.m_strIR01 = m_strIR01
         frm02010603_3.m_strIR02 = m_strIR02
         frm02010603_3.m_strIR03 = m_strIR03
         frm02010603_3.m_strIR04 = m_strIR04
         '2016/10/7 END
         frm02010603_3.Show
         Me.Hide
      Case 1, 2
         If Index = 2 Then
            intLeaveKind = 2
         Else
            intLeaveKind = 1
         End If
         Unload Me
   End Select
End Sub
Private Sub GetOtherInputCaseData()
'TF為馬德里案，另外判斷
If lblSystem = 馬德里案 Then
   GetOtherInputData lblSystem, lblTFCode(0) + IIf(lblTFCode(1) = "", "0", lblTFCode(1)), _
      IIf(lblTFCode(2) = "", "0", lblTFCode(2)), IIf(lblTFCode(3) = "", "00", lblTFCode(3))
Else
   GetOtherInputData lblSystem, lblCode(0), _
      IIf(lblCode(1) = "", "0", lblCode(1)), IIf(lblCode(2) = "", "00", lblCode(2))
End If
End Sub
Private Sub GetOtherInputData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String)
Dim varSaveCursor
Dim nIndex As Integer

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'92.2.18 modify by sonia
'Set grdDataList.Recordset = objPublicData.ReadOtherInputFromCodeRst(intPCaseKind, strGroup, strCode1, strCode2, strCode3, strCode4)
strSql = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(pa09,'020',cpm04,cpm03) 案件性質," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp27") & " 發文日," & SQLDate("cp57") & " 取消收文日期,cp24 結果,cp19 後金,cp64 進度備註,nvl(cp40,nvl(cp50,cp55)) 相關人" + _
         ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " + _
    " from caseprogress,patent,casepropertymap,staff_group where cp01=cpm01 and cp10=cpm02 and sg01=" + CNULL(strGroup) + _
    " and sg02=cp01 and sg03=cp10" + _
    " and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + _
    " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" + _
    " and cp27 is not null order by SORTFIELD desc"
Set grdDataList.Recordset = ClsPDReadRst(strSql)
'92.2.18 end
SetDataListVision grdDataList
intLastRow = 0
If grdDataList.Rows > 1 Then
   'Add By Cheng 2002/07/24
   '不論何系統進入, 發文日由大到小排序
   ' 91.09.13 modify by louis (排序)
   'Me.grdDataList.Col = 5
   'Me.grdDataList.Sort = flexSortGenericDescending
   
   ShowBar grdDataList, intLastRow, intCol
   cmdOK(0).Enabled = True
   cmdOK(0).Default = True
Else
   cmdOK(0).Enabled = False
   cmdOK(1).Default = True
End If
' 90.07.02 modify by louis
For nIndex = 1 To grdDataList.Rows - 1
   Select Case grdDataList.TextMatrix(nIndex, 7)
      Case "1": grdDataList.TextMatrix(nIndex, 7) = "准勝"
      Case "2": grdDataList.TextMatrix(nIndex, 7) = "駁敗"
   End Select
Next nIndex
Screen.MousePointer = varSaveCursor
End Sub
Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

If intPCaseKind = 專利 Then
   varGridWidth = Array(1000, 900, 1000, 900, 900, 900, 1200, 800, 800, 1300, 900)
   intCol = 10
Else
   varGridWidth = Array(1000, 1000, 4500, 1800, 1200)
   intCol = 4
End If
SetGridDataListWidth grdDataList, varGridWidth()
blnOKtoShow = True
End Sub
'Private Sub Form_Activate()
Public Sub QueryData()
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strCustomer As String, strNumber1 As String, strNumber2 As String

On Error GoTo ErrHnd
If lblSystem = 馬德里案 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(lblSystem, lblTFCode(0) + IIf(lblTFCode(1) = "", "0", lblTFCode(1)), _
         IIf(lblTFCode(2) = "", "0", lblTFCode(2)), IIf(lblTFCode(3) = "", "00", lblTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) = False Then GoTo Err1
   If ClsPDCheckCaseCodeIsExist(lblSystem, lblTFCode(0) + IIf(lblTFCode(1) = "", "0", lblTFCode(1)), _
         IIf(lblTFCode(2) = "", "0", lblTFCode(2)), IIf(lblTFCode(3) = "", "00", lblTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) = False Then GoTo err1
Else
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(lblSystem, lblCode(0), _
        IIf(lblCode(1) = "", "0", lblCode(1)), IIf(lblCode(2) = "", "00", lblCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) = False Then GoTo Err1
   If ClsPDCheckCaseCodeIsExist(lblSystem, lblCode(0), _
        IIf(lblCode(1) = "", "0", lblCode(1)), IIf(lblCode(2) = "", "00", lblCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) = False Then GoTo err1
End If
SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
lblNumber1 = strNumber1
lblNumber2 = strNumber2
lblAgent = strCustomer
GetOtherInputData lblSystem, lblCode(0), _
     IIf(lblCode(1) = "", "0", lblCode(1)), IIf(lblCode(2) = "", "00", lblCode(2))

' 90.07.02 modify by louis ' 只有一筆時直接進入下一個畫面
If grdDataList.Rows = 2 Then
   intLastRow = 1
   cmdOK_Click 0
End If

Exit Sub
err1:
Unload Me
Exit Sub
ErrHnd:
ErrorMsg
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Me.Caption = frm02010603_1.Caption
   intLeaveKind = 1
   SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If frm02010603_1.intOpt = 2 Then
      If intLeaveKind = 1 Then
         frm02010603_1.Show
      ElseIf intLeaveKind = 2 Then
         Unload frm02010603_1
      Else
         frm02010603_1.Show
         frm02010603_1.Clear
      End If
   Else
   '   If intLeaveKind = 1 Then
   '      frm02010603_7.Show
   '   ElseIf intLeaveKind = 2 Then
   '      Unload frm02010603_7
   '   End If
   End If
   
   'Add By Sindy 2016/10/11
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   'Add By Cheng 2002/07/18
   Set frm02010603_2 = Nothing
End Sub

Private Sub grdDataList_DblClick()
cmdOK_Click 0
End Sub
Private Sub lblSystem_Change()
If lblSystem = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
Else
   fraTF.Visible = False
   fraElse.Visible = True
End If
End Sub
Private Sub grdDataList_GotFocus()
GridGotFocus grdDataList
End Sub
Private Sub grdDataList_LostFocus()
GridLostFocus grdDataList
End Sub
Private Sub grdDataList_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then grdDataList_DblClick
End Sub
Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      ShowBar grdDataList, intLastRow, intCol
      blnOKtoShow = True
   End If
End If
End Sub
