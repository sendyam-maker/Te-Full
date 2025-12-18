VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm05010401_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "一般來函"
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
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   1470
      Width           =   7275
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12832;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCode 
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   17
      Top             =   600
      Width           =   555
   End
   Begin VB.Label lblCode 
      Height          =   255
      Index           =   1
      Left            =   2295
      TabIndex        =   16
      Top             =   600
      Width           =   210
   End
   Begin VB.Label lblCode 
      Height          =   255
      Index           =   2
      Left            =   2730
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblSystem 
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   600
      Width           =   300
   End
   Begin VB.Label lblNumber1 
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   900
      Width           =   2415
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   1800
      Width           =   3135
      VariousPropertyBits=   27
      Size            =   "5530;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
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
      Caption         =   "本所案號：         -               -        -"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2520
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
Attribute VB_Name = "frm05010401_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (grdDataList,cboCaseName,lblAgent)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

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


Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, bolRt As Boolean

   Select Case Index
      Case 0
         intLeaveKind = 0
         'Add By Sindy 2016/10/7
         frm05010401_3.m_strIR01 = m_strIR01
         frm05010401_3.m_strIR02 = m_strIR02
         frm05010401_3.m_strIR03 = m_strIR03
         frm05010401_3.m_strIR04 = m_strIR04
         '2016/10/7 END
         frm05010401_3.Show   'CFP
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
   GetOtherInputData lblSystem, lblCode(0), _
      IIf(lblCode(1) = "", "0", lblCode(1)), IIf(lblCode(2) = "", "00", lblCode(2))
End Sub

Private Sub GetOtherInputData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String)
Dim varSaveCursor
Dim nIndex As Integer

   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   '92.2.18 modify by sonia
   'Set grdDataList.Recordset = objPublicData.ReadOtherInputFromCodeRst(intPCaseKind, strGroup, strCode1, strCode2, strCode3, strCode4)
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   strSql = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(pa09,'000',cpm03,cpm04) 案件性質," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp27") & " 發文日," & SQLDate("cp57") & " 取消收文日期,cp24 結果,cp19 後金,cp64 進度備註,nvl(cp40,nvl(cp50,cp55)) 相關人" + _
            ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " + _
       " from caseprogress,patent,casepropertymap,staff_group where cp01=cpm01 and cp10=cpm02 and sg01=" + CNULL(strGroup) + _
       " and sg02=cp01 and sg03=cp10" + _
       " and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + _
       " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" + _
       " and cp27 is not null order by SORTFIELD desc"
   Set grdDataList.Recordset = ClsPDReadRst(strSql)
   SetDataListVision grdDataList
   intLastRow = 0
   If grdDataList.Rows > 1 Then
      
      ShowBar grdDataList, intLastRow, intCol
      cmdOK(0).Enabled = True
      cmdOK(0).Default = True
   Else
      cmdOK(0).Enabled = False
      cmdOK(1).Default = True
   End If
   For nIndex = 1 To grdDataList.Rows - 1
      Select Case grdDataList.TextMatrix(nIndex, 7)
         Case "1": grdDataList.TextMatrix(nIndex, 7) = "准勝"
         Case "2": grdDataList.TextMatrix(nIndex, 7) = "駁敗"
      End Select
      Me.grdDataList.TextMatrix(nIndex, 2) = Me.grdDataList.TextMatrix(nIndex, 2) & PUB_GetRelateCasePropertyName(Me.grdDataList.TextMatrix(nIndex, 0), "1")   'ADD BY SONIA 2014/5/13 加相關總收號案件性質
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

   If ClsPDCheckCaseCodeIsExist(lblSystem, lblCode(0), _
        IIf(lblCode(1) = "", "0", lblCode(1)), IIf(lblCode(2) = "", "00", lblCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) = False Then GoTo err1
   
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
   Me.Caption = frm05010401_1.Caption
   intLeaveKind = 1
   SetDataListWidth
   
   'Add By Sindy 2017/12/28
   m_strIR01 = frm05010401_1.m_strIR01
   m_strIR02 = frm05010401_1.m_strIR02
   m_strIR03 = frm05010401_1.m_strIR03
   m_strIR04 = frm05010401_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/28 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2016/10/13
   If Me.m_strIR01 = "" Then
   '2016/10/13 END
      If intLeaveKind = 1 Then
         frm05010401_1.Show
      ElseIf intLeaveKind = 2 Then
         Unload frm05010401_1
      Else
         frm05010401_1.Show
         frm05010401_1.Clear
      End If
   Else
      Unload frm05010401_1
   End If
   Set frm05010401_2 = Nothing
End Sub

Private Sub grdDataList_DblClick()
   cmdOK_Click 0
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
