VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050103_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人案件提申"
   ClientHeight    =   5745
   ClientLeft      =   -600
   ClientTop       =   2985
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9330
   Begin VB.TextBox txtCode 
      Height          =   285
      Index           =   2
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   3
      Top             =   630
      Width           =   492
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Index           =   1
      Left            =   3240
      MaxLength       =   1
      TabIndex        =   2
      Top             =   630
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Index           =   0
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   1
      Top             =   630
      Width           =   1212
   End
   Begin VB.TextBox txtSystem 
      Height          =   285
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   0
      Top             =   630
      Width           =   732
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   1380
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   7435
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "提申資料(&F)"
      Default         =   -1  'True
      Height          =   405
      Index           =   2
      Left            =   7140
      TabIndex        =   6
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   405
      Index           =   0
      Left            =   6300
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   8388
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   1020
      Width           =   8115
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14314;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNation 
      Height          =   255
      Left            =   5940
      TabIndex        =   12
      Top             =   660
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   4980
      TabIndex        =   13
      Top             =   660
      Width           =   975
   End
   Begin MSForms.Label lblCountryName 
      Height          =   255
      Left            =   6300
      TabIndex        =   11
      Top             =   660
      Width           =   2535
      VariousPropertyBits=   27
      Size            =   "4471;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1020
      Width           =   975
   End
End
Attribute VB_Name = "frm050103_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (grdDataList,cboCaseName,lblCountryName)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String
Dim m_Done As Boolean
'2016/10/7 END


Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer, bolRt As Boolean
'Add by Morgan 2005/11/4
Dim stCP10 As String, stCP47 As String

   Select Case Index
      Case 0
         'Add by Morgan 2005/11/4
         stCP10 = grdDataList.TextMatrix(grdDataList.row, 9)
         stCP47 = grdDataList.TextMatrix(grdDataList.row, 8)
         Select Case stCP10
            Case 發明申請, 新型申請, 設計申請, 追加申請, 聯合申請, CIP申請, 分割, CPA申請, 再發行
               If stCP47 <> Empty Then
                  If MsgBox("本收文已輸過提申日，是否要再次輸入？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbNo Then
                     Exit Sub
                  End If
               End If
         End Select
         '2005/11/4 end
         'Add By Sindy 2017/12/27
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtSystem & txtCode(0) & txtCode(1) & txtCode(2) Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
         End If
         '2017/12/27 END
         'Add By Sindy 2016/10/7
         frm050103_2.m_strIR01 = m_strIR01
         frm050103_2.m_strIR02 = m_strIR02
         frm050103_2.m_strIR03 = m_strIR03
         frm050103_2.m_strIR04 = m_strIR04
         frm050103_2.txtCaseField(8) = m_RDate
         '2016/10/7 END
         frm050103_2.Show
         Me.Hide
      Case 1
         Unload Me
      Case 2
         If CheckKeyIn(2) Then GetCaseUpData
   End Select
End Sub

Private Sub GetUpData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String)
Dim varSaveCursor

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
' 91.09.13 modify by louis
'strExc(0) = "select cp09 收文號," & SQLDate("cp05") & " 收文日," & SQLDate("cp27") & " 發文日,cpm03 案件性質,staff.st02 代理人,cp45 彼所案號," & SQLDate("cp46") & " 代理人收達日" + _
'   " from caseprogress,casepropertymap,staff,staff_group,systemkind where sk01=cp01 and cp01=cpm01(+) and cp10=cpm02 and cp44=staff.st01(+) and sg01=" + CNULL(strGroup) + _
'   " and sg02=cp01 and sg03=cp10 and (sk02=" + CNULL(Format(專利)) + " or sk02='5') " + " and sk03=" + CNULL(Format(國外_CF)) + _
'   " and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + _
'   " and cp09<'C'" + _
'   " and cp27 is not null and cp47 is null and cp24 is null and cp61 is null order by 發文日 desc"
'91.11.10 MODIFY BY SONIA 有可能先通知申請日號, 才又寄申請收據
'strExc(0) = "select cp09 收文號," & SQLDate("cp05") & " 收文日," & SQLDate("cp27") & " 發文日,cpm03 案件性質,NVL(fa04,NVL(fa05,fa06)) 代理人,cp45 彼所案號," & SQLDate("cp46") & " 代理人收達日" + _
'            ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " + _
'            " from caseprogress,casepropertymap,FAGENT,staff_group,systemkind where sk01=cp01 and cp01=cpm01(+) and cp10=cpm02 and SUBSTR(cp44,1,8)=fa01(+) and SUBSTR(cp44,9,1)=fa02(+) and sg01=" + CNULL(strGroup) + _
'            " and sg02=cp01 and sg03=cp10 and (sk02=" + CNULL(Format(專利)) + " or sk02='5') " + " and sk03=" + CNULL(Format(國外_CF)) + _
'            " and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + _
'            " and cp09<'C'" + _
'            " and cp27 is not null and cp47 is null and cp24 is null " + _
'            "order by SORTFIELD desc"
'93.3.3 modify byo sonia 再取消 cp24 條件,因為CFP-15430先核准才通知申請案號
'strExc(0) = "select cp09 收文號," & SQLDate("cp05") & " 收文日," & SQLDate("cp27") & " 發文日,cpm03 案件性質,NVL(fa04,NVL(fa05,fa06)) 代理人,cp45 彼所案號," & SQLDate("cp46") & " 代理人收達日" + _
'            ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD " + _
'            " from caseprogress,casepropertymap,FAGENT,staff_group,systemkind where sk01=cp01 and cp01=cpm01(+) and cp10=cpm02 and SUBSTR(cp44,1,8)=fa01(+) and SUBSTR(cp44,9,1)=fa02(+) and sg01=" + CNULL(strGroup) + _
'            " and sg02=cp01 and sg03=cp10 and (sk02=" + CNULL(Format(專利)) + " or sk02='5') " + " and sk03=" + CNULL(Format(國外_CF)) + _
'            " and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + _
'            " and cp09<'C'" + _
'            " and cp27 is not null and cp24 is null " + _
'            "order by SORTFIELD desc"
'Modify by Morgan 2005/11/4 加提申日cp47,案件性質 cp10
strExc(0) = "select cp09 收文號," & SQLDate("cp05") & " 收文日," & SQLDate("cp27") & " 發文日,cpm03 案件性質,NVL(fa04,NVL(fa05,fa06)) 代理人,cp45 彼所案號," & SQLDate("cp46") & " 代理人收達日" + _
            ",DECODE(CP27,19221111,99999999,CP27) AS SORTFIELD,CP47, CP10" + _
            " from caseprogress,casepropertymap,FAGENT,staff_group,systemkind where sk01=cp01 and cp01=cpm01(+) and cp10=cpm02 and SUBSTR(cp44,1,8)=fa01(+) and SUBSTR(cp44,9,1)=fa02(+) and sg01=" + CNULL(strGroup) + _
            " and sg02=cp01 and sg03=cp10 and (sk02=" + CNULL(Format(專利)) + " or sk02='5') " + " and sk03=" + CNULL(Format(國外_CF)) + _
            " and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + _
            " and cp09<'C'" + _
            " and cp27 is not null " + _
            "order by SORTFIELD desc"
'93.3.3 END
'91.11.10 END

If strCode1 = "" Then
   intI = 1
Else
   intI = 0
End If
Set grdDataList.Recordset = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'Set grdDataList.Recordset = rsTemp
SetDataListVision grdDataList, True
'Add by Morgan 2005/11/4
grdDataList.ColWidth(grdDataList.Cols - 1) = 0
grdDataList.ColWidth(grdDataList.Cols - 2) = 0
'2005/11/4

intLastRow = 0
If grdDataList.Rows > 1 Then
   ShowBar grdDataList, intLastRow, 6
   cmdOK(0).Enabled = True
   cmdOK(0).Default = True
   'Add By Sindy 2017/10/17
   If grdDataList.Rows = 2 And m_strIR01 <> "" Then
      cmdOK(0).Value = True
   End If
   '2017/10/17 END
Else
   cmdOK(0).Enabled = False
   cmdOK(2).Default = True
End If
Screen.MousePointer = varSaveCursor
End Sub
Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

varGridWidth = Array(900, 900, 900, 1800, 2200, 1200, 1200)
SetGridDataListWidth grdDataList, varGridWidth()
blnOKtoShow = True
End Sub

Private Sub GetCaseUpData(Optional bolBlank As Boolean = False)
If bolBlank Then
   GetUpData "", "", "", ""
Else
   GetUpData txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
End If
If grdDataList.Rows <= 1 And Not IsEmptyText(txtSystem) And Not IsEmptyText(txtCode(0)) Then
   MsgBox "沒有符合條件的資料", vbOKOnly + vbCritical, "查詢資料"
End If
End Sub

'Private Sub Form_Activate()
'GetCaseUpData True
'End Sub

Public Sub QueryData()
   GetCaseUpData True
End Sub

Private Sub Form_Activate()
   'Added by Sindy 2016/10/7
   If m_strIR01 <> "" And m_Done = False Then
      txtSystem.Text = m_strCP01
      txtCode(0).Text = m_strCP02
      txtCode(1).Text = m_strCP03
      txtCode(2).Text = m_strCP04
      cmdOK(2).Value = True
      m_Done = True
      'Add By Sindy 2017/12/27
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      '2017/12/27 END
   End If
   '2016/10/7 END
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
GetUpData "", "", "", ""
SetDataListWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm050103_1 = Nothing
End Sub

Private Sub grdDataList_DblClick()
cmdOK_Click 0
End Sub
Private Sub lblNation_Change()
Dim strTemp As String
If lblNation = "" Then Exit Sub
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetNation(lblNation, strTemp) Then
If ClsPDGetNation(lblNation, strTemp) Then
   lblCountryName.Caption = strTemp
End If
End Sub

Private Sub txtCode_LostFocus(Index As Integer)
   Select Case Index
      Case 2:
         GetCaseUpData False
         If grdDataList.Rows <= 1 Then
            txtCode(0).SetFocus
         End If
      Case Else:
   End Select
End Sub

Private Sub txtSystem_Change()
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
If grdDataList.Rows > 1 Then GetCaseUpData True
End Sub

Private Sub txtSystem_GotFocus()
txtSystem.SelStart = 0
txtSystem.SelLength = Len(txtSystem.Text)
End Sub
Private Sub txtSystem_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtSystem_Validate(Cancel As Boolean)
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetGroupCase(txtSystem, strGroup) = False Then
If ClsPDGetGroupCase(txtSystem, strGroup) = False Then
   ShowMsg MsgText(1056)
   Cancel = True
   txtSystem_GotFocus
End If
End Sub
'Private Sub txtCode_Change(Index As Integer)
'If cboCaseName.ListCount > 0 Then cboCaseName.Clear
'If grdDataList.Rows > 1 Then
'   GetCaseUpData True
'End If
'End Sub

Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index).Text)
End Sub
Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = False Then
   Cancel = True
End If
End Sub
Private Function CheckKeyIn(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String, strNation As String

If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
   ShowMsg MsgText(9)
ElseIf intIndex = 2 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, , strNation, , , False) Then
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, , strNation, , , False) Then
      lblNation = strNation
      SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
   End If
   CheckKeyIn = True
Else
   CheckKeyIn = True
End If
End Function
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
      ShowBar grdDataList, intLastRow, 6
      blnOKtoShow = True
   End If
End If
End Sub

Public Sub Clear()
   txtSystem = Empty
   txtCode(0) = Empty
   txtCode(1) = Empty
   txtCode(2) = Empty
   lblNation = ""
   lblCountryName = ""
   cboCaseName.Clear
   grdDataList.Rows = 1
   txtSystem.SetFocus
End Sub
