VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010601_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人已收達/已提申"
   ClientHeight    =   5745
   ClientLeft      =   -2910
   ClientTop       =   1200
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   8352
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   6300
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "提申資料(&F)"
      Height          =   405
      Index           =   2
      Left            =   7128
      TabIndex        =   11
      Top             =   70
      Width           =   1200
   End
   Begin VB.Frame fraTrademark 
      BorderStyle     =   0  '沒有框線
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5280
      Width           =   3372
      Begin VB.TextBox txtChoose 
         Height          =   264
         Left            =   600
         TabIndex        =   10
         Top             =   60
         Width           =   492
      End
      Begin VB.Label Label11 
         Caption         =   "結果：              （1：已收達  2：已提申）"
         Height          =   252
         Left            =   0
         TabIndex        =   25
         Top             =   60
         Width           =   3492
      End
   End
   Begin VB.TextBox txtSystem 
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   0
      Top             =   540
      Width           =   732
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   3195
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   5636
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
   Begin VB.Frame fraElse 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
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
      Left            =   1920
      TabIndex        =   20
      Top             =   540
      Width           =   2532
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   2
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   3
         Top             =   0
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   1
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   2
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   264
         Index           =   0
         Left            =   0
         MaxLength       =   6
         TabIndex        =   1
         Top             =   0
         Width           =   1212
      End
   End
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
      Left            =   1920
      TabIndex        =   21
      Top             =   540
      Visible         =   0   'False
      Width           =   2412
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   0
         Left            =   0
         MaxLength       =   5
         TabIndex        =   4
         Top             =   0
         Width           =   852
      End
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   1
         Left            =   960
         MaxLength       =   1
         TabIndex        =   5
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   2
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   6
         Top             =   0
         Width           =   372
      End
      Begin VB.TextBox txtTFCode 
         Height          =   264
         Index           =   3
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   7
         Top             =   0
         Width           =   492
      End
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1080
      TabIndex        =   8
      Top             =   1440
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
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   1740
      Width           =   3975
      VariousPropertyBits=   27
      Size            =   "7011;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNumber1 
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lblNumber2 
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   1140
      Width           =   3975
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "申請人："
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label lblNumber 
      Caption         =   "審定號數："
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   540
      Width           =   975
   End
End
Attribute VB_Name = "frm02010601_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/9 改成Form2.0 (cboCaseName,lblAgent)
'Memo By Morgan 2012/12/17 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
Dim intCols As Integer
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
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
         'Add By Sindy 2017/12/28
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtSystem & txtCode(0) & txtCode(1) & txtCode(2) Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
         End If
         '2017/12/28 END
         'Add By Sindy 2016/10/11
         If Not m_PrevForm Is Nothing Then
            Call frm02010601_2.SetParent(m_PrevForm)
         End If
         '2016/10/11 END
         'Add By Sindy 2016/10/7
         frm02010601_2.m_strIR01 = m_strIR01
         frm02010601_2.m_strIR02 = m_strIR02
         frm02010601_2.m_strIR03 = m_strIR03
         frm02010601_2.m_strIR04 = m_strIR04
         '2016/10/7 END
         frm02010601_2.Show
         Me.Hide
      Case 1
         Unload Me
      Case 2
         If txtSystem = 馬德里案 Then
            bolRt = CheckKeyIn1(3)
         Else
            bolRt = CheckKeyIn2(2)
         End If
         If bolRt Then GetAgentReceivedCaseData
   End Select
End Sub

'Add By Cheng 2002/07/25
Public Sub ClickCmdOk(Index As Integer)
   cmdOK_Click Index
End Sub

Private Sub GetAgentReceivedCaseData(Optional bolBlank As Boolean = False)
'TF為馬德里案，另外判斷
If bolBlank = True Then
   GetAgentReceivedData "", "", "", ""
Else
   If txtSystem = 馬德里案 Then
      GetAgentReceivedData txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
          IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3))
   Else
      GetAgentReceivedData txtSystem, txtCode(0), _
         IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
   End If
End If
End Sub
Private Sub GetAgentReceivedData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String)
Dim varSaveCursor

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/02 不用 dll 了
'Set grdDataList.Recordset = objPublicData.ReadAgentReceivedRst(intPCaseKind, strGroup, strCode1, strCode2, strCode3, strCode4)
Set grdDataList.Recordset = ClsPDReadAgentReceivedRst(intPCaseKind, strGroup, strCode1, strCode2, strCode3, strCode4)
SetDataListVision grdDataList
intLastRow = 0
If grdDataList.Rows > 1 Then
   
   'Add By Cheng 2002/01/09
   ' 91.09.13 modify by louis (排序)
   '依發文日由大至小排序
   'Me.grdDataList.Col = 3
   'Me.grdDataList.Sort = flexSortGenericDescending
   
   ShowBar grdDataList, intLastRow, intCols
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

If intPCaseKind = 專利 Then
   varGridWidth = Array(1500, 1500, 3100, 1500, 3200)
Else
   varGridWidth = Array(1000, 1000, 2500, 1200, 1200, 1200, 1000, 1700)
End If
SetGridDataListWidth grdDataList, varGridWidth()
blnOKtoShow = True
End Sub

Private Sub Form_Activate()
   GetAgentReceivedCaseData True
   'Added by Sindy 2016/10/7
   If m_strIR01 <> "" And m_Done = False Then
      txtSystem.Text = m_strCP01
      txtCode(0).Text = m_strCP02
      txtCode(1).Text = m_strCP03
      txtCode(2).Text = m_strCP04
      cmdOK(2).Value = True
      m_Done = True
      'Add By Sindy 2017/12/28
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      '2017/12/28 END
   End If
   '2016/10/7 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   If intPCaseKind = 專利 Then
      lblNumber = "證書號數："
      If intPWhere = 國外_CF Then
         frm02010601_1.Caption = "代理人已收達"
         fraTrademark.Visible = False
         txtChoose = "1"
      End If
   End If
   If intPCaseKind = 專利 Then
      intCols = 4
   Else
      intCols = 6
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2016/10/11
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   'Add By Cheng 2002/07/18
   Set frm02010601_1 = Nothing
End Sub

Private Sub grdDataList_DblClick()
cmdOK_Click 0
End Sub
Private Sub txtSystem_Change()
If txtSystem.Text = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
Else
   fraTF.Visible = False
   fraElse.Visible = True
End If
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
If grdDataList.Rows > 1 Then GetAgentReceivedCaseData True
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
   ShowMsg MsgText(9171)
   Cancel = True
   txtSystem_GotFocus
End If
End Sub
Private Sub txtCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
If grdDataList.Rows > 1 Then GetAgentReceivedCaseData True
End Sub
Private Sub txtTFCode_Change(Index As Integer)
If cboCaseName.ListCount > 0 Then cboCaseName.Clear
If grdDataList.Rows > 1 Then GetAgentReceivedCaseData True
End Sub
Private Sub txtTFCode_GotFocus(Index As Integer)
txtTFCode(Index).SelStart = 0
txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub
Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn1 (Index)
End Sub
Private Function CheckKeyIn1(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strCustomer As String, strNumber1 As String, strNumber2 As String

If Len(txtTFCode(intIndex)) > 0 And Len(txtTFCode(intIndex)) < txtTFCode(intIndex).MaxLength Then
   ShowMsg MsgText(33)
ElseIf intIndex = 3 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
      SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
      lblNumber1 = strNumber1
      lblNumber2 = strNumber2
      lblAgent = strCustomer
      CheckKeyIn1 = True
   End If
Else
   CheckKeyIn1 = True
End If
End Function
Private Sub txtCode_GotFocus(Index As Integer)
txtCode(Index).SelStart = 0
txtCode(Index).SelLength = Len(txtCode(Index).Text)
End Sub
Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn2 (Index)
End Sub
Private Function CheckKeyIn2(ByRef intIndex As Integer) As Boolean
Dim strCaseName1 As String, strCaseName2 As String, strCaseName3 As String
Dim strCustomer As String, strNumber1 As String, strNumber2 As String

If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
   ShowMsg MsgText(33)
ElseIf intIndex = 2 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
        IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strCaseName1, strCaseName2, strCaseName3, strCustomer, , strNumber1, strNumber2) Then
      SetNameToCombo cboCaseName, strCaseName1, strCaseName2, strCaseName3
      lblNumber1 = strNumber1
      lblNumber2 = strNumber2
      lblAgent = strCustomer
      CheckKeyIn2 = True
   End If
Else
   CheckKeyIn2 = True
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
      ShowBar grdDataList, intLastRow, intCols
      blnOKtoShow = True
   End If
End If
End Sub
'Add By Cheng 2002/01/16
Public Sub Clear()
Me.txtSystem.Text = Empty
Me.txtCode(0).Text = Empty
Me.txtCode(1).Text = Empty
Me.txtCode(2).Text = Empty
Me.lblNumber1.Caption = Empty
Me.lblNumber2.Caption = Empty
Me.cboCaseName.Clear
Me.lblAgent.Caption = Empty
SetDataListWidth
Me.txtSystem.SetFocus
End Sub
