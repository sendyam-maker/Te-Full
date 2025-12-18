VERSION 5.00
Begin VB.Form frm05010401_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "一般來函"
   ClientHeight    =   1860
   ClientLeft      =   930
   ClientTop       =   2385
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form23"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5055
   Begin VB.TextBox txtAppNo 
      Height          =   264
      Left            =   1620
      MaxLength       =   25
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtCaseCode 
      Height          =   264
      Index           =   3
      Left            =   1620
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1470
      Width           =   972
   End
   Begin VB.TextBox txtCaseCode 
      Height          =   264
      Index           =   2
      Left            =   1620
      MaxLength       =   15
      TabIndex        =   14
      Top             =   2610
      Width           =   2052
   End
   Begin VB.Frame fraCode 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   1620
      TabIndex        =   15
      Top             =   1020
      Width           =   3225
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
         Height          =   372
         Left            =   870
         TabIndex        =   17
         Top             =   0
         Width           =   2652
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   2
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   4
            Top             =   0
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   1
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   3
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   0
            Left            =   -30
            MaxLength       =   6
            TabIndex        =   2
            Top             =   0
            Width           =   1212
         End
      End
      Begin VB.Frame fraTF 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   840
         TabIndex        =   16
         Top             =   0
         Width           =   2652
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   3
            Left            =   2040
            TabIndex        =   13
            Top             =   0
            Width           =   492
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   2
            Left            =   1560
            TabIndex        =   12
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   1
            Left            =   1080
            TabIndex        =   11
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   972
         End
      End
      Begin VB.TextBox txtSystem 
         Height          =   288
         Left            =   0
         MaxLength       =   3
         TabIndex        =   1
         Top             =   0
         Width           =   732
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4020
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3144
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox txtCaseCode 
      Height          =   264
      Index           =   1
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   9
      Top             =   2280
      Width           =   1332
   End
   Begin VB.TextBox txtCaseCode 
      Height          =   264
      Index           =   0
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1920
      Width           =   1332
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   495
      TabIndex        =   20
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   495
      TabIndex        =   19
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日："
      Height          =   180
      Left            =   495
      TabIndex        =   18
      Top             =   1500
      Width           =   1080
   End
End
Attribute VB_Name = "frm05010401_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (無)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim m_blnCancel As Boolean
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
'2016/10/7 END


Private Sub cmdOK_Click(Index As Integer)
   Dim i As Integer

   If Index = 0 Then
      If CheckDataValid() = False Then
         GoTo EXITSUB
      End If
      
      '日期
      If CheckKeyIn(3) <> 1 Then
         Screen.MousePointer = vbDefault
         txtCaseCode(3).SetFocus
         txtCaseCode_GotFocus 3
         Exit Sub
      End If
      
      If CheckKeyIn2(2) = False Then
         Screen.MousePointer = vbDefault
         txtSystem.SetFocus
         Exit Sub
      Else
         For i = 0 To 2
            frm05010401_2.lblCode(i) = txtCode(i)
         Next
      End If
         
      'Add By Sindy 2017/12/28
      If m_strIR01 <> "" Then
         If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtSystem & txtCode(0) & txtCode(1) & txtCode(2) Then
            MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
            Exit Sub
         End If
      End If
      '2017/12/28 END
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      frm05010401_2.lblSystem = txtSystem
      'Add By Sindy 2016/10/7
      frm05010401_2.m_strIR01 = m_strIR01
      frm05010401_2.m_strIR02 = m_strIR02
      frm05010401_2.m_strIR03 = m_strIR03
      frm05010401_2.m_strIR04 = m_strIR04
      '2016/10/7 END
      frm05010401_2.Show
      frm05010401_2.Caption = frm05010401_1.Caption
      frm05010401_2.QueryData
      Me.Hide
       
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
   Else
      Unload Me
   End If

EXITSUB:
End Sub

Private Sub Form_Activate()
   'Added by Sindy 2016/10/7
   If m_strIR01 <> "" And m_Done = False Then
      txtSystem.Text = m_strCP01
      txtCode(0).Text = m_strCP02
      txtCode(1).Text = m_strCP03
      txtCode(2).Text = m_strCP04
      txtAppNo.Text = m_AppNo
      txtCaseCode(3).Text = m_RDate
      'cmdOK(0).Value = True
      m_Done = True
      'Add By Sindy 2017/12/28
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      '2017/12/28 END
   End If
   '2016/10/7 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   If intPWhere <> 國外_CF Then
      txtCaseCode(3).MaxLength = 7
   Else
      txtCaseCode(3).MaxLength = 8
   End If
   If intPCaseKind = 專利 And intPWhere = 國外_CF Then
      Label9.Caption = "櫃台收文日:"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm05010401_1 = Nothing
End Sub

Private Sub txtAppNo_GotFocus()
   TextInverse txtAppNo
End Sub

Private Sub txtCaseCode_GotFocus(Index As Integer)
   txtCaseCode(Index).SelStart = 0
   txtCaseCode(Index).SelLength = Len(txtCaseCode(Index))
End Sub

Private Sub txtCaseCode_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = -1 Then
      Cancel = True
      txtCaseCode_GotFocus Index
   End If
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
CheckKeyIn = -1
Select Case intIndex
             Case 0, 1, 2
                        If txtCaseCode(intIndex) = "" Then
                           ShowMsg MsgText(9015)
                        Else
                           CheckKeyIn = 1
                        End If
             Case 3
                        If CheckIsTaiwanDate(txtCaseCode(intIndex)) Then
                           'Modify by Morgan 2010/8/11 百年蟲
                           'If txtCaseCode(intIndex) > GetTaiwanTodayDate Then
                           If Val(txtCaseCode(intIndex)) > Val(strSrvDate(2)) Then
                              ShowMsg MsgText(1050)
                           Else
                              CheckKeyIn = 1
                           End If
                        End If
             Case Else
                        CheckKeyIn = 1
End Select
End Function

Private Sub txtCode_LostFocus(Index As Integer)
Select Case Index
Case 2
   If m_blnCancel = True Then
      Me.txtSystem.SetFocus
   End If
End Select
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
Private Sub txtTFCode_GotFocus(Index As Integer)
txtTFCode(Index).SelStart = 0
txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub
Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
CheckKeyIn1 (Index)
End Sub
Private Function CheckKeyIn1(ByRef intIndex As Integer) As Boolean
If Len(txtTFCode(intIndex)) > 0 And Len(txtTFCode(intIndex)) < txtTFCode(intIndex).MaxLength Then
   ShowMsg MsgText(33)
ElseIf intIndex = 3 Then
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3))) Then
   If ClsPDCheckCaseCodeIsExist(txtSystem, txtTFCode(0) + IIf(txtTFCode(1) = "", "0", txtTFCode(1)), _
         IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3))) Then
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
   Dim Nation As String

   If Len(txtCode(intIndex)) > 0 And Len(txtCode(intIndex)) < txtCode(intIndex).MaxLength Then
      ShowMsg MsgText(33)
   ElseIf intIndex = 2 Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.CheckCaseCodeIsExist(txtSystem, txtCode(0), _
           IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), , , , , Nation) Then
      If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), _
           IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), , , , , Nation) Then
         '92.6.28 add by sonia
         If Nation = 台灣國家代號 Then
            MsgBox "此案件之申請國家為 台灣 !!", vbOKOnly + vbCritical, "檢核資料"
            txtCode(0).SetFocus
            m_blnCancel = True
            Exit Function
         Else
            CheckKeyIn2 = True
            'Add By Cheng 2002/08/27
            m_blnCancel = False
         End If
         '92.6.28 end
      Else
         m_blnCancel = True
      End If
   Else
      CheckKeyIn2 = True
   End If
End Function

' 90.07.12 modify by louis (檢查資料是否輸入完整)
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   CheckDataValid = False
   
   ' 來函收文日
   If IsEmptyText(txtCaseCode(3)) = False Then
      strDate = txtCaseCode(3)
      If CheckIsTaiwanDate(strDate, False) = False Then
         strTit = "檢核資料"
         strMsg = "來函收文日日期格式不正確!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCaseCode(3).SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add by Morgan 2008/5/22 申請案號檢查
   If txtAppNo = "" Then
      If txtSystem = "CFP" Then
         MsgBox "CFP案必須輸入申請案號！"
         txtAppNo.SetFocus
         GoTo EXITSUB
      End If
   Else
      strExc(1) = "N"
      strExc(0) = "select pa02,pa03,pa04 from patent where pa01='CFP' and pa11='" & txtAppNo & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            'Modify by Morgan 2010/1/13 不必檢查多國碼(子案沒有申請號)
            'If txtCode(0) = RsTemp("pa02") And Left(txtCode(1) & "0", 1) = RsTemp("pa03") And Left(txtCode(2) & "00", 2) = RsTemp("pa04") Then
            If txtCode(0) = RsTemp("pa02") And Left(txtCode(1) & "0", 1) = RsTemp("pa03") Then
               strExc(1) = "Y"
               Exit Do
            End If
            .MoveNext
         Loop
         End With
         If strExc(1) = "N" Then
            MsgBox "本所案號輸入錯誤！"
            txtCode(0).SetFocus
            txtCode_GotFocus (0)
            GoTo EXITSUB
         End If
      Else
         MsgBox "申請案號不存在！"
         txtAppNo.SetFocus
         txtAppNo_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Public Sub Clear()
   txtCaseCode(0) = Empty
   txtCaseCode(1) = Empty
   txtSystem = Empty
   txtCode(0) = Empty
   txtCode(1) = Empty
   txtCode(2) = Empty
   txtCaseCode(2) = Empty
   txtCaseCode(3) = Empty
   
   If Me.txtSystem.Enabled Then
      'Modify by Morgan 2008/5/22
      'Me.txtSystem.SetFocus
      txtAppNo = Empty
      txtAppNo.SetFocus
      'end 2008/5/22
   End If
End Sub
