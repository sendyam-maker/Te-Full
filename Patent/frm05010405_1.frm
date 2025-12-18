VERSION 5.00
Begin VB.Form frm05010405_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "年費逾期補繳通知函"
   ClientHeight    =   1890
   ClientLeft      =   930
   ClientTop       =   2385
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form23"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5055
   Begin VB.TextBox txtCode 
      Height          =   288
      Index           =   0
      Left            =   2475
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1020
      Width           =   1212
   End
   Begin VB.TextBox txtCode 
      Height          =   288
      Index           =   1
      Left            =   3780
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1020
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   288
      Index           =   2
      Left            =   4260
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1020
      Width           =   492
   End
   Begin VB.TextBox txtSystem 
      Enabled         =   0   'False
      Height          =   288
      Left            =   1620
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "CFP"
      Top             =   1020
      Width           =   732
   End
   Begin VB.TextBox txtAppNo 
      Height          =   264
      Left            =   1620
      MaxLength       =   25
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtDate 
      Height          =   264
      Left            =   1620
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1470
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   3972
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   495
      TabIndex        =   10
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   495
      TabIndex        =   9
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日："
      Height          =   180
      Left            =   495
      TabIndex        =   8
      Top             =   1500
      Width           =   1080
   End
End
Attribute VB_Name = "frm05010405_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/8 改成Form2.0 (無)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Morgan 2008/12/15
Option Explicit

'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String, m_AppNo As String
Dim m_Done As Boolean
'2016/10/7 END


Private Sub Form_Activate()
   'Added by Sindy 2016/10/7
   If m_strIR01 <> "" And m_Done = False Then
      txtSystem.Text = m_strCP01
      txtCode(0).Text = m_strCP02
      txtCode(1).Text = m_strCP03
      txtCode(2).Text = m_strCP04
      txtAppNo.Text = m_AppNo
      txtDate.Text = m_RDate
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm05010405_1 = Nothing
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
   CheckKeyIn = -1
   Select Case intIndex
      Case 3
         
         If CheckIsTaiwanDate(txtDate) Then
            If Val(txtDate) > Val(strSrvDate(2)) Then
               ShowMsg MsgText(1050)
            Else
               CheckKeyIn = 1
            End If
         End If
      Case Else
         CheckKeyIn = 1
   End Select
End Function

Public Sub Clear()
   txtCode(0) = Empty
   txtCode(1) = Empty
   txtCode(2) = Empty
   txtDate = Empty
   txtAppNo = Empty
   If txtAppNo.Visible = True Then
      txtAppNo.SetFocus
   End If
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
End Sub

Private Sub txtDate_GotFocus()
   TextInverse txtDate
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
   If txtDate = "" Then
      MsgBox "來函收文日不可空白!!", vbExclamation
      Cancel = True
   ElseIf CheckIsTaiwanDate(txtDate) = False Then
      Cancel = True
   ElseIf Val(txtDate) > Val(strSrvDate(2)) Then
      ShowMsg MsgText(1050)
      Cancel = True
   End If
End Sub

Private Sub txtAppNo_GotFocus()
   TextInverse txtAppNo
End Sub


Private Sub cmdOK_Click(Index As Integer)
   If Index = 0 Then
      'Add By Sindy 2017/12/28
      If m_strIR01 <> "" Then
         If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtSystem & txtCode(0) & txtCode(1) & txtCode(2) Then
            MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
            Exit Sub
         End If
      End If
      '2017/12/28 END
      Screen.MousePointer = vbHourglass
      If TxtValidate = True Then
         With frm05010405_2
            .txtSystem = txtSystem
            .txtCode(0) = txtCode(0)
            .txtCode(1) = txtCode(1)
            .txtCode(2) = txtCode(2)
            .lblCaseField(4) = txtDate
            'Add By Sindy 2016/10/7
            .m_strIR01 = m_strIR01
            .m_strIR02 = m_strIR02
            .m_strIR03 = m_strIR03
            .m_strIR04 = m_strIR04
            '2016/10/7 END
            Set .frmParent = Me
            .Show
         End With
         Me.Hide
      End If
      Screen.MousePointer = vbDefault
   Else
      Unload Me
   End If
End Sub

'檢查資料是否輸入完整
Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If txtCode(1) = "" Then txtCode(1) = "0"
   If txtCode(2) = "" Then txtCode(2) = "00"
   
   If txtAppNo = "" Then
      MsgBox "CFP案必須輸入申請案號！"
      txtAppNo.SetFocus
      Exit Function
   ElseIf AppNoCheck = False Then
      Exit Function
   End If
   
   txtDate_Validate bCancel
   If bCancel = True Then
      txtDate.SetFocus
      txtDate_GotFocus
      Exit Function
   End If
   
   TxtValidate = True
End Function

Private Function AppNoCheck() As Boolean
   Dim bFound As Boolean
   
   strExc(0) = "select pa01,pa02,pa03,pa04 from patent where pa11='" & Trim(txtAppNo) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         'Modify by Morgan 2010/1/13 不必檢查多國碼(子案沒有申請號)
         If txtSystem = RsTemp("pa01") And txtCode(0) = RsTemp("pa02") And txtCode(1) = RsTemp("pa03") Then
            bFound = True
            Exit Do
         End If
         .MoveNext
      Loop
      End With
      If bFound = False Then
         MsgBox "本所案號輸入錯誤！"
         txtCode(0).SetFocus
         txtCode_GotFocus 1
      End If
   Else
      MsgBox "申請案號不存在！"
      txtAppNo.SetFocus
      txtAppNo_GotFocus
   End If
   AppNoCheck = bFound
End Function

