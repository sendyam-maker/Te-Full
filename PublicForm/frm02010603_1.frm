VERSION 5.00
Begin VB.Form frm02010603_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人其他來函輸入"
   ClientHeight    =   1635
   ClientLeft      =   930
   ClientTop       =   2385
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form23"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   5055
   Begin VB.TextBox txtCaseCode 
      Height          =   264
      Index           =   3
      Left            =   1620
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1200
      Width           =   972
   End
   Begin VB.TextBox txtCaseCode 
      Height          =   264
      Index           =   2
      Left            =   1620
      MaxLength       =   15
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2052
   End
   Begin VB.Frame fraCode 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   1620
      TabIndex        =   18
      Top             =   750
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
         TabIndex        =   20
         Top             =   0
         Width           =   2652
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   2
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   3
            Top             =   0
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   1
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   2
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Height          =   288
            Index           =   0
            Left            =   -30
            MaxLength       =   6
            TabIndex        =   1
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
         TabIndex        =   19
         Top             =   0
         Width           =   2652
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   3
            Left            =   2040
            TabIndex        =   16
            Top             =   0
            Width           =   492
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   2
            Left            =   1560
            TabIndex        =   15
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Top             =   0
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   288
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   972
         End
      End
      Begin VB.TextBox txtSystem 
         Height          =   288
         Left            =   0
         MaxLength       =   3
         TabIndex        =   0
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
      Left            =   3972
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3144
      TabIndex        =   5
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox txtCaseCode 
      Height          =   264
      Index           =   1
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2010
      Width           =   1332
   End
   Begin VB.TextBox txtCaseCode 
      Height          =   264
      Index           =   0
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1650
      Width           =   1332
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "本所案號："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   2
      Left            =   270
      TabIndex        =   9
      Top             =   780
      Value           =   -1  'True
      Width           =   1212
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "審定號數："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   1
      Left            =   300
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2010
      Width           =   1212
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "本所發文號："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   3
      Left            =   300
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2340
      Width           =   1452
   End
   Begin VB.OptionButton optChoose 
      Caption         =   "申請案號："
      CausesValidation=   0   'False
      Height          =   252
      Index           =   0
      Left            =   300
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1650
      Width           =   1212
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日："
      Height          =   180
      Left            =   540
      TabIndex        =   21
      Top             =   1230
      Width           =   1080
   End
End
Attribute VB_Name = "frm02010603_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/10 改成Form2.0 (無)
'Memo By Morgan 2012/12/17 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

Public intOpt As Integer
Dim bolRun  As Boolean
'Add By Cheng 2002/08/27
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
Dim m_PrevForm As Form 'Add By Sindy 2016/10/11


'Add By Sindy 2016/10/11
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim i As Integer
   
   If Index = 0 Then
      If CheckDataValid() = False Then
         GoTo EXITSUB
      End If
       'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
         If txtCode(1) = "" Then txtCode(1) = "0"
         If txtCode(2) = "" Then txtCode(2) = "00"
         If FMP2open = True Then
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtSystem, txtCode(0), txtCode(1), txtCode(2)) = False Then
              txtCode(0).SetFocus
              GoTo EXITSUB
           End If
         End If
      
      ' 90.07.02 modify by louis
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      
      If CheckKeyIn(3) <> 1 Then
         Screen.MousePointer = vbDefault
         txtCaseCode(3).SetFocus
         txtCaseCode_GotFocus 3
         Exit Sub
      End If
      If intOpt = 2 Then
         'Add By Sindy 2017/12/27
         If m_strIR01 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtSystem & txtCode(0) & txtCode(1) & txtCode(2) Then
               MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Sub
            End If
            Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
         End If
         '2017/12/27 END
         If txtSystem.Text = 馬德里案 Then
            If CheckKeyIn1(3) = False Then
               Screen.MousePointer = vbDefault
               Exit Sub
            Else
               For i = 0 To 3
                      frm02010603_2.lblTFCode(i) = txtTFCode(i)
               Next
            End If
         Else
            If CheckKeyIn2(2) = False Then
               Screen.MousePointer = vbDefault
               txtSystem.SetFocus
               Exit Sub
            Else
               For i = 0 To 2
                  frm02010603_2.lblCode(i) = txtCode(i)
               Next
            End If
         End If
         frm02010603_2.lblSystem = txtSystem
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
         'Add By Sindy 2016/10/11
         If Not m_PrevForm Is Nothing Then
            Call frm02010603_2.SetParent(m_PrevForm)
         End If
         '2016/10/11 END
         'Add By Sindy 2016/10/7
         frm02010603_2.m_strIR01 = m_strIR01
         frm02010603_2.m_strIR02 = m_strIR02
         frm02010603_2.m_strIR03 = m_strIR03
         frm02010603_2.m_strIR04 = m_strIR04
         '2016/10/7 END
         frm02010603_2.Show
         frm02010603_2.Caption = frm02010603_1.Caption
         frm02010603_2.QueryData
      Else
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
   '      frm02010603_7.Show
      End If
      Me.Hide
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
      txtCaseCode(3).Text = m_RDate
      optChoose(2).Value = True
      cmdOK(0).Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
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
   'Add By Cheng 2002/07/24
   If intPCaseKind = 專利 And intPWhere = 國外_CF Then
      Label9.Caption = "櫃台收文日:"
   End If
   bolRun = False
   intOpt = 2
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2016/10/11
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   'Add By Cheng 2002/07/18
   Set frm02010603_1 = Nothing
End Sub

Private Sub optChoose_Click(Index As Integer)
   intOpt = Index
   txtCaseCode(0).Enabled = False
   txtCaseCode(1).Enabled = False
   txtCaseCode(2).Enabled = False
   fraCode.Enabled = False
   Select Case Index
                Case 0
                           txtCaseCode(0).Enabled = True
                           txtCaseCode(0).SetFocus
                Case 1
                           txtCaseCode(1).Enabled = True
                           txtCaseCode(1).SetFocus
                Case 2
                           fraCode.Enabled = True
                           txtSystem.SetFocus
                Case 3
                           txtCaseCode(2).Enabled = True
                           txtCaseCode(2).SetFocus
   End Select
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
                           'Modify by Morgan 2010/8/18 百年蟲
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

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_LostFocus(Index As Integer)
Select Case Index
Case 2
   'Add By Cheng 2002/08/27
   If m_blnCancel = True Then
      Me.txtSystem.SetFocus
   End If
End Select
End Sub

Private Sub txtSystem_Change()
If txtSystem.Text = 馬德里案 Then
   fraTF.Visible = True
   fraElse.Visible = False
Else
   fraTF.Visible = False
   fraElse.Visible = True
End If
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
   'Add By Cheng 2002/01/08
   '確保作業處理完游標停在應該停的位置
   If Me.txtSystem.Enabled Then Me.txtSystem.SetFocus
End Sub
