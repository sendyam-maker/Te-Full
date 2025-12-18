VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090801_14 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "對造資料"
   ClientHeight    =   8544
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8544
   ScaleWidth      =   8220
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox textCRL157 
      Height          =   330
      Index           =   4
      Left            =   1110
      MaxLength       =   600
      TabIndex        =   9
      Top             =   7460
      Width           =   6900
   End
   Begin VB.TextBox textCRL157 
      Height          =   330
      Index           =   3
      Left            =   1110
      MaxLength       =   600
      TabIndex        =   8
      Top             =   5830
      Width           =   6900
   End
   Begin VB.TextBox textCRL157 
      Height          =   330
      Index           =   2
      Left            =   1110
      MaxLength       =   600
      TabIndex        =   7
      Top             =   4140
      Width           =   6900
   End
   Begin VB.TextBox textCRL157 
      Height          =   330
      Index           =   1
      Left            =   1110
      MaxLength       =   600
      TabIndex        =   6
      Top             =   2412
      Width           =   6900
   End
   Begin VB.TextBox textCRL157 
      Height          =   330
      Index           =   0
      Left            =   1110
      MaxLength       =   120
      TabIndex        =   5
      Top             =   800
      Width           =   6900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回上一頁"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   21
      Top             =   30
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   6060
      TabIndex        =   19
      Top             =   30
      Width           =   930
   End
   Begin MSForms.TextBox txtIsCmp 
      Height          =   336
      Index           =   4
      Left            =   4200
      TabIndex        =   51
      Top             =   8160
      Visible         =   0   'False
      Width           =   300
      VariousPropertyBits=   679493659
      MaxLength       =   70
      Size            =   "529;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtIsCmp 
      Height          =   336
      Index           =   3
      Left            =   4200
      TabIndex        =   50
      Top             =   6550
      Visible         =   0   'False
      Width           =   300
      VariousPropertyBits=   679493659
      MaxLength       =   70
      Size            =   "529;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtIsCmp 
      Height          =   336
      Index           =   2
      Left            =   4200
      TabIndex        =   49
      Top             =   4850
      Visible         =   0   'False
      Width           =   300
      VariousPropertyBits=   679493659
      MaxLength       =   70
      Size            =   "529;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtIsCmp 
      Height          =   336
      Index           =   1
      Left            =   4200
      TabIndex        =   48
      Top             =   3120
      Visible         =   0   'False
      Width           =   300
      VariousPropertyBits=   679493659
      MaxLength       =   70
      Size            =   "529;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtIsCmp 
      Height          =   336
      Index           =   0
      Left            =   4200
      TabIndex        =   40
      Top             =   1500
      Visible         =   0   'False
      Width           =   300
      VariousPropertyBits=   679493659
      MaxLength       =   70
      Size            =   "529;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCRL159 
      Height          =   336
      Index           =   4
      Left            =   2160
      TabIndex        =   47
      Top             =   8160
      Width           =   2004
      VariousPropertyBits=   679493659
      MaxLength       =   70
      Size            =   "3528;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "身份證字號 / 統一編號 :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   46
      Top             =   8160
      Width           =   2004
   End
   Begin MSForms.TextBox textCRL159 
      Height          =   336
      Index           =   3
      Left            =   2160
      TabIndex        =   45
      Top             =   6550
      Width           =   2004
      VariousPropertyBits=   679493659
      MaxLength       =   70
      Size            =   "3528;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "身份證字號 / 統一編號 :"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   44
      Top             =   6550
      Width           =   2004
   End
   Begin VB.Label lbl1 
      Caption         =   "對造５  :"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   43
      Top             =   6850
      Width           =   804
   End
   Begin MSForms.TextBox textCRL159 
      Height          =   336
      Index           =   2
      Left            =   2160
      TabIndex        =   42
      Top             =   4850
      Width           =   2000
      VariousPropertyBits=   679493659
      MaxLength       =   70
      Size            =   "3528;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "身份證字號 / 統一編號 :"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   41
      Top             =   4850
      Width           =   2004
   End
   Begin VB.Label lbl1 
      Caption         =   "對造４ :"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   39
      Top             =   5160
      Width           =   800
   End
   Begin VB.Label lbl1 
      Caption         =   "對造３ :"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   38
      Top             =   3480
      Width           =   800
   End
   Begin MSForms.TextBox textCRL159 
      Height          =   336
      Index           =   1
      Left            =   2160
      TabIndex        =   37
      Top             =   3120
      Width           =   2000
      VariousPropertyBits=   679493659
      MaxLength       =   70
      Size            =   "3528;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "身份證字號 / 統一編號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   36
      Top             =   3120
      Width           =   2004
   End
   Begin VB.Label Label1 
      Caption         =   "中文名稱 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   480
      Width           =   900
   End
   Begin VB.Label lbl1 
      Caption         =   "對造２ :"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   34
      Top             =   1800
      Width           =   800
   End
   Begin MSForms.TextBox textCRL159 
      Height          =   336
      Index           =   0
      Left            =   2160
      TabIndex        =   15
      Top             =   1500
      Width           =   2000
      VariousPropertyBits=   679493659
      MaxLength       =   70
      Size            =   "3528;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "身份證字號 / 統一編號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   1500
      Width           =   2004
   End
   Begin VB.Label Label3 
      Caption         =   "日文名稱:"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   32
      Top             =   7820
      Width           =   900
   End
   Begin MSForms.TextBox textCRL158 
      Height          =   336
      Index           =   4
      Left            =   1110
      TabIndex        =   14
      Top             =   7820
      Width           =   6900
      VariousPropertyBits=   679493659
      MaxLength       =   120
      Size            =   "12171;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "日文名稱 :"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   31
      Top             =   6190
      Width           =   900
   End
   Begin MSForms.TextBox textCRL158 
      Height          =   336
      Index           =   3
      Left            =   1110
      TabIndex        =   13
      Top             =   6190
      Width           =   6900
      VariousPropertyBits=   679493659
      MaxLength       =   120
      Size            =   "12171;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "日文名稱:"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   30
      Top             =   4500
      Width           =   900
   End
   Begin MSForms.TextBox textCRL158 
      Height          =   336
      Index           =   2
      Left            =   1110
      TabIndex        =   12
      Top             =   4500
      Width           =   6900
      VariousPropertyBits=   679493659
      MaxLength       =   120
      Size            =   "12171;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "日文名稱:"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   2760
      Width           =   900
   End
   Begin MSForms.TextBox textCRL158 
      Height          =   336
      Index           =   1
      Left            =   1110
      TabIndex        =   11
      Top             =   2760
      Width           =   6900
      VariousPropertyBits=   679493659
      MaxLength       =   120
      Size            =   "12171;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "英文名稱 :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   28
      Top             =   7460
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "英文名稱 :"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Top             =   5830
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "英文名稱 :"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   26
      Top             =   4140
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "英文名稱 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   2412
      Width           =   900
   End
   Begin MSForms.TextBox textCRL156 
      Height          =   336
      Index           =   4
      Left            =   1110
      TabIndex        =   4
      Top             =   7110
      Width           =   6900
      VariousPropertyBits=   679493659
      MaxLength       =   120
      Size            =   "12171;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCRL156 
      Height          =   336
      Index           =   3
      Left            =   1110
      TabIndex        =   3
      Top             =   5480
      Width           =   6900
      VariousPropertyBits=   679493659
      MaxLength       =   120
      Size            =   "12171;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCRL156 
      Height          =   336
      Index           =   2
      Left            =   1110
      TabIndex        =   2
      Top             =   3792
      Width           =   6900
      VariousPropertyBits=   679493659
      MaxLength       =   120
      Size            =   "12171;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCRL156 
      Height          =   336
      Index           =   1
      Left            =   1110
      TabIndex        =   1
      Top             =   2076
      Width           =   6900
      VariousPropertyBits=   679493659
      MaxLength       =   120
      Size            =   "12171;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "中文名稱 :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   24
      Top             =   7110
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "中文名稱 :"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   5480
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "中文名稱 :"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   3792
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "中文名稱 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   2076
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "日文名稱 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   1150
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "英文名稱 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   800
      Width           =   900
   End
   Begin VB.Label lbl1 
      Caption         =   "對造１ :"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   216
      Width           =   800
   End
   Begin MSForms.TextBox textCRL158 
      Height          =   336
      Index           =   0
      Left            =   1110
      TabIndex        =   10
      Top             =   1150
      Width           =   6900
      VariousPropertyBits=   679493659
      MaxLength       =   120
      Size            =   "12171;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCRL156 
      Height          =   336
      Index           =   0
      Left            =   1100
      TabIndex        =   0
      Top             =   450
      Width           =   6900
      VariousPropertyBits=   679493659
      MaxLength       =   120
      Size            =   "12171;593"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm090801_14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Amy 2024/01/10
Option Explicit

Public m_stCRL01 As String, m_stSysKind As String, m_stCRL159 As String '接洽單號/系統別/對造身份證號or統編
Public m_stCRL156 As String, m_stCRL157 As String, m_stCRL158 As String '對造 中、英、日 名稱
Public bolSendRiskChkMail As Boolean, m_RiskMsg As String '是否寄風險檢查對象通知信/風險檢查對象訊息

Dim m_PrevForm As Form '前一畫面
Dim i As Integer

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim IsRisk As Boolean, stSysKind As String, stAllRisk As String, stTP(1) As String, txt As Object
   Dim stCRL156 As String, stCRL157 As String, stCRL158 As String, stCRL159 As String
   
   PUB_FilterFormText Me
   If m_RiskMsg <> MsgText(601) Then
      stSysKind = m_RiskMsg
      m_RiskMsg = ""
   End If
   
   '確定
   If Index = 0 Then
      If FormCheck = False Then Exit Sub
      
      bolSendRiskChkMail = False '重新確認是否為風險檢查對象名單
      '中
      For Each txt In textCRL156
         If txt.Text <> MsgText(601) Then
            stCRL156 = stCRL156 & "☆" & txt.Text
         End If
      Next
      If stCRL156 <> MsgText(601) Then stCRL156 = Mid(stCRL156, 2)
      
      '英 (Memo 英文會用到,;故以☆區隔)
      For Each txt In textCRL157
         If txt.Text <> MsgText(601) Then
            stCRL157 = stCRL157 & "☆" & txt.Text
         End If
      Next
      If stCRL157 <> MsgText(601) Then stCRL157 = Mid(stCRL157, 2)
      
      '日
      For Each txt In textCRL158
         If txt.Text <> MsgText(601) Then
            stCRL158 = stCRL158 & "☆" & txt.Text
         End If
      Next
      If stCRL158 <> MsgText(601) Then stCRL158 = Mid(stCRL158, 2)
      
      stTP(0) = "": stTP(1) = ""
      If stCRL156 <> MsgText(601) Then stTP(0) = stTP(0) & "中-" & stCRL156
      If stCRL157 <> MsgText(601) Then
         If stTP(0) <> MsgText(601) Then stTP(0) = stTP(0) & "★"
         stTP(0) = stTP(0) & "英-" & stCRL157
      End If
      If stCRL159 <> MsgText(601) Then
          If stTP(0) <> MsgText(601) Then stTP(0) = stTP(0) & "★"
          stTP(0) = stTP(0) & "日-" & stCRL159
      End If
      
      IsRisk = ChkRiskData(2, Me.Name, stSysKind, , stTP(0), stTP(1))
      If IsRisk = True Then
         bolSendRiskChkMail = True
         m_RiskMsg = m_RiskMsg & "☆" & stTP(1)
      End If
      If m_RiskMsg <> MsgText(601) Then m_RiskMsg = Replace(Mid(m_RiskMsg, 5), "<br>", vbCrLf)
      
      '身份證字號/統一編號
      For Each txt In textCRL159
         If txt.Text <> MsgText(601) Then
            stCRL159 = stCRL159 & "☆" & txt.Text
         End If
      Next
      If stCRL159 <> MsgText(601) Then stCRL159 = Mid(stCRL159, 2)
      
      m_PrevForm.m_stCRL156 = stCRL156
      m_PrevForm.m_stCRL157 = stCRL157
      m_PrevForm.m_stCRL158 = stCRL158
      m_PrevForm.m_stCRL159 = stCRL159
      m_PrevForm.bolSendRiskChkMail = bolSendRiskChkMail
      m_PrevForm.m_RiskMsg = m_RiskMsg
   End If
   'Screen.MousePointer = m_MousePointer
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Move 0, 0
   FormClear
   cmdOK(0).Visible = False '確定
   cmdOK(1).Visible = False '回前畫面
   If UCase(TypeName(m_PrevForm)) <> UCase("Nothing") Then
      If TypeName(m_PrevForm) = "frm090801_New" Then
         If m_PrevForm.cmdOK(0).Caption = "新增" Or m_PrevForm.cmdOK(0).Caption = "存檔" Then
            cmdOK(0).Visible = True
            cmdOK(0).Left = 7080
         Else
            cmdOK(1).Visible = True
         End If
      Else
         cmdOK(1).Visible = True
      End If
      SetData
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090801_14 = Nothing
End Sub

'身份證字號/統一編號
Private Sub textCRL159_GotFocus(Index As Integer)
   InverseTextBox textCRL159(Index)
End Sub

Private Sub textCRL156_GotFocus(Index As Integer)
   InverseTextBox textCRL156(Index)
End Sub

Private Sub textCRL157_GotFocus(Index As Integer)
   InverseTextBox textCRL157(Index)
End Sub

Private Sub textCRL158_GotFocus(Index As Integer)
   InverseTextBox textCRL158(Index)
End Sub

'身份證字號/統一編號
Private Sub textCRL159_Validate(Index As Integer, Cancel As Boolean)
   Dim stMsg As Boolean, ii As Integer
   
   If textCRL159(Index) = MsgText(601) Then Exit Sub
   
   textCRL159(Index) = Trim(PUB_StringFilter(textCRL159(Index))) '複製貼上時多貼到空白格
   '國籍 台灣 or 身份證號 長度10個字
   If Pub_CheckIDAll(0, Me.Name, textCRL159(Index)) = False Then
      Cancel = True
      textCRL159(Index).SetFocus
   End If
End Sub

Private Sub FormClear()
   Dim txt As Object
   
   '中
   For Each txt In textCRL156
      txt.Text = ""
   Next
   '英
   For Each txt In textCRL157
      txt.Text = ""
   Next
   '日
   For Each txt In textCRL158
      txt.Text = ""
   Next
   '身份證字號/統一編號
   For Each txt In textCRL159
      txt.Text = ""
   Next
End Sub

Private Sub SetTextLock(ByVal bEnable As Boolean)
   Dim txt As Object
   
   '中
   For Each txt In textCRL156
      txt.Locked = bEnable
   Next
   '英
   For Each txt In textCRL157
      txt.Locked = bEnable
   Next
   '日
   For Each txt In textCRL158
      txt.Locked = bEnable
   Next
   '身份證字號/統一編號
   For Each txt In textCRL159
      txt.Locked = bEnable
   Next
   
End Sub

Private Function FormCheck() As Boolean
   Dim bolHasData As Boolean, bCancel As Boolean, stMsg(1) As String, txt As Object
   
   FormCheck = False
   '檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
      Exit Function
   End If
   
   '中
   For Each txt In textCRL156
      If Trim(txt.Text) <> MsgText(601) Then
         bolHasData = True
         '中文名稱有輸判斷公司(同客戶檔檢查),統編必輸
         If GetTextLength(textCRL156(txt.Index).Text) > 6 Then
            If textCRL159(txt.Index) = MsgText(601) Then
               stMsg(1) = stMsg(1) & ",對造" & txt.Index + 1
            End If
         End If
      End If
   Next
   If bolHasData = False Then
      '英
      For Each txt In textCRL157
         If Trim(txt.Text) <> MsgText(601) Then
            bolHasData = True
            Exit For
         End If
      Next
   End If
   If bolHasData = False Then
      '日
      For Each txt In textCRL158
         If Trim(txt.Text) <> MsgText(601) Then
            bolHasData = True
            Exit For
         End If
      Next
   End If
   
   '名稱 中/英/日 需擇一輸入
   If bolHasData = False Then
      stMsg(0) = "對造中/英/日 名稱需擇一輸入"
      MsgBox stMsg(0), , MsgText(5)
      textCRL156(0).SetFocus
      Exit Function
   End If
   
   '中文名稱有輸判斷大於6個字(公司),統編必輸
   If stMsg(1) <> MsgText(601) Then
      MsgBox Mid(stMsg(1), 2) & vbCrLf & _
                      "有中文對造名稱統一編號不可為空！", , MsgText(5)
      Exit Function
   Else
      '身份證/統編 檢查
      For Each txt In textCRL159
         If txt.Text <> MsgText(601) Then
            Call textCRL159_Validate(txt.Index, bCancel)
            If bCancel = True Then
               Exit Function
            End If
         End If
      Next
   End If
   
   FormCheck = True
End Function

Private Sub SetData()
   Dim arrTmp
   
   '中
   If m_stCRL156 <> MsgText(601) Then
      arrTmp = Split(m_stCRL156, "☆")
      For i = LBound(arrTmp) To UBound(arrTmp)
         textCRL156(i).Text = arrTmp(i)
      Next i
   End If
   '英
   If m_stCRL157 <> MsgText(601) Then
      arrTmp = Split(m_stCRL157, "☆")
      For i = LBound(arrTmp) To UBound(arrTmp)
         textCRL157(i).Text = arrTmp(i)
      Next
   End If
   '日
   If m_stCRL158 <> MsgText(601) Then
      arrTmp = Split(m_stCRL158, "☆")
      For i = LBound(arrTmp) To UBound(arrTmp)
         textCRL158(i).Text = arrTmp(i)
      Next
   End If
   '身份證/統編
   If m_stCRL159 <> MsgText(601) Then
      arrTmp = Split(m_stCRL159, "☆")
      For i = LBound(arrTmp) To UBound(arrTmp)
         textCRL159(i).Text = arrTmp(i)
      Next
   End If
End Sub

