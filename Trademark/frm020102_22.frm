VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_22 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請人及公司負責人輸入"
   ClientHeight    =   5748
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6996
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   6996
   Begin VB.CommandButton cmdok 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   315
      Left            =   5880
      TabIndex        =   9
      Top             =   120
      Width           =   1035
   End
   Begin VB.TextBox txtCU10 
      Height          =   264
      Left            =   30
      TabIndex        =   16
      Top             =   5340
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCU40 
      Height          =   264
      Left            =   1410
      MaxLength       =   60
      TabIndex        =   7
      Top             =   4800
      Width           =   5535
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3540
      Top             =   45
   End
   Begin VB.TextBox txtCU112 
      Height          =   264
      Left            =   1665
      MaxLength       =   5
      TabIndex        =   5
      Top             =   4155
      Width           =   1335
   End
   Begin VB.TextBox txtCU90 
      Height          =   615
      Left            =   30
      MaxLength       =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   3
      Top             =   2550
      Width           =   6915
   End
   Begin VB.TextBox txtCU89 
      Height          =   615
      Left            =   30
      MaxLength       =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   2
      Top             =   1890
      Width           =   6915
   End
   Begin VB.TextBox txtCU88 
      Height          =   615
      Left            =   30
      MaxLength       =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   1
      Top             =   1230
      Width           =   6915
   End
   Begin VB.TextBox txtCU05 
      Height          =   615
      Left            =   30
      MaxLength       =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   0
      Top             =   570
      Width           =   6915
   End
   Begin VB.TextBox txtCU103 
      Height          =   615
      Left            =   30
      MaxLength       =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   4
      Top             =   3480
      Width           =   6915
   End
   Begin MSForms.TextBox txtCU39 
      Height          =   300
      Left            =   1410
      TabIndex        =   6
      Top             =   4500
      Width           =   5535
      VariousPropertyBits=   679493659
      MaxLength       =   40
      Size            =   "9763;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCU41 
      Height          =   300
      Left            =   1410
      TabIndex        =   8
      Top             =   5055
      Width           =   5535
      VariousPropertyBits=   679493659
      MaxLength       =   40
      Size            =   "9763;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   195
      Left            =   810
      TabIndex        =   18
      Top             =   60
      Width           =   4995
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "8811;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "申請人："
      Height          =   195
      Left            =   30
      TabIndex        =   17
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "代表人1（中）："
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   15
      Top             =   4530
      Width           =   1350
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "代表人1（英）："
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   14
      Top             =   4830
      Width           =   1350
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "代表人1（日）："
      Height          =   180
      Index           =   2
      Left            =   30
      TabIndex        =   13
      Top             =   5085
      Width           =   1350
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "中文地址郵遞區號："
      Height          =   180
      Index           =   31
      Left            =   30
      TabIndex        =   12
      Top             =   4200
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請人英文名稱"
      Height          =   180
      Left            =   30
      TabIndex        =   11
      Top             =   330
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "公司負責人英文名稱"
      Height          =   180
      Left            =   30
      TabIndex        =   10
      Top             =   3240
      Width           =   1620
   End
End
Attribute VB_Name = "frm020102_22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo by Amy 2021/12/28 Form2.0已修改 label4/txtCU39/txtCU41
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit
'Modified by Lydia 2024/07/03 改變數
'Public oNextForm As Form
Dim oNextForm As Form
Dim mAppNo As String  '申請人編號
Dim m_CU112 As String, m_CU103 As String, m_CU05 As String, m_CU88 As String, m_CU89 As String
Dim m_CU90 As String, m_CU39 As String, m_CU40 As String, m_CU41 As String, m_CU10 As String
'end 2024/07/03

'Added by Lydia 2024/07/03 SetParent
Public Sub SetParent(ByVal pFrm As Form, ByVal pCustNo As String)
   Set oNextForm = pFrm
   mAppNo = pCustNo
End Sub

Private Sub cmdOK_Click()
'Add by Amy 2021/12/28檢查畫面的 TextBox是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Sub
End If

'Added by Lydia 2022/11/23 申請人英文名稱1檢查第1碼非英文,彈提醒
Dim strTmp As String
If txtCU05 <> "" Then '可空白
   If Asc(Left(UCase(txtCU05), 1)) < 65 Or Asc(Left(UCase(txtCU05), 1)) > 90 Then strTmp = "Y"
End If
If strTmp <> "" Then
   If MsgBox("申請人英文名稱輸入非英文字，是否修改畫面上的資料？", vbInformation + vbYesNo + vbDefaultButton1, "申請人英文名稱檢查") = vbYes Then
      txtCU05.SetFocus
      txtCU05_GotFocus
      Exit Sub
   End If
End If
'end 2022/11/23


oNextForm.m_CU103 = txtCU103.Text
oNextForm.m_CU05 = txtCU05.Text
oNextForm.m_CU88 = txtCU88.Text
oNextForm.m_CU89 = txtCU89.Text
oNextForm.m_CU90 = txtCU90.Text
'add by nickc 2006/01/20
oNextForm.m_CU112 = txtCU112.Text
'Add By Sindy 2012/2/7
oNextForm.m_CU39 = txtCU39.Text
oNextForm.m_CU40 = txtCU40.Text
oNextForm.m_CU41 = txtCU41.Text
'2012/2/7 End
'Add By Sindy 2012/10/31
txtCU10 = pub_NationByName(Trim(txtCU05 & txtCU88 & txtCU89 & txtCU90), Trim(txtCU10))
oNextForm.m_CU10 = txtCU10.Text
'2012/10/31 End
Unload Me
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Modifie by Lydia 2024/07/03 因為再次發生同一天同客戶不同案件之發文並且修改客戶名稱後留下重複的dml_log，調整如下：
   '1.讀取資料的模組改成回傳變數，而非直接變更發文表單的Public變數
   '2.輸入畫面(frm020102_22)改成開啟表單重新讀取資料，不是用發文畫面的共用變數。
                      'ex.113/7/1 X1714900發文T-191291,T-191301;另外一個疑問，當時的113/7/1 X1714900的dml_log雖然是記錄英文名稱改為" SHIN CHIAO SHOES CO., LTD."
                      '但是Customer還是未修改前的" SHIN CHIAO SHOES  CO., LTD."(CO. 前有2個空白)，還是測不出來。
   'add by nickc 2006/01/20
   'txtCU112.Text = oNextForm.m_CU112
   'txtCU103.Text = oNextForm.m_CU103
   'txtCU05.Text = oNextForm.m_CU05
   'txtCU88.Text = oNextForm.m_CU88
   'txtCU89.Text = oNextForm.m_CU89
   'txtCU90.Text = oNextForm.m_CU90
   ''add by nickc 2006/02/07
   ''Add By Sindy 2012/2/7
   'txtCU39.Text = oNextForm.m_CU39
   'txtCU40.Text = oNextForm.m_CU40
   'txtCU41.Text = oNextForm.m_CU41
   '2012/2/7 End
   'Add By Sindy 2012/10/31
   'txtCU10.Text = oNextForm.m_CU10
   '2012/10/31 End
'   GetData mAppNo
   Dim strTmp(1 To 10)
   Call Pub_GetDataFrm020102(mAppNo, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
   txtCU103 = m_CU103
   txtCU05 = m_CU05
   txtCU88 = m_CU88
   txtCU89 = m_CU89
   txtCU90 = m_CU90
   txtCU112 = m_CU112
   txtCU39 = m_CU39
   txtCU40 = m_CU40
   txtCU41 = m_CU41
   txtCU10 = m_CU10
   'end 2024/07/03
  
   Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm020102_22 = Nothing
End Sub


'add by nickc 2006/02/07
Private Sub Timer1_Timer()
'Modify By Sindy 2012/2/9
'If txtCU112.Visible And txtCU103.Visible And txtCU05.Visible Then
If txtCU112.Visible And txtCU103.Visible And txtCU05.Visible And txtCU39.Visible Then
   If txtCU112.Text = "" Then txtCU112.SetFocus
   If txtCU103.Text = "" Then txtCU103.SetFocus
   If txtCU05.Text = "" Then txtCU05.SetFocus
   If txtCU39.Text = "" Then txtCU39.SetFocus 'Add By Sindy 2012/2/9
   Screen.MousePointer = vbDefault
   Timer1.Interval = 0
End If
End Sub

Private Sub txtCU103_GotFocus()
InverseTextBox txtCU103
End Sub

Private Sub txtCU103_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(txtCU103, 30) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "英文名稱內容太長"
      txtCU103_GotFocus
   End If
End Sub

Private Sub txtCU05_GotFocus()
InverseTextBox txtCU05
End Sub

Private Sub txtCU05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(txtCU05, 30) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "英文名稱內容太長"
      txtCU05_GotFocus
   End If
End Sub

Private Sub txtCU112_GotFocus()
InverseTextBox txtCU112
End Sub

Private Sub txtCU112_KeyPress(KeyAscii As Integer)
KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub txtCU88_GotFocus()
InverseTextBox txtCU88
End Sub

Private Sub txtCU88_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(txtCU88, 30) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "英文名稱內容太長"
      txtCU88_GotFocus
   End If
End Sub

Private Sub txtCU89_GotFocus()
InverseTextBox txtCU89
End Sub

Private Sub txtCU89_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(txtCU89, 30) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "英文名稱內容太長"
      txtCU89_GotFocus
   End If
End Sub

Private Sub txtCU90_GotFocus()
InverseTextBox txtCU90
End Sub

Private Sub txtCU90_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(txtCU90, 30) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "英文名稱內容太長"
      txtCU90_GotFocus
   End If
End Sub

'Add By Sindy 2012/2/7
Private Sub txtCU39_GotFocus()
   OpenIme
   TextInverse txtCU39
End Sub

'Add By Sindy 2012/2/7
Private Sub txtCU39_Validate(Cancel As Boolean)
   If txtCU39.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCU39, txtCU39.MaxLength) Then
      Cancel = True
      txtCU39_GotFocus
   End If
End Sub

'Add By Sindy 2012/2/7
Private Sub txtCU40_GotFocus()
   CloseIme
   TextInverse txtCU40
End Sub

'Add By Sindy 2012/2/7
Private Sub txtCU41_GotFocus()
   OpenIme
   TextInverse txtCU41
End Sub

'Add By Sindy 2012/2/7
Private Sub txtCU41_Validate(Cancel As Boolean)
   If txtCU41.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCU41, txtCU41.MaxLength) Then
      Cancel = True
      txtCU41_GotFocus
   End If
End Sub
