VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_b_3 
   BackColor       =   &H80000004&
   BorderStyle     =   1  '單線固定
   Caption         =   "申請書"
   ClientHeight    =   4110
   ClientLeft      =   1440
   ClientTop       =   2310
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8085
   Begin VB.TextBox txtStartDate 
      Height          =   300
      Left            =   3990
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CheckBox checkCE09 
      Height          =   180
      Left            =   1830
      TabIndex        =   22
      Top             =   1800
      Width           =   252
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5400
      TabIndex        =   6
      Top             =   180
      Width           =   800
   End
   Begin VB.TextBox textCE04 
      Height          =   300
      Left            =   1380
      MaxLength       =   9
      TabIndex        =   0
      Top             =   2130
      Width           =   1212
   End
   Begin VB.TextBox textCE05 
      Height          =   300
      Left            =   1380
      MaxLength       =   9
      TabIndex        =   1
      Top             =   2445
      Width           =   1212
   End
   Begin VB.TextBox textCE06 
      Height          =   300
      Left            =   1380
      MaxLength       =   9
      TabIndex        =   2
      Top             =   2760
      Width           =   1212
   End
   Begin VB.TextBox textCE07 
      Height          =   300
      Left            =   1380
      MaxLength       =   9
      TabIndex        =   3
      Top             =   3075
      Width           =   1212
   End
   Begin VB.TextBox textCE08 
      Height          =   300
      Left            =   1380
      MaxLength       =   9
      TabIndex        =   4
      Top             =   3390
      Width           =   1212
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   6330
      TabIndex        =   7
      Top             =   180
      Width           =   800
   End
   Begin MSForms.TextBox textCE04_2 
      Height          =   300
      Left            =   2700
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2130
      Width           =   4815
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8493;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCE05_2 
      Height          =   300
      Left            =   2700
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2445
      Width           =   4815
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8493;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCE06_2 
      Height          =   300
      Left            =   2700
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2760
      Width           =   4815
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8493;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCE07_2 
      Height          =   300
      Left            =   2700
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3075
      Width           =   4815
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8493;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCE08_2 
      Height          =   300
      Left            =   2700
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3390
      Width           =   4815
      VariousPropertyBits=   679493663
      BackColor       =   14737632
      MaxLength       =   35
      Size            =   "8493;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "申請人5："
      Height          =   270
      Index           =   4
      Left            =   240
      TabIndex        =   30
      Top             =   3390
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "申請人4："
      Height          =   270
      Index           =   3
      Left            =   240
      TabIndex        =   29
      Top             =   3081
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "申請人3："
      Height          =   270
      Index           =   2
      Left            =   240
      TabIndex        =   28
      Top             =   2774
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "申請人2："
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   27
      Top             =   2467
      Width           =   1095
   End
   Begin VB.Label lblStart 
      Caption         =   "授權起始日期："
      Height          =   255
      Left            =   2700
      TabIndex        =   26
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label lblKind 
      Caption         =   "紙本送件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3900
      TabIndex        =   25
      Top             =   210
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "備註：欲按申請書時，建議不要同時使用Word軟體，因程式執行中會使用到Word。"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   24
      Top             =   3840
      Width           =   7890
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   3
      Left            =   1365
      TabIndex        =   23
      Top             =   315
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "123"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "是否要變更申請人："
      Height          =   270
      Left            =   180
      TabIndex        =   21
      Top             =   1800
      Width           =   2355
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   9
      Left            =   1365
      TabIndex        =   20
      Top             =   975
      Width           =   3045
      VariousPropertyBits=   27
      Caption         =   "123"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   7
      Left            =   1365
      TabIndex        =   19
      Top             =   660
      Width           =   1830
      VariousPropertyBits=   27
      Caption         =   "123"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   270
      Index           =   15
      Left            =   420
      TabIndex        =   18
      Top             =   1305
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   270
      Index           =   18
      Left            =   420
      TabIndex        =   17
      Top             =   975
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   270
      Index           =   19
      Left            =   420
      TabIndex        =   16
      Top             =   660
      Width           =   915
   End
   Begin MSForms.Label lbl1 
      Height          =   270
      Index           =   15
      Left            =   1365
      TabIndex        =   15
      Top             =   1305
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "123"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "申請人1："
      Height          =   270
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "總收文號："
      Height          =   270
      Index           =   0
      Left            =   420
      TabIndex        =   8
      Top             =   315
      Width           =   915
   End
End
Attribute VB_Name = "frm090201_b_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 lbl1(3)/lbl1(7)/lbl1(9)/lbl1(15)/textCE04_2/textCE05_2/textCE06_2/textCE07_2/textCE08_2
'Create By Sindy 2013/7/19
Option Explicit
Public m_CP10 As String
Public bolCP118 As Boolean 'Added by Lydia 2020/10/07 是否電子送件

Private Sub checkCE09_Click()
   If checkCE09.Value = 1 Then
      textCE04.Enabled = True
      textCE05.Enabled = True
      textCE06.Enabled = True
      textCE07.Enabled = True
      textCE08.Enabled = True
   Else
      textCE04.Enabled = False
      textCE05.Enabled = False
      textCE06.Enabled = False
      textCE07.Enabled = False
      textCE08.Enabled = False
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim rs As ADODB.Recordset
   
   If m_CP10 = "501" And textCE04 = "" Then
      MsgBox "請輸入受讓人 !!!"
      textCE04.SetFocus
      Exit Sub
   'Added by Lydia 2020/10/21
   ElseIf m_CP10 = "502" Then
      If textCE04 = "" Then
          MsgBox "請輸入被授權人 !!!"
          textCE04.SetFocus
          Exit Sub
      ElseIf Trim(txtStartDate) = "" Then
          MsgBox "請輸入授權起始日期 !!!"
          txtStartDate.SetFocus
          Exit Sub
      End If
   'end 2020/10/21
   ElseIf checkCE09.Visible = True And checkCE09.Value = 1 And textCE04 = "" Then
      textCE04.SetFocus
      Exit Sub
   End If
   
   '申請人必須依序輸入
   If textCE04 <> "" Or textCE05 <> "" Or textCE06 <> "" Or textCE07 <> "" Or textCE08 <> "" Then
      If (Trim(textCE05) <> "" And Trim(textCE04) = "") Or _
         (Trim(textCE06) <> "" And Trim(textCE05) = "") Or _
         (Trim(textCE07) <> "" And Trim(textCE06) = "") Or _
         (Trim(textCE08) <> "" And Trim(textCE07) = "") Then
         MsgBox "請依序輸入！", vbExclamation
         Exit Sub
      End If
      If (textCE05 <> "" And Trim(textCE05) = Trim(textCE04)) Or _
         (textCE06 <> "" And Trim(textCE06) = Trim(textCE05)) Or _
         (textCE07 <> "" And Trim(textCE07) = Trim(textCE06)) Or _
         (textCE08 <> "" And Trim(textCE08) = Trim(textCE07)) Then
         MsgBox "資料重覆！", vbExclamation
         Exit Sub
      End If
   End If
   
   If TxtValidate = False Then Exit Sub
   
   '更新變更事項檔
   If checkCE09.Visible = True And checkCE09.Value = 1 Then
      '檢查是否有此筆文號變更資料
      strExc(0) = "select ce01 from changeevent where ce01='" & Trim(Lbl1(3)) & "'"
      intI = 1
      Set rs = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "update changeevent" & _
                  " set ce04='" & textCE04 & "'" & _
                  " ,ce05='" & textCE05 & "'" & _
                  " ,ce06='" & textCE06 & "'" & _
                  " ,ce07='" & textCE07 & "'" & _
                  " ,ce08='" & textCE08 & "'" & _
                  " WHERE CE01='" & Trim(Lbl1(3)) & "'"
      Else
         strSql = "insert into changeevent(ce01,ce04,ce05,ce06,ce07,ce08) values(" & _
                  CNULL(Trim(Lbl1(3))) & "," & CNULL(textCE04) & "," & CNULL(textCE05) & "," & _
                  CNULL(textCE06) & "," & CNULL(textCE07) & "," & CNULL(textCE08) & ")"
      End If
      cnnConnection.Execute strSql
   End If
   
   '產生申請書
   'Added by Lydia 2020/10/07 電子送件申請書
   If bolCP118 = True Then
      Call GetApplBook_T(Lbl1(7), Lbl1(3).Caption, m_CP10)
   Else
   'end 2020/10/07
      'Modified by Lydia 2019/03/28 +傳入收文號 =>lbl1(3).Caption
      If PUB_GetApplBook(Lbl1(7), m_CP10, _
      IIf(textCE04.Enabled = True And Trim(textCE04) <> "", textCE04, ""), _
      IIf(textCE05.Enabled = True And Trim(textCE05) <> "", textCE05, ""), _
      IIf(textCE06.Enabled = True And Trim(textCE06) <> "", textCE06, ""), _
      IIf(textCE07.Enabled = True And Trim(textCE07) <> "", textCE07, ""), _
      IIf(textCE08.Enabled = True And Trim(textCE08) <> "", textCE08, ""), Lbl1(3).Caption) = True Then
         Call cmdExit_Click
      End If
   End If 'Added by Lydia 2020/10/07
   
   Set rs = Nothing
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'Modifed by Lydia 2022/04/19 +延展102
   If m_CP10 = "301" Or m_CP10 = "102" Then
      Label3.Caption = "是否要變更申請人："
      checkCE09.Visible = True
      textCE04.Enabled = False
      textCE05.Enabled = False
      textCE06.Enabled = False
      textCE07.Enabled = False
      textCE08.Enabled = False
      checkCE09.TabIndex = 0
   Else
      Label3.Caption = "請輸入受讓人："
      checkCE09.Visible = False
      textCE04.Enabled = True
      textCE05.Enabled = True
      textCE06.Enabled = True
      textCE07.Enabled = True
      textCE08.Enabled = True
      textCE04.TabIndex = 0
   End If
   'Added by Lydia 2020/10/21
   If m_CP10 = "502" Then
       Label3.Caption = "請輸入被授權人："
       lblStart.Visible = True
       txtStartDate.Visible = True
       For intI = 0 To 4
           Label2(intI).Caption = "被授權人" & intI + 1 & "："
       Next intI
   Else
       lblStart.Visible = False
       txtStartDate.Visible = False
       For intI = 0 To 4
           Label2(intI).Caption = "申請人" & intI + 1 & "："
       Next intI
   End If
   'end 2020/10/21
   textCE04.Text = ""
   textCE04_2.Text = ""
   textCE05.Text = ""
   textCE05_2.Text = ""
   textCE06.Text = ""
   textCE06_2.Text = ""
   textCE07.Text = ""
   textCE07_2.Text = ""
   textCE08.Text = ""
   textCE08_2.Text = ""
   'Added by Lydia 2020/10/07 顯示紙本送件/電子送件
   If bolCP118 = True Then
       lblKind.Caption = "電子送件"
   Else
       lblKind.Caption = "紙本送件"
   End If
   'end 2020/10/07
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090201_b_3 = Nothing
End Sub

Private Sub textCE04_GotFocus()
   InverseTextBox textCE04
End Sub

Private Sub textCE04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCE04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE04_2 = Empty
   If IsEmptyText(textCE04) = False Then
      '補滿9碼
      Me.textCE04.Text = Left(Me.textCE04.Text & "000000000", 9)
      '檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      oState = True
      textCE04_2 = GetCustomerNameAndState(textCE04, "0", oState)
      If oState = False Then
         Cancel = True
         Exit Sub
      End If
      If textCE04_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE04 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
'      '顯示申請人地址
'      textCE23.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "1")
'      textCE24.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "2")
'      textCE25.Text = PUB_GetCustEachAdd(Me.textCE04.Text, "3")
   End If
End Sub

Private Sub textCE05_GotFocus()
   InverseTextBox textCE05
End Sub

Private Sub textCE05_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCE05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE05_2 = Empty
   If IsEmptyText(textCE05) = False Then
      Me.textCE05.Text = Left(Me.textCE05.Text & "000000000", 9)
      Dim oState As Boolean
      oState = True
      textCE05_2 = GetCustomerNameAndState(textCE05, "0", oState)
      If oState = False Then
         Cancel = True
         Exit Sub
      End If
      If textCE05_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE05 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
'      '顯示申請人地址
'      textCE26.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "1")
'      textCE27.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "2")
'      textCE28.Text = PUB_GetCustEachAdd(Me.textCE05.Text, "3")
   End If
End Sub

Private Sub textCE06_GotFocus()
   InverseTextBox textCE06
End Sub

Private Sub textCE06_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCE06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE06_2 = Empty
   If IsEmptyText(textCE06) = False Then
      Me.textCE06.Text = Left(Me.textCE06.Text & "000000000", 9)
      Dim oState As Boolean
      oState = True
      textCE06_2 = GetCustomerNameAndState(textCE06, "0", oState)
      If oState = False Then
         Cancel = True
         Exit Sub
      End If
      If textCE06_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE06 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
'      '顯示申請人地址
'      textCE29.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "1")
'      textCE30.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "2")
'      textCE31.Text = PUB_GetCustEachAdd(Me.textCE06.Text, "3")
   End If
End Sub

Private Sub textCE07_GotFocus()
   InverseTextBox textCE07
End Sub

Private Sub textCE07_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCE07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE07_2 = Empty
   If IsEmptyText(textCE07) = False Then
      Me.textCE07.Text = Left(Me.textCE07.Text & "000000000", 9)
      Dim oState As Boolean
      oState = True
      textCE07_2 = GetCustomerNameAndState(textCE07, "0", oState)
      If oState = False Then
         Cancel = True
         Exit Sub
      End If
      If textCE07_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE07 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
'      '顯示申請人地址
'      textCE32.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "1")
'      textCE33.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "2")
'      textCE34.Text = PUB_GetCustEachAdd(Me.textCE07.Text, "3")
   End If
End Sub

Private Sub textCE08_GotFocus()
   InverseTextBox textCE08
End Sub

Private Sub textCE08_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCE08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCE08_2 = Empty
   If IsEmptyText(textCE08) = False Then
      Me.textCE08.Text = Left(Me.textCE08.Text & "000000000", 9)
      Dim oState As Boolean
      oState = True
      textCE08_2 = GetCustomerNameAndState(textCE08, "0", oState)
      If oState = False Then
         Cancel = True
         Exit Sub
      End If
      If textCE08_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textCE08 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
'      '顯示申請人地址
'      textCE35.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "1")
'      textCE36.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "2")
'      textCE37.Text = PUB_GetCustEachAdd(Me.textCE08.Text, "3")
   End If
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If Me.textCE04.Enabled = True Then
      Cancel = False
      textCE04_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCE05.Enabled = True Then
      Cancel = False
      textCE05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCE06.Enabled = True Then
      Cancel = False
      textCE06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCE07.Enabled = True Then
      Cancel = False
      textCE07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCE08.Enabled = True Then
      Cancel = False
      textCE08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

'Added by Lydia 2020/10/07 產生電子送件申請書
Private Sub GetApplBook_T(ByVal pCaseNo As String, ByVal pCP09 As String, ByVal pCP10 As String)
Dim bolChk As Boolean
Dim tm() As String '商標基本檔
Dim intWhere As Integer
Dim strFolder As String, strFileName As String
Dim ET01 As String, ET03 As String, ET03_1 As String
Dim strChkVal As String
Dim m_CaseNo As String
Dim strContent As String

    ReDim tm(TF_TM)
    Call ChgCaseNo(Replace(pCaseNo, "-", ""), tm)
    intWhere = 國內
    If ClsPDReadTrademarkDatabase(tm(), intWhere) Then
    End If
        
    Screen.MousePointer = vbHourglass
    
    m_CaseNo = PUB_FCPCaseNo2FileName(tm(1), tm(2), tm(3), tm(4))

    strLetterDate = strSrvDate(1)
    
    '2.申請書
    ET01 = "90"
    ET03 = ""
    strExc(0) = ""
    If pCP10 = "301" Then '變更
        If tm(15) = "" Then '註冊前變更
            ET03 = "01"     '申請書
            ET03_1 = "03"  '基本資料表(註冊前變更)
        Else                    '註冊變更
            ET03 = "02"     '申請書
            'Modified by Lydia 2020/11/23 基本資料表不同一般,增加【身分類別】
            'ET03_1 = "00"  '基本資料表(一般)
            ET03_1 = "04"
        End If
    ElseIf pCP10 = "501" Then '移轉
            ET03 = "01"     '申請書
            ET03_1 = "02"  '基本資料表(移轉)
    'Added by Lydia 2020/10/21
    ElseIf pCP10 = "502" Then '授權
            ET03 = "01"     '申請書
            ET03_1 = "02"  '基本資料表(授權)
    'end 2020/10/21
    'Added by Lydia 2022/04/19 延展可同時變更申請人
    ElseIf pCP10 = "102" Then '延展 (處理狀況的編號與FCT案一致)
             ET03 = "25"
             ET03_1 = "00"
    'end 2022/04/19
    End If
    If ET03 <> "" Then
         '申請書
         If StartLetter2(tm, m_CaseNo, ET01, ET03, pCP09, "2") = False Then Exit Sub
         NowPrint pCP09, ET01, ET03, False, strUserNum, , strContent, True, strContent
    End If
    
    '基本資料表
    If ET03_1 <> "" Then
        If StartLetter2(tm, m_CaseNo, ET01, ET03_1, pCP09, "1") = False Then Exit Sub
        NowPrint pCP09, ET01, ET03_1, False, strUserNum, , strContent, True, strContent
    End If
    strContent = Replace(strContent, vbCrLf & Chr(12), vbCrLf & "|#(分節)#|")    '換頁符號Chr(12)替換為分節符號 "|#(分節)#|"
    Call PUB_MakeDoc(strContent, strFileName, , , , , True)  '分節處理頁碼
    
    MsgBox "資料已產生完畢!!!"
    
ExitSub1:
    Screen.MousePointer = vbDefault
End Sub

'Added by Lydia 2020/10/07 電子送件-申請書
Private Function StartLetter2(ByRef iTM() As String, ByVal iCaseNo As String, ByVal iET01 As String, _
   ByVal iET03 As String, ByVal iCp09 As String, ByVal iKind As String) As Boolean
Dim strTxt(1 To 30) As String
Dim ii As Integer, jj As Integer
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim intA As Integer
Dim iCP10 As String
Dim iCP17 As String
Dim iCP14 As String, iCP14ext As String
Dim iCP110 As String, iCP08 As String
Dim rsAD As New ADODB.Recordset
Dim TempList As ListBox
Dim strTmp As String
   
   strSql = "select cp08,cp09,cp10,cp14,cp17,cp110,ed01 from caseprogress,ExtensionData where cp09='" & iCp09 & "' and cp14=ed02(+) "
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strSql)
   If intA = 1 Then
      iCP10 = "" & rsAD.Fields("cp10")
      iCP14 = "" & rsAD.Fields("cp14")
      iCP14ext = "" & rsAD.Fields("ed01")
      iCP17 = "" & rsAD.Fields("cp17")
      iCP110 = "" & rsAD.Fields("cp110")
      iCP08 = "" & rsAD.Fields("cp08")
   End If
   
   EndLetter iET01, iCp09, iET03, strUserNum
   
   ii = 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','本所案號','" & iCaseNo & "')"
   
   '申請人資料
   strExc(0) = ""
    If textCE04.Enabled = True And textCE04.Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textCE04.Text)
    If textCE05.Enabled = True And textCE05.Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textCE05.Text)
    If textCE06.Enabled = True And textCE06.Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textCE06.Text)
    If textCE07.Enabled = True And textCE07.Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textCE07.Text)
    If textCE08.Enabled = True And textCE08.Text <> "" Then strExc(0) = strExc(0) & "@" & ChangeCustomerL(textCE08.Text)
    If strExc(0) <> "" Then strExc(0) = Mid(strExc(0), 2)
   Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, iCP10, iTM(), False, strExc(0), , iTM(1))
      
   '出名代理人: 改成共用模組取得資料
   If iCP110 = "" Then
       Call PUB_SetOurAgent(TempList, iTM, iCP110, iCP10)
   End If
   strExc(0) = PUB_GetAgentCP110(iCp09, iCP110, "T", "4")
   If strExc(0) <> "" Then
       tmpArr1 = Empty
       tmpArr1 = Split(strExc(0), "|")
       'Added by Lydia 2020/10/21
       strExc(1) = "代理人"
       If iCP10 = "502" Then
           strExc(1) = "授權人之代理人"
           strExc(2) = "被授權人之代理人"
       End If
       'end 2020/10/21
       For jj = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(jj)) <> "" Then
               tmpArr2 = Empty
               tmpArr2 = Split(tmpArr1(jj), ",")
                ii = ii + 1
                'Modified by Lydia 2020/10/21 代理人=>改用變數strExc(1)
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(1) & jj + 1 & "-證書字號','" & tmpArr2(0) & "')"
                ii = ii + 1
                'Modified by Lydia 2020/10/21 代理人=>改用變數strExc(1)
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(1) & jj + 1 & "-ID','" & tmpArr2(1) & "')"
                ii = ii + 1
                'Modified by Lydia 2020/10/21 代理人=>改用變數strExc(1)
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(1) & jj + 1 & "-中文姓名','" & PUB_ConvertNameFormat("" & tmpArr2(2)) & "')"
                'Added by Lydia 2020/10/21 授權502，預設被授權人之代理人=授權人之代理人
                If iCP10 = "502" Then
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(2) & jj + 1 & "-證書字號','" & tmpArr2(0) & "')"
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(2) & jj + 1 & "-ID','" & tmpArr2(1) & "')"
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','" & strExc(2) & jj + 1 & "-中文姓名','" & PUB_ConvertNameFormat("" & tmpArr2(2)) & "')"
                End If
                'end 2020/10/21
           End If
       Next jj
   End If
   
   If iKind = "1" Then '基本資料表
        ii = ii + 1
        '內商承辦分機
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','內商承辦分機','" & iCP14ext & "')"
   End If
   
   If iKind = "2" Then '電子送件申請書
        'Added by Lydia 2020/10/21 授權502：商品服務名稱
        If iCP10 = "502" Then
            strExc(1) = "": strExc(2) = "": strExc(3) = ""
            strExc(0) = BeforePrintGetDBData("TMGoods:" & iTM(1) & "-" & iTM(2) & "-" & iTM(3) & "-" & iTM(4) & "-||區隔", True)
            If Trim(strExc(0)) <> "" Then
                tmpArr1 = Empty
                tmpArr1 = Split(strExc(0), "||")
                jj = 1
                For intA = 0 To UBound(tmpArr1)
                    strExc(1) = Trim(tmpArr1(intA))
                    If strExc(1) <> "" Then
                        strExc(2) = strExc(2) & _
                                         "【部分授權" & jj & "】  " & vbCrLf & _
                                         "　　【類別】　　　　　　　　　" & Mid(strExc(1), 1, InStr(strExc(1), "：") - 1) & vbCrLf & _
                                         "　　【商品服務名稱】　　　　　" & Mid(strExc(1), InStr(strExc(1), "：") + 1) & vbCrLf
                        jj = jj + 1
                    End If
                Next intA
            ElseIf iTM(9) <> "" Then
                 tmpArr1 = Empty
                 tmpArr1 = Split(iTM(9), ",")
                 jj = 1
                 For intA = 0 To UBound(tmpArr1)
                     strExc(1) = Trim(tmpArr1(intA))
                     If strExc(1) <> "" Then
                          strExc(2) = strExc(2) & _
                                           "【部分授權" & jj & "】  " & vbCrLf & _
                                           "　　【類別】　　　　　　　　　" & strExc(1) & vbCrLf & _
                                           "　　【商品服務名稱】　　　　　" & vbCrLf
                          jj = jj + 1
                     End If
                 Next intA
            Else
                     strExc(2) = "【部分授權1】  " & vbCrLf & _
                                      "　　【類別】　　　　　　　　　" & vbCrLf & _
                                      "　　【商品服務名稱】　　　　　" & vbCrLf
            End If
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','部分授權','" & ChgSQL(strExc(2)) & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','授權起始日期','" & ChangeTStringToTDateString(txtStartDate) & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','授權終止日期','" & ChangeTStringToTDateString(iTM(22)) & "')"
        End If
        'end 2020/10/21
        
        ii = ii + 1
        '繳費金額
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','繳費金額','" & Val(iCP17) & "')"
       '收據抬頭(內商才用)
        strExc(1) = ""
        strExc(1) = GetPrjPeople1(ChangeCustomerL(iTM(23)))
        For intI = 78 To 81 '申請人2~4
            If iTM(intI) <> "" Then
               strExc(1) = strExc(1) & "、" & GetPrjPeople1(ChangeCustomerL(iTM(intI)))
            End If
        Next intI
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','收據抬頭', " & CNULL(ChgSQL(strExc(1))) & ")"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-基本資料表', '" & iCaseNo & ".contact.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-委任書', '" & iCaseNo & ".poa.pdf')"
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

'Added by Lydia 2020/10/21
Private Sub txtStartDate_GotFocus()
   TextInverse txtStartDate
End Sub

Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtStartDate_Validate(Cancel As Boolean)
   Cancel = False
   If IsEmptyText(txtStartDate) = False Then
      If CheckIsTaiwanDate(txtStartDate, False) = False Then
          MsgBox "日期格式不正確!!!", vbCritical, "檢核資料"
          Cancel = True
      End If
   End If
   If Cancel = True Then
      txtStartDate.SetFocus
      txtStartDate_GotFocus
   End If
End Sub
