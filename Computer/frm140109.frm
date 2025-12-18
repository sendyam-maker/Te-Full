VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140109 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人日文資料維護作業"
   ClientHeight    =   3735
   ClientLeft      =   570
   ClientTop       =   975
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8955
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   350
      Index           =   2
      Left            =   3120
      TabIndex        =   1
      Top             =   330
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   5700
      TabIndex        =   8
      Top             =   120
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   7620
      TabIndex        =   10
      Top             =   120
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&C)"
      Height          =   400
      Index           =   3
      Left            =   6660
      TabIndex        =   9
      Top             =   120
      Width           =   912
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   330
      Width           =   1245
      VariousPropertyBits=   679495707
      MaxLength       =   9
      Size            =   "2196;529"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textFA78 
      Height          =   300
      Left            =   1800
      TabIndex        =   7
      Top             =   2670
      Width           =   7095
      VariousPropertyBits=   679495707
      MaxLength       =   60
      Size            =   "12515;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textFA54 
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   2328
      Width           =   7095
      VariousPropertyBits=   679495707
      MaxLength       =   60
      Size            =   "12515;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textFA09 
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   1986
      Width           =   7095
      VariousPropertyBits=   679495707
      MaxLength       =   60
      Size            =   "12515;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textFA58 
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   1644
      Width           =   7095
      VariousPropertyBits=   679495707
      MaxLength       =   60
      Size            =   "12515;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textFA23 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   1302
      Width           =   7095
      VariousPropertyBits=   679495707
      MaxLength       =   70
      Size            =   "12515;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textFA06 
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   7095
      VariousPropertyBits=   679495707
      MaxLength       =   80
      Size            =   "12515;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   1830
      TabIndex        =   18
      Top             =   720
      Width           =   7095
      Size            =   "12515;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label41 
      Caption         =   "聯絡人２(日)："
      Height          =   255
      Index           =   0
      Left            =   555
      TabIndex        =   17
      Top             =   2325
      Width           =   1275
   End
   Begin VB.Label Label38 
      Caption         =   "聯絡人１(日)："
      Height          =   255
      Left            =   555
      TabIndex        =   16
      Top             =   2010
      Width           =   1275
   End
   Begin VB.Label Label62 
      Caption         =   "聯絡人部門(日)："
      Height          =   255
      Left            =   375
      TabIndex        =   15
      Top             =   2670
      Width           =   1455
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "實體聯絡人日文名稱："
      Height          =   180
      Index           =   2
      Left            =   30
      TabIndex        =   14
      Top             =   1710
      Width           =   1800
   End
   Begin VB.Label Label13 
      Caption         =   "代理人地址(日)："
      Height          =   255
      Left            =   375
      TabIndex        =   13
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label Label30 
      Caption         =   "代理人名稱(日)："
      Height          =   255
      Left            =   375
      TabIndex        =   12
      Top             =   1020
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人編號："
      Height          =   180
      Index           =   0
      Left            =   690
      TabIndex        =   11
      Top             =   390
      Width           =   1125
   End
End
Attribute VB_Name = "frm140109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2018/11/05 改成Form2.0 (Label2和Textbox)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_FA04 As String, m_FA05 As String

Dim bolMsgRight As Boolean 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效
Dim SyxMsg As String 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效(記錄前一位置)

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bUpdate = IsUserHasRightOfFunction("frm140109", strEdit, False)
   
   MoveFormToCenter Me
   CmdLock 1
   
   'Added by Lydia 2018/11/20 模組-抓DB中的欄位實際長度
   textFA06.MaxLength = PUB_GetFieldDefSize("FAGENT", "FA06")
   textFA09.MaxLength = PUB_GetFieldDefSize("FAGENT", "FA09")
   textFA23.MaxLength = PUB_GetFieldDefSize("FAGENT", "FA23")
   textFA54.MaxLength = PUB_GetFieldDefSize("FAGENT", "FA54")
   textFA58.MaxLength = PUB_GetFieldDefSize("FAGENT", "FA58")
   textFA78.MaxLength = PUB_GetFieldDefSize("FAGENT", "FA78")
End Sub

Private Sub CmdLock(TF As Integer)
   Select Case TF
      Case 0
         Command1(2).Enabled = False
         Command1(0).Enabled = True
         Command1(3).Enabled = True
         Text1(0).Locked = True
      Case 1
         Command1(2).Enabled = True
         Command1(0).Enabled = False
         Command1(3).Enabled = False
         Text1(0).Locked = False
   End Select
End Sub

Private Sub ClearAll(bClearPk As Boolean)
   If bClearPk = True Then
      Text1(0) = Empty
   End If
   textFA06 = Empty
   textFA09 = Empty
   textFA23 = Empty
   textFA54 = Empty
   textFA58 = Empty
   textFA78 = Empty
   m_FA04 = Empty
   m_FA05 = Empty
   Label2.Caption = Empty
End Sub

Private Sub Command1_Click(Index As Integer)
On Error GoTo ErrHand
   Select Case Index
      Case 0 '確定
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            On Error GoTo ErrorHandler
            cnnConnection.BeginTrans
            strExc(1) = "UPDATE Fagent " & _
                                       "Set FA06='" & ChgSQL(textFA06) & "'," & _
                                             "FA09='" & textFA09 & "'," & _
                                             "FA23='" & ChgSQL(textFA23) & "'," & _
                                             "FA54='" & ChgSQL(textFA54) & "'," & _
                                             "FA58='" & ChgSQL(textFA58) & "'," & _
                                             "FA78='" & textFA78 & "' " & _
                              "WHERE FA01='" & Left(Text1(0), 8) & "' and FA02='" & Mid(Text1(0), 9, 1) & "' "
            Pub_SeekTbLog strExc(1)
            cnnConnection.Execute strExc(1)
            cnnConnection.CommitTrans
            'Modified by Lydia 2018/02/21 vbCritical=> vbInformation
            MsgBox "存檔完成 !", vbInformation
            CmdLock 1
            Call ClearAll(True)
            Text1(0).SetFocus
         Else
            Exit Sub
         End If
      Case 1 '結束
         Unload frm140109
         Set frm140109 = Nothing
      Case 2 '尋找
         If Text1(0) = "" Then
            MsgBox "請輸入代理人編號 !", vbCritical
            Exit Sub
         End If
         Text1(0) = Left(Trim(Text1(0)) & "000000000", 9)
         Call ClearAll(False)
         intI = 1
         strExc(0) = "SELECT * FROM Fagent WHERE FA01='" & Left(Text1(0), 8) & "' and FA02='" & Mid(Text1(0), 9, 1) & "' "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.RecordCount > 0 Then
               Label2.Caption = "" & RsTemp.Fields("FA04") & "" & RsTemp.Fields("FA05") & "" & RsTemp.Fields("FA63") & "" & RsTemp.Fields("FA64") & "" & RsTemp.Fields("FA65")
               If Not IsNull(RsTemp.Fields("FA06")) Then textFA06 = RsTemp.Fields("FA06")
               If Not IsNull(RsTemp.Fields("FA09")) Then textFA09 = RsTemp.Fields("FA09")
               If Not IsNull(RsTemp.Fields("FA23")) Then textFA23 = RsTemp.Fields("FA23")
               If Not IsNull(RsTemp.Fields("FA54")) Then textFA54 = RsTemp.Fields("FA54")
               If Not IsNull(RsTemp.Fields("FA58")) Then textFA58 = RsTemp.Fields("FA58")
               If Not IsNull(RsTemp.Fields("FA78")) Then textFA78 = RsTemp.Fields("FA78")
               If Not IsNull(RsTemp.Fields("FA04")) Then m_FA04 = RsTemp.Fields("FA04")
               If Not IsNull(RsTemp.Fields("FA05")) Then m_FA05 = RsTemp.Fields("FA05")
            End If
            CmdLock 0
         Else
            MsgBox "代理人編號錯誤，請重新輸入 !", vbCritical
            Text1(0).SetFocus
         End If
      Case 3 '取消
         If MsgBox("你並未存檔，確定離開嗎 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
         CmdLock 1
         Call ClearAll(True)
         Text1(0).SetFocus
   End Select
   Exit Sub
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbInformation
    Exit Sub
ErrorHandler:
    cnnConnection.RollbackTrans
    MsgBox "更新資料失敗，請洽系統管理員 !", vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140109 = Nothing
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim strTmp As String
Dim nResponse
   CheckDataValid = False
   
   ' 中文名稱, 英文名稱, 日文名稱不可全為空白
   If IsEmptyText(m_FA04) = True And IsEmptyText(m_FA05) = True And IsEmptyText(textFA06) = True Then
      strTit = "檢核資料"
      strMsg = "日文名稱不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textFA06.SetFocus
      GoTo EXITSUB
   End If

    'Added by Lydia 2021/04/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True) = False Then
        Exit Function
    End If
    'end 2021/04/14
    
   CheckDataValid = True
EXITSUB:
End Function

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If Me.textFA06.Enabled = True Then
      Cancel = False
      textFA06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textFA09.Enabled = True Then
      Cancel = False
      textFA09_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textFA23.Enabled = True Then
      Cancel = False
      textFA23_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textFA54.Enabled = True Then
      Cancel = False
      textFA54_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textFA58.Enabled = True Then
      Cancel = False
      textFA58_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textFA78.Enabled = True Then
      Cancel = False
      textFA78_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

Private Sub Text1_GotFocus(Index As Integer)
   InverseTextBox Text1(Index)
End Sub

'Modified by Lydia 2018/11/05 改成Form2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    KeyAscii = UpperCase(KeyAscii)
End Sub

' 代理人編號
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTmp As String, i As Integer
Dim strAgent As String
Dim strAgentName As String
   Cancel = False
   Select Case Index
      Case 0
         If Not IsEmptyText(Text1(Index)) Then
            If Mid(Text1(Index), 1, 1) <> "Y" Then
               Cancel = True
               MsgBox "代理人編號開頭必須為Y! ", vbCritical + vbOKOnly, "檢核資料"
               Call Text1_GotFocus(Index)
               Exit Sub
            End If
            strAgent = Text1(Index) & String(9 - Len(Text1(Index)), "0")
            strAgentName = GetFAgentName(strAgent)
            If IsEmptyText(strAgentName) Then
               Cancel = True
               MsgBox "無此代理人編號! ", vbCritical + vbOKOnly, "檢核資料"
               Call Text1_GotFocus(Index)
               Exit Sub
            End If
         End If
   End Select
End Sub

Private Sub textFA06_GotFocus()
   InverseTextBox textFA06
   OpenIme
End Sub

Private Sub textFA09_GotFocus()
   InverseTextBox textFA09
   OpenIme
End Sub

Private Sub textFA23_GotFocus()
   InverseTextBox textFA23
   OpenIme
End Sub

Private Sub textFA54_GotFocus()
   InverseTextBox textFA54
   OpenIme
End Sub

Private Sub textFA58_GotFocus()
   InverseTextBox textFA58
   OpenIme
End Sub

Private Sub textFA78_GotFocus()
   InverseTextBox textFA78
   OpenIme
End Sub

'日文地址要轉全形
'Modified by Lydia 2018/11/05 改成Form2.0
'Private Sub textFA23_KeyPress(KeyAscii As Integer)
Private Sub textFA23_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = ChangeZIP(KeyAscii)
End Sub

' 代理人名稱(日)
Private Sub textFA06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA06) = False Then
      If StrLength(textFA06) > 80 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人名稱(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA06_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

' 聯絡人1(日)
Private Sub textFA09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA09) = False Then
      If StrLength(textFA09) > textFA09.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "聯絡人1(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA09_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

' 代理人地址(日)
Private Sub textFA23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA23) = False Then
      If StrLength(textFA23) > 70 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "代理人地址(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA23_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

' 聯絡人2(日)
Private Sub textFA54_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA54) = False Then
      If StrLength(textFA54) > textFA54.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "聯絡人2(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA54_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

' 實體聯絡人日文名稱
Private Sub textFA58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA58) = False Then
      If StrLength(textFA58) > 20 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "實體聯絡人日文名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA58_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

' 聯絡人部門(日)
Private Sub textFA78_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textFA78) = False Then
      If StrLength(textFA78) > textFA78.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "聯絡人部門(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFA78_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'Added by Lydia 2018/11/21
Private Sub textFA06_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textFA06" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textFA06"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textFA09_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textFA09" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textFA09"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textFA23_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textFA23" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textFA23"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textFA54_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textFA54" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textFA54"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textFA58_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textFA58" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textFA58"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textFA78_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textFA78" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textFA78"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub
