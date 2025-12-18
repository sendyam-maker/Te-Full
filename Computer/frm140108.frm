VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140108 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶日文資料維護作業"
   ClientHeight    =   5745
   ClientLeft      =   570
   ClientTop       =   975
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   350
      Index           =   2
      Left            =   2970
      TabIndex        =   1
      Top             =   330
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   5700
      TabIndex        =   14
      Top             =   120
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   7620
      TabIndex        =   16
      Top             =   120
      Width           =   912
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&C)"
      Height          =   400
      Index           =   3
      Left            =   6660
      TabIndex        =   15
      Top             =   120
      Width           =   912
   End
   Begin MSForms.TextBox textCU56 
      Height          =   300
      Left            =   1680
      TabIndex        =   13
      Top             =   4824
      Width           =   5295
      VariousPropertyBits=   679495707
      Size            =   "9340;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU53 
      Height          =   300
      Left            =   1680
      TabIndex        =   12
      Top             =   4492
      Width           =   5295
      VariousPropertyBits=   679495707
      Size            =   "9340;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU50 
      Height          =   300
      Left            =   1680
      TabIndex        =   11
      Top             =   4164
      Width           =   5295
      VariousPropertyBits=   679495707
      Size            =   "9340;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU47 
      Height          =   300
      Left            =   1680
      TabIndex        =   10
      Top             =   3836
      Width           =   5295
      VariousPropertyBits=   679495707
      Size            =   "9340;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU44 
      Height          =   300
      Left            =   1680
      TabIndex        =   9
      Top             =   3508
      Width           =   5295
      VariousPropertyBits=   679495707
      Size            =   "9340;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU41 
      Height          =   300
      Left            =   1680
      TabIndex        =   8
      Top             =   3180
      Width           =   5295
      VariousPropertyBits=   679495707
      Size            =   "9340;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU93 
      Height          =   300
      Left            =   1680
      TabIndex        =   7
      Top             =   2835
      Width           =   6645
      VariousPropertyBits=   679495707
      Size            =   "11721;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU114 
      Height          =   300
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   6645
      VariousPropertyBits=   679495707
      Size            =   "11721;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU63 
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      Top             =   2190
      Width           =   6645
      VariousPropertyBits=   679495707
      Size            =   "11721;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU60 
      Height          =   300
      Left            =   1680
      TabIndex        =   4
      Top             =   1860
      Width           =   6645
      VariousPropertyBits=   679495707
      Size            =   "11721;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU29 
      Height          =   516
      Left            =   1680
      TabIndex        =   3
      Top             =   1290
      Width           =   7275
      VariousPropertyBits=   -1466939365
      Size            =   "12832;910"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCU06 
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   6972
      VariousPropertyBits=   679495707
      Size            =   "12298;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   330
      Width           =   1245
      VariousPropertyBits=   679495707
      MaxLength       =   9
      Size            =   "2196;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   1680
      TabIndex        =   30
      Top             =   720
      Width           =   7095
      Size            =   "12515;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "代表人1（日）："
      Height          =   180
      Index           =   2
      Left            =   300
      TabIndex        =   29
      Top             =   3210
      Width           =   1350
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "代表人2（日）："
      Height          =   180
      Index           =   5
      Left            =   300
      TabIndex        =   28
      Top             =   3525
      Width           =   1350
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "代表人3（日）："
      Height          =   180
      Index           =   8
      Left            =   300
      TabIndex        =   27
      Top             =   3870
      Width           =   1350
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "代表人4（日）："
      Height          =   180
      Index           =   11
      Left            =   300
      TabIndex        =   26
      Top             =   4200
      Width           =   1350
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "代表人5（日）："
      Height          =   180
      Index           =   14
      Left            =   300
      TabIndex        =   25
      Top             =   4515
      Width           =   1350
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "代表人6（日）："
      Height          =   180
      Index           =   17
      Left            =   300
      TabIndex        =   24
      Top             =   4860
      Width           =   1350
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "日文地址："
      Height          =   180
      Index           =   28
      Left            =   750
      TabIndex        =   23
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人1（日）："
      Height          =   180
      Index           =   2
      Left            =   300
      TabIndex        =   22
      Top             =   1890
      Width           =   1350
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人2（日）："
      Height          =   180
      Index           =   5
      Left            =   300
      TabIndex        =   21
      Top             =   2220
      Width           =   1350
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "實體聯絡人（日）："
      Height          =   180
      Index           =   8
      Left            =   30
      TabIndex        =   20
      Top             =   2880
      Width           =   1620
   End
   Begin VB.Label Label63 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人部門（日）："
      Height          =   180
      Index           =   16
      Left            =   30
      TabIndex        =   19
      Top             =   2550
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶名稱（日）："
      Height          =   180
      Index           =   4
      Left            =   210
      TabIndex        =   18
      Top             =   990
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   0
      Left            =   750
      TabIndex        =   17
      Top             =   390
      Width           =   900
   End
End
Attribute VB_Name = "frm140108"
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
Dim m_CU04 As String, m_CU05 As String, m_CU88 As String, m_CU89 As String
Dim m_CU90 As String, m_CU23 As String, m_CU24 As String, m_CU25 As String
Dim m_CU26 As String, m_CU27 As String, m_CU28 As String, m_CU102 As String

Dim bolMsgRight As Boolean 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效
Dim SyxMsg As String 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效(記錄前一位置)

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bUpdate = IsUserHasRightOfFunction("frm140108", strEdit, False)
   
   MoveFormToCenter Me
   CmdLock 1
   
   'Added by Lydia 2018/11/20 模組-抓DB中的欄位實際長度
   textCU06.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU06")
   textCU29.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU29")
   textCU41.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU41")
   textCU44.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU44")
   textCU47.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU47")
   textCU50.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU50")
   textCU53.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU53")
   textCU56.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU56")
   textCU60.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU60")
   textCU63.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU63")
   textCU93.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU93")
   textCU114.MaxLength = PUB_GetFieldDefSize("CUSTOMER", "CU114")
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
   textCU06 = Empty
   textCU29 = Empty
   textCU41 = Empty
   textCU44 = Empty
   textCU47 = Empty
   textCU50 = Empty
   textCU53 = Empty
   textCU56 = Empty
   textCU60 = Empty
   textCU63 = Empty
   textCU93 = Empty
   textCU114 = Empty
   m_CU04 = Empty
   m_CU05 = Empty
   m_CU88 = Empty
   m_CU89 = Empty
   m_CU90 = Empty
   m_CU23 = Empty
   m_CU24 = Empty
   m_CU25 = Empty
   m_CU26 = Empty
   m_CU27 = Empty
   m_CU28 = Empty
   m_CU102 = Empty
   'Modified by Lydia 2018/11/05 改成Form2.0
   'Label1(1).Caption = Empty
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
            strExc(1) = "UPDATE Customer " & _
                                       "SET CU06='" & textCU06 & "'," & _
                                               "CU29='" & textCU29 & "'," & _
                                               "CU41='" & textCU41 & "'," & _
                                               "CU44='" & textCU44 & "'," & _
                                               "CU47='" & textCU47 & "'," & _
                                               "CU50='" & textCU50 & "'," & _
                                               "CU53='" & textCU53 & "'," & _
                                               "CU56='" & textCU56 & "'," & _
                                               "CU60='" & textCU60 & "'," & _
                                               "CU63='" & textCU63 & "'," & _
                                               "CU93='" & textCU93 & "'," & _
                                               "CU114='" & textCU114 & "' " & _
                              "WHERE CU01='" & Left(Text1(0), 8) & "' and CU02='" & Mid(Text1(0), 9, 1) & "' "
            Pub_SeekTbLog strExc(1)
            cnnConnection.Execute strExc(1)
            cnnConnection.CommitTrans
            'Add by Amy 2017/12/08 修改名稱發信給秀玲
            If textCU06.Tag <> CheckStr(textCU06) And textCU06.Tag <> MsgText(601) Then
                PUB_SendMail strUserNum, "83002", "", Text1(0), " 客戶名稱修改！", "日：" & textCU06.Tag & " --> " & textCU06 & vbCrLf
            End If
            textCU06.Tag = textCU06
            'end 2017/12/08
            'Modified by Lydia 2018/02/21 vbCritical=> vbInformation
            MsgBox "存檔完成 !", vbInformation
            CmdLock 1
            Call ClearAll(True)
            Text1(0).SetFocus
         Else
            Exit Sub
         End If
      Case 1 '結束
         Unload frm140108
         Set frm140108 = Nothing
      Case 2 '尋找
         If Text1(0) = "" Then
            MsgBox "請輸入客戶編號 !", vbCritical
            Exit Sub
         End If
         Text1(0) = Left(Trim(Text1(0)) & "000000000", 9)
         Call ClearAll(False)
         intI = 1
         strExc(0) = "SELECT * FROM Customer WHERE CU01='" & Left(Text1(0), 8) & "' and CU02='" & Mid(Text1(0), 9, 1) & "' "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.RecordCount > 0 Then
               'Modified by Lydia 2018/11/05 改成Form2.0
               'Label1(1).Caption = "" & RsTemp.Fields("CU04") & "" & RsTemp.Fields("CU05") & "" & RsTemp.Fields("CU88") & "" & RsTemp.Fields("CU89") & "" & RsTemp.Fields("CU90") & ""
               Label2.Caption = "" & RsTemp.Fields("CU04") & "" & RsTemp.Fields("CU05") & "" & RsTemp.Fields("CU88") & "" & RsTemp.Fields("CU89") & "" & RsTemp.Fields("CU90") & ""
               If Not IsNull(RsTemp.Fields("CU06")) Then textCU06 = RsTemp.Fields("CU06"): textCU06.Tag = textCU06 'Add by Amy 2017/12/05
               If Not IsNull(RsTemp.Fields("CU29")) Then textCU29 = RsTemp.Fields("CU29")
               If Not IsNull(RsTemp.Fields("CU41")) Then textCU41 = RsTemp.Fields("CU41")
               If Not IsNull(RsTemp.Fields("CU44")) Then textCU44 = RsTemp.Fields("CU44")
               If Not IsNull(RsTemp.Fields("CU47")) Then textCU47 = RsTemp.Fields("CU47")
               If Not IsNull(RsTemp.Fields("CU50")) Then textCU50 = RsTemp.Fields("CU50")
               If Not IsNull(RsTemp.Fields("CU53")) Then textCU53 = RsTemp.Fields("CU53")
               If Not IsNull(RsTemp.Fields("CU56")) Then textCU56 = RsTemp.Fields("CU56")
               If Not IsNull(RsTemp.Fields("CU60")) Then textCU60 = RsTemp.Fields("CU60")
               If Not IsNull(RsTemp.Fields("CU63")) Then textCU63 = RsTemp.Fields("CU63")
               If Not IsNull(RsTemp.Fields("CU93")) Then textCU93 = RsTemp.Fields("CU93")
               If Not IsNull(RsTemp.Fields("CU114")) Then textCU114 = RsTemp.Fields("CU114")
               If Not IsNull(RsTemp.Fields("CU04")) Then m_CU04 = RsTemp.Fields("CU04")
               If Not IsNull(RsTemp.Fields("CU05")) Then m_CU05 = RsTemp.Fields("CU05")
               If Not IsNull(RsTemp.Fields("CU88")) Then m_CU88 = RsTemp.Fields("CU88")
               If Not IsNull(RsTemp.Fields("CU89")) Then m_CU89 = RsTemp.Fields("CU89")
               If Not IsNull(RsTemp.Fields("CU90")) Then m_CU90 = RsTemp.Fields("CU90")
               If Not IsNull(RsTemp.Fields("CU23")) Then m_CU23 = RsTemp.Fields("CU23")
               If Not IsNull(RsTemp.Fields("CU24")) Then m_CU24 = RsTemp.Fields("CU24")
               If Not IsNull(RsTemp.Fields("CU25")) Then m_CU25 = RsTemp.Fields("CU25")
               If Not IsNull(RsTemp.Fields("CU26")) Then m_CU26 = RsTemp.Fields("CU26")
               If Not IsNull(RsTemp.Fields("CU27")) Then m_CU27 = RsTemp.Fields("CU27")
               If Not IsNull(RsTemp.Fields("CU28")) Then m_CU28 = RsTemp.Fields("CU28")
               If Not IsNull(RsTemp.Fields("CU102")) Then m_CU102 = RsTemp.Fields("CU102")
            End If
            CmdLock 0
         Else
            MsgBox "客戶編號錯誤，請重新輸入 !", vbCritical
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
   Set frm140108 = Nothing
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTmp  As String
   CheckDataValid = False
   If m_CU04 = "" And m_CU05 = "" And m_CU88 = "" And m_CU89 = "" _
      And m_CU90 = "" And textCU06 = "" Then
      ShowMsg "日文客戶名稱不可為空白 !"
      textCU06.SetFocus
      Exit Function
   End If
   '客戶名稱(日)
   If Not CheckLengthIsOK(textCU06, 80) Then
      textCU06.SetFocus
      textCU06_GotFocus
      Exit Function
   End If
   '聯絡人1(日)
   If Not CheckLengthIsOK(textCU60, 20) Then
      textCU60.SetFocus
      textCU60_GotFocus
      Exit Function
   End If
   '聯絡人2(日)
   If Not CheckLengthIsOK(textCU63, 20) Then
      textCU63.SetFocus
      textCU63_GotFocus
      Exit Function
   End If
   '實體聯絡人(日)
   If Not CheckLengthIsOK(textCU93, 20) Then
      textCU93.SetFocus
      textCU93_GotFocus
      Exit Function
   End If
   '日文地址
   If Not CheckLengthIsOK(textCU29, 70) Then
      textCU29.SetFocus
      textCU29_GotFocus
      Exit Function
   End If
   '代表人1(中), 代表人1(日), 代表人2(中), 代表人2(日), ...
   If Not CheckLengthIsOK(textCU41, 40) Then
      textCU41.SetFocus
      textCU41_GotFocus
      Exit Function
   End If
   If Not CheckLengthIsOK(textCU44, 40) Then
      textCU44.SetFocus
      textCU44_GotFocus
      Exit Function
   End If
   If Not CheckLengthIsOK(textCU47, 40) Then
      textCU47.SetFocus
      textCU47_GotFocus
      Exit Function
   End If
   If Not CheckLengthIsOK(textCU50, 40) Then
      textCU50.SetFocus
      textCU50_GotFocus
      Exit Function
   End If
   If Not CheckLengthIsOK(textCU53, 40) Then
      textCU53.SetFocus
      textCU53_GotFocus
      Exit Function
   End If
   If Not CheckLengthIsOK(textCU56, 40) Then
      textCU56.SetFocus
      textCU56_GotFocus
      Exit Function
   End If
   If m_CU23 = "" And m_CU24 = "" And m_CU25 = "" And m_CU26 = "" _
      And m_CU27 = "" And m_CU28 = "" And textCU29 = "" And m_CU102 = "" Then
      ShowMsg "日文地址不可為空白 !"
      textCU29.SetFocus
      Exit Function
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
   
   If Me.textCU06.Enabled = True Then
      Cancel = False
      textCU06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU29.Enabled = True Then
      Cancel = False
      textCU29_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU41.Enabled = True Then
      Cancel = False
      textCU41_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU44.Enabled = True Then
      Cancel = False
      textCU44_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU47.Enabled = True Then
      Cancel = False
      textCU47_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU50.Enabled = True Then
      Cancel = False
      textCU50_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU53.Enabled = True Then
      Cancel = False
      textCU53_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU56.Enabled = True Then
      Cancel = False
      textCU56_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU60.Enabled = True Then
      Cancel = False
      textCU60_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU63.Enabled = True Then
      Cancel = False
      textCU63_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU93.Enabled = True Then
      Cancel = False
      textCU93_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCU114.Enabled = True Then
      Cancel = False
      textCU114_Validate Cancel
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

' 客戶編號
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTemp As String
   Cancel = False
   Select Case Index
      Case 0
         If IsEmptyText(Text1(Index)) = False Then
            If Mid(Text1(Index), 1, 1) <> "X" Then
               Cancel = True
               MsgBox "客戶編號開頭必須為X! ", vbCritical + vbOKOnly, "檢核資料"
               Call Text1_GotFocus(Index)
               Exit Sub
            End If
            strTemp = GetCustomerName(Text1(Index), 0)
            If IsEmptyText(strTemp) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "客戶編號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               Call Text1_GotFocus(Index)
               Exit Sub
            End If
         End If
   End Select
End Sub

Private Sub textCU06_GotFocus()
OpenIme
TextInverse textCU06
End Sub

Private Sub textCU29_GotFocus()
OpenIme
'TextInverse textCU29 'Remove by Lydia 2018/11/21 取消反白(Enter換行)
End Sub

Private Sub textCU41_GotFocus()
OpenIme
TextInverse textCU41
End Sub

Private Sub textCU44_GotFocus()
OpenIme
TextInverse textCU44
End Sub

Private Sub textCU47_GotFocus()
OpenIme
TextInverse textCU47
End Sub

Private Sub textCU50_GotFocus()
OpenIme
TextInverse textCU50
End Sub

Private Sub textCU53_GotFocus()
OpenIme
TextInverse textCU53
End Sub

Private Sub textCU56_GotFocus()
OpenIme
TextInverse textCU56
End Sub

Private Sub textCU60_GotFocus()
OpenIme
TextInverse textCU60
End Sub

Private Sub textCU63_GotFocus()
OpenIme
TextInverse textCU63
End Sub

Private Sub textCU93_GotFocus()
OpenIme
TextInverse textCU93
End Sub

Private Sub textCU114_GotFocus()
CloseIme
TextInverse textCU114
End Sub

'Modified by Lydia 2018/11/05 改成Form2.0
'Private Sub textCU29_KeyPress(KeyAscii As Integer)
Private Sub textCU29_KeyPress(KeyAscii As MSForms.ReturnInteger)
KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub textCU06_Validate(Cancel As Boolean)
   If textCU06.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU06, textCU06.MaxLength - 1) Then
      Cancel = True
   End If
End Sub

Private Sub textCU29_Validate(Cancel As Boolean)
   If textCU29.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU29, textCU29.MaxLength - 1) Then
      Cancel = True
   End If
End Sub

Private Sub textCU41_Validate(Cancel As Boolean)
   If textCU41.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU41, textCU41.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU44_Validate(Cancel As Boolean)
   If textCU44.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU44, textCU44.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU47_Validate(Cancel As Boolean)
   If textCU47.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU47, textCU47.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU50_Validate(Cancel As Boolean)
   If textCU50.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU50, textCU50.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU53_Validate(Cancel As Boolean)
   If textCU53.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU53, textCU53.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU56_Validate(Cancel As Boolean)
   If textCU56.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textCU56, textCU56.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textCU60_Validate(Cancel As Boolean)
   If textCU60.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU60, textCU60.MaxLength - 1) Then
      Cancel = True
   End If
End Sub

Private Sub textCU63_Validate(Cancel As Boolean)
   If textCU63.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU63, textCU63.MaxLength - 1) Then
      Cancel = True
   End If
End Sub

Private Sub textCU93_Validate(Cancel As Boolean)
   If textCU93.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU93, textCU93.MaxLength - 1) Then
      Cancel = True
   End If
End Sub

Private Sub textCU114_Validate(Cancel As Boolean)
   If textCU114.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textCU114, textCU114.MaxLength - 1) Then
      Cancel = True
   End If
End Sub

'Added by Lydia 2018/11/21 有換行輸入
Private Sub textCU29_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
     Call PUB_HandleForm2TextBox(Me.textCU29, Me.Text1(0), KeyCode, Shift)  '模組化-統一控制
End Sub

'Added by Lydia 2018/11/21
Private Sub textCU29_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU29" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU29"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU06_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU06" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU06"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU114_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU114" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU114"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU41_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU41" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU41"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU44_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU44" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU44"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU47_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU47" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU47"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU50_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU50" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU50"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU53_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU53" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU53"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU56_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU56" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU56"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU60_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU60" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU60"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU63_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU63" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU63"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub

Private Sub textCU93_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SyxMsg <> "textCU93" Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "textCU93"
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
End Sub
