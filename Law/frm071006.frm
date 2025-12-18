VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071006 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文"
   ClientHeight    =   6504
   ClientLeft      =   5160
   ClientTop       =   972
   ClientWidth     =   9060
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6504
   ScaleWidth      =   9060
   Begin VB.CommandButton CmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7020
      TabIndex        =   23
      Top             =   30
      Width           =   1100
   End
   Begin VB.CommandButton CmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6192
      TabIndex        =   22
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton CmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8148
      TabIndex        =   24
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton Command1 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   5064
      TabIndex        =   21
      Top             =   30
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4395
      Left            =   144
      TabIndex        =   39
      Top             =   2070
      Width           =   8784
      Begin VB.TextBox txtCp44 
         Height          =   285
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   4
         Top             =   470
         Width           =   972
      End
      Begin VB.CommandButton CmdDot 
         Caption         =   "工作點數分配"
         Height          =   255
         Left            =   3660
         TabIndex        =   2
         Top             =   162
         Width           =   1305
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   6900
         MaxLength       =   1
         TabIndex        =   17
         Top             =   2010
         Width           =   375
      End
      Begin VB.TextBox textPrint 
         Height          =   285
         Left            =   3570
         MaxLength       =   1
         TabIndex        =   16
         Top             =   2010
         Width           =   732
      End
      Begin VB.TextBox txtCP113 
         Height          =   285
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   1
         Top             =   162
         Width           =   600
      End
      Begin VB.TextBox txtcp20 
         Height          =   285
         Left            =   6900
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1702
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "輸入"
         Height          =   255
         Left            =   7860
         TabIndex        =   8
         Top             =   778
         Width           =   615
      End
      Begin VB.TextBox txtNextProgress 
         Height          =   285
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1394
         Width           =   735
      End
      Begin VB.TextBox txtRule 
         Height          =   285
         Left            =   3570
         MaxLength       =   7
         TabIndex        =   14
         Top             =   1702
         Width           =   1095
      End
      Begin VB.TextBox txtDays 
         Height          =   285
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1394
         Width           =   615
      End
      Begin VB.TextBox txtCustomer 
         Height          =   285
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   9
         Top             =   1086
         Width           =   972
      End
      Begin VB.TextBox txtSubNum 
         Height          =   285
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   6
         Top             =   778
         Width           =   1620
      End
      Begin VB.TextBox txtUDate 
         Height          =   285
         Left            =   5400
         MaxLength       =   7
         TabIndex        =   7
         Top             =   778
         Width           =   1215
      End
      Begin VB.TextBox txtMon 
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1394
         Width           =   615
      End
      Begin VB.TextBox txtDispatch 
         Height          =   285
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   0
         Top             =   162
         Width           =   975
      End
      Begin VB.TextBox txtLimtDate 
         Height          =   285
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   13
         Top             =   1702
         Width           =   975
      End
      Begin MSForms.ComboBox cboGov 
         Height          =   324
         Left            =   5856
         TabIndex        =   3
         Top             =   120
         Width           =   2844
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "5016;572"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCp130 
         Height          =   285
         Left            =   1470
         TabIndex        =   5
         Top             =   470
         Visible         =   0   'False
         Width           =   7215
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   100
         Size            =   "12726;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNextMemo 
         Height          =   460
         Left            =   144
         TabIndex        =   18
         Top             =   2340
         Width           =   8535
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "15055;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtMome1 
         Height          =   465
         Left            =   120
         TabIndex        =   19
         Top             =   3105
         Width           =   8535
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "15055;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtMome2 
         Height          =   465
         Left            =   120
         TabIndex        =   20
         Top             =   3870
         Width           =   8535
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "15055;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "修改通知函內容：           (Y:WORD)"
         Height          =   285
         Index           =   1
         Left            =   5415
         TabIndex        =   65
         Top             =   2010
         Width           =   2745
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿：                   (N:不印)"
         Height          =   285
         Left            =   2625
         TabIndex        =   64
         Top             =   2010
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作時數："
         Height          =   285
         Index           =   12
         Left            =   2112
         TabIndex        =   63
         Top             =   162
         Width           =   1056
      End
      Begin VB.Label lblcp44 
         Height          =   255
         Left            =   2100
         TabIndex        =   62
         Top             =   405
         Width           =   4515
      End
      Begin VB.Label Label16 
         Caption         =   "是否向客戶收款：           (N:不收)"
         Height          =   180
         Left            =   5415
         TabIndex        =   61
         Top             =   1702
         Width           =   2745
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "CF代理人："
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Top             =   470
         Width           =   930
      End
      Begin VB.Label Label29 
         Caption         =   "收件人資料："
         Height          =   285
         Left            =   6780
         TabIndex        =   59
         Top             =   778
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "下一程序備註："
         Height          =   285
         Left            =   120
         TabIndex        =   58
         Top             =   2100
         Width           =   1335
      End
      Begin VB.Label lbeNextProperty 
         Height          =   216
         Left            =   6228
         TabIndex        =   57
         Top             =   1188
         Width           =   2136
      End
      Begin VB.Label lbeNextLimt 
         Height          =   252
         Left            =   6720
         TabIndex        =   52
         Top             =   684
         Width           =   1632
      End
      Begin VB.Label Label8 
         Caption         =   "案件備註："
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   3630
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "本所期限："
         Height          =   285
         Left            =   120
         TabIndex        =   50
         Top             =   1702
         Width           =   972
      End
      Begin VB.Label Label9 
         Caption         =   "下一程序："
         Height          =   285
         Left            =   4536
         TabIndex        =   49
         Top             =   1394
         Width           =   972
      End
      Begin VB.Label Label6 
         Caption         =   "法定期限："
         Height          =   285
         Left            =   2625
         TabIndex        =   48
         Top             =   1702
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "發  文  日："
         Height          =   285
         Left            =   120
         TabIndex        =   47
         Top             =   162
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "機關代號："
         Height          =   285
         Left            =   5010
         TabIndex        =   46
         Top             =   162
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "分所案號："
         Height          =   285
         Left            =   120
         TabIndex        =   45
         Top             =   778
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "管制期限：              月                天"
         Height          =   288
         Left            =   120
         TabIndex        =   44
         Top             =   1392
         Width           =   3516
      End
      Begin VB.Label Label11 
         Caption         =   "當  事  人："
         Height          =   285
         Left            =   120
         TabIndex        =   43
         Top             =   1086
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "催審期限："
         Height          =   285
         Left            =   4536
         TabIndex        =   42
         Top             =   778
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "進度備註："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   41
         Top             =   2850
         Width           =   1815
      End
      Begin MSForms.Label lbeCusName 
         Height          =   285
         Left            =   2115
         TabIndex        =   40
         Top             =   1086
         Width           =   5985
         VariousPropertyBits=   27
         Size            =   "2037;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   345
      Left            =   1230
      TabIndex        =   69
      Top             =   1710
      Width           =   7725
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13626;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label23 
      Caption         =   "協辦人員："
      Height          =   285
      Left            =   4650
      TabIndex        =   68
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label lbeCP29 
      Height          =   285
      Left            =   5670
      TabIndex        =   67
      Top             =   1380
      Width           =   735
   End
   Begin MSForms.Label lbeCP29Name 
      Height          =   285
      Left            =   6450
      TabIndex        =   66
      Top             =   1380
      Width           =   1575
      VariousPropertyBits=   27
      Size            =   "2037;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label19 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   210
      TabIndex        =   56
      Top             =   1710
      Width           =   975
   End
   Begin MSForms.Label lbeSaleName 
      Height          =   285
      Left            =   2010
      TabIndex        =   55
      Top             =   1065
      Width           =   1575
      VariousPropertyBits=   27
      Size            =   "2037;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbePropertyName 
      Height          =   285
      Left            =   6084
      TabIndex        =   54
      Top             =   750
      Width           =   2412
   End
   Begin MSForms.Label lbeEngName 
      Height          =   285
      Left            =   2010
      TabIndex        =   53
      Top             =   1380
      Width           =   1575
      VariousPropertyBits=   27
      Size            =   "2037;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbeSale 
      Height          =   285
      Left            =   1230
      TabIndex        =   38
      Top             =   1065
      Width           =   735
   End
   Begin VB.Label lbePoint 
      Height          =   285
      Left            =   5670
      TabIndex        =   37
      Top             =   1065
      Width           =   615
   End
   Begin VB.Label lbeEng 
      Height          =   285
      Left            =   1224
      TabIndex        =   36
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label lbeProperty 
      Height          =   285
      Left            =   5664
      TabIndex        =   35
      Top             =   750
      Width           =   372
   End
   Begin VB.Label lbeCaseNum 
      Height          =   285
      Left            =   1230
      TabIndex        =   34
      Top             =   750
      Width           =   2055
   End
   Begin VB.Label lbeDate 
      Height          =   285
      Left            =   5664
      TabIndex        =   33
      Top             =   435
      Width           =   1572
   End
   Begin VB.Label lbeNum 
      Height          =   285
      Left            =   1230
      TabIndex        =   32
      Top             =   435
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "點        數："
      Height          =   285
      Left            =   4650
      TabIndex        =   31
      Top             =   1065
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "收  文  日："
      Height          =   285
      Left            =   4650
      TabIndex        =   30
      Top             =   435
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   285
      Left            =   210
      TabIndex        =   29
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號："
      Height          =   285
      Index           =   0
      Left            =   210
      TabIndex        =   28
      Top             =   435
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "案件性質："
      Height          =   285
      Left            =   4650
      TabIndex        =   27
      Top             =   750
      Width           =   972
   End
   Begin VB.Label Label24 
      Caption         =   "智權人員："
      Height          =   285
      Left            =   210
      TabIndex        =   26
      Top             =   1065
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "承  辦  人："
      Height          =   285
      Left            =   210
      TabIndex        =   25
      Top             =   1380
      Width           =   975
   End
End
Attribute VB_Name = "frm071006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/04/28 Form2.0已修改; cboCaseName、lbeSaleName、lbeEngName、lbeCP29Name、lbeCusName、txtNextMemo、txtMome1、txtMome2、txtCp130
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
'2006/1/5整理
Option Explicit

Dim strCP09() As String
Dim t As Integer
Dim Worklc15 As String
Dim blnIsSave As Boolean, LcTmp As String, lc01 As String, lc02 As String, lc03 As String, lc04 As String
Dim rsAD As New ADODB.Recordset
Dim m_Receiver() As String
Dim m_CP12 As String
Dim m_CP11 As String
Dim m_CP29 As String
Dim m_rsRD As New ADODB.Recordset
Dim m_strExSql As String
Dim m_LC15 As String '相關國家
Dim m_strCust1 As String '申請人1
Dim m_LC22 As String 'FC代理人
'2006/1/4 ADD BY SONIA
Dim m_cp40 As String '對造名稱
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_LC40 As String
Dim m_LC41 As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim m_PrevForm As Form '前一畫面 Add By Sindy 2015/10/20
Dim bolACS112 As Boolean 'Added by Lydia 2021/04/28 ACS案件若曾有收文智財顧問112
'Added by Lydia 2024/09/30 (113/11/01上線)
Dim bolHaveCaseLawer As Boolean '是否有出庭律師
Dim bChkPaid As Boolean, m_CCP60 As String  '是否已付款, 收款之收據/請款單號
Dim m_LOS15 As String '案源單號
Dim m_LOS01 As String '案源總收文號
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '案源總收文號之本所案號
Dim m_LOS02 As String '案源案件類型
Dim m_LOS01fa As String '案源之FC代理人
'end 2024/09/30 (113/11/01上線)
Dim m_GovListNew As String, m_GovListDef As String  'Added by Lydia 2025/11/18

'Add By Sindy 2015/10/20
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdBack_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
          Exit Sub
      End If
   End If
   'Add By Sindy 2015/10/20
   If TypeName(m_PrevForm) <> "Nothing" Then
      m_PrevForm.Show
      m_PrevForm.QueryData
   Else
   '2015/10/20 END
      If intForm = 4 Then Tmpfrm071004.Show Else Tmpfrm071005.Show
   End If
   Unload frm071007
   Unload Me
End Sub

Private Sub cmdEnd_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   'Add By Sindy 2015/10/20
   If TypeName(m_PrevForm) <> "Nothing" Then
      m_PrevForm.Show
      m_PrevForm.QueryData
   Else
   '2015/10/20 END
      Unload Tmpfrm071005
      Unload Tmpfrm071004
   End If
   Unload frm071007
   Unload Me
End Sub

Private Sub cmdSure_Click()
On Error GoTo ErrHand
   If m_LC15 <> 台灣國家代號 Then
      If Len(Me.txtCp44.Text) <= 0 Then
         MsgBox "當申請國家非台灣時，CF代理人欄不可為空白!!!", vbExclamation
         txtCp44_GotFocus
         Exit Sub
      End If
   '2008/9/16 ADD BY SONIA
   Else
      If Me.txtCp44.Text <> "" Then
         MsgBox "申請國家為台灣時，CF代理人欄不可輸入!!!", vbExclamation
         txtCp44_GotFocus
         Exit Sub
      End If
   '2008/9/16 END
   End If
   
   'Add By Sindy 2018/11/15
   '承辦人或協辦人員至少要輸一個才可發文
   If lbeEng = "" And lbeCP29 = "" Then
      MsgBox "請先至分案輸入承辦人或協辦人員！", vbExclamation
      Exit Sub
   End If
   'Added by Lydia 2020/03/31 檢查承辦人不可為空白
   If strSrvDate(1) >= 智慧所更名日 And lbeEng = "" Then
      MsgBox "請先至分案輸入承辦人！", vbExclamation
      Exit Sub
   End If
   'end 2020/03/31
   
   If AllTextBeforeSaveCheck Then Exit Sub
   
   If TxtValidate = False Then Exit Sub
   
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Sub
   End If
   
   'Added by Lydia 2020/04/10 若發文時未按工作點數分配按鈕也直接依規則先寫入ACC1N0之工作點數資料。
   'Modified by Lydia 2020/04/17 +啟用工作點數,才自動分配
   If Val(lbePoint.Caption) > 0 And CmdDot.Visible = True Then
      If PUB_GetLawPointAuto(lbeNum.Caption, True, False) = False Then
      End If
   End If
   'end 2020/04/10
   
   If Not SaveData Then DataErrorMessage (3): Exit Sub
   
   'Add By Sindy 2011/6/28
   If lbeProperty <> "9001" And lbeProperty <> "9002" Then
      ' 列印定稿
      If textPrint <> "N" Then
         PrintLetter
      End If
   End If
   
   '若有未發文資料顯示警告
   PUB_GetCPunIssueDatas "" & Me.lbeCaseNum.Caption
   
   'Add By Sindy 2015/10/20
   If TypeName(m_PrevForm) <> "Nothing" Then
      m_PrevForm.Show
      m_PrevForm.QueryData
   Else
   '2015/10/20 END
      Tmpfrm071004.Show
   End If
   Unload frm071006
   Exit Sub
   
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Sub Command1_Click()
Dim i As Integer
Dim strNum As String
Dim strTmp As String
   
   Call Forms(0).SetTmpfrm1103_2 'Add By Sindy 2015/10/20
   With Tmpfrm1103_2
      strTmp = lbeCaseNum.Caption
      If strTmp = "" Then
         Exit Sub
      End If
      i = InStr(strTmp, "-")
      If i <> 0 Then
         strNum = Left(strTmp, i - 1)
         strTmp = Mid(strTmp, i + 1)
      End If
      .intWhereComeFrom = 1
      Set .m_form = Me
      .lblSystem = strNum
      i = InStr(strTmp, "-")
      If i <> 0 Then
         strNum = Left(strTmp, i - 1)
         If strTmp <> "" Then
            strTmp = Mid(strTmp, i + 1)
         End If
         .lblCode(0) = strNum
      Else
         .lblCode(0) = strTmp
         strTmp = ""
      End If
      If i <> 0 Then
         i = InStr(strTmp, "-")
         If i <> 0 Then
            strNum = Left(strTmp, i - 1)
            If strTmp <> "" Then
               strTmp = Mid(strTmp, i + 1)
            End If
            .lblCode(1) = strNum
         Else
            .lblCode(1) = strTmp
         End If
      Else
            .lblCode(1) = "0"
      End If
      
      If strTmp <> "" Then
         .lblCode(2) = strTmp
      Else
         .lblCode(2) = "00"
      End If
      
      .Show
      Me.Hide
   End With
End Sub

Private Sub Command2_Click()
   frm071007.Show
'   Me.Hide
End Sub

Private Sub Form_Load()
Dim i As Integer, n As Integer

   MoveFormToCenter Me
   
   'Add By Sindy 2015/10/20
   If TypeName(m_PrevForm) <> "Nothing" Then
      n = 0
      ReDim Preserve strCP09(n)
      strCP09(n) = m_PrevForm.m_CP09
   Else
      Call Forms(0).SetTmpfrm071004
      Call Forms(0).SetTmpfrm071005
   '2015/10/20 END
      If intForm = 4 Then
         n = 0
         If Tmpfrm071004.Option1.Value = True Then
            ReDim Preserve strCP09(n)
            strCP09(n) = Tmpfrm071004.txtDNum
         ElseIf Tmpfrm071004.Option2.Value = True Then
            ReDim Preserve strCP09(n)
            strCP09(n) = Tmpfrm071004.m_CP09
         End If
      Else
         With Tmpfrm071005.MSHFlexGrid1
            n = 0
            For i = 1 To .Rows - 1
               .row = i
               .col = 0
               If .Text = "v" Then
                  .col = 2
                  ReDim Preserve strCP09(n)
                  strCP09(n) = .Text
                  n = n + 1
               End If
            Next
         End With
      End If
   End If
   
   GetData
   
   'Mark by Lydia 2021/08/18 法務系統的工作點數分配功能先上線
   'CmdDot.Visible = False 'Added by Lydia 2020/04/20 先隱藏
End Sub

Private Sub GetData()
Dim i As Integer, LcTmp As String, cp01 As String
   
   For i = 0 To UBound(strCP09)
      If i = 0 Then
          LcTmp = LcTmp + "'" + strCP09(i) + "'"
      Else
          LcTmp = LcTmp + "," + "'" + strCP09(i) + "'"
      End If
   Next
   Unload frm071007
   
   m_LC15 = ""
      
   'edit by nickc 2007/02/07 不用 dll 了
   'If objPublicData.CheckRecieveCode(strCP09(0), strCP01, strCP02, strCP03, strCP04) Then
   '2009/9/9 MODIFY BY SONIA
   'If ClsPDCheckRecieveCode(strCP09(0), strCP01, strCP02, strCP03, strCP04) Then
   If CheckReceive(strCP09(0)) Then
   '2009/9/9 END
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      '2019/7/29
      If strCP01 = "L" Or strCP01 = "FCL" Or strCP01 = "CFL" Or strCP01 = "LIN" Or strCP01 = "ACS" Then
         '2006/1/5 MODIFY BY SONIA 加對造名稱
         'strExc(1) = "select cp05,cp06,cp07,cp08,cp09,cp10,cp01,cp02,cp03,cp04,cp13,cp14,cp18,cp27,cp29,cp64,cp71," + _
         '   " cp50,lc05,lc06,lc07,lc09,lc11,lc15,lc16,lc27,cp12,cp11,LC22  from caseprogress,lawcase where cp09 = (" + LcTmp + ") and" + _
         '   " CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and cp27 is null  and cp57 is null"
         'edit by nickc 2008/02/22
         'strExc(1) = "select cp05,cp06,cp07,cp08,cp09,cp10,cp01,cp02,cp03,cp04,cp13,cp14,cp18,cp27,cp29,cp64,cp71," + _
            " cp50,NVL(CP40,NVL(CP41,CP42)) AS CP40,lc05,lc06,lc07,lc09,lc11,lc15,lc16,lc27,cp12,cp11,LC22  from caseprogress,lawcase where cp09 = (" + LcTmp + ") and" + _
            " CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and cp27 is null  and cp57 is null"
         strExc(1) = "select cp05,cp06,cp07,cp08,cp09,cp10,cp01,cp02,cp03,cp04,cp13,cp14,cp18,cp27,cp29,cp64,cp71," + _
            " cp50,NVL(CP40,NVL(CP41,CP42)) AS CP40,lc05,lc06,lc07,lc09,lc11,lc15,lc16,lc27,cp12,cp11,LC22,cp44,lc40,lc41,cp116,cp113,cp130 " + _
            " from caseprogress,lawcase where cp09 = (" + LcTmp + ") and" + _
            " CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and cp27 is null  and cp57 is null"
      ElseIf strCP01 = "LA" Then
         '2006/1/5 MODIFY BY SONIA 加對造名稱
         'strExc(1) = "select cp05,cp06,cp07,cp08,cp09,cp10,cp01,cp02,cp03,cp04,cp13,cp14,cp18,cp27,cp29,cp64,cp71," + _
         '   " cp50,hc06,hc05 as lc11,hc07 as lc16,hc12 as lc27,cp12,cp11,'' AS LC22  from caseprogress,hirecase where cp09 = (" + LcTmp + ") and" + _
         '   " CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 and cp27 is null and cp57 is null  "
         'edit by nickc 2008/02/22
         'strExc(1) = "select cp05,cp06,cp07,cp08,cp09,cp10,cp01,cp02,cp03,cp04,cp13,cp14,cp18,cp27,cp29,cp64,cp71," + _
            " cp50,NVL(CP40,NVL(CP41,CP42)) AS CP40,hc06,hc05 as lc11,hc07 as lc16,hc12 as lc27,cp12,cp11,'' AS LC22  from caseprogress,hirecase where cp09 = (" + LcTmp + ") and" + _
            " CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 and cp27 is null and cp57 is null  "
         strExc(1) = "select cp05,cp06,cp07,cp08,cp09,cp10,cp01,cp02,cp03,cp04,cp13,cp14,cp18,cp27,cp29,cp64,cp71," + _
            " cp50,NVL(CP40,NVL(CP41,CP42)) AS CP40,hc06,hc05 as lc11,hc07 as lc16,hc12 as lc27,cp12,cp11,'' AS LC22,cp44,'' as LC40,'' as LC41,cp116,cp113,cp130 " + _
            " from caseprogress,hirecase where cp09 = (" + LcTmp + ") and" + _
            " CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 and cp27 is null and cp57 is null  "
      End If
      intI = 0
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(1))
      If intI = 1 Then
         PutDataInObject
      Else
         IsNoExistData = True
      End If
   End If
   
   bolACS112 = False 'Added by Lydia 2021/04/28
   
   'Add By Sindy 2020/10/12
   If lc01 = "ACS" Then
      '上方固定欄位取消協辦人欄
      Label23.Visible = False
      lbeCP29.Visible = False
      lbeCP29Name.Visible = False
      '下方輸入欄位取消機關代號、CF代理人、收件人資料及輸入按鈕
      '機關代號
      Label4.Visible = False
      'Modified by Lydia 2025/11/18 改成下拉選單
      'txtGov.Visible = False
      'lbeGov.Visible = False
      cboGov.Visible = False
      'end 2025/11/18
      'CF代理人
      txtCp44.Visible = False
      lblcp44.Visible = False
      '收件人資料及輸入按鈕
      Label29.Visible = False
      Command2.Visible = False
      '+主管機關：CP130，預設為案件國家收費表之主管機關CF10，可修改也可空白。
      Label15.Caption = "主管機關名稱："
      txtCP130.Visible = True
      'Added by Lydia 2021/04/28 ACS案件若曾有收文智財顧問
      strExc(1) = "select cp09 from caseprogress where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "' and cp10='112' and cp159=0 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
      If intI = 1 Then
          bolACS112 = True
      End If
      'end 2021/04/28
   End If
   '2020/10/12 END
   
    'Added by Lydia 2024/09/30 (113/11/01上線) 出庭費領取：檢查
    bolHaveCaseLawer = False
    bChkPaid = False
    m_CCP60 = ""
    bChkPaid = PUB_ChkIsPaid(lbeNum, m_CCP60) '是否已請款、已付款
    Call ReadLOS
    If lc01 <> "LA" And InStr(lc01, "L") > 0 Then
      If Pub_ChkPtyCL(lc01, lbeProperty) = True Then  '1.新增會計科目為2201131案件性質發文檢查沒有CaseLawer的設定(Pub_ChkPtyCL內含「出庭費特殊性質」檢查)
         '2.檢查案源是否可以輸入出庭律師
         If Pub_ChkLosToCL(lbeNum.Caption, False, strExc(1)) = False Then
            bolHaveCaseLawer = False
            '另外檢查是否存在出庭費
            If strExc(1) <> "" Then
               strExc(0) = "select cl01,cl02,cl03 from caselawer where cl01='" & Trim(lbeNum.Caption) & "' and nvl(cl03,0)> 0 "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox strExc(1), vbExclamation + vbOKOnly, "資料稽核"
                  cmdSure.Enabled = False
               End If
            End If
         Else
            bolHaveCaseLawer = True
         End If
         If bolHaveCaseLawer = True Then
            strExc(0) = "select cl01,cl02,cl03 from caselawer where cl01='" & Trim(lbeNum.Caption) & "' and nvl(cl03,0)>0 "
            '3.增加特定案件性質可輸出庭費，但也可以輸0表示有輸過。(出庭費可以輸入0的狀況1)
            strExc(2) = Pub_GetSpecMan("出庭費特殊性質")
            If InStr(lc01, "L") > 0 And InStr(";" & strExc(2) & ";", ";" & lbeProperty & ";") > 0 Then
               '出庭費特殊性質的控制，因為此類案件性質多數不必輸出庭費，請改為開放可輸入出庭費，但不必檢查設定出庭費=0
            Else
               '4.案源為商標且有FC代理人之法務案34行政訴訟程序若已輸入0則不必再提醒。(出庭費可以輸入0的狀況2)
               If m_LOS01cp01 <> "" And m_LOS01cp01 <> "TT" And InStr(m_LOS01cp01, "T") > 0 And m_LOS01fa <> "" And lbeProperty = "34" Then
                   strExc(0) = Replace(strExc(0), "and nvl(cl03,0)>0", "")
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  MsgBox "請先到分案作業，執行〔出庭律師〕設定人員和出庭費！", vbCritical, "資料稽核"
                  cmdSure.Enabled = False
               End If
            End If
         End If 'If bolHaveCaseLawer = True Then
      End If  'If Pub_ChkPtyCL(lc01, lbeProperty) = True Then
    End If
    'end 2024/09/30 (113/11/01上線)
End Sub

Private Function SaveData() As Boolean
Dim j As Integer
Dim strCP50 As Variant, strNewNum As String
Dim strPaperNum As String
Dim strNP08 As String 'Add By Sindy 2015/10/22

On Error GoTo ErrorHandler
   
   SaveData = True
   cnnConnection.BeginTrans
   
   LcTmp = lc01 + lc02 + lc03 + lc04
   'Modify By Sindy 2020/10/12 + ",cp130=" + CNULL(ChgSQL(txtCp130))
   'Modified by Lydia 2025/11/18 ",cp71=" + CNULL(txtGov) => IIf(cboGov.Visible = True, ",cp71='" & Trim(Left(cboGov.Text, 3)) & "' ", "")
   strExc(1) = "update caseprogress set cp27=" + CNULL(ChangeTStringToWString(txtDispatch)) + _
      IIf(cboGov.Visible = True, ",cp71='" & Trim(Left(cboGov.Text, 3)) & "' ", "") + ",cp64=" + CNULL(ChgSQL(txtMome1)) + ",cp130=" + CNULL(ChgSQL(txtCP130))
   '2005/10/3 MODIFY BY SONIA
   'If lc01 = "CFL" Then
   '   strExc(1) = strExc(1) & ",cp44=" + CNULL(GetNewFagent(txtCp44)) + ",cp20=" + CNULL(txtcp20.Text)
   'End If
   strExc(1) = strExc(1) & ",cp20=" + CNULL(txtcp20.Text)
   If txtCp44 <> "" Then
      strExc(1) = strExc(1) & ",cp44=" + CNULL(GetNewFagent(txtCp44))
      'add by nickc 2008/02/22
      m_CP44New = GetNewFagent(txtCp44)
   End If
   '2005/10/3 END
   '2009/12/8 add by sonia
   If txtCP113 <> "" Then
      strExc(1) = strExc(1) & ",cp113=" + CNULL(txtCP113)
   End If
   '2009/12/8 END
   strExc(1) = strExc(1) & " where cp09='" + lbeNum + "'"
   cnnConnection.Execute strExc(1)
   
   'Added by Lydia 2025/02/24 TIPS分配比例管制：與ACS案有關之智財協作發文時一併產生TIPS案請款階段分配比例
   If lc01 = "L" Then
      Call PUB_InsertACS_TIPS_Rate(lc01, lc02, lc03, lc04, Me.lbeNum, Trim(Me.lbeProperty))
   End If
   'end 2025/02/24
   
   If lc01 <> "LA" Then
      strExc(2) = "update lawcase set lc11=" + CNULL(ChangeCustomerL(txtCustomer)) + _
         ",lc16=" + CNULL(txtSubNum) + ",lc27=" + CNULL(ChgSQL(txtMome2)) + _
         " where " & ChgLawcase(LcTmp)
        cnnConnection.Execute strExc(2)
   Else
      strExc(2) = "update hirecase set hc05=" + CNULL(ChangeCustomerL(txtCustomer)) + _
         ",hc07=" + CNULL(txtSubNum) + ",hc12=" + CNULL(ChgSQL(txtMome2)) + _
         " where " & ChgHirecase(LcTmp)
        cnnConnection.Execute strExc(2)
   End If
   
   '下一程序
   If txtNextProgress <> "" Then
      strExc(1) = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP15," & _
         "NP08,NP09,NP10,NP22) values (" & CNULL(lbeNum) & ",'" & _
         lc01 & "','" & lc02 & "','" & lc03 & "','" & lc04 & "'," & _
         CNULL(txtNextProgress) & "," & _
         CNULL(txtNextMemo) & "," & _
         CNULL(ChangeTStringToWString(txtLimtDate)) & "," & _
         CNULL(ChangeTStringToWString(txtRule)) & "," & CNULL(lbeSale) & _
         "," & GetNextProgressNo() & ")"
      cnnConnection.Execute strExc(1)
   End If
   '催審期限
   If txtUDate.Text <> "" Then
      'Modify by Morgan 2005/9/7 智權人員改抓操作人員
      '...CNULL(lbeSale) & "," & Format(j) & ")"
      strExc(1) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07," & _
         "NP08,NP09,NP10,NP22) values (" & CNULL(lbeNum) & ",'" & _
         lc01 & "','" & lc02 & "','" & lc03 & "','" & lc04 & "','6001'," & _
         CNULL(ChangeTStringToWString(txtUDate.Text)) & "," & _
         CNULL(ChangeTStringToWString(txtUDate.Text)) & "," & _
         CNULL(strUserNum) & "," & GetNextProgressNo() & ")"
      cnnConnection.Execute strExc(1)
   End If
    
   'Add By Sindy 2015/10/22
   If Trim(lbeProperty) = "901" Then '催款
      strNP08 = DBDATE(DateAdd("m", 2, ChangeWStringToWDateString(DBDATE(txtDispatch)))) '本所期限=發文日+2個月
      '期限的智權人員欄位應掛承辦人非使用者
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES (" & CNULL(lbeNum) & ",'" & lc01 & "','" & lc02 & "','" & lc03 & "','" & lc04 & "','901'," & _
                          strNP08 & "," & strNP08 & "," & CNULL(lbeEng) & "," & GetNextProgressNo() & ")"
      cnnConnection.Execute strSql
   End If
   '2015/10/22 END
   
   '檢查案件國家收費表若有收達或提申天數, 則分別新增至下一程序檔
   m_strExSql = "SELECT CF23,CF11 FROM CASEFEE WHERE CF01='" & lc01 & "' AND CF02='" & m_LC15 & "' AND CF03='" & Me.lbeProperty.Caption & "'"
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set m_rsRD = objLawDll.ReadRstMsg(intI, m_strExSql)
   Set m_rsRD = ClsLawReadRstMsg(intI, m_strExSql)
   If intI = 1 Then
      If Not IsNull(m_rsRD.Fields(0)) And m_rsRD.Fields(0) <> 0 Then
         strExc(1) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
            "NP09,NP10,NP22) VALUES ('" & Me.lbeNum.Caption & "','" & lc01 & "','" & lc02 & _
            "','" & lc03 & "','" & lc04 & "'," & 收達 & "," & _
            CompDate(2, m_rsRD.Fields(0), TransDate(Me.txtDispatch.Text, 2)) & "," & _
            CompDate(2, m_rsRD.Fields(0), TransDate(Me.txtDispatch.Text, 2)) & ",'" & _
            strUserNum & "'," & GetNextProgressNo() & ")"
         cnnConnection.Execute strExc(1)
      End If
      If Not IsNull(m_rsRD.Fields(1)) And m_rsRD.Fields(1) <> 0 Then
         strExc(1) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
            "NP09,NP10,NP22) VALUES ('" & Me.lbeNum.Caption & "','" & lc01 & "','" & lc02 & _
            "','" & lc03 & "','" & lc04 & "'," & 提申 & "," & _
            CompDate(2, m_rsRD.Fields(0), TransDate(Me.txtDispatch.Text, 2)) & "," & _
            CompDate(2, m_rsRD.Fields(0), TransDate(Me.txtDispatch.Text, 2)) & ",'" & _
            strUserNum & "'," & GetNextProgressNo() & ")"
         cnnConnection.Execute strExc(1)
      End If
   End If
   
   'add by sonia 2024/1/3 TIPS案件發文時收據上不列印
   If PUB_ChkACSforTIPS(lc01 & lc02 & lc03 & lc04) = True Then
      strSql = "update acc0k0 set a0k32='Z' where a0k01 in (select distinct a0j13 from acc0j0 where a0j01 in '" & Me.lbeNum.Caption & "') and a0k32 in ('N','Y')"
      cnnConnection.Execute strSql, intI
   End If
   'end 2024/1/3
   
   'Added by Lydia 2024/09/30 (113/11/01上線) 出庭費領取：Email通知承辦律師確認是否領取出庭費
   If lc01 <> "LA" And InStr(lc01, "L") > 0 Then
      If bolHaveCaseLawer = True Then
         strExc(0) = "select cl02 from caselawer,staff where cl01='" & Trim(lbeNum) & "' and nvl(cl03,0) > 0 and cl02=st01(+) and st04='1' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(2) = RsTemp.GetString(adClipString, , , ";")
            If Right(strExc(2), 1) = ";" Then strExc(2) = Mid(strExc(2), 1, Len(strExc(2)) - 1)
            strExc(0) = lc01 & "-" & lc02 & IIf(lc03 & lc03 <> "000", "-" & lc03 & "-" & lc04, "") & IIf(Left(m_LOS02, 1) = "B", "(案源案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & "-" & m_LOS01cp03 & "-" & m_LOS01cp04 & ")", "") & "通知領取出庭費事"
            strExc(1) = "本所案號：" & lc01 & "-" & lc02 & IIf(lc03 & lc03 <> "000", "-" & lc03 & "-" & lc04, "") & vbCrLf & _
                        IIf(m_LOS01cp01 <> "" And m_LOS01cp01 <> "TT", "案源案號：" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 & m_LOS01cp03 <> "000", "-" & m_LOS01cp03 & "-" & m_LOS01cp04, "") & vbCrLf, "") & _
                        "案件名稱：" & Mid(cboCaseName.List(0), 3) & vbCrLf & _
                        "案件性質：" & Trim(lbePropertyName) & vbCrLf
            If m_CCP60 <> "" Then strExc(1) = strExc(1) & IIf(Left(m_CCP60, 1) = "E", "收據號碼：", "請款單號：") & m_CCP60 & vbCrLf
            strExc(1) = strExc(1) & "收款狀態：" & IIf(m_CCP60 = "", "未請款", IIf(bChkPaid = True, "已收款", "未收款")) & vbCrLf
            strExc(1) = strExc(1) & vbCrLf & "已完成委任狀發文程序，請至【法務系統->內法->資料處理->出庭費確認維護】確認開庭費領取事宜。"

            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13) " & _
                       "values('" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') " & _
                       ",'" & ChgSQL(strExc(0)) & "','" & ChgSQL(strExc(1)) & "',null,'" & Trim(lbeNum) & "') "
            cnnConnection.Execute strSql
            strSql = "Update CaseLawer Set CL07=sqldatet(to_char(sysdate,'yyyymmdd'))||decode(cl07,null,'',',')||cl07 Where CL01='" & Trim(lbeNum) & "' and instr ('" & strExc(2) & "',cl02) > 0"
            cnnConnection.Execute strSql
         End If
      End If
   End If
   'end 2024/09/30 (113/11/01上線)
   
   If strPublicTemp <> "" Then '收件人
      strCP50 = Split(strPublicTemp, ",")
      For j = 0 To UBound(strCP50) - 1
         '2006/1/5 MODIFY BY SONIA 收件人與對造名稱相同時更新收件人不另新增B類收文
         'If objPublicData.GetAutoNumber("B", strNewNum, True, True) Then
         '   strPaperNum = "B" + CStr(Year(Date) - 1911) + strNewNum
         'End If
         'strExc(j + 1) = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp50," & _
         '   "cp10,cp13,cp14,cp05,cp43,cp27,cp64,cp71,cp32,cp20,cp11,cp12,cp29,CP44) values ('" & strPaperNum & "','" & _
         '   lc01 & "','" & lc02 & "','" & lc03 & "','" & lc04 & "','" & _
         '   strCP50(j) & "','47','" & lbeSale.Caption & _
         '   "','" & lbeEng.Caption & "','" & Format(Date, "YYYYMMDD") & "','" & _
         '   lbeNum.Caption & "','" & ChangeTStringToWString(txtDispatch) & "','" & ChgSQL(txtMome1.Text) & "','" & _
         '   txtGov.Text & "','N','N','" & m_cp11 & "','" & m_CP12 & "','" & m_cp29 & "','" & CNULL(GetNewFagent(txtCp44)) & "')"
         If strCP50(j) = m_cp40 And m_cp40 <> "" Then
            strExc(j + 1) = "update caseprogress set cp50='" & strCP50(j) & "' where cp09='" + lbeNum + "'"
         Else
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetAutoNumber("B", strNewNum, True, True) Then
            If ClsPDGetAutoNumber("B", strNewNum, True, True) Then
               'Modify By Sindy 2010/8/18 比對自動編號年度
               'strPaperNum = "B" + CStr(Year(Date) - 1911) + strNewNum
               strPaperNum = "B" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) + strNewNum
            End If
            'Modified by Lydia 2025/11/18 txtGov.Text => IIf(cboGov.Visible = True, Trim(Left(cboGov.Text, 3)), "")
            strExc(j + 1) = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp50," & _
               "cp10,cp13,cp14,cp05,cp43,cp27,cp64,cp71,cp32,cp20,cp11,cp12,cp29,CP44) values ('" & strPaperNum & "','" & _
               lc01 & "','" & lc02 & "','" & lc03 & "','" & lc04 & "','" & _
               strCP50(j) & "','47','" & lbeSale.Caption & _
               "','" & lbeEng.Caption & "','" & Format(Date, "YYYYMMDD") & "','" & _
               lbeNum.Caption & "','" & ChangeTStringToWString(txtDispatch) & "','" & ChgSQL(txtMome1.Text) & "','" & _
               IIf(cboGov.Visible = True, Trim(Left(cboGov.Text, 3)), "") & "','N','N','" & m_CP11 & "','" & m_CP12 & "','" & m_CP29 & "'," & CNULL(ChangeCustomerL(txtCp44)) & ")"
         End If
         '2006/1/5 END
         cnnConnection.Execute strExc(j + 1)
      Next
   End If
   If SaveData Then blnIsSave = True Else blnIsSave = False

cnnConnection.CommitTrans

'2015/1/5 modify by sonia 僅投資法務仍要補,故原call basQuery的PUB_CheckEMail, 改移到此form內
'    'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
'    PUB_CheckEMail m_CP44New, m_CP116
'    PUB_CheckEMail m_LC22, m_LC40
'    If m_LC41 <> "" Then
'       PUB_CheckEMail m_LC22, m_LC41
'    End If
'    'end 2008/02/22
    CheckEMail m_CP44New, m_CP116
    CheckEMail m_LC22, m_LC40
    If m_LC41 <> "" Then
       CheckEMail m_LC22, m_LC41
    End If
'2015/1/5 end

Exit Function
ErrorHandler:
   cnnConnection.RollbackTrans
   SaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)

   Call PUB_SendMailCache 'Added by Lydia 2024/09/30 (113/11/01上線)
   strPublicTemp = ""
   Set m_PrevForm = Nothing 'Add By Sindy 2015/10/20
   Set frm071006 = Nothing
End Sub

Private Sub Text7_GotFocus()
   TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
       KeyAscii = 0
   End If
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case "N"
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入 N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

'2009/12/8 ADD BY SONIA
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   
End Sub
'2009/12/3 END

Private Sub txtcp20_GotFocus()
   TextInverse txtcp20
End Sub

Private Sub txtcp20_LostFocus()
   txtcp20.Text = UCase(txtcp20.Text)
   If txtcp20.Text <> "" Then
      If txtcp20.Text <> "N" Then
         MsgBox "只可空白或'N'", vbExclamation, "發文"
         txtcp20.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub txtCp44_GotFocus()
   TextInverse txtCp44
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtCp44.IMEMode = 2
   CloseIme
End Sub

Private Sub txtCp44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCp44_LostFocus()
   If txtCp44.Text = "" Then
      Exit Sub
   End If
  
   txtCp44.Text = UCase(txtCp44.Text)
   
   If Left(txtCp44.Text, 1) = "X" Then
      txtCp44 = "Y" & Mid(txtCp44.Text, 2)
   ElseIf Left(txtCp44.Text, 1) <> "Y" Then
      MsgBox "代理人代碼輸入錯誤!", vbExclamation, "發文"
      txtCp44.SetFocus
      Exit Sub
   End If
   ReadFagent (txtCp44.Text)

End Sub

'Add By Sindy 2013/5/20
' 代理人
Private Sub txtCp44_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTempName As String
   
   Cancel = False
   lblcp44 = Empty
   If IsEmptyText(txtCp44) = False Then
      '聯絡人
      If InStr(txtCp44, "-") > 0 Then
         If ClsPDGetContact(txtCp44, strTempName) Then
            lblcp44 = strTempName
         Else
            If txtCp44.Locked = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "聯絡人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               txtCp44_GotFocus
            End If
         End If
      Else
         If PUB_GetAgentName(strCP01, Me.txtCp44.Text, strTempName) = True Then
            lblcp44 = strTempName
         Else
            lblcp44 = ""
         End If
         If IsEmptyText(lblcp44) = True Then
'            Select Case m_EditMode
'               Case 1, 2:
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "代理人代號不存在"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  txtCp44_GotFocus
'            End Select
         End If
      End If
   End If
End Sub

'Add By Sindy 2020/10/12
Private Sub txtCp130_GotFocus()
   TextInverse txtCP130
   CloseIme
End Sub
'Remove by Lydia 2021/09/14
'Private Sub txtCp130_KeyPress(KeyAscii As Integer)
'   'KeyAscii = UpperCase(KeyAscii)
'End Sub
'2020/10/12 END

Private Sub txtCustomer_Change()
   If txtCustomer = "" Then lbeCusName = ""
End Sub

Private Sub txtCustomer_GotFocus()
   TextInverse txtCustomer
End Sub

Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
Dim StrCusName As String
   
   If txtCustomer <> "" Then
      txtCustomer = UCase(txtCustomer)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      'If objPublicData.GetCustomer(txtCustomer, StrCusName) Then lbeCusName = StrCusName Else Cancel = True: lbeCusName = ""
      'Modify By Sindy 2015/8/27 +strCP01
      If GetCustomerAndState(txtCustomer, StrCusName, , , , strCP01) Then lbeCusName = StrCusName Else Cancel = True: lbeCusName = ""
   Else
      '2005/12/6 MODIFY BY SONIA
      'MsgBox "當事人不可空白", vbCritical
      'lbeCusName = ""
      'Cancel = True
      If m_LC22 = "" Then
         MsgBox "代理人和當事人不可同時空白", vbCritical
         lbeCusName = ""
         Cancel = True
      End If
      '2005/12/6 END
   End If
   
   If Cancel = False Then
      If m_strCust1 <> Me.txtCustomer.Text Then
         If Not PUB_EditCustOk(Me.lbeNum.Caption, lc01, lc02, lc03, lc04) Then Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtCustomer

End Sub

Private Sub txtDays_GotFocus()
   TextInverse txtDays
End Sub

Private Sub txtDays_Validate(Cancel As Boolean)
Dim strDate As String
   
   If Val(txtDays) > 31 Then
       MsgBox "天數不可大於31天"
       Cancel = True
   ElseIf txtDays <> "" Then
         'Modified by Lydia 2018/04/30 計算有問題
'         StrDate = DateAdd("d", Val(txtDays), DateSerial(Val(Left(txtDispatch, 2)) + 1911, Mid(txtDispatch, 3, 2), Right(txtDispatch, 2)))
'         If txtMon <> "" Then
'            StrDate = DateAdd("m", Val(txtMon), StrDate)
'         End If
'            txtRule = ChangeWDateStringToTString(StrDate)
'            StrDate = DateAdd("d", -4, StrDate)
'            txtLimtDate = ChangeWDateStringToTString(StrDate)
         Call GetNewDate
   End If
   If Cancel Then TextInverse txtDays

End Sub

Private Sub txtDispatch_GotFocus()
   TextInverse txtDispatch
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtDispatch.IMEMode = 2
   CloseIme
End Sub

Private Sub txtDispatch_Validate(Cancel As Boolean)
   If txtDispatch <> "" Then
      If CheckIsTaiwanDate(txtDispatch) Then
         If Val(GetTaiwanTodayDate) - Val(txtDispatch) < 0 Then
            MsgBox "輸入日期大於系統日", vbCritical
            Cancel = True
         Else
            If lc01 = "LA" Then
               GetCaseFee lc01, "000", lbeProperty
            Else
               GetCaseFee lc01, Worklc15, lbeProperty
            End If
         End If
       Else
          Cancel = True
       End If
   End If
   If Cancel Then TextInverse txtDispatch

End Sub

'Mark by Lydia 2025/11/18
'Private Sub txtGov_Change()
'   If txtGov = "" Then lbeGov = ""
'
'End Sub
'
'Private Sub txtGov_GotFocus()
'   TextInverse txtGov
'End Sub
'
'Private Sub txtGov_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
'Private Sub txtGov_Validate(Cancel As Boolean)
'Dim strTempName As String
'
'   If txtGov <> "" Then
'      'edit by nickc 2007/02/07 不用 dll 了
'      'If objLawDll.GetGovName(txtGov, strTempName) Then lbeGov = strTempName Else Cancel = True
'      If ClsPDGetGovName(txtGov, strTempName) Then lbeGov = strTempName Else Cancel = True
'   Else
'    lbeGov = ""
'   End If
'   If Cancel Then TextInverse txtGov
'
'End Sub
'end 2025/11/18

Private Sub txtLimtDate_GotFocus()
   TextInverse txtLimtDate
End Sub

Private Sub txtLimtDate_Validate(Cancel As Boolean)
   If txtLimtDate <> "" Then
      If CheckIsTaiwanDate(txtLimtDate) Then
         If txtRule <> "" Then
            If Val(txtRule) - Val(txtLimtDate) < 0 Then
               DataErrorMessage 13
               Cancel = True
            End If
         End If
      Else
           MsgBox "輸入日期非民國日期", vbCritical
           Cancel = True
      End If
   ElseIf txtNextProgress <> "" And txtLimtDate = "" Then
      MsgBox "本所期限不可空白"
      Cancel = True
   End If
   If Cancel Then TextInverse txtLimtDate

End Sub

Private Sub txtMome1_GotFocus()
   TextInverse txtMome1
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtMome1.IMEMode = 1
   OpenIme
End Sub

Private Sub txtMome1_Validate(Cancel As Boolean)
   If txtMome1 <> "" Then
     If CheckLengthIsOK(txtMome1, 2000) = False Then
         Cancel = True
         txtMome1.SetFocus
     End If
   End If
   'edit by nickc 2007/06/11  切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub txtMome2_GotFocus()
   TextInverse txtMome2
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtMome2.IMEMode = 1
   OpenIme
End Sub

Private Sub txtMome2_Validate(Cancel As Boolean)
   If txtMome2 <> "" Then
      If CheckLengthIsOK(txtMome2, 2000) = False Then
         Cancel = True
         txtMome2.SetFocus
      End If
   End If
   'edit by nickc 2007/06/11  切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub txtMon_GotFocus()
   TextInverse txtMon
End Sub

Private Sub txtMon_Validate(Cancel As Boolean)
Dim strDate As String

   If txtMon <> "" Then
      If txtMon > 12 Or txtMon < 0 Then
         MsgBox "輸入錯誤!", vbCritical
         txtMon = ""
         Cancel = True
      Else
         'Modified by Lydia 2018/04/30 計算有問題
'         strDate = DateAdd("m", Val(txtMon), DateSerial(Val(Left(txtDispatch, 2)) + 1911, Mid(txtDispatch, 3, 2), Right(txtDispatch, 2)))
'         If txtDays <> "" Then
'             strDate = DateAdd("d", txtDays, strDate)
'         End If
'         txtRule = ChangeWDateStringToTString(strDate)
'         strDate = DateAdd("d", -4, strDate)
'         txtLimtDate = ChangeWDateStringToTString(strDate)
         Call GetNewDate
      End If
   End If
   If Cancel Then TextInverse txtMon
End Sub

Private Sub txtNextMemo_GotFocus()
   TextInverse txtNextMemo
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtNextMemo.IMEMode = 1
   OpenIme
End Sub

Private Sub txtNextMemo_Validate(Cancel As Boolean)
   If txtNextMemo <> "" Then
     If CheckLengthIsOK(txtNextMemo, 2000) = False Then
         Cancel = True
         txtNextMemo.SetFocus
     End If
   End If
   'edit by nickc 2007/06/11  切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub txtNextProgress_Change()
   If txtNextProgress = "" Then lbeNextProperty = ""
End Sub

Private Sub txtNextProgress_GotFocus()
   TextInverse txtNextProgress
End Sub

Private Sub txtNextProgress_Validate(Cancel As Boolean)
Dim strTemp1 As String, strTemp2 As String
 
   If txtNextProgress <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetCaseProperty(GetCaseNumSysKind(lbeCaseNum), txtNextProgress, strTemp1, False) Then lbeNextProperty = strTemp1 Else Cancel = True
      If ClsPDGetCaseProperty(GetCaseNumSysKind(lbeCaseNum), txtNextProgress, strTemp1, False) Then lbeNextProperty = strTemp1 Else Cancel = True
   Else
      lbeNextProperty = ""
   End If
   If Cancel Then TextInverse txtNextProgress

End Sub

Private Sub txtRule_GotFocus()
   TextInverse txtRule
End Sub

Private Sub txtRule_Validate(Cancel As Boolean)
   If txtRule <> "" Then
      If CheckIsTaiwanDate(txtRule) Then
         If txtLimtDate <> "" Then
            If Val(txtRule) - Val(txtLimtDate) < 0 Then
               DataErrorMessage 12
               Cancel = True
            End If
         End If
      Else
         Cancel = True
      End If
   ElseIf txtNextProgress <> "" And txtLimtDate = "" Then
      MsgBox "法定期限不可空白"
      Cancel = True
   End If
   If Cancel Then TextInverse txtRule

End Sub

Private Function DataMap(i As Integer, strData As String, Optional IsCustomer As Boolean) As String
Dim strTempName As String
Dim strTemp As String 'Added by Lydia 2023/12/27

   Select Case i
      Case 0
         DataMap = ChangeTStringToTDateString(ChangeWStringToTString(strData))
      Case 1
         DataMap = ChangeWStringToTString(strData)
      Case 2
         'edit by nickc 2007/02/07 不用 dll 了
         'If objLawDll.GetGovName(strData, strTempName) Then DataMap = strTempName
         If ClsPDGetGovName(strData, strTempName) Then DataMap = strTempName
      'Modified by Lydia 2023/12/27 區分人員: CP13=3, cp14=14,cp29=29
      Case 3, 14, 29
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(strData, strTempName) Then DataMap = strTempName
         'Modified by Lydia 2023/12/27 修改智權人員欄或CP14承辦人欄己離職，A：仍要帶出姓名、B：彈出的訊息請帶出是智權人員或是承辦人或是協辦人員。
         'If ClsPDGetStaff(strData, strTempName) Then DataMap = strTempName
         strTempName = GetStaffName(strData, True, , , strTemp)
         If strTemp <> "1" Then
             Select Case i
                Case 3: strTemp = "智權人員"
                Case 14: strTemp = "承辦人員"
                Case 29: strTemp = "協辦人員"
             End Select
             MsgBox strTemp & "已離職！"
         End If
         DataMap = strTempName
      Case 4
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCaseProperty(RsTemp.Fields!CP01, strData, strTempName, False) Then DataMap = strTempName
         If ClsPDGetCaseProperty(RsTemp.Fields!cp01, strData, strTempName, False) Then DataMap = strTempName
      Case 5
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCustomer(strData, strTempName) Then DataMap = strTempName
         If ClsPDGetCustomer(strData, strTempName) Then DataMap = strTempName
         If IsCustomer Then txtCustomer = strData
   End Select

End Function

Private Sub PutDataInObject()
Dim strTemp1 As String, strTemp2 As String
Dim strCF10 As String
   
   lbeCaseNum = GiveSymbol(RsTemp.Fields!cp01, RsTemp.Fields!cp02, RsTemp.Fields!cp03, RsTemp.Fields!cp04, LcTmp)
   lc01 = RsTemp.Fields!cp01
   lc02 = RsTemp.Fields!cp02
   lc03 = RsTemp.Fields!cp03
   lc04 = RsTemp.Fields!cp04
   lbeNum = RsTemp.Fields!CP09

   If lc01 = "LA" Then
       If Not IsNull(RsTemp.Fields!hc06) Then cboCaseName.AddItem "中:" + RsTemp.Fields!hc06
       If Not IsNull(RsTemp.Fields!CP10) Then ChkUData
       m_LC15 = "000"
   Else
       If Not IsNull(RsTemp.Fields!lc05) Then cboCaseName.AddItem "中:" + RsTemp.Fields!lc05
       If Not IsNull(RsTemp.Fields!lc06) Then cboCaseName.AddItem "英:" + RsTemp.Fields!lc06
       If Not IsNull(RsTemp.Fields!lc07) Then cboCaseName.AddItem "日:" + RsTemp.Fields!lc07
       If Not IsNull(RsTemp.Fields!lc15) Then
          If Not IsNull(RsTemp.Fields!CP10) Then ChkUData
       End If
       If Not IsNull(RsTemp.Fields!lc15) Then Worklc15 = RsTemp.Fields!lc15
       m_LC15 = "" & RsTemp.Fields!lc15
   End If
   If cboCaseName.ListCount <> 0 Then
      cboCaseName.ListIndex = 0
   End If

   If IsNull(RsTemp.Fields!LC11) Then txtCustomer = "" Else txtCustomer = RsTemp.Fields!LC11: lbeCusName = DataMap(5, txtCustomer, True)
   '2005/12/6 ADD BY SONIA
   If IsNull(RsTemp.Fields!LC22) Then m_LC22 = "" Else m_LC22 = RsTemp.Fields!LC22
   '2005/12/6 END
   m_strCust1 = "" & Me.txtCustomer.Text
   
   If IsNull(RsTemp.Fields!lc16) Then txtSubNum = "" Else txtSubNum = RsTemp.Fields!lc16
   If IsNull(RsTemp.Fields!cp05) Then lbeDate = "" Else lbeDate = DataMap(0, RsTemp.Fields!cp05)
   If IsNull(RsTemp.Fields!CP10) Then lbeProperty = "" Else lbeProperty = RsTemp.Fields!CP10:    lbePropertyName = DataMap(4, RsTemp.Fields!CP10)
   If Not IsNull(RsTemp.Fields!cp13) Then lbeSale = RsTemp.Fields!cp13: lbeSaleName = DataMap(3, RsTemp.Fields!cp13)
   'Modified by Lydia 2023/12/27 DataMap(3) 改為DataMap(14)
   If IsNull(RsTemp.Fields!cp14) Then lbeEng = "" Else lbeEng = RsTemp.Fields!cp14: lbeEngName = DataMap(14, RsTemp.Fields!cp14)
   'Modified by Lydia 2023/12/27 DataMap(3) 改為DataMap(29)
   If IsNull(RsTemp.Fields!cp29) Then lbeCP29 = "" Else lbeCP29 = RsTemp.Fields!cp29: lbeCP29Name = DataMap(29, RsTemp.Fields!cp29) 'Add By Sindy 2018/11/15
   If IsNull(RsTemp.Fields!cp18) Then lbePoint = "" Else lbePoint = RsTemp.Fields!cp18
   If IsNull(RsTemp.Fields!Cp27) Then txtDispatch = "" Else txtDispatch = DataMap(1, RsTemp.Fields!Cp27)
   If IsNull(RsTemp.Fields!CP64) Then txtMome1 = "" Else txtMome1 = RsTemp.Fields!CP64
   'Modifie by Lydia 2025/11/18 改成下拉選單
   'If IsNull(RsTemp.Fields!cp71) Then txtGov = "" Else txtGov = RsTemp.Fields!cp71: lbeGov = DataMap(2, RsTemp.Fields!cp71)
   Call PUB_SetGovCmb(Me.cboGov, m_GovListNew, "" & RsTemp.Fields("cp71"))
   m_GovListDef = m_GovListNew
   Me.cboGov.Tag = Me.cboGov.Text
   'end 2025/11/18
   If IsNull(RsTemp.Fields!lc27) Then txtMome2 = "" Else txtMome2 = RsTemp.Fields!lc27
   If IsNull(RsTemp.Fields!cp11) Then m_CP11 = "" Else m_CP11 = RsTemp.Fields!cp11
   If IsNull(RsTemp.Fields!cp12) Then m_CP12 = "" Else m_CP12 = RsTemp.Fields!cp12
   If IsNull(RsTemp.Fields!cp29) Then m_CP29 = "" Else m_CP29 = RsTemp.Fields!cp29
   '2006/1/6 ADD BY SONIA
   If IsNull(RsTemp.Fields!cp40) Then m_cp40 = "" Else m_cp40 = RsTemp.Fields!cp40
   '2008/9/16 ADD BY SONIA
   If IsNull(RsTemp.Fields!CP44) Then
      txtCp44 = ""
   Else
      txtCp44 = RsTemp.Fields("CP44")
      ReadFagent (txtCp44.Text)
   End If
   '2008/9/22 END
   If IsNull(RsTemp.Fields!CP113) Then txtCP113 = "" Else txtCP113 = RsTemp.Fields!CP113   '2009/12/8 ADD BY SONIA
      
   'ADD BY Sindy 2020/10/12
   If lc01 = "ACS" Then
      If IsNull(RsTemp.Fields!CP130) Then
         Call GetCaseFeeByNick(lc01, m_LC15, lbeProperty.Caption, strCF10) '預設值
         txtCP130 = strCF10
      Else
         txtCP130 = RsTemp.Fields("CP130")
      End If
   End If
   '2020/10/12 END
   'Added by Lydia 2025/10/27各專業部的智財協作發文，都預設不出定稿
   If lc01 = "L" And lbeProperty = "7601" Then
      textPrint = "N"
   End If
   'end 2025/10/27
   
   'add by nickc 2008/02/22
    m_CP116 = CheckStr(RsTemp.Fields("CP116"))
    m_CP44New = CheckStr(RsTemp.Fields("CP44"))
    m_LC22 = CheckStr(RsTemp.Fields("LC22"))
    m_LC40 = CheckStr(RsTemp.Fields("LC40"))
    m_LC41 = CheckStr(RsTemp.Fields("LC41"))
End Sub

Private Function ChkUData() As Boolean
Dim dblTempDays As Double, strDate As Variant

   If lc01 = "LA" Then
      GetCaseFee lc01, "000", RsTemp.Fields!CP10
   Else
      GetCaseFee lc01, RsTemp.Fields!lc15, RsTemp.Fields!CP10
   End If
End Function

Private Sub GetCaseFee(ByVal CF01 As String, ByVal CF02 As String, ByVal CF03 As String)
   intI = 1
   strExc(0) = "SELECT CF05,CF11 FROM CASEFEE WHERE CF01='" & CF01 & "' AND " & _
      "CF02='" & CF02 & "' AND CF03='" & CF03 & "'"
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set rsAD = objLawDll.ReadRstMsg(intI, strExc(0))
   Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(rsAD.Fields(0)) Then
         txtUDate = TransDate(CompDate(2, Val(rsAD.Fields(0)), TransDate(txtDispatch, 2)), 1)
      End If
   End If
End Sub

Private Sub txtSubNum_GotFocus()
   TextInverse txtSubNum
End Sub

Private Sub txtSubNum_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUDate_GotFocus()
   TextInverse txtUDate
End Sub

Private Sub txtUDate_Validate(Cancel As Boolean)
   If txtUDate <> "" Then
      If Not CheckIsTaiwanDate(txtUDate) Then
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtUDate

End Sub

Private Function AllTextBeforeSaveCheck() As Boolean
Dim i As Integer, StrCusName As String
   
   If txtDispatch <> " " Then
      If CheckIsTaiwanDate(txtDispatch) Then
         If Val(GetTaiwanTodayDate) - Val(txtDispatch) < 0 Then
            MsgBox "輸入日期大於系統日", vbCritical
            txtDispatch.SetFocus
            AllTextBeforeSaveCheck = True
            Exit Function
         End If
      Else
         txtDispatch.SetFocus
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
   End If
   If txtCustomer <> "" Then
      txtCustomer = UCase(txtCustomer)
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetCustomer(txtCustomer, StrCusName) Then
      If ClsPDGetCustomer(txtCustomer, StrCusName) Then
         lbeCusName = StrCusName
      Else
         TextInverse txtCustomer
         lbeCusName = ""
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
   Else
      '2005/12/6 MODIFY BY SONIA
      'MsgBox "當事人不可空白", vbCritical
      'TextInverse txtCustomer
      'lbeCusName = ""
      'Exit Function
      If m_LC22 = "" Then
         MsgBox "FC代理人和當事人不可同時空白", vbCritical
         TextInverse txtCustomer
         lbeCusName = ""
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
      '2005/12/6 END
   End If
   
   'Add By Sindy 2011/6/9
   If lc01 = "LA" And txtCP113 = "" Then
      If MsgBox("是否輸入工作時數？", vbExclamation + vbYesNo) = vbYes Then
         txtCP113.SetFocus
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
   End If
   
   'Added by Lydia 2021/04/28 ACS智財顧問專業分配比例管制：ACS案件若曾有收文智財顧問112，每道進度發文都一定要輸入工作時數，輸0也可以
   If bolACS112 = True And Trim(txtCP113) = "" And strSrvDate(1) >= ACS_PFrateStart Then
       MsgBox "請輸入工作時數！", vbInformation
       txtCP113.SetFocus
       txtCP113_GotFocus
       AllTextBeforeSaveCheck = True
       Exit Function
   End If
   'end 2021/04/28
   
    'Added by Lydia 2021/07/16 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
    If Pub_ChkACS112isNull(lc01, lc02, lc03, lc04, txtCP113) = True Then
          txtCP113.SetFocus
          txtCP113_GotFocus
          AllTextBeforeSaveCheck = True
          Exit Function
    End If
    'end 2021/07/16
    
   AllTextBeforeSaveCheck = False
End Function

Private Sub ReadFagent(ByVal strCP44 As String)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strFA01 As String

   strFA01 = GetNewFagent(strCP44)
   'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
   'strSQL = "SELECT nvl(fa05,nvl(fa04,fa06)) FROM FAGENT " & _
            "WHERE FA01 = '" & Left(strCP44, 8) & "'"
   strSql = "SELECT nvl(fa05,nvl(fa04,fa06)),fa69 FROM FAGENT " & _
            "WHERE FA01 = '" & Left(strCP44, 8) & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      If Not IsNull(rsTmp.Fields(0)) Then
         lblcp44.Caption = rsTmp.Fields(0)
      End If
      'add by nick 2004/07/22
      If CheckStr(rsTmp.Fields(1).Value) = "不再使用" Then
              MsgBox "此代理人資料已不再使用，請確認！！", , MsgText(5)
              txtCp44.SetFocus
      End If
   Else
      MsgBox "代理人代碼不存在!", vbExclamation, "發文"
      txtCp44.SetFocus
   End If
   rsTmp.Close
   Set rsTmp = Nothing

End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   If Me.txtCustomer.Enabled = True Then
      Cancel = False
      txtCustomer_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtDays.Enabled = True Then
      Cancel = False
      txtDays_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtDispatch.Enabled = True Then
      Cancel = False
      txtDispatch_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   '2009/12/8 ADD BY SONIA
   txtCP113_Validate Cancel
   If Cancel = True Then
      txtCP113.SetFocus
      Exit Function
   End If
   '2009/12/8 END
   
   'Modified by Lydia 2025/11/18 改成下拉選單
   'If Me.txtGov.Enabled = True Then
   '   Cancel = False
   '   txtGov_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   If Me.cboGov.Visible = True Then
      Cancel = False
      cboGov_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2025/11/18
   
   If Me.txtLimtDate.Enabled = True Then
      Cancel = False
      txtLimtDate_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtMome1.Enabled = True Then
      Cancel = False
      txtMome1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtMome2.Enabled = True Then
      Cancel = False
      txtMome2_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtMon.Enabled = True Then
      Cancel = False
      txtMon_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtNextMemo.Enabled = True Then
      Cancel = False
      txtNextMemo_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtNextProgress.Enabled = True Then
      Cancel = False
      txtNextProgress_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtRule.Enabled = True Then
      Cancel = False
      txtRule_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtUDate.Enabled = True Then
      Cancel = False
      txtUDate_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2013/5/20
   If Me.txtCp44.Enabled = True Then
      Cancel = False
      txtCp44_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2013/5/20 End
   
   TxtValidate = True
End Function

'2009/9/9 ADD BY SONIA
Private Function CheckReceive(strNum As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM CASEPROGRESS WHERE CP09 ='" & strNum & "'" & _
            " AND (CP27 IS NULL OR CP27 ='') AND (CP57 IS NULL OR CP57 = '')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      CheckReceive = True
      strCP01 = rsTmp.Fields("CP01"): strCP02 = rsTmp.Fields("CP02"): strCP03 = rsTmp.Fields("CP03"): strCP04 = rsTmp.Fields("CP04")
   Else
      CheckReceive = False
   End If
End Function
'2009/9/9 END

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintLetter()
Dim bolChk As Boolean
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   If Text7 = "Y" Then
      bolChk = True
   Else
      bolChk = False
   End If
   
   Select Case lc01
      Case "L", "LA"
         NowPrint lbeNum, "01", "01", bolChk, strUserNum, 0
   End Select
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Select Case lc01
      Case "L", "LA"
         ' 清除定稿例外欄位檔原有資料
         EndLetter "01", lbeNum, "01", strUserNum
   End Select
End Sub

'2015/1/5 sonia自basQuery的PUB_CheckEMail複製過來,因只剩投資法務要補輸
'檢查若沒有EMail時提醒補輸並更新
Private Sub CheckEMail(ByVal stAgentNo As String, Optional ByVal stContactNo As String)
Dim stSQL As String, intR As Integer, stMsg As String, stEMail As String
Dim iTbNo As Integer
   
   If stAgentNo = "" Then Exit Sub
   
On Error GoTo ErrHnd

   stAgentNo = Left(stAgentNo & "000", 9)
   If stContactNo <> "" Then
      iTbNo = 1
      stMsg = "聯絡人【 " & stAgentNo & "-" & stContactNo
      stSQL = "select NVL(PCC05,NVL(PCC03,PCC04)) FNAME from potcustcont where pcc01='" & Left(stAgentNo, 8) & "' and pcc02='" & stContactNo & "' and pcc08 is null"
   Else
      If Left(stAgentNo, 1) = "Y" Then
         iTbNo = 2
         stMsg = "代理人【 " & stAgentNo
         stSQL = "select NVL(FA04,NVL(RTRIM(FA05||' '||FA63||' '||FA64||' '||FA65),FA06)) FNAME from fagent where fa01='" & Left(stAgentNo, 8) & "' and fa02='" & Mid(stAgentNo, 9) & "' and fa16 is null"
      Else
         iTbNo = 3
         stMsg = "客戶【 " & stAgentNo
         stSQL = "select NVL(CU04,NVL(RTRIM(CU05||' '||CU88||' '||CU89||' '||CU90),CU06)) FNAME from customer where cu01='" & Left(stAgentNo, 8) & "' and cu02='" & Mid(stAgentNo, 9) & "' and cu20 is null"
      End If
   End If
   intR = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      stMsg = stMsg & " " & AdoRecordSet3(0) & " 】" & "尚無 EMail 資料，請輸入！"
      Do
         stEMail = InputBox(stMsg, "Email 輸入", stEMail)
         If stEMail = "" Then
            If MsgBox("是否確定這次不補 EMail 資料？", vbYesNo + vbDefaultButton2, "Email 輸入") = vbYes Then
               Exit Do
            End If
         Else
            If PUB_CheckMail(stEMail) = True Then
               If iTbNo = 1 Then
                  stSQL = "update potcustcont set pcc08='" & ChgSQL(stEMail) & "' where pcc01='" & Left(stAgentNo, 8) & "' and pcc02='" & stContactNo & "'"
               ElseIf iTbNo = 2 Then
                  stSQL = "update fagent set fa16='" & ChgSQL(stEMail) & "' where fa01='" & Left(stAgentNo, 8) & "' and fa02='" & Mid(stAgentNo, 9) & "'"
               Else
                  stSQL = "update customer set cu20='" & ChgSQL(stEMail) & "' where cu01='" & Left(stAgentNo, 8) & "' and cu02='" & Mid(stAgentNo, 9) & "'"
               End If
               cnnConnection.BeginTrans
On Error GoTo ErrHnd2
               Pub_SeekTbLog stSQL
               cnnConnection.Execute stSQL, intR
               cnnConnection.CommitTrans
               Exit Do
            End If
         End If
      Loop
   End If
   Exit Sub
   
ErrHnd2:
   cnnConnection.RollbackTrans
ErrHnd:
   MsgBox Err.Description
   
End Sub
'2015/1/5 end

'Added by Lydia 2018/04/30 計算所限和法限
Private Sub GetNewDate()
Dim strDate As String
      
      If Trim(txtDispatch) = "" Then Exit Sub
      If Val(txtMon) > 0 Then
           strDate = CompDate(1, Val(txtMon), TransDate(txtDispatch, 2))
      Else
           strDate = TransDate(txtDispatch, 2)
      End If
      If Val(txtDays) > 0 Then
           strDate = CompDate(2, Val(txtDays), strDate)
      End If
      txtRule = TransDate(strDate, 1)
      strDate = CompDate(2, -4, strDate)
      txtLimtDate = TransDate(strDate, 1)
End Sub

'Memo by Lydia 2021/08/31 法務系統的工作點數分配功能先上線(110/9/1)
'Memo by Lydia 2020/04/10 啟用
'Memo by Lydia 2016/?/? 已經隱藏工作點數按鈕
'Added by Lydia 2015/06/01 新增-法務工作點數分配
Private Sub CmdDot_Click()
   If Val(lbePoint.Caption) > 0 Then
        'Added by Lydia 2021/08/18
        If PUB_CheckFormExist("frm071021") Then
             MsgBox "請先關閉【法務工作點數分配】畫面！", vbExclamation
             Exit Sub
        End If
        'end 2021/08/18
        If PUB_GetLawPointAuto(lbeNum.Caption, True, False) = True Then 'Added by Lydia 2020/04/10 內外法發文按工作點數分配按鈕時，若無ACC1N0之工作點數資料則直接依規則先寫入ACC1N0之工作點數資料
            Set frm071021.m_PrevForm = Me
            frm071021.m_bolPrev = True
            frm071021.m_KeyList = lbeNum.Caption
            Me.Hide
            frm071021.Show
        End If  'Added by Lydia 2020/04/10
   Else
        MsgBox "無點數可供分配!", vbExclamation
   End If
End Sub

'Added by Lydia 2024/09/30 (113/11/01上線) 法律所案源收文：讀取法務案源檔
Private Sub ReadLOS()
Dim stSQL As String, intQ As Integer
Dim RsQ As ADODB.Recordset

   m_LOS01 = "": m_LOS01cp01 = "": m_LOS01cp02 = "": m_LOS01cp03 = "": m_LOS01cp04 = ""
   m_LOS02 = "": m_LOS15 = "": m_LOS01fa = ""
   
   stSQL = "select c1.cp09 as lcp09,c1.cp16 as lcp16, los02,los15,c2.cp01 as ocp01, c2.cp02 as ocp02, c2.cp03 as ocp03,c2.cp04 as ocp04,c2.cp09 as ocp09,c2.cp16 as ocp16,NVL(PA75,TM44) FAGENT " & _
           "from caseprogress c1,lawofficesource, caseprogress c2, patent, trademark " & _
           "where c1.cp09='" & Trim(lbeNum.Caption) & "' and c1.cp162=los15(+) and los01=c2.cp09(+) " & _
           "and c2.cp01=pa01(+) and c2.cp02=pa02(+) and c2.cp03=pa03(+) and c2.cp04=pa04(+) and c2.cp01=tm01(+) and c2.cp02=tm02(+) and c2.cp03=tm03(+) and c2.cp04=tm04(+) "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      m_LOS01cp01 = "" & RsQ.Fields("ocp01")
      m_LOS01cp02 = "" & RsQ.Fields("ocp02")
      m_LOS01cp03 = "" & RsQ.Fields("ocp03")
      m_LOS01cp04 = "" & RsQ.Fields("ocp04")
      m_LOS02 = "" & RsQ.Fields("los02")
      m_LOS15 = "" & RsQ.Fields("los15")
      m_LOS01fa = "" & RsQ.Fields("FAGENT")
   End If
      
   Set RsQ = Nothing
   
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_GotFocus()
   TextInverse cboGov
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_Validate(Cancel As Boolean)

   If Trim(cboGov.Text) <> "" And cboGov.Tag <> cboGov.Text Then
      If PUB_ChkGovIsExist(IIf(Val(Trim(Left(cboGov.Text, 3))) > 0, Trim(Left(cboGov.Text, 3)), Trim(cboGov.Text)), strExc(3), strExc(4)) = True Then
         cboGov.Text = strExc(3) & " " & strExc(4)
      Else
         Cancel = True
         cboGov.SetFocus
         cboGov_GotFocus
         Exit Sub
      End If
   End If
   cboGov.Tag = cboGov.Text
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_DropButtonClick()
   If cboGov.Text <> "" Then
      If Val(Trim(Left(cboGov.Text, 3))) > 0 Then
      Else  '依輸入文字模糊比對
         Call PUB_SetGovCmb(cboGov, m_GovListNew, , Trim(cboGov.Text))
         If m_GovListNew = "" Then
            Call PUB_SetGovCmb(cboGov, m_GovListNew)
         End If
      End If
   Else
      If m_GovListNew <> m_GovListDef Then
         Call PUB_SetGovCmb(cboGov, m_GovListNew)
      End If
   End If
End Sub

