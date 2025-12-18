VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090801_7 
   BorderStyle     =   4  '單線固定工具視窗
   Caption         =   "特殊收據"
   ClientHeight    =   6030
   ClientLeft      =   50
   ClientTop       =   290
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   405
      Left            =   2280
      TabIndex        =   72
      Top             =   -60
      Width           =   3300
      Begin VB.CheckBox Check2 
         Caption         =   "４"
         Height          =   250
         Index           =   3
         Left            =   2760
         TabIndex        =   77
         Top             =   120
         Width           =   420
      End
      Begin VB.CheckBox Check2 
         Caption         =   "３"
         Height          =   250
         Index           =   2
         Left            =   2160
         TabIndex        =   76
         Top             =   120
         Width           =   420
      End
      Begin VB.CheckBox Check2 
         Caption         =   "２"
         Height          =   250
         Index           =   1
         Left            =   1560
         TabIndex        =   75
         Top             =   120
         Width           =   420
      End
      Begin VB.CheckBox Check2 
         Caption         =   "１"
         Height          =   250
         Index           =   0
         Left            =   960
         TabIndex        =   74
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "境外公司:"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   73
         Top             =   160
         Width           =   1005
      End
   End
   Begin VB.TextBox txtCRL132 
      Height          =   285
      Left            =   5040
      MaxLength       =   60
      TabIndex        =   34
      Top             =   4800
      Width           =   3816
   End
   Begin VB.TextBox txtCRL131 
      Height          =   285
      Left            =   2700
      MaxLength       =   20
      TabIndex        =   33
      Top             =   4800
      Width           =   1245
   End
   Begin VB.TextBox txtCRL130 
      Height          =   285
      Left            =   690
      MaxLength       =   20
      TabIndex        =   32
      Top             =   4800
      Width           =   1245
   End
   Begin VB.TextBox txtCRL129 
      Height          =   285
      Left            =   5040
      MaxLength       =   60
      TabIndex        =   26
      Top             =   3648
      Width           =   3816
   End
   Begin VB.TextBox txtCRL128 
      Height          =   285
      Left            =   2700
      MaxLength       =   20
      TabIndex        =   25
      Top             =   3648
      Width           =   1245
   End
   Begin VB.TextBox txtCRL127 
      Height          =   285
      Left            =   690
      MaxLength       =   20
      TabIndex        =   24
      Top             =   3648
      Width           =   1245
   End
   Begin VB.TextBox txtCRL126 
      Height          =   285
      Left            =   5040
      MaxLength       =   60
      TabIndex        =   18
      Top             =   2448
      Width           =   3816
   End
   Begin VB.TextBox txtCRL124 
      Height          =   285
      Left            =   2700
      MaxLength       =   20
      TabIndex        =   17
      Top             =   2448
      Width           =   1245
   End
   Begin VB.TextBox txtCRL117 
      Height          =   285
      Left            =   690
      MaxLength       =   20
      TabIndex        =   16
      Top             =   2448
      Width           =   1245
   End
   Begin VB.TextBox txtCRL116 
      Height          =   285
      Left            =   5040
      MaxLength       =   60
      TabIndex        =   10
      Top             =   1248
      Width           =   3816
   End
   Begin VB.TextBox txtCRL115 
      Height          =   285
      Left            =   2700
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1248
      Width           =   1245
   End
   Begin VB.TextBox txtCRL114 
      Height          =   285
      Left            =   690
      MaxLength       =   20
      TabIndex        =   8
      Top             =   1248
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   1620
      TabIndex        =   56
      Top             =   30
      Visible         =   0   'False
      Width           =   4884
      Begin VB.TextBox txtCRL03 
         Height          =   285
         Left            =   2910
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   2
         Top             =   90
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "frm090801_7.frx":0000
         Left            =   870
         List            =   "frm090801_7.frx":0010
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   10
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblCRL03 
         Height          =   225
         Left            =   3720
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label8 
         Caption         =   "智權人員:"
         Height          =   225
         Left            =   2100
         TabIndex        =   58
         Top             =   120
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label7 
         Caption         =   "收據公司:"
         Height          =   228
         Left            =   60
         TabIndex        =   57
         Top             =   36
         Visible         =   0   'False
         Width           =   828
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "同上"
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   30
      Top             =   4536
      Width           =   675
   End
   Begin VB.TextBox txtCRL111 
      Height          =   285
      Left            =   7584
      MaxLength       =   10
      TabIndex        =   28
      Top             =   3936
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Caption         =   "同上"
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   22
      Top             =   3348
      Width           =   675
   End
   Begin VB.TextBox txtCRL107 
      Height          =   285
      Left            =   7584
      MaxLength       =   10
      TabIndex        =   20
      Top             =   2736
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Caption         =   "同上"
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   14
      Top             =   2160
      Width           =   675
   End
   Begin VB.TextBox txtCRL103 
      Height          =   285
      Left            =   7584
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1548
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Caption         =   "同上"
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Top             =   948
      Width           =   675
   End
   Begin VB.TextBox txtCRL99 
      Height          =   285
      Left            =   7584
      MaxLength       =   10
      TabIndex        =   4
      Top             =   375
      Width           =   1275
   End
   Begin VB.TextBox txtCRL97 
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   0
      Top             =   10
      Width           =   435
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   37
      Top             =   0
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   6660
      TabIndex        =   36
      Top             =   0
      Width           =   930
   End
   Begin MSForms.ComboBox cboCRL110 
      Height          =   300
      Left            =   1080
      TabIndex        =   27
      Top             =   3936
      Width           =   5340
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "9419;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCRL106 
      Height          =   300
      Left            =   1080
      TabIndex        =   19
      Top             =   2736
      Width           =   5340
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "9419;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCRL102 
      Height          =   300
      Left            =   1080
      TabIndex        =   11
      Top             =   1548
      Width           =   5340
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "9419;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCRL98 
      Height          =   300
      Left            =   1080
      TabIndex        =   3
      Top             =   372
      Width           =   5340
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "9419;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCRL118 
      Height          =   648
      Left            =   24
      TabIndex        =   35
      Top             =   5352
      Width           =   8868
      VariousPropertyBits=   -1462747109
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "15642;1143"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCRL112 
      Height          =   300
      Left            =   1740
      TabIndex        =   31
      Top             =   4524
      Width           =   7092
      VariousPropertyBits=   675301403
      MaxLength       =   80
      Size            =   "12509;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCRL113 
      Height          =   300
      Left            =   1080
      TabIndex        =   29
      Top             =   4236
      Width           =   7776
      VariousPropertyBits=   675301403
      MaxLength       =   80
      Size            =   "13716;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCRL108 
      Height          =   300
      Left            =   1740
      TabIndex        =   23
      Top             =   3348
      Width           =   7092
      VariousPropertyBits=   675301403
      MaxLength       =   80
      Size            =   "12509;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCRL109 
      Height          =   300
      Left            =   1080
      TabIndex        =   21
      Top             =   3036
      Width           =   7776
      VariousPropertyBits=   675301403
      MaxLength       =   80
      Size            =   "13716;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCRL104 
      Height          =   300
      Left            =   1740
      TabIndex        =   15
      Top             =   2160
      Width           =   7092
      VariousPropertyBits=   675301403
      MaxLength       =   80
      Size            =   "12509;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCRL105 
      Height          =   300
      Left            =   1080
      TabIndex        =   13
      Top             =   1848
      Width           =   7776
      VariousPropertyBits=   675301403
      MaxLength       =   80
      Size            =   "13716;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCRL100 
      Height          =   300
      Left            =   1740
      TabIndex        =   7
      Top             =   948
      Width           =   7092
      VariousPropertyBits=   675301403
      MaxLength       =   80
      Size            =   "12509;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCRL101 
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Top             =   672
      Width           =   7776
      VariousPropertyBits=   675301403
      MaxLength       =   80
      Size            =   "13716;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      Caption         =   "特殊ID統 一編號輸8個0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   168
      Left            =   6600
      TabIndex        =   78
      Top             =   5148
      Width           =   2208
   End
   Begin VB.Label Label4 
      Caption         =   "E-Mail(財務):"
      Height          =   228
      Index           =   15
      Left            =   3960
      TabIndex        =   71
      Top             =   4848
      Width           =   1068
   End
   Begin VB.Label Label4 
      Caption         =   "傳真４:"
      Height          =   228
      Index           =   14
      Left            =   1980
      TabIndex        =   70
      Top             =   4848
      Width           =   708
   End
   Begin VB.Label Label4 
      Caption         =   "電話４:"
      Height          =   228
      Index           =   13
      Left            =   36
      TabIndex        =   69
      Top             =   4848
      Width           =   648
   End
   Begin VB.Label Label4 
      Caption         =   "E-Mail(財務):"
      Height          =   228
      Index           =   12
      Left            =   3960
      TabIndex        =   68
      Top             =   3684
      Width           =   1068
   End
   Begin VB.Label Label4 
      Caption         =   "傳真３:"
      Height          =   228
      Index           =   11
      Left            =   1980
      TabIndex        =   67
      Top             =   3684
      Width           =   708
   End
   Begin VB.Label Label4 
      Caption         =   "電話３:"
      Height          =   228
      Index           =   10
      Left            =   36
      TabIndex        =   66
      Top             =   3684
      Width           =   648
   End
   Begin VB.Label Label4 
      Caption         =   "E-Mail(財務):"
      Height          =   228
      Index           =   9
      Left            =   3960
      TabIndex        =   65
      Top             =   2496
      Width           =   1068
   End
   Begin VB.Label Label4 
      Caption         =   "傳真２:"
      Height          =   228
      Index           =   8
      Left            =   1980
      TabIndex        =   64
      Top             =   2496
      Width           =   708
   End
   Begin VB.Label Label4 
      Caption         =   "電話２:"
      Height          =   228
      Index           =   7
      Left            =   36
      TabIndex        =   63
      Top             =   2496
      Width           =   648
   End
   Begin VB.Label Label4 
      Caption         =   "E-Mail(財務):"
      Height          =   228
      Index           =   6
      Left            =   3960
      TabIndex        =   62
      Top             =   1272
      Width           =   1068
   End
   Begin VB.Label Label4 
      Caption         =   "傳真１:"
      Height          =   228
      Index           =   5
      Left            =   1980
      TabIndex        =   61
      Top             =   1272
      Width           =   708
   End
   Begin VB.Label Label4 
      Caption         =   "電話１:"
      Height          =   228
      Index           =   4
      Left            =   36
      TabIndex        =   60
      Top             =   1272
      Width           =   648
   End
   Begin VB.Label Label6 
      Caption         =   "收據案件性質或金額劃分或其他說明:（收據抬頭超過4個時，請在此欄加註！）"
      Height          =   168
      Left            =   36
      TabIndex        =   55
      Top             =   5148
      Width           =   6372
   End
   Begin VB.Label Label5 
      Caption         =   "郵寄地址４:"
      Height          =   228
      Index           =   3
      Left            =   36
      TabIndex        =   54
      Top             =   4584
      Width           =   1008
   End
   Begin VB.Label Label4 
      Caption         =   "營業地址４:"
      Height          =   228
      Index           =   3
      Left            =   36
      TabIndex        =   53
      Top             =   4272
      Width           =   1008
   End
   Begin VB.Label Label3 
      Caption         =   "統一編號４:"
      Height          =   228
      Index           =   3
      Left            =   6540
      TabIndex        =   52
      Top             =   3996
      Width           =   1008
   End
   Begin VB.Label Label2 
      Caption         =   "收據抬頭４:"
      Height          =   228
      Index           =   3
      Left            =   36
      TabIndex        =   51
      Top             =   3984
      Width           =   1008
   End
   Begin VB.Label Label5 
      Caption         =   "郵寄地址３:"
      Height          =   228
      Index           =   2
      Left            =   36
      TabIndex        =   50
      Top             =   3372
      Width           =   1008
   End
   Begin VB.Label Label4 
      Caption         =   "營業地址３:"
      Height          =   228
      Index           =   2
      Left            =   36
      TabIndex        =   49
      Top             =   3060
      Width           =   1008
   End
   Begin VB.Label Label3 
      Caption         =   "統一編號３:"
      Height          =   228
      Index           =   2
      Left            =   6540
      TabIndex        =   48
      Top             =   2796
      Width           =   1008
   End
   Begin VB.Label Label2 
      Caption         =   "收據抬頭３:"
      Height          =   228
      Index           =   2
      Left            =   36
      TabIndex        =   47
      Top             =   2796
      Width           =   1008
   End
   Begin VB.Label Label5 
      Caption         =   "郵寄地址２:"
      Height          =   228
      Index           =   1
      Left            =   36
      TabIndex        =   46
      Top             =   2172
      Width           =   1008
   End
   Begin VB.Label Label4 
      Caption         =   "營業地址２:"
      Height          =   228
      Index           =   1
      Left            =   36
      TabIndex        =   45
      Top             =   1872
      Width           =   1008
   End
   Begin VB.Label Label3 
      Caption         =   "統一編號２:"
      Height          =   228
      Index           =   1
      Left            =   6540
      TabIndex        =   44
      Top             =   1584
      Width           =   1008
   End
   Begin VB.Label Label2 
      Caption         =   "收據抬頭２:"
      Height          =   228
      Index           =   1
      Left            =   36
      TabIndex        =   43
      Top             =   1584
      Width           =   1008
   End
   Begin VB.Label Label5 
      Caption         =   "郵寄地址１:"
      Height          =   228
      Index           =   0
      Left            =   36
      TabIndex        =   42
      Top             =   984
      Width           =   1008
   End
   Begin VB.Label Label4 
      Caption         =   "營業地址１:"
      Height          =   228
      Index           =   0
      Left            =   36
      TabIndex        =   41
      Top             =   696
      Width           =   1008
   End
   Begin VB.Label Label3 
      Caption         =   "統一編號１:"
      Height          =   228
      Index           =   0
      Left            =   6540
      TabIndex        =   40
      Top             =   408
      Width           =   1008
   End
   Begin VB.Label Label2 
      Caption         =   "收據抬頭１:"
      Height          =   228
      Index           =   0
      Left            =   36
      TabIndex        =   39
      Top             =   408
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "拆收據張數:"
      Height          =   228
      Left            =   36
      TabIndex        =   38
      Top             =   84
      Width           =   1008
   End
End
Attribute VB_Name = "frm090801_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/22 改成Form2.0 (cboCRL98,txtCRL101,txtCRL101...)
'Create By Sindy 2014/1/29
Option Explicit

Public m_stCRL01 As String
Public m_stCRL97 As String, m_stCRL118 As String
Public m_stCRL98 As String, m_stCRL99 As String, m_stCRL100 As String, m_stCRL101 As String
Public m_stCRL102 As String, m_stCRL103 As String, m_stCRL104 As String, m_stCRL105 As String
Public m_stCRL106 As String, m_stCRL107 As String, m_stCRL108 As String, m_stCRL109 As String
Public m_stCRL110 As String, m_stCRL111 As String, m_stCRL112 As String, m_stCRL113 As String

'電話,傳真,E-Mail
Public m_stCRL114 As String, m_stCRL115 As String, m_stCRL116 As String
Public m_stCRL117 As String, m_stCRL124 As String, m_stCRL126 As String
Public m_stCRL127 As String, m_stCRL128 As String, m_stCRL129 As String
Public m_stCRL130 As String, m_stCRL131 As String, m_stCRL132 As String

Public m_stCRL120 As String, m_stCRL121 As String, m_stCRL122 As String, m_stCRL123 As String
Dim m_PrevForm As Form '前一畫面
Dim m_MousePointer As Integer
Dim m_bolActivated As Boolean
'Add by Amy 2016/09/01
Dim i As Integer
Public stCaseNo1 As String, stCaseNo2 As String, stCaseNo3 As String, stCaseNo4 As String
Dim bolfrm090801_1_Show As Boolean 'Add By Sindy 2023/10/12


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub Check1_Click(Index As Integer)
Dim strText As String
   If Check1(Index).Value = 1 Then
      If Index = 0 Then
         If txtCRL101.Enabled = True Then
            strText = txtCRL101.Text
         Else
            Check1(Index).Value = 0
            Exit Sub
         End If
      ElseIf Index = 1 Then
         If txtCRL105.Enabled = True Then
            strText = txtCRL105.Text
         Else
            Check1(Index).Value = 0
            Exit Sub
         End If
      ElseIf Index = 2 Then
         If txtCRL109.Enabled = True Then
            strText = txtCRL109.Text
         Else
            Check1(Index).Value = 0
            Exit Sub
         End If
      ElseIf Index = 3 Then
         If txtCRL113.Enabled = True Then
            strText = txtCRL113.Text
         Else
            Check1(Index).Value = 0
            Exit Sub
         End If
      End If
      If strText = "" Then
         MsgBox "請輸入營業地址" & IIf(Index = 0, "1", IIf(Index = 1, "2", IIf(Index = 2, "3", "4"))) & "！", vbExclamation
         Check1(Index).Value = 0
         If Index = 0 Then txtCRL101.SetFocus
         If Index = 1 Then txtCRL105.SetFocus
         If Index = 2 Then txtCRL109.SetFocus
         If Index = 3 Then txtCRL113.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Check2_Click(Index As Integer)
    Dim strTmp As String
    
    If Check2(Index).Value = 1 Then
        strTmp = "境外"
    Else
        strTmp = ""
    End If
    
    Select Case Index
        Case 0
            txtCRL99 = strTmp '統一編號
        Case 1
            txtCRL103 = strTmp
        Case 2
            txtCRL107 = strTmp
        Case 3
            txtCRL111 = strTmp
    End Select
End Sub


Private Sub cmdok_Click(Index As Integer)
   Dim bCancel As Boolean, iIdx As Integer
   
   PUB_FilterFormText Me 'Add By Sindy 2017/12/5 去掉跳行符號
   
   If Index = 0 Then '確定
   
      'Added by Morgan 2022/1/22 檢查畫面輸入欄位是否含有Unicode文字
      If PUB_ChkUniText(Me, , True) = False Then
          Exit Sub
      End If
      'end 2022/1/22
   
      If Trim(txtCRL97) = "" Then
         MsgBox "請輸入拆收據張數！", vbExclamation
         txtCRL97.SetFocus
         Exit Sub
      End If
      '收據抬頭至少輸入一項
      If Trim(cboCRL98) = "" And Trim(cboCRL102) = "" And Trim(cboCRL106) = "" And Trim(cboCRL110) = "" Then
         MsgBox "請至少輸入一項收據抬頭！", vbExclamation
         Exit Sub
      End If
      '檢查是否有依照順序輸入抬頭
      If (cboCRL98 = "" And (cboCRL102 <> "" Or cboCRL106 <> "" Or cboCRL110 <> "")) Or _
         (cboCRL102 = "" And (cboCRL106 <> "" Or cboCRL110 <> "")) Or _
         (cboCRL106 = "" And (cboCRL110 <> "")) Then
         MsgBox "請依序輸入收據抬頭資料！", vbExclamation
         Exit Sub
      End If
      '檢查是否重覆輸入
      If cboCRL110 <> "" And (cboCRL110 = cboCRL98 Or cboCRL110 = cboCRL102 Or cboCRL110 = cboCRL106) Then
         MsgBox "收據抬頭重覆輸入！", vbExclamation
         cboCRL110.SetFocus
         Exit Sub
      End If
      If cboCRL106 <> "" And (cboCRL106 = cboCRL98 Or cboCRL106 = cboCRL102 Or cboCRL106 = cboCRL110) Then
         MsgBox "收據抬頭重覆輸入！", vbExclamation
         cboCRL106.SetFocus
         Exit Sub
      End If
      If cboCRL102 <> "" And (cboCRL102 = cboCRL98 Or cboCRL102 = cboCRL106 Or cboCRL102 = cboCRL110) Then
         MsgBox "收據抬頭重覆輸入！", vbExclamation
         cboCRL102.SetFocus
         Exit Sub
      End If
      If cboCRL98 <> "" And (cboCRL98 = cboCRL102 Or cboCRL98 = cboCRL106 Or cboCRL98 = cboCRL110) Then
         MsgBox "收據抬頭重覆輸入！", vbExclamation
         cboCRL98.SetFocus
         Exit Sub
      End If
      '收據抬頭>=4碼時則統一編號欄必須有值
      '營業地址及郵寄地址不可空白
      If Trim(cboCRL98.Text) <> "" Then
         If Len(Trim(cboCRL98.Text)) >= 4 Then
            If txtCRL99.Enabled = True And Len(Trim(txtCRL99.Text)) = 0 Then
               MsgBox "請輸入統一編號1！", vbExclamation
               txtCRL99.SetFocus
               Exit Sub
            End If
            'Modify by Amy 2016/05/23 +統一編號欄位不是"境外"才檢查
            If txtCRL101.Enabled = True And Trim(txtCRL101.Text) = "" And Trim(txtCRL99.Text) <> "境外" Then
               MsgBox "請輸入營業地址1！", vbExclamation
               txtCRL101.SetFocus
               Exit Sub
            End If
            If txtCRL100.Enabled = True And Trim(txtCRL100.Text) = "" And Check1(0).Value = 0 Then
               MsgBox "請輸入郵寄地址1！", vbExclamation
               txtCRL100.SetFocus
               Exit Sub
            End If
         End If
         'Add By Sindy 2015/9/10
         If txtCRL114.Enabled = True And Len(Trim(txtCRL114.Text)) = 0 Then
            MsgBox "請輸入電話1！", vbExclamation
            txtCRL114.SetFocus
            Exit Sub
         End If
         '2015/9/10 END
      End If
      If Trim(cboCRL102.Text) <> "" Then
         If Len(Trim(cboCRL102.Text)) >= 4 Then
            If txtCRL103.Enabled = True And Len(Trim(txtCRL103.Text)) = 0 Then
               MsgBox "請輸入統一編號2！", vbExclamation
               txtCRL103.SetFocus
               Exit Sub
            End If
            'Modify by Amy 2016/05/23 +統一編號欄位不是"境外"才檢查
            If txtCRL105.Enabled = True And Trim(txtCRL105.Text) = "" And Trim(txtCRL103.Text) <> "境外" Then
               MsgBox "請輸入營業地址2！", vbExclamation
               txtCRL105.SetFocus
               Exit Sub
            End If
            If txtCRL104.Enabled = True And Trim(txtCRL104.Text) = "" And Check1(1).Value = 0 Then
               MsgBox "請輸入郵寄地址2！", vbExclamation
               txtCRL104.SetFocus
               Exit Sub
            End If
         End If
         'Add By Sindy 2015/9/10
         If txtCRL117.Enabled = True And Len(Trim(txtCRL117.Text)) = 0 Then
            MsgBox "請輸入電話2！", vbExclamation
            txtCRL117.SetFocus
            Exit Sub
         End If
         '2015/9/10 END
      End If
      If Trim(cboCRL106.Text) <> "" Then
         If Len(Trim(cboCRL106.Text)) >= 4 Then
            If txtCRL107.Enabled = True And Len(Trim(txtCRL107.Text)) = 0 Then
               MsgBox "請輸入統一編號3！", vbExclamation
               txtCRL107.SetFocus
               Exit Sub
            End If
            'Modify by Amy 2016/05/23 +統一編號欄位不是"境外"才檢查
            If txtCRL109.Enabled = True And Trim(txtCRL109.Text) = "" And Trim(txtCRL107.Text) <> "境外" Then
               MsgBox "請輸入營業地址3！", vbExclamation
               txtCRL109.SetFocus
               Exit Sub
            End If
            If txtCRL108.Enabled = True And Trim(txtCRL108.Text) = "" And Check1(2).Value = 0 Then
               MsgBox "請輸入郵寄地址3！", vbExclamation
               txtCRL108.SetFocus
               Exit Sub
            End If
         End If
         'Add By Sindy 2015/9/10
         If txtCRL127.Enabled = True And Len(Trim(txtCRL127.Text)) = 0 Then
            MsgBox "請輸入電話3！", vbExclamation
            txtCRL127.SetFocus
            Exit Sub
         End If
         '2015/9/10 END
      End If
      If Trim(cboCRL110.Text) <> "" Then
         If Len(Trim(cboCRL110.Text)) >= 4 Then
            If txtCRL111.Enabled = True And Len(Trim(txtCRL111.Text)) = 0 Then
               MsgBox "請輸入統一編號4！", vbExclamation
               txtCRL111.SetFocus
               Exit Sub
            End If
            'Modify by Amy 2016/05/23 +統一編號欄位不是"境外"才檢查
            If txtCRL113.Enabled = True And Trim(txtCRL113.Text) = "" And Trim(txtCRL111.Text) <> "境外" Then
               MsgBox "請輸入營業地址4！", vbExclamation
               txtCRL113.SetFocus
               Exit Sub
            End If
            If txtCRL112.Enabled = True And Trim(txtCRL112.Text) = "" And Check1(3).Value = 0 Then
               MsgBox "請輸入郵寄地址4！", vbExclamation
               txtCRL112.SetFocus
               Exit Sub
            End If
         End If
         'Add By Sindy 2015/9/10
         If txtCRL130.Enabled = True And Len(Trim(txtCRL130.Text)) = 0 Then
            MsgBox "請輸入電話4！", vbExclamation
            txtCRL130.SetFocus
            Exit Sub
         End If
         '2015/9/10 END
      End If
      '檢查各欄位資料輸入是否正確:
      cboCRL98_Validate bCancel
      If bCancel = True Then
         cboCRL98.SetFocus
         Exit Sub
      End If
      txtCRL99_Validate bCancel
      If bCancel = True Then
         txtCRL99.SetFocus
         Exit Sub
      End If
      txtCRL101_Validate bCancel
      If bCancel = True Then
         txtCRL101.SetFocus
         Exit Sub
      End If
      txtCRL100_Validate bCancel
      If bCancel = True Then
         txtCRL100.SetFocus
         Exit Sub
      End If
      cboCRL102_Validate bCancel
      If bCancel = True Then
         cboCRL102.SetFocus
         Exit Sub
      End If
      txtCRL103_Validate bCancel
      If bCancel = True Then
         txtCRL103.SetFocus
         Exit Sub
      End If
      txtCRL105_Validate bCancel
      If bCancel = True Then
         txtCRL105.SetFocus
         Exit Sub
      End If
      txtCRL104_Validate bCancel
      If bCancel = True Then
         txtCRL104.SetFocus
         Exit Sub
      End If
      cboCRL106_Validate bCancel
      If bCancel = True Then
         cboCRL106.SetFocus
         Exit Sub
      End If
      txtCRL107_Validate bCancel
      If bCancel = True Then
         txtCRL107.SetFocus
         Exit Sub
      End If
      txtCRL109_Validate bCancel
      If bCancel = True Then
         txtCRL109.SetFocus
         Exit Sub
      End If
      txtCRL108_Validate bCancel
      If bCancel = True Then
         txtCRL108.SetFocus
         Exit Sub
      End If
      cboCRL110_Validate bCancel
      If bCancel = True Then
         cboCRL110.SetFocus
         Exit Sub
      End If
      txtCRL111_Validate bCancel
      If bCancel = True Then
         txtCRL111.SetFocus
         Exit Sub
      End If
      txtCRL113_Validate bCancel
      If bCancel = True Then
         txtCRL113.SetFocus
         Exit Sub
      End If
      txtCRL112_Validate bCancel
      If bCancel = True Then
         txtCRL112.SetFocus
         Exit Sub
      End If
      txtCRL118_Validate bCancel
      If bCancel = True Then
         txtCRL118.SetFocus
         Exit Sub
      End If
      txtCRL116_Validate bCancel
      If bCancel = True Then
         txtCRL116.SetFocus
         Exit Sub
      End If
      txtCRL126_Validate bCancel
      If bCancel = True Then
         txtCRL126.SetFocus
         Exit Sub
      End If
      txtCRL129_Validate bCancel
      If bCancel = True Then
         txtCRL129.SetFocus
         Exit Sub
      End If
      txtCRL132_Validate bCancel
      If bCancel = True Then
         txtCRL132.SetFocus
         Exit Sub
      End If
      'Add by Amy 2015/10/23 +前畫面為接洽記錄單則需判特殊收據抬頭
      'Modify By Sindy 2022/9/1 + Or UCase(m_PrevForm.Name) = UCase("frm090801_new")
      If UCase(m_PrevForm.Name) = "FRM090801" Or UCase(m_PrevForm.Name) = UCase("frm090801_new") Then
            'Modify by Amy 2020/02/14 +未勾特殊收據才判斷
            If m_PrevForm.Check9.Value = 0 And m_PrevForm.Option2(1).Value = True And m_PrevForm.cboTitle <> cboCRL98 Then
                MsgBox "收據抬頭與接洽單收據抬頭不符，請確認！", vbExclamation
                cboCRL98.SetFocus
                Exit Sub
            End If
      End If
      'end 2015/10/23
      
      '填入變數值:
      m_PrevForm.m_stCRL97 = txtCRL97.Text
      m_PrevForm.m_stCRL98 = cboCRL98.Text
      m_PrevForm.m_stCRL99 = txtCRL99.Text
      '若勾選同上時,則預設營業地址至郵寄地址:
      '郵寄地址
      If Check1(0).Value = 1 Then
         m_PrevForm.m_stCRL100 = txtCRL101.Text
      Else
         m_PrevForm.m_stCRL100 = txtCRL100.Text
      End If
      m_PrevForm.m_stCRL101 = txtCRL101.Text
      m_PrevForm.m_stCRL102 = cboCRL102.Text
      m_PrevForm.m_stCRL103 = txtCRL103.Text
      '郵寄地址
      If Check1(1).Value = 1 Then
         m_PrevForm.m_stCRL104 = txtCRL105.Text
      Else
         m_PrevForm.m_stCRL104 = txtCRL104.Text
      End If
      m_PrevForm.m_stCRL105 = txtCRL105.Text
      m_PrevForm.m_stCRL106 = cboCRL106.Text
      m_PrevForm.m_stCRL107 = txtCRL107.Text
      '郵寄地址
      If Check1(2).Value = 1 Then
         m_PrevForm.m_stCRL108 = txtCRL109.Text
      Else
         m_PrevForm.m_stCRL108 = txtCRL108.Text
      End If
      m_PrevForm.m_stCRL109 = txtCRL109.Text
      m_PrevForm.m_stCRL110 = cboCRL110.Text
      m_PrevForm.m_stCRL111 = txtCRL111.Text
      '郵寄地址
      If Check1(3).Value = 1 Then
         m_PrevForm.m_stCRL112 = txtCRL113.Text
      Else
         m_PrevForm.m_stCRL112 = txtCRL112.Text
      End If
      m_PrevForm.m_stCRL113 = txtCRL113.Text
      m_PrevForm.m_stCRL114 = txtCRL114.Text
      m_PrevForm.m_stCRL115 = txtCRL115.Text
      m_PrevForm.m_stCRL116 = txtCRL116.Text
      m_PrevForm.m_stCRL117 = txtCRL117.Text
      m_PrevForm.m_stCRL118 = txtCRL118.Text
      m_PrevForm.m_stCRL120 = m_stCRL120
      m_PrevForm.m_stCRL121 = m_stCRL121
      m_PrevForm.m_stCRL122 = m_stCRL122
      m_PrevForm.m_stCRL123 = m_stCRL123
      m_PrevForm.m_stCRL124 = txtCRL124.Text
      m_PrevForm.m_stCRL126 = txtCRL126.Text
      m_PrevForm.m_stCRL127 = txtCRL127.Text
      m_PrevForm.m_stCRL128 = txtCRL128.Text
      m_PrevForm.m_stCRL129 = txtCRL129.Text
      m_PrevForm.m_stCRL130 = txtCRL130.Text
      m_PrevForm.m_stCRL131 = txtCRL131.Text
      m_PrevForm.m_stCRL132 = txtCRL132.Text
      'Exit For
   End If
   Screen.MousePointer = m_MousePointer
   Unload Me
End Sub

Private Sub Form_Activate()
Dim i As Integer
'Add by Amy 2016/09/01
Dim RsQ As New ADODB.Recordset
Dim strQ As String, strA0k02 As String, intQ As Integer, intReceipt As Integer
'Add by Amy 2017/03/30
Dim strCRL02 As String, strCRL120 As String, strCRL121 As String, strCRL122 As String, strCRL123 As String, strA4210(3) As String
Dim bolChkA4210(3) As Boolean
   
   If m_bolActivated = False Then
      If m_stCRL01 <> "" Then
         '讀取資料
         strSql = "SELECT * From consultrecordlist WHERE CRL01='" & m_stCRL01 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            Me.Caption = "特殊收據 (接洽單編號:" & m_stCRL01 & ")" & IIf(InStr(UCase(m_PrevForm.Name), UCase("Frmacc112")) > 0, IIf("" & RsTemp.Fields("crL119") = "Y", "有特殊收據", "無特殊收據"), "") 'Add By Sindy 2022/12/29
            strCRL02 = "" & RsTemp.Fields("crl02") 'Add by Amy 2017/03/30 CreateDate
            txtCRL97.Text = "" & RsTemp.Fields("crl97")
            cboCRL98.Text = "" & RsTemp.Fields("crl98") & IIf(cboCRL98.Text <> "" And "" & RsTemp.Fields("crl120") = "A", "(非客戶)", IIf(cboCRL98.Text <> "" And "" & RsTemp.Fields("crl120") = "", "(非客戶新抬頭)", ""))
            txtCRL99.Text = "" & RsTemp.Fields("crl99")
            txtCRL100.Text = "" & RsTemp.Fields("crl100")
            txtCRL101.Text = "" & RsTemp.Fields("crl101")
            cboCRL102.Text = "" & RsTemp.Fields("crl102") & IIf(cboCRL102.Text <> "" And "" & RsTemp.Fields("crl121") = "A", "(非客戶)", IIf(cboCRL102.Text <> "" And "" & RsTemp.Fields("crl121") = "", "(非客戶新抬頭)", ""))
            txtCRL103.Text = "" & RsTemp.Fields("crl103")
            txtCRL104.Text = "" & RsTemp.Fields("crl104")
            txtCRL105.Text = "" & RsTemp.Fields("crl105")
            cboCRL106.Text = "" & RsTemp.Fields("crl106") & IIf(cboCRL106.Text <> "" And "" & RsTemp.Fields("crl122") = "A", "(非客戶)", IIf(cboCRL106.Text <> "" And "" & RsTemp.Fields("crl122") = "", "(非客戶新抬頭)", ""))
            txtCRL107.Text = "" & RsTemp.Fields("crl107")
            txtCRL108.Text = "" & RsTemp.Fields("crl108")
            txtCRL109.Text = "" & RsTemp.Fields("crl109")
            cboCRL110.Text = "" & RsTemp.Fields("crl110") & IIf(cboCRL110.Text <> "" And "" & RsTemp.Fields("crl123") = "A", "(非客戶)", IIf(cboCRL110.Text <> "" And "" & RsTemp.Fields("crl123") = "", "(非客戶新抬頭)", ""))
            txtCRL111.Text = "" & RsTemp.Fields("crl111")
            txtCRL112.Text = "" & RsTemp.Fields("crl112")
            txtCRL113.Text = "" & RsTemp.Fields("crl113")
            txtCRL114.Text = "" & RsTemp.Fields("crl114")
            txtCRL115.Text = "" & RsTemp.Fields("crl115")
            txtCRL116.Text = "" & RsTemp.Fields("crl116")
            txtCRL117.Text = "" & RsTemp.Fields("crl117")
            txtCRL118.Text = "" & RsTemp.Fields("crl118")
            'Add by Amy 2017/03/30
            strCRL120 = "" & RsTemp.Fields("crl120")
            strCRL121 = "" & RsTemp.Fields("crl121")
            strCRL122 = "" & RsTemp.Fields("crl122")
            strCRL123 = "" & RsTemp.Fields("crl123")
            'end 2017/03/30
            txtCRL124.Text = "" & RsTemp.Fields("crl124")
            txtCRL126.Text = "" & RsTemp.Fields("crl126")
            txtCRL127.Text = "" & RsTemp.Fields("crl127")
            txtCRL128.Text = "" & RsTemp.Fields("crl128")
            txtCRL129.Text = "" & RsTemp.Fields("crl129")
            txtCRL130.Text = "" & RsTemp.Fields("crl130")
            txtCRL131.Text = "" & RsTemp.Fields("crl131")
            txtCRL132.Text = "" & RsTemp.Fields("crl132")
            txtCRL03.Text = "" & RsTemp.Fields("crl03"): lblCRL03.Caption = GetStaffName(txtCRL03.Text)
            Combo1.ListIndex = Val("" & RsTemp.Fields("crl49"))
         End If
         'Add By Sindy 2023/8/8
         If TypeName(m_PrevForm) = "frm090801_New" Then
            If m_PrevForm.cmdOK(0).Caption = "存檔" Or m_PrevForm.cmdOK(0).Caption = "重送" Then
               cmdOK(0).Visible = True
            Else
               cmdOK(0).Visible = False
            End If
         Else
         '2023/8/8 END
            cmdOK(0).Visible = False
         End If
         'Add By Sindy 2023/8/8
         If cmdOK(0).Visible = False Then
         '2023/8/8 END
            Frame1.Visible = True
            '鎖住各欄位
            txtCRL97.Locked = True
            cboCRL98.Locked = True
            txtCRL99.Locked = True
            txtCRL100.Locked = True
            txtCRL101.Locked = True
            cboCRL102.Locked = True
            txtCRL103.Locked = True
            txtCRL104.Locked = True
            txtCRL105.Locked = True
            cboCRL106.Locked = True
            txtCRL107.Locked = True
            txtCRL108.Locked = True
            txtCRL109.Locked = True
            cboCRL110.Locked = True
            txtCRL111.Locked = True
            txtCRL112.Locked = True
            txtCRL113.Locked = True
            txtCRL114.Locked = True
            txtCRL115.Locked = True
            txtCRL116.Locked = True
            txtCRL117.Locked = True
            txtCRL118.Locked = True
            txtCRL124.Locked = True
            txtCRL126.Locked = True
            txtCRL127.Locked = True
            txtCRL128.Locked = True
            txtCRL129.Locked = True
            txtCRL130.Locked = True
            txtCRL131.Locked = True
            txtCRL132.Locked = True
         End If
         'Add by Amy 2017/03/30 非客戶新抬頭標註顏色 ex:P-114070 開收據彈特殊收據畫面新抬頭要標註-瑞婷
         Label2(0).ForeColor = &H80000012
         Label2(1).ForeColor = &H80000012
         Label2(2).ForeColor = &H80000012
         Label2(3).ForeColor = &H80000012
         If cboCRL98 <> "" And strCRL120 <> "C" Then bolChkA4210(0) = ChkA4210(cboCRL98, strA4210(0))
         If cboCRL102 <> "" And strCRL121 <> "C" Then bolChkA4210(1) = ChkA4210(cboCRL102, strA4210(1))
         If cboCRL106 <> "" And strCRL122 <> "C" Then bolChkA4210(2) = ChkA4210(cboCRL106, strA4210(2))
         If cboCRL110 <> "" And strCRL123 <> "C" Then bolChkA4210(3) = ChkA4210(cboCRL110, strA4210(3))
         If bolChkA4210(0) = True And Val(strA4210(0)) + 19110000 >= Val(strCRL02) Then Label2(0).ForeColor = &HC0&
         If bolChkA4210(1) = True And Val(strA4210(1)) + 19110000 >= Val(strCRL02) Then Label2(1).ForeColor = &HC0&
         If bolChkA4210(2) = True And Val(strA4210(2)) + 19110000 >= Val(strCRL02) Then Label2(2).ForeColor = &HC0&
         If bolChkA4210(3) = True And Val(strA4210(3)) + 19110000 >= Val(strCRL02) Then Label2(3).ForeColor = &HC0&
         'end 2017/03/30
      Else
         '讀取下拉選單資料
         If m_PrevForm.cboTitle.ListCount <> 0 Then
            For i = 0 To m_PrevForm.cboTitle.ListCount - 1
               If Trim(m_PrevForm.cboTitle.List(i)) <> "" Then
                  cboCRL98.AddItem m_PrevForm.cboTitle.List(i)
                  cboCRL102.AddItem m_PrevForm.cboTitle.List(i)
                  cboCRL106.AddItem m_PrevForm.cboTitle.List(i)
                  cboCRL110.AddItem m_PrevForm.cboTitle.List(i)
               End If
            Next i
         End If
         txtCRL97.Text = m_stCRL97
         
        'Add by Amy 2016/09/01 '預帶上次開過的收據抬頭
         If m_stCRL97 = MsgText(601) Then
            If cboCRL98.Text <> "" Then strQ = "And a0k04<>'" & cboCRL98.Text & "' "
            strQ = "Select a0k04,A0k02 From acc0k0 Where a0k01 in " & _
                    "(Select a0j13 From acc0j0 where a0j02='" & stCaseNo1 & stCaseNo2 & stCaseNo3 & stCaseNo4 & "') " & _
                        strQ & " Order by a0k02 Desc,a0k01"
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
            If intQ = 1 And RsQ.RecordCount > 0 Then
               intReceipt = 1
               RsQ.MoveFirst
               strA0k02 = "" & RsQ.Fields("a0k02")
               For i = 0 To RsQ.RecordCount - 1
                   If strA0k02 <> "" & RsQ.Fields("a0k02") Then
                      Exit For
                    Else
                        Select Case i
                            Case 0
                               m_stCRL102 = RsQ.Fields("a0k04")
                            Case 1
                               m_stCRL106 = RsQ.Fields("a0k04")
                            Case 2
                               m_stCRL110 = RsQ.Fields("a0k04")
                        End Select
                        intReceipt = intReceipt + 1
                   End If
                   strA0k02 = "" & RsQ.Fields("a0k02")
                   RsQ.MoveNext
               Next i
               txtCRL97.Text = intReceipt
            End If
         End If
         'end 2016/09/01
        
        If m_stCRL98 = "" Then
            '預設
            cboCRL98.Text = m_PrevForm.cboTitle.Text
        Else
            cboCRL98.Text = m_stCRL98
        End If
        'Call cboCRL98_LostFocus
        If m_stCRL120 = "" Then
            txtCRL99.Text = m_stCRL99
            txtCRL100.Text = m_stCRL100
            txtCRL101.Text = m_stCRL101
            txtCRL114.Text = m_stCRL114
            txtCRL115.Text = m_stCRL115
            txtCRL116.Text = m_stCRL116
        End If
            
        cboCRL102.Text = m_stCRL102
        'Call cboCRL102_LostFocus
        If m_stCRL121 = "" Then
            txtCRL103.Text = m_stCRL103
            txtCRL104.Text = m_stCRL104
            txtCRL105.Text = m_stCRL105
            txtCRL117.Text = m_stCRL117
            txtCRL124.Text = m_stCRL124
            txtCRL126.Text = m_stCRL126
        End If
            
        cboCRL106.Text = m_stCRL106
        'Call cboCRL106_LostFocus
        If m_stCRL122 = "" Then
            txtCRL107.Text = m_stCRL107
            txtCRL108.Text = m_stCRL108
            txtCRL109.Text = m_stCRL109
            txtCRL127.Text = m_stCRL127
            txtCRL128.Text = m_stCRL128
            txtCRL129.Text = m_stCRL129
        End If
            
        cboCRL110.Text = m_stCRL110
        'Call cboCRL110_LostFocus
        If m_stCRL123 = "" Then
            txtCRL111.Text = m_stCRL111
            txtCRL112.Text = m_stCRL112
            txtCRL113.Text = m_stCRL113
            txtCRL130.Text = m_stCRL130
            txtCRL131.Text = m_stCRL131
            txtCRL132.Text = m_stCRL132
        End If
         
         txtCRL118.Text = m_stCRL118
         'Add By Sindy 2024/2/15
         If txtCRL97.Enabled = True Then
         '2024/2/15 END
            txtCRL97.SetFocus
         End If
         Frame1.Visible = False
         cmdOK(0).Visible = True
      End If
      m_bolActivated = True
   End If
End Sub

Private Sub Form_Load()
   'Modify By Sindy 2014/3/7 畫面不置中,改在最左上方
   'MoveFormToCenter Me
   Me.Move 0, 0
   '2014/3/7 END
   m_MousePointer = Screen.MousePointer
   'Screen.MousePointer = vbDefault
   cboCRL98.Clear
   cboCRL102.Clear
   cboCRL106.Clear
   cboCRL110.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2023/11/27
   If UCase(TypeName(m_PrevForm)) <> UCase("Nothing") Then
   '2023/11/27 END
      If UCase(m_PrevForm.Name) = UCase("Frmacc1121") Or _
         UCase(m_PrevForm.Name) = UCase("Frmacc1125") Then
         m_PrevForm.m_CallForm = ""
      'Modify By Sindy 2022/9/1 + Or UCase(m_PrevForm.Name) = UCase("frm090801_new")
      ElseIf UCase(m_PrevForm.Name) = UCase("frm090801") Or _
             UCase(m_PrevForm.Name) = UCase("frm090801_new") Then
         'Add by Amy 2017/03/30 拆收據張數與抬頭數不對不可繼續
         Dim intCRL97 As Integer
         If Trim(cboCRL98) <> "" Then intCRL97 = intCRL97 + 1
         If Trim(cboCRL102) <> "" Then intCRL97 = intCRL97 + 1
         If Trim(cboCRL106) <> "" Then intCRL97 = intCRL97 + 1
         If Trim(cboCRL110) <> "" Then intCRL97 = intCRL97 + 1
         If Val(txtCRL97) <> intCRL97 Then
           Cancel = True
           MsgBox "拆收據張數與抬頭數不合,請確認！"
           Exit Sub
         End If
         'Add by Amy 2020/02/14 勾選特殊收據且有輸抬頭,則清空抬頭-瑞婷
         If m_PrevForm.Check9.Value = 1 And m_PrevForm.Option2(1).Value = True And m_PrevForm.cboTitle <> MsgText(601) Then
              m_PrevForm.bolNotClsVal = True
              m_PrevForm.cboTitle = ""
              m_PrevForm.bolNotClsVal = False
         End If
         'm_PrevForm.Show vbModal
      End If
   End If
   'Add by Amy 2016/09/01
   stCaseNo1 = "": stCaseNo2 = "": stCaseNo3 = "": stCaseNo4 = ""
   Set frm090801_7 = Nothing
End Sub

Private Sub ReadCust(Index As Integer)
Dim strKey As String
Dim m_blnOneRec As Boolean
Dim m_strCustCode As String
   
   If bolfrm090801_1_Show = True Then Exit Sub 'Add By Sindy 2023/10/12
   Select Case Index
      Case 1
         If cboCRL98.Locked = True Then Exit Sub
         strKey = cboCRL98.Text
         If cboCRL98.Tag <> cboCRL98.Text Then
            m_stCRL120 = ""
            Check1(0).Value = 0
            txtCRL99.Text = ""
            txtCRL100.Text = ""
            txtCRL101.Text = ""
            txtCRL114.Text = ""
            txtCRL115.Text = ""
            txtCRL116.Text = ""
         End If
         If Trim(strKey) = "" Then
            txtCRL99.Enabled = False
            txtCRL100.Enabled = False
            txtCRL101.Enabled = False
            txtCRL114.Enabled = False
            txtCRL115.Enabled = False
            txtCRL116.Enabled = False
         Else
            txtCRL99.Enabled = True
            txtCRL100.Enabled = True
            txtCRL101.Enabled = True
            txtCRL114.Enabled = True
            txtCRL115.Enabled = True
            txtCRL116.Enabled = True
         End If
         cboCRL98.Tag = cboCRL98.Text
      Case 2
         If cboCRL102.Locked = True Then Exit Sub
         strKey = cboCRL102.Text
         If cboCRL102.Tag <> cboCRL102.Text Then
            m_stCRL121 = ""
            Check1(1).Value = 0
            txtCRL103.Text = ""
            txtCRL104.Text = ""
            txtCRL105.Text = ""
            txtCRL117.Text = ""
            txtCRL124.Text = ""
            txtCRL126.Text = ""
         End If
         If Trim(strKey) = "" Then
            txtCRL103.Enabled = False
            txtCRL104.Enabled = False
            txtCRL105.Enabled = False
            txtCRL117.Enabled = False
            txtCRL124.Enabled = False
            txtCRL126.Enabled = False
         Else
            txtCRL103.Enabled = True
            txtCRL104.Enabled = True
            txtCRL105.Enabled = True
            txtCRL117.Enabled = True
            txtCRL124.Enabled = True
            txtCRL126.Enabled = True
         End If
         cboCRL102.Tag = cboCRL102.Text
      Case 3
         If cboCRL106.Locked = True Then Exit Sub
         strKey = cboCRL106.Text
         If cboCRL106.Tag <> cboCRL106.Text Then
            m_stCRL122 = ""
            Check1(2).Value = 0
            txtCRL107.Text = ""
            txtCRL108.Text = ""
            txtCRL109.Text = ""
            txtCRL127.Text = ""
            txtCRL128.Text = ""
            txtCRL129.Text = ""
         End If
         If Trim(strKey) = "" Then
            txtCRL107.Enabled = False
            txtCRL108.Enabled = False
            txtCRL109.Enabled = False
            txtCRL127.Enabled = False
            txtCRL128.Enabled = False
            txtCRL129.Enabled = False
         Else
            txtCRL107.Enabled = True
            txtCRL108.Enabled = True
            txtCRL109.Enabled = True
            txtCRL127.Enabled = True
            txtCRL128.Enabled = True
            txtCRL129.Enabled = True
         End If
         cboCRL106.Tag = cboCRL106.Text
      Case 4
         If cboCRL110.Locked = True Then Exit Sub
         strKey = cboCRL110.Text
         If cboCRL110.Tag <> cboCRL110.Text Then
            m_stCRL123 = ""
            Check1(3).Value = 0
            txtCRL111.Text = ""
            txtCRL112.Text = ""
            txtCRL113.Text = ""
            txtCRL130.Text = ""
            txtCRL131.Text = ""
            txtCRL132.Text = ""
         End If
         If Trim(strKey) = "" Then
            txtCRL111.Enabled = False
            txtCRL112.Enabled = False
            txtCRL113.Enabled = False
            txtCRL130.Enabled = False
            txtCRL131.Enabled = False
            txtCRL132.Enabled = False
         Else
            txtCRL111.Enabled = True
            txtCRL112.Enabled = True
            txtCRL113.Enabled = True
            txtCRL130.Enabled = True
            txtCRL131.Enabled = True
            txtCRL132.Enabled = True
         End If
         cboCRL110.Tag = cboCRL110.Text
   End Select
   If Trim(strKey) = "" Then
      Exit Sub
   'Add By Sindy 2017/2/14
   Else
      strKey = Trim(strKey)
   End If
   '2017/2/14 END
   
'   If m_bolActivated = True Then
      'Modify By Sindy 2014/3/25 若A4202='04150022'不顯示到畫面上,並控制不必再請USER補輸
      'Modify By Sindy 2014/8/11 +and (cu80 is null or cu80='其他' or cu80='業務自行處理') and cu02='0'
      'Modify By Sindy 2017/4/6 + 再加英文名稱和日文名稱做比對,及全部轉換大寫比對
      'Modify By Sindy 2023/2/24 客戶狀態CU80，除了空白及其他、業務自行處理外，
      '                          要再加不得代理專利、不得代理商標、解除對造、國內同業；
      'Modify By Sindy 2025/6/27 +or cu80='其他' or cu80='業務自行處理' or cu80='不得代理專利' or cu80='不得代理商標' or cu80='解除對造' or cu80='國內同業'
      '                          改抓常變數
      strSql = "SELECT '' As V,st02 As 智權人員,CU04 As 收據抬頭,CU11 As 統一編號,CU23 As 營業地址,CU31 As 郵寄地址" & _
               ",CU16 As 電話,CU18 As 傳真,CU115 As 財務Mail,1 as sort" & _
               " From Customer,staff" & _
               " WHERE (upper(cu04)=upper('" & ChgSQL(strKey) & "') or upper(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90))=upper('" & ChgSQL(strKey) & "') or upper(cu06)=upper('" & ChgSQL(strKey) & "'))" & _
               " and (cu80 is null or instr('" & 客戶及代理人可讀取的狀態 & "',cu80)>0) and cu02='0'" & _
               " and cu13=st01(+)" & _
               " union SELECT '' As V,st02 As 智權人員,a4201 As 收據抬頭,a4202 As 統一編號,a4215 As 營業地址,a4203 As 郵寄地址" & _
               ",a4204 As 電話,a4205 As 傳真,a4218 As 財務Mail,2 as sort" & _
               " From Acc420,staff" & _
               " WHERE upper(a4201)=upper('" & ChgSQL(strKey) & "')" & _
               " and a4206=st01(+)" & _
               " order by sort asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Add By Sindy 2023/8/7 若大於1筆,開視窗讓人員選擇要帶入那一筆資料 ex:楊惠如有 2 筆(張宜萱的客戶)
         If RsTemp.RecordCount > 1 Then
            frm090801_1.m_Type = "4"
            frm090801_1.Caption = "收據抬頭資料查詢"
            frm090801_1.Label1(1).Caption = "收據抬頭："
            frm090801_1.m_strCustChnName = strKey
            frm090801_1.lblName.Caption = frm090801_1.m_strCustChnName
            m_blnOneRec = False
            m_strCustCode = ""
            Call frm090801_1.StrMenu2(RsTemp)
            If frm090801_1.m_blnOneRec = False Then
               bolfrm090801_1_Show = True 'Add By Sindy 2023/10/12
               frm090801_1.Show vbModal
            End If
            m_blnOneRec = frm090801_1.m_blnOneRec
            m_strCustCode = frm090801_1.m_strCustCode
            Unload frm090801_1
            If m_blnOneRec = True And Val(m_strCustCode) > 0 Then
               '移動至欲查詢出來的資料列
               RsTemp.MoveFirst
               For i = 2 To Val(m_strCustCode)
                  RsTemp.MoveNext
               Next
            End If
         Else
            m_blnOneRec = True
         End If
         If m_blnOneRec = False Then Exit Sub
         '2023/8/7 END
         Select Case Index
            Case 1
               m_stCRL120 = IIf(RsTemp.Fields("sort") = 1, "C", "A")
               txtCRL99.Enabled = False
               txtCRL100.Enabled = False
               txtCRL101.Enabled = False
               txtCRL114.Enabled = False
               txtCRL115.Enabled = False
               txtCRL116.Enabled = False
               'Modify By Sindy 2014/3/25 統一編號
               If "" & RsTemp.Fields("統一編號") = "04150022" Then
                  txtCRL99.Enabled = False
               Else
               '2014/3/25 END
                  txtCRL99.Text = "" & RsTemp.Fields("統一編號")
               End If
               txtCRL100.Text = "" & RsTemp.Fields("郵寄地址") '郵寄地址
               txtCRL101.Text = "" & RsTemp.Fields("營業地址") '營業地址
               txtCRL114.Text = "" & RsTemp.Fields("電話") '電話1
               txtCRL115.Text = "" & RsTemp.Fields("傳真") '傳真1
               txtCRL116.Text = "" & RsTemp.Fields("財務Mail") 'E-Mail(財務)
            Case 2
               m_stCRL121 = IIf(RsTemp.Fields("sort") = 1, "C", "A")
               txtCRL103.Enabled = False
               txtCRL104.Enabled = False
               txtCRL105.Enabled = False
               txtCRL117.Enabled = False
               txtCRL124.Enabled = False
               txtCRL126.Enabled = False
               'Modify By Sindy 2014/3/25 統一編號
               If "" & RsTemp.Fields("統一編號") = "04150022" Then
                  txtCRL103.Enabled = False
               Else
               '2014/3/25 END
                  txtCRL103.Text = "" & RsTemp.Fields("統一編號")
               End If
               txtCRL104.Text = "" & RsTemp.Fields("郵寄地址") '郵寄地址
               txtCRL105.Text = "" & RsTemp.Fields("營業地址") '營業地址
               txtCRL117.Text = "" & RsTemp.Fields("電話") '電話1
               txtCRL124.Text = "" & RsTemp.Fields("傳真") '傳真1
               txtCRL126.Text = "" & RsTemp.Fields("財務Mail") 'E-Mail(財務)
            Case 3
               m_stCRL122 = IIf(RsTemp.Fields("sort") = 1, "C", "A")
               txtCRL107.Enabled = False
               txtCRL108.Enabled = False
               txtCRL109.Enabled = False
               txtCRL127.Enabled = False
               txtCRL128.Enabled = False
               txtCRL129.Enabled = False
               'Modify By Sindy 2014/3/25 統一編號
               If "" & RsTemp.Fields("統一編號") = "04150022" Then
                  txtCRL107.Enabled = False
               Else
               '2014/3/25 END
                  txtCRL107.Text = "" & RsTemp.Fields("統一編號")
               End If
               txtCRL108.Text = "" & RsTemp.Fields("郵寄地址") '郵寄地址
               txtCRL109.Text = "" & RsTemp.Fields("營業地址") '營業地址
               txtCRL127.Text = "" & RsTemp.Fields("電話") '電話1
               txtCRL128.Text = "" & RsTemp.Fields("傳真") '傳真1
               txtCRL129.Text = "" & RsTemp.Fields("財務Mail") 'E-Mail(財務)
            Case 4
               m_stCRL123 = IIf(RsTemp.Fields("sort") = 1, "C", "A")
               txtCRL111.Enabled = False
               txtCRL112.Enabled = False
               txtCRL113.Enabled = False
               txtCRL130.Enabled = False
               txtCRL131.Enabled = False
               txtCRL132.Enabled = False
               'Modify By Sindy 2014/3/25 統一編號
               If "" & RsTemp.Fields("統一編號") = "04150022" Then
                  txtCRL111.Enabled = False
               Else
               '2014/3/25 END
                  txtCRL111.Text = "" & RsTemp.Fields("統一編號")
               End If
               txtCRL112.Text = "" & RsTemp.Fields("郵寄地址") '郵寄地址
               txtCRL113.Text = "" & RsTemp.Fields("營業地址") '營業地址
               txtCRL130.Text = "" & RsTemp.Fields("電話") '電話1
               txtCRL131.Text = "" & RsTemp.Fields("傳真") '傳真1
               txtCRL132.Text = "" & RsTemp.Fields("財務Mail") 'E-Mail(財務)
         End Select
      End If
'   End If
   bolfrm090801_1_Show = False 'Add By Sindy 2023/10/12
End Sub

Private Sub txtCRL100_GotFocus()
   OpenIme
   TextInverse txtCRL100
End Sub

Private Sub txtCRL100_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, txtCRL100)
End Sub

Private Sub txtCRL100_Validate(Cancel As Boolean)
   If txtCRL100.Enabled = False Then Exit Sub

   If txtCRL100.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL100, txtCRL100.MaxLength) Then
      Call txtCRL100_GotFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL104_GotFocus()
   OpenIme
   TextInverse txtCRL104
End Sub

Private Sub txtCRL104_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, txtCRL104)
End Sub

Private Sub txtCRL104_Validate(Cancel As Boolean)
   If txtCRL104.Enabled = False Then Exit Sub

   If txtCRL104.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL104, txtCRL104.MaxLength) Then
      Call txtCRL104_GotFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL108_GotFocus()
   OpenIme
   TextInverse txtCRL108
End Sub

Private Sub txtCRL108_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, txtCRL108)
End Sub

Private Sub txtCRL108_Validate(Cancel As Boolean)
   If txtCRL108.Enabled = False Then Exit Sub

   If txtCRL108.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL108, txtCRL108.MaxLength) Then
      Call txtCRL108_GotFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL112_GotFocus()
   OpenIme
   TextInverse txtCRL112
End Sub

Private Sub txtCRL112_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, txtCRL112)
End Sub

Private Sub txtCRL112_Validate(Cancel As Boolean)
   If txtCRL112.Enabled = False Then Exit Sub

   If txtCRL112.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL112, txtCRL112.MaxLength) Then
      Call txtCRL112_GotFocus
      Cancel = True
   End If
End Sub

'Add By Sindy 2015/8/28
Private Sub txtCRL114_GotFocus()
   CloseIme
   TextInverse txtCRL114
End Sub
Private Sub txtCRL115_GotFocus()
   CloseIme
   TextInverse txtCRL115
End Sub
Private Sub txtCRL116_GotFocus()
   CloseIme
   TextInverse txtCRL116
End Sub
Private Sub txtCRL116_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Email輸入字元檢查
End Sub
Private Sub txtCRL116_Validate(Cancel As Boolean)
   If txtCRL116.Enabled = False Then Exit Sub
   
   If txtCRL116.Text = "" Then Exit Sub
   Cancel = Not PUB_CheckMail(txtCRL116.Text)
End Sub

Private Sub txtCRL117_GotFocus()
   CloseIme
   TextInverse txtCRL117
End Sub
Private Sub txtCRL124_GotFocus()
   CloseIme
   TextInverse txtCRL124
End Sub
Private Sub txtCRL126_GotFocus()
   CloseIme
   TextInverse txtCRL126
End Sub
Private Sub txtCRL126_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Email輸入字元檢查
End Sub
Private Sub txtCRL126_Validate(Cancel As Boolean)
   If txtCRL126.Enabled = False Then Exit Sub
   
   If txtCRL126.Text = "" Then Exit Sub
   Cancel = Not PUB_CheckMail(txtCRL126.Text)
End Sub

Private Sub txtCRL127_GotFocus()
   CloseIme
   TextInverse txtCRL127
End Sub
Private Sub txtCRL128_GotFocus()
   CloseIme
   TextInverse txtCRL128
End Sub
Private Sub txtCRL129_GotFocus()
   CloseIme
   TextInverse txtCRL129
End Sub
Private Sub txtCRL129_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Email輸入字元檢查
End Sub
Private Sub txtCRL129_Validate(Cancel As Boolean)
   If txtCRL129.Enabled = False Then Exit Sub
   
   If txtCRL129.Text = "" Then Exit Sub
   Cancel = Not PUB_CheckMail(txtCRL129.Text)
End Sub

Private Sub txtCRL130_GotFocus()
   CloseIme
   TextInverse txtCRL130
End Sub
Private Sub txtCRL131_GotFocus()
   CloseIme
   TextInverse txtCRL131
End Sub
Private Sub txtCRL132_GotFocus()
   CloseIme
   TextInverse txtCRL132
End Sub
Private Sub txtCRL132_KeyPress(KeyAscii As Integer)
   PUB_EMailFilter KeyAscii 'Email輸入字元檢查
End Sub
Private Sub txtCRL132_Validate(Cancel As Boolean)
   If txtCRL132.Enabled = False Then Exit Sub
   
   If txtCRL132.Text = "" Then Exit Sub
   Cancel = Not PUB_CheckMail(txtCRL132.Text)
End Sub
'2015/8/28 END

Private Sub txtCRL101_GotFocus()
   OpenIme
   TextInverse txtCRL101
End Sub

Private Sub txtCRL101_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, txtCRL101)
End Sub

Private Sub txtCRL101_Validate(Cancel As Boolean)
   If txtCRL101.Enabled = False Then Exit Sub

   If txtCRL101.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL101, txtCRL101.MaxLength) Then
      Call txtCRL101_GotFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL105_GotFocus()
   OpenIme
   TextInverse txtCRL105
End Sub

Private Sub txtCRL105_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, txtCRL105)
End Sub

Private Sub txtCRL105_Validate(Cancel As Boolean)
   If txtCRL105.Enabled = False Then Exit Sub

   If txtCRL105.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL105, txtCRL105.MaxLength) Then
      Call txtCRL105_GotFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL109_GotFocus()
   OpenIme
   TextInverse txtCRL109
End Sub

Private Sub txtCRL109_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, txtCRL109)
End Sub

Private Sub txtCRL109_Validate(Cancel As Boolean)
   If txtCRL109.Enabled = False Then Exit Sub

   If txtCRL109.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL109, txtCRL109.MaxLength) Then
      Call txtCRL109_GotFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL113_GotFocus()
   OpenIme
   TextInverse txtCRL113
End Sub

Private Sub txtCRL113_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, txtCRL113)
End Sub

Private Sub txtCRL113_Validate(Cancel As Boolean)
   If txtCRL113.Enabled = False Then Exit Sub

   If txtCRL113.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL113, txtCRL113.MaxLength) Then
      Call txtCRL113_GotFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL118_GotFocus()
   OpenIme
   TextInverse txtCRL118
End Sub

Private Sub txtCRL118_Validate(Cancel As Boolean)
   If txtCRL118.Enabled = False Then Exit Sub

   If txtCRL118.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(txtCRL118, txtCRL118.MaxLength) Then
      Call txtCRL118_GotFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL97_GotFocus()
   TextInverse txtCRL97
   CloseIme
End Sub

Private Sub txtCRL97_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub cboCRL98_Click()
   Call cboCRL98_LostFocus
End Sub

Private Sub cboCRL98_LostFocus()
   Call ReadCust(1)
End Sub

Private Sub cboCRL98_Validate(Cancel As Boolean)
   If cboCRL98.Enabled = False Then Exit Sub
   
   If cboCRL98.Text = "" Then
      m_stCRL120 = ""
      Check1(0).Value = 0
      txtCRL99.Text = ""
      txtCRL101.Text = ""
      txtCRL100.Text = ""
      txtCRL114.Text = ""
      txtCRL115.Text = ""
      txtCRL116.Text = ""
      Exit Sub
   End If
   If Not CheckLengthIsOK(cboCRL98, 100) Then
      cboCRL98.SetFocus
      Cancel = True
   End If
End Sub

Private Sub cboCRL102_Click()
   Call cboCRL102_LostFocus
End Sub

Private Sub cboCRL102_LostFocus()
   Call ReadCust(2)
End Sub

Private Sub cboCRL102_Validate(Cancel As Boolean)
   If cboCRL102.Enabled = False Then Exit Sub
   
   If cboCRL102.Text = "" Then
      m_stCRL121 = ""
      Check1(1).Value = 0
      txtCRL103.Text = ""
      txtCRL105.Text = ""
      txtCRL104.Text = ""
      txtCRL117.Text = ""
      txtCRL124.Text = ""
      txtCRL126.Text = ""
      Exit Sub
   End If
   If Not CheckLengthIsOK(cboCRL102, 100) Then
      cboCRL102.SetFocus
      Cancel = True
   End If
End Sub

Private Sub cboCRL106_Click()
   Call cboCRL106_LostFocus
End Sub

Private Sub cboCRL106_LostFocus()
   Call ReadCust(3)
End Sub

Private Sub cboCRL106_Validate(Cancel As Boolean)
   If cboCRL106.Enabled = False Then Exit Sub
   
   If cboCRL106.Text = "" Then
      m_stCRL122 = ""
      Check1(2).Value = 0
      txtCRL107.Text = ""
      txtCRL109.Text = ""
      txtCRL108.Text = ""
      txtCRL127.Text = ""
      txtCRL128.Text = ""
      txtCRL129.Text = ""
      Exit Sub
   End If
   If Not CheckLengthIsOK(cboCRL106, 100) Then
      cboCRL106.SetFocus
      Cancel = True
   End If
End Sub

Private Sub cboCRL110_Click()
   Call cboCRL110_LostFocus
End Sub

Private Sub cboCRL110_LostFocus()
   Call ReadCust(4)
End Sub

Private Sub cboCRL110_Validate(Cancel As Boolean)
   If cboCRL110.Enabled = False Then Exit Sub
   
   If cboCRL110.Text = "" Then
      m_stCRL123 = ""
      Check1(3).Value = 0
      txtCRL111.Text = ""
      txtCRL113.Text = ""
      txtCRL112.Text = ""
      txtCRL130.Text = ""
      txtCRL131.Text = ""
      txtCRL132.Text = ""
      Exit Sub
   End If
   If Not CheckLengthIsOK(cboCRL110, 100) Then
      cboCRL110.SetFocus
      Cancel = True
   End If
End Sub

Private Sub txtCRL99_GotFocus()
   TextInverse txtCRL99
   CloseIme
End Sub

Private Sub txtCRL99_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCRL99_Validate(Cancel As Boolean)
Dim strTmp As String
   
   If txtCRL99.Enabled = False Then Exit Sub

   If txtCRL99.Text = "" Then Exit Sub
   If Trim(txtCRL99) = "境外" Then Exit Sub 'Add by Amy 2016/05/23
   
   txtCRL99.Text = Trim(PUB_StringFilter(txtCRL99.Text)) 'Add By Sindy 2014/4/11 瑞婷反應智權同仁在複製貼上時多貼到空白格
   If GetTextLength(txtCRL99.Text) <> 8 Then
      Call txtCRL99_GotFocus
      strTmp = "統編必須是8碼 ! 請確定 ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckID(1, txtCRL99.Text) = False Then
      Call txtCRL99_GotFocus
      strTmp = "統一編號錯誤，是否確定 ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtCRL103_GotFocus()
   TextInverse txtCRL103
   CloseIme
End Sub

Private Sub txtCRL103_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCRL103_Validate(Cancel As Boolean)
Dim strTmp As String
   
   If txtCRL103.Enabled = False Then Exit Sub

   If txtCRL103.Text = "" Then Exit Sub
   If Trim(txtCRL103) = "境外" Then Exit Sub 'Add by Amy 2016/05/23
   
   txtCRL103.Text = Trim(PUB_StringFilter(txtCRL103.Text)) 'Add By Sindy 2014/4/11 瑞婷反應智權同仁在複製貼上時多貼到空白格
   If GetTextLength(txtCRL103.Text) <> 8 Then
      Call txtCRL103_GotFocus
      strTmp = "統編必須是8碼 ! 請確定 ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckID(1, txtCRL103.Text) = False Then
      Call txtCRL103_GotFocus
      strTmp = "統一編號錯誤，是否確定 ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtCRL107_GotFocus()
   TextInverse txtCRL107
   CloseIme
End Sub

Private Sub txtCRL107_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCRL107_Validate(Cancel As Boolean)
Dim strTmp As String
   
   If txtCRL107.Enabled = False Then Exit Sub

   If txtCRL107.Text = "" Then Exit Sub
   If Trim(txtCRL107) = "境外" Then Exit Sub 'Add by Amy 2016/05/23
   
   txtCRL107.Text = Trim(PUB_StringFilter(txtCRL107.Text)) 'Add By Sindy 2014/4/11 瑞婷反應智權同仁在複製貼上時多貼到空白格
   If GetTextLength(txtCRL107.Text) <> 8 Then
      Call txtCRL107_GotFocus
      strTmp = "統編必須是8碼 ! 請確定 ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckID(1, txtCRL107.Text) = False Then
      Call txtCRL107_GotFocus
      strTmp = "統一編號錯誤，是否確定 ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

Private Sub txtCRL111_GotFocus()
   TextInverse txtCRL111
   CloseIme
End Sub

Private Sub txtCRL111_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCRL111_Validate(Cancel As Boolean)
Dim strTmp As String
   
   If txtCRL111.Enabled = False Then Exit Sub

   If txtCRL111.Text = "" Then Exit Sub
   If Trim(txtCRL111) = "境外" Then Exit Sub 'Add by Amy 2016/05/23
   
   txtCRL111.Text = Trim(PUB_StringFilter(txtCRL111.Text)) 'Add By Sindy 2014/4/11 瑞婷反應智權同仁在複製貼上時多貼到空白格
   If GetTextLength(txtCRL111.Text) <> 8 Then
      Call txtCRL111_GotFocus
      strTmp = "統編必須是8碼 ! 請確定 ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
         Exit Sub
      End If
   End If
   If CheckID(1, txtCRL111.Text) = False Then
      Call txtCRL111_GotFocus
      strTmp = "統一編號錯誤，是否確定 ?"
      If MsgBox(strTmp, vbYesNo + vbCritical) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

Private Function ChkA4210(ByVal strA4201 As String, ByRef strA4210 As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    ChkA4210 = False: strA4210 = ""
    'Modify By Sindy 2025/6/9 + ChgSQL ex:Sophie's Bionutrients Pte. Ltd.
    strQ = "Select  A4210 From Acc420 Where A4201='" & ChgSQL(strA4201) & "'"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        strA4210 = "" & RsQ.Fields("A4210")
        ChkA4210 = True
    End If
    RsQ.Close
End Function
