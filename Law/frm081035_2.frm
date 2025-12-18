VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081035_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文"
   ClientHeight    =   6324
   ClientLeft      =   5160
   ClientTop       =   972
   ClientWidth     =   9060
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6324
   ScaleWidth      =   9060
   Begin VB.TextBox textCP60 
      Height          =   300
      Left            =   7104
      TabIndex        =   58
      Top             =   1080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7020
      TabIndex        =   16
      Top             =   30
      Width           =   1100
   End
   Begin VB.CommandButton CmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6192
      TabIndex        =   15
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton CmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8148
      TabIndex        =   17
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton Command1 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   5064
      TabIndex        =   14
      Top             =   30
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   4176
      Left            =   168
      TabIndex        =   32
      Top             =   2112
      Width           =   8784
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   684
         Left            =   48
         TabIndex        =   47
         Top             =   1536
         Width           =   8556
         Begin VB.TextBox textAmt 
            Height          =   300
            Left            =   7752
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   11
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtMailC 
            Height          =   300
            Left            =   7080
            MaxLength       =   1
            TabIndex        =   9
            Top             =   24
            Width           =   375
         End
         Begin VB.TextBox txtMailF 
            Height          =   300
            Left            =   3648
            MaxLength       =   1
            TabIndex        =   8
            Top             =   24
            Width           =   375
         End
         Begin VB.TextBox textCP144 
            Height          =   300
            Left            =   1008
            MaxLength       =   8
            TabIndex        =   7
            Top             =   24
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "補收款："
            Height          =   240
            Left            =   6984
            TabIndex        =   55
            Top             =   384
            Width           =   732
         End
         Begin VB.Label Label9 
            Caption         =   "(請款項目會帶入定稿和Email內文)"
            Height          =   240
            Left            =   4008
            TabIndex        =   54
            Top             =   384
            Width           =   2748
         End
         Begin MSForms.TextBox textItem 
            Height          =   300
            Left            =   1008
            TabIndex        =   10
            Top             =   360
            Width           =   2976
            VariousPropertyBits=   671105051
            BackColor       =   16777215
            MaxLength       =   30
            Size            =   "5249;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label4 
            Caption         =   "請款項目："
            Height          =   240
            Left            =   72
            TabIndex        =   53
            Top             =   384
            Width           =   972
         End
         Begin VB.Label lblWord 
            AutoSize        =   -1  'True
            Caption         =   "產生對內定稿郵件：           (N:不產生)"
            Height          =   180
            Left            =   5400
            TabIndex        =   50
            Top             =   76
            Width           =   2976
         End
         Begin VB.Label lbe1 
            Caption         =   "產生對外定稿郵件：           (N:不產生)"
            Height          =   240
            Left            =   1992
            TabIndex        =   49
            Top             =   46
            Width           =   3156
         End
         Begin VB.Label lblCP144 
            Caption         =   "請款金額："
            Height          =   240
            Left            =   72
            TabIndex        =   48
            Top             =   46
            Width           =   972
         End
      End
      Begin VB.CommandButton CmdDot 
         Caption         =   "工作點數分配"
         Height          =   255
         Left            =   6120
         TabIndex        =   2
         Top             =   185
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox txtCP113 
         Height          =   300
         Left            =   5496
         MaxLength       =   4
         TabIndex        =   1
         Top             =   162
         Width           =   600
      End
      Begin VB.TextBox txtRule 
         Height          =   300
         Left            =   5496
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1176
         Width           =   975
      End
      Begin VB.TextBox txtCustomer 
         Height          =   300
         Left            =   1056
         MaxLength       =   9
         TabIndex        =   4
         Top             =   852
         Width           =   972
      End
      Begin VB.TextBox txtDispatch 
         Height          =   300
         Left            =   1056
         MaxLength       =   7
         TabIndex        =   0
         Top             =   162
         Width           =   975
      End
      Begin VB.TextBox txtLimtDate 
         Height          =   300
         Left            =   1056
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin MSForms.TextBox txtCP130 
         Height          =   300
         Left            =   1440
         TabIndex        =   3
         Top             =   504
         Width           =   7212
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   100
         Size            =   "12721;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCP64 
         Height          =   600
         Left            =   96
         TabIndex        =   12
         Top             =   2544
         Width           =   8532
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "15049;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtLC27 
         Height          =   600
         Left            =   120
         TabIndex        =   13
         Top             =   3456
         Width           =   8532
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "15049;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作時數："
         Height          =   180
         Index           =   12
         Left            =   4524
         TabIndex        =   45
         Top             =   228
         Width           =   1056
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "主管機關名稱："
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   552
         Width           =   1260
      End
      Begin VB.Label Label8 
         Caption         =   "案件備註："
         Height          =   252
         Left            =   120
         TabIndex        =   39
         Top             =   3216
         Width           =   1812
      End
      Begin VB.Label Label13 
         Caption         =   "本所期限："
         Height          =   240
         Left            =   120
         TabIndex        =   38
         Top             =   1224
         Width           =   972
      End
      Begin VB.Label Label6 
         Caption         =   "法定期限："
         Height          =   240
         Left            =   4524
         TabIndex        =   37
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label Label3 
         Caption         =   "發  文  日："
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   222
         Width           =   972
      End
      Begin VB.Label Label11 
         Caption         =   "當  事  人："
         Height          =   288
         Left            =   120
         TabIndex        =   35
         Top             =   852
         Width           =   972
      End
      Begin VB.Label Label18 
         Caption         =   "進度備註："
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   2280
         Width           =   1812
      End
      Begin MSForms.Label lbeCusName 
         Height          =   288
         Left            =   2064
         TabIndex        =   33
         Top             =   864
         Width           =   5988
         BackColor       =   16777152
         VariousPropertyBits=   27
         Size            =   "10562;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Label Label16 
      Caption         =   "請款年度："
      Height          =   288
      Left            =   7176
      TabIndex        =   57
      Top             =   1392
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblCP115 
      Height          =   288
      Left            =   8160
      TabIndex        =   56
      Top             =   1392
      Visible         =   0   'False
      Width           =   756
   End
   Begin VB.Label lblCP156 
      Height          =   288
      Left            =   5664
      TabIndex        =   52
      Top             =   1392
      Width           =   1356
   End
   Begin VB.Label Label7 
      Caption         =   "請款階段："
      Height          =   288
      Left            =   4656
      TabIndex        =   51
      Top             =   1380
      Width           =   972
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   345
      Left            =   1230
      TabIndex        =   46
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
   Begin VB.Label Label19 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   210
      TabIndex        =   43
      Top             =   1710
      Width           =   975
   End
   Begin MSForms.Label lbeSaleName 
      Height          =   285
      Left            =   2010
      TabIndex        =   42
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
      TabIndex        =   41
      Top             =   750
      Width           =   2412
   End
   Begin MSForms.Label lbeEngName 
      Height          =   285
      Left            =   2010
      TabIndex        =   40
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
      TabIndex        =   31
      Top             =   1065
      Width           =   735
   End
   Begin VB.Label lbePoint 
      Height          =   285
      Left            =   5664
      TabIndex        =   30
      Top             =   1065
      Width           =   615
   End
   Begin VB.Label lbeEng 
      Height          =   285
      Left            =   1230
      TabIndex        =   29
      Top             =   1380
      Width           =   735
   End
   Begin VB.Label lbeProperty 
      Height          =   285
      Left            =   5664
      TabIndex        =   28
      Top             =   750
      Width           =   372
   End
   Begin VB.Label lbeCaseNum 
      Height          =   285
      Left            =   1230
      TabIndex        =   27
      Top             =   750
      Width           =   2055
   End
   Begin VB.Label lbeDate 
      Height          =   285
      Left            =   5664
      TabIndex        =   26
      Top             =   435
      Width           =   1572
   End
   Begin VB.Label lbeNum 
      Height          =   285
      Left            =   1230
      TabIndex        =   25
      Top             =   435
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "點        數："
      Height          =   285
      Left            =   4650
      TabIndex        =   24
      Top             =   1065
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "收  文  日："
      Height          =   285
      Left            =   4650
      TabIndex        =   23
      Top             =   435
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   285
      Left            =   210
      TabIndex        =   22
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號："
      Height          =   285
      Index           =   0
      Left            =   210
      TabIndex        =   21
      Top             =   435
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "案件性質："
      Height          =   285
      Left            =   4650
      TabIndex        =   20
      Top             =   750
      Width           =   972
   End
   Begin VB.Label Label24 
      Caption         =   "智權人員："
      Height          =   285
      Left            =   210
      TabIndex        =   19
      Top             =   1065
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "承  辦  人："
      Height          =   285
      Left            =   210
      TabIndex        =   18
      Top             =   1380
      Width           =   975
   End
End
Attribute VB_Name = "frm081035_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2024/03/25 Form2.0已修改; cboCaseName、lbeSaleName、lbeEngName、lbeCusName、txtCP64、txtLC27、txtCP130
Option Explicit

Dim t As Integer
Dim m_LC01 As String, m_LC02 As String, m_LC03 As String, m_LC04 As String
Dim blnIsSave As Boolean, LcTmp As String
Dim RsAdo As New ADODB.Recordset
Dim rsQuery As New ADODB.Recordset

Dim m_CP09 As String
Dim m_CP16 As String
Dim m_LC15 As String '相關國家
Dim m_LC48 As String  '特殊出名公司
Dim m_PrevForm As Form '前一畫面
Dim bolACS112 As Boolean ' ACS案件若曾有收文智財顧問112
Dim m_CP156 As String, m_CP144 As String  '請款階段CP156,請款金額CP144
Dim m_CP43T As String, m_CP43Tcp10 As String, m_CP43Tcp10name As String, m_CP43Tcp16 As String  '相關收文號,案件性質,費用
Dim m_TIPSamt0 As String '其他請款金額
Dim bolTIPSamt As Boolean 'ACS案TIPS請款階段作業(1~3)+代收代付706
Dim m_stET03 As String '定稿處理狀況
Dim m_stRtnData As String '對外定稿內文
Dim m_AttPath As String, m_AttFile As String    '附件資料夾,檔案
'Mark by Lydia 2025/08/20 改成共用ACSforLetter
'Private Const m_strMailPty = "'115','124','125','126','127'"  'Added by Lydia 2025/06/03 ACS案針對其他案件性質增加請款發文定稿通知=>【115(教育訓練)、124(營業秘密管理輔導)、125(研發管理輔導)、126(商標管理輔導)、(127)程序管理輔導】

'Added by Lydia 2025/07/22
Dim lC() As String
Dim m_CP156Max As String 'Added by Lydia 2025/09/02 該案最大的請款階段

Public Sub SetParent(ByRef fm As Form, ByVal pCP09 As String)
   Set m_PrevForm = fm
   m_CP09 = pCP09
End Sub

Private Sub cmdBack_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
          Exit Sub
      End If
   End If

   Unload Me
End Sub

'法務工作點數分配(隱瞞)
Private Sub CmdDot_Click()
   If Val(lbePoint.Caption) > 0 Then
        If PUB_CheckFormExist("frm071021") Then
             MsgBox "請先關閉【法務工作點數分配】畫面！", vbExclamation
             Exit Sub
        End If

        If PUB_GetLawPointAuto(lbeNum.Caption, True, False) = True Then '內外法發文按工作點數分配按鈕時，若無ACC1N0之工作點數資料則直接依規則先寫入ACC1N0之工作點數資料
            Set frm071021.m_PrevForm = Me
            frm071021.m_bolPrev = True
            frm071021.m_KeyList = lbeNum.Caption
            Me.Hide
            frm071021.Show
        End If
   Else
        MsgBox "無點數可供分配!", vbExclamation
   End If
End Sub

Private Sub cmdEnd_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If

   Unload Me
End Sub

Private Sub cmdSure_Click()
On Error GoTo ErrHand
   
   '檢查承辦人不可為空白
   If lbeEng = "" Then
      MsgBox "請先至分案輸入承辦人！", vbExclamation
      Exit Sub
   End If
   
   If AllTextBeforeSaveCheck Then Exit Sub
   
   If TxtValidate = False Then Exit Sub
   
   '檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Sub
   End If
   
   'Added by Lydia 2024/05/10
   textItem = PUB_StringFilter(textItem)
   If Frame2.Visible = True And Trim(textItem) = "" Then
      MsgBox "請款項目不可空白！", vbExclamation
      Exit Sub
   End If
   'end 2024/05/10
   
   '啟用工作點數,才自動分配
   If Val(lbePoint.Caption) > 0 And CmdDot.Visible = True Then
      If PUB_GetLawPointAuto(lbeNum.Caption, True, False) = False Then
      End If
   End If
   
   'Added by Lydia 2025/10/21 因為人員有時會忘記修改請款項目，預設彈提醒---By 黃教威
   If Frame2.Visible = True Then
      If MsgBox("請確認「請款金額」和「請款項目」名稱是否與合約一致，" & vbCrLf & "是否繼續發文？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   'end 2025/10/25
   
   'ACS案TIPS請款階段作業:檢查總金額
   If bolTIPSamt = True Then
      If Val(textCP144) = 0 Then
         MsgBox "請輸入請款金額!!", vbExclamation + vbOKOnly, "TIPS請款階段作業檢查"
         textCP144.SetFocus
         textCP144_GotFocus
         Exit Sub
      End If
      If Val(m_CP43Tcp16) = 0 And Val(m_CP156) > 0 Then
         MsgBox "【TIPS" & m_CP43Tcp10name & "】費用為0", vbExclamation + vbOKOnly, "TIPS請款階段作業檢查"
         textCP144.SetFocus
         textCP144_GotFocus
         Exit Sub
      End If
      If Val(m_CP156) > 0 Then
         m_TIPSamt0 = Pub_GetCP144Val(m_LC01, m_LC02, m_LC03, m_LC04, "1", lbeNum) '更新金額
         'Modified by Lydia 2025/05/19 TIPS案請款階段設定：增加「請款金額」欄位，輸入後自動寫入相對應收據號；所以排除已有收據號+ And textCP60 = ""
         If Val(m_CP43Tcp16) < Val(m_TIPSamt0) + Val(textCP144) And textCP60 = "" Then
            '抓案件全部TIPS收文總金額：舊制度可能分不同階段有收文金額，新制度只有在第一道收文有收文金額
            MsgBox "請款總金額超過收文費用＝" & m_CP43Tcp16 & _
                  IIf(Val(m_TIPSamt0) > 0, vbCrLf & "其他階段請款金額：" & m_TIPSamt0, "") & _
                  IIf(Val(textCP144) > 0, vbCrLf & "請款金額：" & textCP144, ""), vbExclamation + vbOKOnly, "TIPS請款階段作業檢查"
            textCP144.SetFocus
            textCP144_GotFocus
            Exit Sub
         End If
         'Added by Lydia 2025/04/18 TIPS分配比例管制：請款年度和收據檢查(寫入)
         If Val(lblCP115) = 0 Then
            MsgBox "請到TIPS案請款階段設定，輸入請款年度！", vbExclamation + vbOKOnly, "TIPS請款階段作業檢查"
            Exit Sub
         End If
         If textCP60 = "" Then
            'Modified by Lydia 2025/05/19 原程式改成共用模組
            textCP60 = Pub_ACS_TIPS_GetCp60("1", m_LC01, m_LC02, m_LC03, m_LC04, textCP144)
            If textCP60 = "" Then
            'end 2025/05/19
               MsgBox "請款金額與收據金額不符，請洽財務處出納人員！", vbExclamation + vbOKOnly, "TIPS請款階段作業檢查"
               Exit Sub
            End If
         End If
         'end 2025/04/18
      End If
   'Added by Lydia 2025/06/03 ACS案：增加請款發文定稿通知
   ElseIf Frame2.Visible = True Then
      If Val(textCP144) = 0 Then
         If Left(lbeNum, 1) = "A" Then
            strExc(1) = "select count(*) cnt from caseprogress where cp01='" & m_LC01 & "' and cp02='" & m_LC02 & "' and cp03='" & m_LC03 & "' and cp04='" & m_LC04 & "' and cp159=0 "
            intI = 1
            Set rsQuery = ClsLawReadRstMsg(intI, strExc(1))
            If intI = 1 Then
               If Val("" & rsQuery.Fields("cnt")) = 1 Then
                  MsgBox "請款金額不可為0！", vbExclamation, "發文請款金額檢查"
                  textCP144.SetFocus
                  textCP144_GotFocus
                  Exit Sub
               End If
            End If
         End If
         If MsgBox("請款金額為0，是否繼續發文作業？", vbInformation + vbYesNo + vbDefaultButton2, "發文請款金額檢查") = vbNo Then
            textCP144.SetFocus
            textCP144_GotFocus
            Exit Sub
         End If
      Else
          strExc(0) = GetCP144Val_2(m_LC01, m_LC02, m_LC03, m_LC04, "1", lbeNum)
          If (Val(m_CP43Tcp16) > 0 And Val(m_CP43Tcp16) < Val(strExc(0)) + Val(textCP144)) Or (Val(m_CP16) > 0 And Val(m_CP16) < Val(strExc(0)) + Val(textCP144)) Then
              MsgBox "請款總金額超過收文費用＝" & IIf(Val(m_CP43Tcp16) > 0, m_CP43Tcp16, m_CP16) & _
                    IIf(Val(m_TIPSamt0) > 0, vbCrLf & "其他請款金額：" & strExc(0), "") & _
                    IIf(Val(textCP144) > 0, vbCrLf & "請款金額：" & textCP144, ""), vbExclamation + vbOKOnly, "發文請款金額檢查"
              textCP144.SetFocus
              textCP144_GotFocus
              Exit Sub
          End If
          
          If textCP60 = "" Then
             'Modified by Lydia 2025/10/03
             'textCP60 = Pub_ACS_TIPS_GetCp60("1", m_LC01, m_LC02, m_LC03, m_LC04, textCP144, , , ACSforLetter)
             textCP60 = Pub_ACS_TIPS_GetCp60("1", m_LC01, m_LC02, m_LC03, m_LC04, textCP144, , , Mid(IIf(ACSforLetter <> "", ",", "") & ACSforLetter & IIf(ACSforTIPSAdd <> "", ",", "") & ACSforTIPSAdd, 2))
             If textCP60 = "" Then
                MsgBox "請款金額與收據金額不符，請洽財務處出納人員！", vbExclamation + vbOKOnly, "發文請款金額檢查"
                Exit Sub
             End If
          End If
      End If
   'end 2025/06/03
   End If
   
   If Frame2.Visible = True And Val(textCP144) > 0 Then  'Added by Lydia 2025/06/03
      'Move by Lydia 2025/06/03 從「ACS案TIPS請款階段作業:檢查總金額」移出來
      '先檢查檔案是否存在
      m_AttPath = App.path & "\" & strUserNum
      Call Pub_ChkExcelPath(m_AttPath)
      m_AttFile = PUB_CaseNo2FileName(m_LC01, m_LC02, m_LC03, m_LC04) & "." & lbeProperty & ".CUS.pdf"
      If Dir(m_AttPath & "\" & m_AttFile) <> "" Then
         If PUB_ChkFileOpening(m_AttFile & "\" & m_AttFile, True) = True Then
            Exit Sub
         End If
         If PUB_DelPCOrgFile(m_AttPath & "\" & m_AttFile) = False Then
            Exit Sub
         End If
      End If
      '保留; 原本定稿是開啟Word維護(frm1105_1)，再發Email
      'If txtMailC = "Y" And txtMailF <> "N" Then
      '   If PUB_CheckFormExist("frm1105_1") = True Then
      '      MsgBox "定稿維護畫面已開啟，不得繼續作業!", vbCritical + vbOKOnly
      '      Exit Sub
      '   End If
      'End If
   End If
   m_stRtnData = ""
   Me.Enabled = False
   If Not SaveData Then DataErrorMessage (3): Exit Sub
   
   '(保留)若有未發文資料顯示警告; ACS案一定有未發文
   'PUB_GetCPunIssueDatas "" & lbeCaseNum.Caption
      
   ' 列印定稿
   m_stRtnData = ""
   If txtMailF <> "N" Then
      PrintLetter
   End If

   'ACS-TIPS案請款作業通知Email
   'Modified by Lydia 2025/06/03
   'If m_LC01 = "ACS" And textCP144.Visible = True Then
   If m_LC01 = "ACS" And Frame2.Visible = True And Val(textCP144) > 0 Then
      If txtMailF <> "N" Or txtMailC <> "N" Then
         Call ShowMailACS(lbeNum, m_stRtnData)
      End If
   End If
   Me.Enabled = True
   
   Unload Me
   Exit Sub
   
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbCritical
   CmdSure.Enabled = False
End Sub

Private Sub Command1_Click()
Dim i As Integer
Dim strNum As String
Dim strTmp As String
   
   Call Forms(0).SetTmpfrm1103_2
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

Private Sub Form_Activate()
   txtDispatch.SetFocus
End Sub

Private Sub Form_Load()
Dim i As Integer, n As Integer

   MoveFormToCenter Me
   ReDim lC(TF_LC) 'Added by Lydia 2025/07/22
   
   lbeCusName.BackColor = &H8000000F
   Frame2.BackColor = &H8000000F
   GetData

End Sub

Private Sub GetData()

   If CheckReceive(m_CP09) Then
      'Modified by Lydia 2025/04/18 +cp60,cp115
      strExc(1) = "select cp05,cp06,cp07,cp08,cp09,cp10,cp01,cp02,cp03,cp04,cp13,cp14,cp16,cp18,cp27,cp64,cp71," + _
         " cp50,lc05,lc06,lc07,lc09,lc11,lc15,lc16,lc27,cp12,cp11," + _
         " cp116,cp113,cp130,CP43,CP156,CP144,lc48,cp60,cp115" + _
         " from caseprogress,lawcase where cp09 = (" + CNULL(m_CP09) + ") and" + _
         " CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 and cp27 is null  and cp57 is null"
      intI = 0
      Set RsAdo = ClsLawReadRstMsg(intI, strExc(1))
      If intI = 1 Then
         PutDataInObject
      Else
         IsNoExistData = True
      End If
   End If
     
   '與frm071006差異：取消「分所案號、催審期限、機關代號、CF代理人、收件人資料及輸入按鈕、管制期限、下一程序、是否向客戶收款、下一程序備註」，工作點數分配隱藏


   'ACS案件若曾有收文智財顧問
   bolACS112 = False
   strExc(1) = "select cp09 from caseprogress where cp01='" & m_LC01 & "' and cp02='" & m_LC02 & "' and cp03='" & m_LC03 & "' and cp04='" & m_LC04 & "' and cp10='112' and cp159=0 "
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(1))
   If intI = 1 Then
       bolACS112 = True
   End If

   'ACS案TIPS請款階段作業(1~3)+代收代付706
   m_TIPSamt0 = "0"
   bolTIPSamt = False
   txtMailF.Enabled = True
   textCP144.Enabled = True
   textAmt = "" 'Added by Lydia 2024/07/31
   'Modified by Lydia 2024/04/12 +131前置自行申請首次驗證,141諮詢再驗證,142諮詢抽驗Ｌ,143諮詢抽驗Ｓ
   'If m_LC01 = "ACS" And (m_CP43Tcp10 = "101" Or m_CP43Tcp10 = "1012" Or m_CP43Tcp10 = "1013" Or m_CP43Tcp10 = "1014") _
      And (m_CP156 <> "" Or lbeProperty = "706") Then
   If m_LC01 = "ACS" And InStr(ACSforTIPSstep, "'" & m_CP43Tcp10 & "'") > 0 And (m_CP156 <> "" Or lbeProperty = "706") Then
      bolTIPSamt = True
      Frame2.Visible = True
      Label7.Visible = True
      lblCP156.Visible = True
      textItem.Text = lbePropertyName 'Added by Lydia 2024/05/10
      Label16.Visible = True: lblCP115.Visible = True  'Added by Lydia 2025/04/18
      
      Call ClsPDGetCaseProperty(m_LC01, m_CP43Tcp10, m_CP43Tcp10name, False)
      If m_CP156 <> "" Then
         m_TIPSamt0 = Pub_GetCP144Val(m_LC01, m_LC02, m_LC03, m_LC04, "1", lbeNum)
         m_CP144 = Pub_GetCP144Val(m_LC01, m_LC02, m_LC03, m_LC04, "2", m_CP144)
         If Val(m_CP144) > 0 Then
            textCP144 = m_CP144
         End If
         lblCP156 = m_CP156
         If Val(m_CP156) > 1 Then
            strExc(1) = "select count(*) as cnt from caseprogress where  cp01='" & m_LC01 & "' and cp02='" & m_LC02 & "' and cp03='" & m_LC03 & "' and cp04='" & m_LC04 & "' and cp43='" & m_CP43T & "' and cp156<'" & m_CP156 & "' and cp158=0 and cp159=0 "
            intI = 1
            Set rsQuery = ClsLawReadRstMsg(intI, strExc(1))
            If intI = 1 Then
               If Val("" & rsQuery.Fields("cnt")) > 0 Then
                  MsgBox "前面的請款階段尚未發文，請注意 !", vbExclamation + vbOKOnly, ""
               End If
            End If
         End If
      Else
         lblCP156 = lbePropertyName
         textCP144 = m_CP16
         textCP144.Enabled = False
      End If
      'Added by Lydia 2024/07/31 在發文畫面另外顯示相關收文之補收款704的金額，代收代付彈訊息請款金額內含補收款，非代收代付彈訊息請人員自行加入。
      strExc(1) = "select sum(cp16) amt from caseprogress where cp43='" & lbeNum & "' and cp159=0 "
      intI = 1
      Set rsQuery = ClsLawReadRstMsg(intI, strExc(1))
      If intI = 1 Then
         textAmt = "" & rsQuery.Fields("amt")
         If Val(textAmt) <> 0 Then
            If lbeProperty = "706" Then
               textCP144 = Val(textCP144) + Val(textAmt)
               MsgBox "請款金額已包含補收款！", vbInformation + vbOKOnly
            Else
               MsgBox "輸入的請款金額請包含補收款！", vbInformation + vbOKOnly
            End If
         End If
      End If
      'end 2024/07/31
      
      'J公司才開放對外定稿郵件及定稿；LC48為空白為非J公司時則先彈提醒「此案非智權公司出名，不產生對外定稿 !」，對外定稿欄預設為N且鎖住。
      If m_LC48 <> "J" Then
         MsgBox "此案非智權公司出名，不產生對外定稿 !", vbExclamation + vbOKOnly, ""
         txtMailF = "N"
         txtMailF.Enabled = False
      End If
   'Added by Lydia 2025/06/03 ACS案：增加請款發文定稿通知
   'Modified by Lydia 2025/10/03 增加請款發文定稿通知性質
   'ElseIf m_LC01 = "ACS" And ((InStr(ACSforLetter, lbeProperty) > 0 And Val(m_CP16) > 0) Or (m_CP43Tcp10 <> "" And InStr(ACSforLetter, m_CP43Tcp10) > 0 And Val(m_CP43Tcp16) > 0)) Then
   ElseIf m_LC01 = "ACS" And ((InStr(ACSforLetter & "," & ACSforTIPSAdd, lbeProperty) > 0 And Val(m_CP16) > 0) Or (m_CP43Tcp10 <> "" And InStr(ACSforLetter & "," & ACSforTIPSAdd, m_CP43Tcp10) > 0 And Val(m_CP43Tcp16) > 0)) Then
      Frame2.Visible = True
      textItem.Text = lbePropertyName
   'end 2025/06/03
   Else
      Frame2.Visible = False
      Label7.Visible = False
      lblCP156.Visible = False
      Label16.Visible = False: lblCP115.Visible = False  'Added by Lydia 2025/04/18
   End If

End Sub

Private Function SaveData() As Boolean
Dim strUpd As String
Dim strNewNum As String
Dim strPaperNum As String
Dim strNP08 As String
   
On Error GoTo ErrorHandler
   
   SaveData = True
   cnnConnection.BeginTrans
   
   If txtCP113 <> "" Then
      strUpd = strUpd & ",cp113=" + CNULL(txtCP113)
   End If
   'ACS案TIPS請款階段作業(1~3)+代收代付706
   If bolTIPSamt = True Then
      strExc(1) = Pub_GetCP144Val(m_LC01, m_LC02, m_LC03, m_LC04, "0", textCP144)
      strUpd = strUpd & ",cp144=" & CNULL(ChgSQL(strExc(1)))
      txtCP64 = ChangeWStringToTDateString(DBDATE(txtDispatch)) & " " & Format(strExc(1), "##,##0") & txtCP64
      'Added by Lydia 2025/04/18
      strUpd = strUpd & ",cp60=" & CNULL(textCP60)
   'Added by Lydia 2025/06/03 ACS案：增加請款發文定稿通知
   ElseIf Frame2.Visible = True And Val(textCP144) > 0 Then
      strExc(1) = GetCP144Val_2(m_LC01, m_LC02, m_LC03, m_LC04, "0", textCP144)
      strUpd = strUpd & ",cp144=" & CNULL(ChgSQL(strExc(1)))
      txtCP64 = ChangeWStringToTDateString(DBDATE(txtDispatch)) & Format(strExc(1), "##,##0") & txtCP64
      If Left(m_CP09, 1) > "A" Then '因為案件有可能只有一次請款，不會拆收據
         strUpd = strUpd & ",cp60=" & CNULL(textCP60)
      End If
   'end 2025/06/03
   End If

   
   LcTmp = m_LC01 + m_LC02 + m_LC03 + m_LC04
   
   strExc(1) = "update caseprogress set cp27=" + CNULL(ChangeTStringToWString(txtDispatch)) + _
      ",cp64=" + CNULL(ChgSQL(txtCP64)) + ",cp130=" + CNULL(ChgSQL(txtCP130)) + strUpd
   
   strExc(1) = strExc(1) & " where cp09='" + lbeNum + "'"
   cnnConnection.Execute strExc(1)

   If ChangeCustomerL(txtCustomer) <> ChangeCustomerL(txtCustomer.Tag) Or txtLC27.Text <> txtLC27.Tag Then
      strExc(2) = "update lawcase set lc11=" + CNULL(ChangeCustomerL(txtCustomer)) + _
         ",lc27=" + CNULL(ChgSQL(txtLC27)) + _
         " where " & ChgLawcase(LcTmp)
        If txtLC27.Text <> txtLC27.Tag Then
           Pub_SeekTbLog strExc(2)
        End If
        cnnConnection.Execute strExc(2)
   End If
   
   If Trim(lbeProperty) = "901" Then '催款
      strNP08 = DBDATE(DateAdd("m", 2, ChangeWStringToWDateString(DBDATE(txtDispatch)))) '本所期限=發文日+2個月
      '期限的智權人員欄位應掛承辦人非使用者
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES (" & CNULL(lbeNum) & ",'" & m_LC01 & "','" & m_LC02 & "','" & m_LC03 & "','" & m_LC04 & "','901'," & _
                          strNP08 & "," & strNP08 & "," & CNULL(lbeEng) & "," & GetNextProgressNo() & ")"
      cnnConnection.Execute strSql
   End If
   
   '檢查案件國家收費表若有收達或提申天數, 則分別新增至下一程序檔
   strExc(0) = "SELECT CF23,CF11 FROM CASEFEE WHERE CF01='" & m_LC01 & "' AND CF02='" & m_LC15 & "' AND CF03='" & lbeProperty.Caption & "'"
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(rsQuery.Fields(0)) And rsQuery.Fields(0) <> 0 Then
         strExc(1) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
            "NP09,NP10,NP22) VALUES ('" & lbeNum.Caption & "','" & m_LC01 & "','" & m_LC02 & _
            "','" & m_LC03 & "','" & m_LC04 & "'," & 收達 & "," & _
            CompDate(2, rsQuery.Fields(0), TransDate(txtDispatch.Text, 2)) & "," & _
            CompDate(2, rsQuery.Fields(0), TransDate(txtDispatch.Text, 2)) & ",'" & _
            strUserNum & "'," & GetNextProgressNo() & ")"
         cnnConnection.Execute strExc(1)
      End If
      If Not IsNull(rsQuery.Fields(1)) And rsQuery.Fields(1) <> 0 Then
         strExc(1) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
            "NP09,NP10,NP22) VALUES ('" & lbeNum.Caption & "','" & m_LC01 & "','" & m_LC02 & _
            "','" & m_LC03 & "','" & m_LC04 & "'," & 提申 & "," & _
            CompDate(2, rsQuery.Fields(0), TransDate(txtDispatch.Text, 2)) & "," & _
            CompDate(2, rsQuery.Fields(0), TransDate(txtDispatch.Text, 2)) & ",'" & _
            strUserNum & "'," & GetNextProgressNo() & ")"
         cnnConnection.Execute strExc(1)
      End If
   End If
   
   'TIPS案件發文時收據上不列印
   If PUB_ChkACSforTIPS(m_LC01 & m_LC02 & m_LC03 & m_LC04) = True Then
      strSql = "update acc0k0 set a0k32='Z' where a0k01 in (select distinct a0j13 from acc0j0 where a0j01 in '" & lbeNum.Caption & "') and a0k32 in ('N','Y')"
      cnnConnection.Execute strSql, intI
   End If
   
   If SaveData Then blnIsSave = True Else blnIsSave = False

cnnConnection.CommitTrans

Exit Function
ErrorHandler:
   cnnConnection.RollbackTrans
   SaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)

   Set rsQuery = Nothing
   Set RsAdo = Nothing

   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
      m_PrevForm.Show
      If UCase(TypeName(m_PrevForm)) = "FRM081035_1" Then
         '(保留)ACS案一定有未發文,所以不要再選,直接回第一畫面
         'm_PrevForm.QueryData m_LC01 & m_LC02 & m_LC03 & m_LC04
         'If IsNoExistData = True Then
         '   Unload m_PrevForm
         'End If
         Unload m_PrevForm
      End If
   End If
   
   Set frm081035_2 = Nothing
End Sub

Private Sub txtLC27_GotFocus()
   TextInverse txtLC27
   OpenIme
End Sub

Private Sub txtLC27_Validate(Cancel As Boolean)
   If txtLC27 <> "" Then
      If CheckLengthIsOK(txtLC27, 2000) = False Then
         Cancel = True
         txtLC27.SetFocus
         txtLC27_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub txtMailC_GotFocus()
   TextInverse txtMailC
End Sub

Private Sub txtMailC_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtMailF_GotFocus()
   InverseTextBox txtMailF
End Sub

Private Sub txtMailF_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
       KeyAscii = 0
   End If
End Sub

' 列印定稿
Private Sub txtMailF_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(txtMailF) = False Then
      Select Case txtMailF
         Case "N"
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入 N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            txtMailF_GotFocus
      End Select
   End If
End Sub

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

Private Sub txtCp130_GotFocus()
   TextInverse txtCP130
   CloseIme
End Sub

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

      If GetCustomerAndState(txtCustomer, StrCusName, , , , m_LC01) Then lbeCusName = StrCusName Else Cancel = True: lbeCusName = ""
   Else
      MsgBox "當事人不可空白", vbCritical
      lbeCusName = ""
   End If
   
   If Cancel = False Then
      If txtCustomer.Tag <> Me.txtCustomer.Text Then
         If Not PUB_EditCustOk(lbeNum.Caption, m_LC01, m_LC02, m_LC03, m_LC04) Then Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtCustomer

End Sub

Private Sub txtDispatch_GotFocus()
   TextInverse txtDispatch
   CloseIme
End Sub

Private Sub txtDispatch_Validate(Cancel As Boolean)
   If txtDispatch <> "" Then
      If CheckIsTaiwanDate(txtDispatch) Then
         If Val(GetTaiwanTodayDate) - Val(txtDispatch) < 0 Then
            MsgBox "輸入日期大於系統日", vbCritical
            Cancel = True
         End If
       Else
          Cancel = True
       End If
   End If
   If Cancel Then TextInverse txtDispatch

End Sub


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
   End If
   If Cancel Then TextInverse txtLimtDate

End Sub

Private Sub txtCP64_GotFocus()
   TextInverse txtCP64
   OpenIme
End Sub

Private Sub txtCP64_Validate(Cancel As Boolean)
   If txtCP64 <> "" Then
     If CheckLengthIsOK(txtCP64, 2000) = False Then
         Cancel = True
         txtCP64.SetFocus
     End If
   End If

   If Cancel = False Then CloseIme
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
   End If
   If Cancel Then TextInverse txtRule

End Sub

Private Sub PutDataInObject()
Dim strTemp1 As String
   
   lbeCaseNum = GiveSymbol("" & RsAdo.Fields!cp01, "" & RsAdo.Fields!cp02, "" & RsAdo.Fields!cp03, "" & RsAdo.Fields!cp04, LcTmp)
   m_LC01 = "" & RsAdo.Fields!cp01
   m_LC02 = "" & RsAdo.Fields!cp02
   m_LC03 = "" & RsAdo.Fields!cp03
   m_LC04 = "" & RsAdo.Fields!cp04
   lbeNum = "" & RsAdo.Fields!CP09
   If Not IsNull("" & RsAdo.Fields!lc05) Then cboCaseName.AddItem "中:" + "" & RsAdo.Fields!lc05
   If Not IsNull("" & RsAdo.Fields!lc06) Then cboCaseName.AddItem "英:" + "" & RsAdo.Fields!lc06
   If Not IsNull("" & RsAdo.Fields!lc07) Then cboCaseName.AddItem "日:" + "" & RsAdo.Fields!lc07

   m_LC15 = "" & RsAdo.Fields!lc15
   If cboCaseName.ListCount <> 0 Then
      cboCaseName.ListIndex = 0
   End If
   'Added by Lydia 2025/07/22
   lC(1) = m_LC01
   lC(2) = m_LC02
   lC(3) = m_LC03
   lC(4) = m_LC04
   If ClsPDReadLawCaseDatabase(lC()) = True Then
      If lC(11) <> "" Then
         lC(11) = ChangeCustomerL(lC(11))
      End If
      If lC(43) <> "" Then
         lC(43) = ChangeCustomerL(lC(43))
      End If
      If lC(44) <> "" Then
         lC(44) = ChangeCustomerL(lC(44))
      End If
      If lC(45) <> "" Then
         lC(45) = ChangeCustomerL(lC(45))
      End If
      If lC(46) <> "" Then
         lC(46) = ChangeCustomerL(lC(46))
      End If
   End If
   'end 2025/07/22
   
   If IsNull("" & RsAdo.Fields!LC11) Then txtCustomer = "" Else
   txtCustomer = "" & RsAdo.Fields!LC11
   Call ClsPDGetCustomer(txtCustomer, strTemp1)
   lbeCusName = strTemp1
   txtCustomer.Tag = "" & txtCustomer.Text
   
   lbeDate = ChangeWStringToTDateString("" & RsAdo.Fields!cp05)
   lbeProperty = "" & RsAdo.Fields!CP10
   Call ClsPDGetCaseProperty(RsAdo.Fields!cp01, lbeProperty, strTemp1, IIf(m_LC15 <> "000", True, False))
   lbePropertyName = strTemp1
   lbeSale = "" & RsAdo.Fields!cp13
   strTemp1 = GetStaffName(lbeSale, True, , , strExc(1))
   If strExc(1) <> "1" Then
      MsgBox "智權人員已離職！"
   End If
   lbeSaleName = strTemp1
   lbeEng = "" & RsAdo.Fields!cp14
   strTemp1 = GetStaffName(lbeEng, True, , , strExc(1))
   If strExc(1) <> "1" Then
      MsgBox "承辦人員已離職！"
   End If
   lbeEngName = strTemp1
   lbePoint = "" & RsAdo.Fields!cp18
   txtDispatch = ChangeWStringToTDateString("" & RsAdo.Fields!Cp27)
   txtCP64 = "" & RsAdo.Fields!CP64
   txtLC27 = "" & RsAdo.Fields!lc27
   txtLC27.Tag = txtLC27
   
   m_CP16 = "" & RsAdo.Fields!cp16
   txtCP113 = "" & RsAdo.Fields!CP113
   
   If IsNull("" & RsAdo.Fields!CP130) Then
      Call GetCaseFeeByNick(m_LC01, m_LC15, lbeProperty.Caption, strTemp1) '預設值
      txtCP130 = strTemp1
   Else
      txtCP130 = "" & RsAdo.Fields("CP130")
   End If

   m_LC48 = "" & RsAdo.Fields("LC48")  '特殊出名公司
    
   m_CP156 = "" & RsAdo.Fields("CP156")    '請款階段
   m_CP144 = "" & RsAdo.Fields("CP144")    '請款金額
   m_CP43T = "" & RsAdo.Fields("CP43")
   'Added by Lydia 2025/04/18
   textCP60 = "" & RsAdo.Fields("CP60")
   textCP60.Tag = textCP60.Text
   lblCP115 = "" & RsAdo.Fields("CP115")
   'end 2025/04/18
   'Added by Lydia 2025/09/02 該案最大的請款階段
   m_CP156Max = ""
   If m_CP156 <> "" Then
      strExc(0) = "select nvl(max(cp156),0) mno from caseprogress where cp01='" & lC(1) & "' and cp02='" & lC(2) & "' and cp03='" & lC(3) & "' and cp04='" & lC(4) & "' and cp159=0 and nvl(cp156,0) > 0 "
      intI = 1
      Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_CP156Max = "" & rsQuery.Fields("mno")
      End If
   End If
   'end 2025/09/02
   
   strTemp1 = m_CP43T
   If m_CP43T <> "" Then
JumpToPreCP:
      If Mid(m_CP43T, 1, 1) > "A" Then
         strExc(0) = "select cp43 from caseprogress where cp09=" & CNULL(m_CP43T)
         intI = 1
         Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If "" & rsQuery.Fields("cp43") <> "" Then
               m_CP43T = "" & rsQuery.Fields("cp43")
               If Mid(m_CP43T, 1, 1) = "A" Then
                  GoTo JumpToGet
               Else
                  GoTo JumpToPreCP
               End If
            Else
               m_CP43Tcp10 = ""
               GoTo JumpToNext
            End If
         End If
      Else
JumpToGet:
         strExc(0) = "select cp09,cp10,nvl(cp16,0) cp16 from caseprogress where cp09=" & CNULL(m_CP43T)
      End If
      intI = 1
      Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         
         m_CP43Tcp10 = "" & rsQuery.Fields("cp10") '113/3/27 定稿顯示階段依照相關總收文號的性質帶出來
         
         'Modified by Lydia 2025/06/03 增加請款發文定稿+"," & ACSforLetter
         'Modified by Lydia 2025/10/03 增加請款發文定稿+"," & ACSforTIPSAdd
         If InStr(ACSforTIPSstep & "," & ACSforLetter & "," & ACSforTIPSAdd, m_CP43Tcp10) > 0 Then
            '抓契約第一道收文的費用,其他沒有費用
            '保留: 113/3/27 抓案件全部TIPS收文總金額：舊制度可能分不同階段有收文金額，新制度只有在第一道收文有收文金額
            'If "" & rsQuery.Fields("cp16") > 0 Then
            '   m_CP43Tcp16 = "" & rsQuery.Fields("cp16")
            'Else
            '   strExc(0) = "select cp09,cp10,cp16 from caseprogress where cp01='" & m_LC01 & "' and cp02='" & m_LC02 & "' and cp03='" & m_LC03 & "' and cp04='" & m_LC04 & "' " & _
            '               "and cp159=0 and cp10 in (" & ACSforTIPSstep & ") and nvl(cp16,0)>0 order by cp09 "
            '   intI = 1
            '   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
            '   If intI = 1 Then
            '      m_CP43Tcp10 = "" & rsQuery.Fields("cp10")
            '      m_CP43Tcp16 = "" & rsQuery.Fields("cp16")
            '   End If
            'End If
            strExc(0) = "select sum(nvl(cp16,0)) tot1 from caseprogress where cp01='" & m_LC01 & "' and cp02='" & m_LC02 & "' and cp03='" & m_LC03 & "' and cp04='" & m_LC04 & "' and cp159=0 "
            'Added by Lydia 2025/06/03 區分"請款發文定稿"
            If InStr(ACSforLetter, m_CP43Tcp10) > 0 Then
               strExc(0) = strExc(0) & "and cp10 in (" & ACSforLetter & ") "
            'Added by Lydia 2025/10/03 增加請款發文定稿通知性質
            ElseIf InStr(ACSforTIPSAdd, m_CP43Tcp10) > 0 Then
               strExc(0) = strExc(0) & "and cp10 in (" & ACSforTIPSAdd & ") "
            'end 2025/10/03
            Else
            'end 2025/06/03
               strExc(0) = strExc(0) & "and cp10 in (" & ACSforTIPSstep & ") "
            End If
            
            intI = 1
            Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_CP43Tcp16 = "" & rsQuery.Fields("tot1")
            End If
            
         Else
            m_CP43Tcp10 = ""
         End If
      End If
   End If
   
JumpToNext:
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
      If ClsPDGetCustomer(txtCustomer, StrCusName) Then
         lbeCusName = StrCusName
      Else
         TextInverse txtCustomer
         lbeCusName = ""
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
   Else
      MsgBox "當事人不可空白", vbCritical
      TextInverse txtCustomer
      lbeCusName = ""
   End If

   '智財顧問專業分配比例管制：ACS案件若曾有收文智財顧問112，每道進度發文都一定要輸入工作時數，輸0也可以
   If bolACS112 = True And Trim(txtCP113) = "" And strSrvDate(1) >= ACS_PFrateStart Then
       MsgBox "請輸入工作時數！", vbInformation
       txtCP113.SetFocus
       txtCP113_GotFocus
       AllTextBeforeSaveCheck = True
       Exit Function
   End If
   
    'ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
    If Pub_ChkACS112isNull(m_LC01, m_LC02, m_LC03, m_LC04, txtCP113) = True Then
          txtCP113.SetFocus
          txtCP113_GotFocus
          AllTextBeforeSaveCheck = True
          Exit Function
    End If
    
   AllTextBeforeSaveCheck = False
End Function


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
   
  
   If Me.txtDispatch.Enabled = True Then
      Cancel = False
      txtDispatch_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If

   txtCP113_Validate Cancel
   If Cancel = True Then
      txtCP113.SetFocus
      Exit Function
   End If

   If Me.txtLimtDate.Enabled = True Then
      Cancel = False
      txtLimtDate_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtCP64.Enabled = True Then
      Cancel = False
      txtCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtLC27.Enabled = True Then
      Cancel = False
      txtLC27_Validate Cancel
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

  
   TxtValidate = True
End Function

Private Function CheckReceive(strNum As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM CASEPROGRESS WHERE CP09 ='" & strNum & "'" & _
            " AND (CP27 IS NULL OR CP27 ='') AND (CP57 IS NULL OR CP57 = '')"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.EOF = False Then
      CheckReceive = True
      m_LC01 = rsTmp.Fields("CP01"): m_LC02 = rsTmp.Fields("CP02"): m_LC03 = rsTmp.Fields("CP03"): m_LC04 = rsTmp.Fields("CP04")
   Else
      CheckReceive = False
   End If
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintLetter()
Dim bolChk As Boolean

   Dim strFullFileName As String
   Dim oFileSys As New FileSystemObject
   Dim oFile As File
   Dim stET02 As String
   m_stET03 = ""
   m_stRtnData = ""
   If bolTIPSamt = True Then
      If Val(m_CP156) > 0 Then
         m_stET03 = "01"
      Else
         m_stET03 = "10" '代收代付
      End If
   'Added by Lydia 2025/06/03 ACS案：增加請款發文定稿通知
   ElseIf Frame2.Visible = True And Val(textCP144) > 0 Then
      m_stET03 = "11" '非TIPS案之發文請款
   'end 2025/06/03
   End If
   

   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   If m_stET03 <> "" Then
   
      StartLetter "01", lbeNum, m_stET03
      
      NowPrint lbeNum, "01", m_stET03, False, strUserNum, , , True, m_stRtnData
      '人工去掉Word符號
      m_stRtnData = Replace(m_stRtnData, "|#(中字,靠右)(", "")
      m_stRtnData = Replace(m_stRtnData, ")#|", "")
      
      '上傳到卷宗區
      PUB_DelFtpFile2 lbeNum, " and upper(cpp02)='" & UCase(m_AttFile) & "'"  '檔案改放 FTP,必須在DB資料刪除前執行
      strSql = "delete from CasePaperPDF where cpp01='" & lbeNum & "' and upper(cpp02)='" & UCase(m_AttFile) & "'"
      cnnConnection.Execute strSql
      If PUB_PrintLetter(lbeNum, , , True, strFullFileName, False) = True Then
         Call PUB_ChkFileStatus(strFullFileName, False)  '判斷檔案是否存在, 超過時間就繼續;
         Set oFile = oFileSys.GetFile(strFullFileName)
         If SaveAttFile_PDF(lbeNum, strFullFileName, m_AttFile, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) = False Then
            MsgBox "定稿上傳卷宗區失敗!" & vbCrLf & Err.Description, vbExclamation
            Exit Sub
         End If
         Kill strFullFileName
      End If
   End If

'保留: 原本定稿是開啟Word維護(frm1105_1)，再發Email
'   Select Case m_LC01
'      Case "L", "LA"
'         'Modified by Lydia 2024/03/04 改用變數"01"->m_stET03
'         NowPrint lbeNum, "01", m_stET03, bolChk, strUserNum, 0
'      'Added by Lydia 2024/03/04
'      Case Else
'         If m_stET03 <> "" Then
'            If bolTIPSamt = True Then
'               If bolChk = True Then
'                  '預設開啟Word=>上傳到卷宗區
'                  PUB_DelFtpFile2 lbeNum, " and upper(cpp02)='" & UCase(m_AttFile) & "'"  '檔案改放 FTP,必須在DB資料刪除前執行
'                  strSql = "delete from CasePaperPDF where cpp01='" & lbeNum & "' and upper(cpp02)='" & UCase(m_AttFile) & "'"
'                  cnnConnection.Execute strSql
'                  NowPrint lbeNum, "01", m_stET03, bolChk, strUserNum, 0
'                  frm1105_1.m_RecNo = lbeNum
'                  frm1105_1.m_PdfName = m_AttFile '會直接上傳到卷宗區
'                  Set frm1105_1.m_PrevForm = Me
'                  frm1105_1.Show
'                  Me.Hide
'               Else
'                  NowPrint lbeNum, "01", m_stET03, False, strUserNum, , , True, m_stRtnData
'                  PUB_DelFtpFile2 lbeNum, " and upper(cpp02)='" & UCase(m_AttFile) & "'"  '檔案改放 FTP,必須在DB資料刪除前執行
'                  strSql = "delete from CasePaperPDF where cpp01='" & lbeNum & "' and upper(cpp02)='" & UCase(m_AttFile) & "'"
'                  cnnConnection.Execute strSql
'                  If PUB_PrintLetter(lbeNum, , , True, strFullFileName, False) = True Then
'                     Call PUB_ChkFileStatus(strFullFileName, False)  '判斷檔案是否存在, 超過時間就繼續;
'                     Set oFile = oFileSys.GetFile(strFullFileName)
'                     If SaveAttFile_PDF(lbeNum, strFullFileName, m_AttFile, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) = False Then
'                        MsgBox "定稿上傳卷宗區失敗!" & vbCrLf & Err.Description, vbExclamation
'                        Exit Sub
'                     End If
'                     Kill strFullFileName
'                  End If
'               End If
'            Else
'               NowPrint lbeNum, "01", m_stET03, bolChk, strUserNum, 0
'            End If
'         End If
'   End Select
End Sub

'計算所限和法限
Private Sub GetNewDate()
Dim strDate As String
      
   If Trim(txtDispatch) = "" Then Exit Sub
   strDate = TransDate(txtDispatch, 2)
   txtRule = TransDate(strDate, 1)
   strDate = CompDate(2, -4, strDate)
   txtLimtDate = TransDate(strDate, 1)

End Sub

Private Sub textCP144_GotFocus()
   TextInverse textCP144
End Sub

Private Sub textCP144_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
Dim strTxt(1 To 10) As String, strTmp As String
Dim ii As Integer
   
   ' 清除定稿例外欄位檔原有資料
   EndLetter ET01, lbeNum, ET03, strUserNum
   If m_LC01 = "ACS" And bolTIPSamt = True Then
      Select Case m_CP43Tcp10
         'Modified by Lydia Lydia 2024/04/12 +前置自行申請首次驗證131
         Case "101", "131"
            strTmp = "TIPS首次驗證"
         'Modified by Lydia Lydia 2024/04/12 +諮詢再驗證141
         Case "1012", "141"
            strTmp = "TIPS再驗證"
         'Modified by Lydia Lydia 2024/04/12 +諮詢抽驗Ｌ142,諮詢抽驗Ｓ143
         Case "1013", "1014", "142", "143"
            strTmp = "TIPS抽驗"
         'Added by Lydia 2025/02/24 增加：124營業秘密管理輔導,125研發管理輔導,126商標管理輔導,127程序管理輔導
         Case Else
            strTmp = "TIPS" & m_CP43Tcp10name
         'end 2025/02/24
      End Select
      
      If Val(m_CP156) > 0 Then
         ii = ii + 1
         'Modified by Lydia 2025/09/02 若該案只有一個請款階段，不顯示第Ｘ階段
         'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
              "','請款階段','" & ChgSQL(PUB_ChgNumber2Chinese(m_CP156)) & "')"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
              "','請款階段','" & IIf(m_CP156 = "1" And m_CP156 = m_CP156Max, "", ChgSQL("第" & PUB_ChgNumber2Chinese(m_CP156) & "階段")) & "')"
         If InStr(strTmp, "輔導") = 0 Then 'Added by Lydia 2025/02/24
            strTmp = strTmp & "輔導"
         End If 'Added by Lydia 2025/02/24
      Else
         strTmp = strTmp & "申請費用"
      End If
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','案件種類','" & ChgSQL(strTmp) & "')"
      
      'Mark by Lydia 2024/04/02
      'If lbeSale > "6" And lbeSale < "F" Then
       '  strTmp = GetStaffName(lbeSale, True, strExc(1))
       '  If strTmp <> "" Then
       '     ii = ii + 1
       '     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       '          "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
       '          "','智權人員資料','" & ChgSQL("智權部　 " & strTmp) & "')"
        ' End If
      'End If
      
      'If lbeEngName <> "" Then
       '  ii = ii + 1
        ' strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
              "','承辦人員資料','" & ChgSQL("顧問服務組　 " & lbeEngName) & "')"
      'End If
      'end 2024/04/02

   End If
   
   'Added by Lydia 2025/06/03
   If Frame2.Visible = True And Val(textCP144) > 0 Then
      'Move by Lydia 2025/06/03 從上面bolTIPSamt = True移過 來
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','請款費用','" & ChgSQL(textCP144) & "')"
      
      ii = ii + 1
      strTmp = "瑞興商業銀行長安分行，戶名：台一智權股份有限公司，帳號：0075211607750。"
      'Added by Lydia 2025/06/03 改用收據公司別帶出帳號
      If textCP60 <> "" Then
          strExc(1) = "select a0802,a0826 from acc0k0, acc080 where a0k01='" & textCP60 & "' and a0k11=a0801(+)"
          intI = 1
          Set rsQuery = ClsLawReadRstMsg(intI, strExc(1))
          If intI = 1 Then
             strTmp = "瑞興商業銀行長安分行，戶名：" & rsQuery.Fields("a0802") & "，帳號：" & rsQuery.Fields("a0826") & "。"
          End If
      End If
      'end 2025/06/03
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','帳戶ID資料','" & ChgSQL(strTmp) & "')"
      'Added by Lydia 2024/05/10 請款項目取代<案件進度檔-案件性質>
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
           "','請款項目','" & ChgSQL(Trim(textItem)) & "')"
      'end 2024/05/10
   End If
   'end 2025/06/03
   
   If ii <> 0 Then
      If Not ClsLawExecSQL(ii, strTxt) Then
         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
      End If
   End If
End Sub

'***** ACS-TIPS案請款作業通知Email *****
Private Sub ShowMailACS(ByVal m_CP09 As String, ByVal pContent As String)
Dim strSubject As String, strAttFile As String, strContent As String
Dim strTitle1 As String, strTitle2 As String
Dim pbolDone1 As Boolean, pbolDone2 As Boolean
Dim intB As Integer, strB1 As String
Dim rsBD As New ADODB.Recordset
Dim strLetterEnd As String 'Added by Lydia 2024/04/02 Email內文信尾簽名
Dim strB2 As String 'Added by Lydia 2024/07/05
Dim strEx01 As String 'Added by Lydia 2025/07/22

   '對內Email
   'Modified by Lydia Lydia 2024/04/12 +前置自行申請首次驗證131
   If m_CP43Tcp10 = "101" Or m_CP43Tcp10 = "131" Then
      strTitle1 = "TIPS首次驗證"
   'Modified by Lydia Lydia 2024/04/12 +諮詢再驗證141
   ElseIf m_CP43Tcp10 = "1012" Or m_CP43Tcp10 = "141" Then
      strTitle1 = "TIPS再驗證"
   'Modified by Lydia Lydia 2024/04/12 +諮詢抽驗Ｌ142,諮詢抽驗Ｓ143
   ElseIf m_CP43Tcp10 = "1013" Or m_CP43Tcp10 = "1014" Or m_CP43Tcp10 = "142" Or m_CP43Tcp10 = "143" Then
      strTitle1 = "TIPS抽驗"
   'Added by Lydia 2025/02/24 TIPS分配比例管制：先開放案件性質=124營業秘密管理輔導,125研發管理輔導,126商標管理輔導,127程序管理輔導
   Else
      strTitle1 = "TIPS" & m_CP43Tcp10name
   'end 2025/02/24
   End If
   
   'Added by Lydia 2024/04/02 信尾:簽名檔請加入ACS案件承辦人=發文人員
   strExc(0) = strUserNum
   strExc(1) = GetStaffName(strExc(0), True, strExc(2))
   strExc(3) = Pub_GetStaffExtn(strExc(0))
   strLetterEnd = strExc(2) & "　　" & strExc(1) & "(#" & strExc(3) & ")　敬上" & vbCrLf
   strLetterEnd = "以上說明，" & vbCrLf & _
                  "如有相關問題，" & vbCrLf & _
                  "敬請不吝提出，" & vbCrLf & _
                  "非常感謝。" & vbCrLf & vbCrLf & strLetterEnd
   'end 2024/04/02
   
   'Modified by Lydia 2025/07/22 對內Email內文調整
   'strContent = Trim(lbeSaleName) & "您好，" & vbCrLf & vbCrLf
   strContent = vbCrLf & vbCrLf
   If bolTIPSamt = True Then 'Added by Lydia 2025/06/03 判斷TIPS請款
      If Val(m_CP156) > 0 Then
         strContent = strContent & "緣" & Trim(lbeCusName) & "業經" & Val(Mid(txtDispatch, 1, 3)) & "年" & Mid(txtDispatch, 4, 2) & "月" & Mid(txtDispatch, 6, 2) & "日" & "已T2，"
         'Added by Lydia 2025/07/22 世界先進(X85936000)、華邦電子(X36279010)因該兩家公司對於發票內容有特殊需求，須特別管控由智權人員通知財務處發票內容才開立發票，故針對TIPS請款階段發文作業通知調整
         If InStr("X85936000,X36279010", lC(11)) > 0 Then
             strEx01 = "Y"
             'Modified by Lydia 2025/09/02 若該案只有一個請款階段，不顯示第Ｘ階段
             'strContent = strContent & "TIPS輔導專案第" & PUB_ChgNumber2Chinese(m_CP156) & "階段工作項目已完成，惠請協助處理後續請款相關事宜。請向客戶確認發票開立內容，並通知財務處確認完成，於收到電子發票後，將發票寄予客戶。" & vbCrLf & vbCrLf
             strContent = strContent & "TIPS輔導專案" & IIf(m_CP156 = "1" And m_CP156 = m_CP156Max, "", "第" & PUB_ChgNumber2Chinese(m_CP156) & "階段工作項目") & "已完成，惠請協助處理後續請款相關事宜。請向客戶確認發票開立內容，並通知財務處確認完成，於收到電子發票後，將發票寄予客戶。" & vbCrLf & vbCrLf
         Else
         'end 2025/07/22
             'Modified by Lydia 2025/09/02 若該案只有一個請款階段，不顯示第Ｘ階段
             'strContent = strContent & "TIPS輔導專案第" & PUB_ChgNumber2Chinese(m_CP156) & "階段工作項目已完成，惠請協助處理後續請款相關事宜，並請於收到電子發票後，將發票寄予客戶。" & vbCrLf & vbCrLf
             strContent = strContent & "TIPS輔導專案" & IIf(m_CP156 = "1" And m_CP156 = m_CP156Max, "", "第" & PUB_ChgNumber2Chinese(m_CP156) & "階段工作項目") & "已完成，惠請協助處理後續請款相關事宜，並請於收到電子發票後，將發票寄予客戶。" & vbCrLf & vbCrLf
         End If 'Added by Lydia 2025/07/22
         strContent = strContent & "電子發票內容如下:" & vbCrLf
         strContent = strContent & "品名：T1/T2" & vbCrLf
         strContent = strContent & "金額：" & Format(textCP144, "##,##0") & vbCrLf
         'Added by Lydia 2025/02/24 通知財務處連結的收據號碼
         If textCP60 <> "" Then
            strContent = strContent & "收據號碼：" & textCP60 & vbCrLf
         End If
         'end 2025/02/24
         strContent = strContent & "備註：" & vbCrLf
         If Val(m_CP156) < 3 Then
            'Modified by Lydia 2024/05/10 請款項目取代案件性質; lbePropertyName>>textItem
            strTitle2 = "完成" & Trim(textItem)
         Else
            'Modified by Lydia 2024/05/10 請款項目取代案件性質; lbePropertyName>>textItem
            strTitle2 = "提出" & Trim(textItem)
         End If
         strContent = Replace(Replace(strContent, "T1", strTitle1), "T2", strTitle2)
      Else  '代收代付706
         'Modified by Lydia 2024/07/31 +匯款證明、
         strContent = strContent & "檢送" & Trim(lbeCusName) & "TIPS驗證費用之匯款證明、請款單如附件，" & vbCrLf
         strContent = strContent & "惠請協助處理後續請款相關事宜，" & vbCrLf
         'Modified by Lydia 2024/05/10 請款項目取代案件性質; lbePropertyName>>textItem
         strContent = strContent & "並請財務處協助開立本所" & Trim(textItem) & "驗證費用之請款單，" & vbCrLf
         'Modified by Lydia 2024/07/31 +、匯款證明
         strContent = strContent & "續行再請將本所請款單、匯款證明及資策會開立之電子發票一併寄予客戶請款。" & vbCrLf & vbCrLf
         strContent = strContent & "請款單內容如下:" & vbCrLf
         'Modified by Lydia 2024/05/10 請款項目取代案件性質; lbePropertyName>>textItem
         strContent = strContent & "品名：" & Trim(textItem) & strTitle1 & "/驗證申請費用" & vbCrLf
         strContent = strContent & "金額：" & Format(textCP144, "##,##0") & vbCrLf
         strContent = strContent & "備註：" & vbCrLf
      End If
   'Added by Lydia 2025/06/03
   Else  '非TIPS請款
      strTitle1 = lbePropertyName
      strContent = strContent & "緣" & Trim(lbeCusName) & "業經" & Val(Mid(txtDispatch, 1, 3)) & "年" & Mid(txtDispatch, 4, 2) & "月" & Mid(txtDispatch, 6, 2) & "日" & "已完成「T1」，"
      strContent = strContent & "惠請協助處理後續請款相關事宜，並請於收到電子發票後，將發票寄予客戶。" & vbCrLf & vbCrLf
      strContent = strContent & "電子發票內容如下:" & vbCrLf
      strContent = strContent & "品名：完成「T1」" & vbCrLf
      strContent = strContent & "金額：" & Format(textCP144, "##,##0") & vbCrLf
      If textCP60 <> "" Then
         strContent = strContent & "收據號碼：" & textCP60 & vbCrLf
      End If

      strContent = strContent & "備註：" & vbCrLf
      strContent = Replace(strContent, "T1", strTitle1)
   End If
   'end 2025/06/03
   
   If strTitle1 <> "" Then
      strContent = strContent & vbCrLf & strLetterEnd 'Added by Lydia 2024/04/02 +信尾

      Screen.MousePointer = vbHourglass
      'E-Mail呼叫 frm880019:要將寄信的內容及寄信的成功時間儲存在資料庫中，便於事後查詢。
      frm880019.m_bolSaveMail = True
      frm880019.m_CP01 = m_LC01
      frm880019.m_CP02 = m_LC02
      frm880019.m_CP03 = m_LC03
      frm880019.m_CP04 = m_LC04
      frm880019.m_CP09 = m_CP09
      frm880019.m_CP10 = lbeProperty
      'Modified by Lydia 2024/07/05 預設寄件人strUserNum
      frm880019.SetParent Me, strUserNum

      strSubject = m_LC01 & "-" & m_LC02 & IIf(m_LC03 & m_LC04 <> "000", "-" & m_LC03 & "-" & m_LC04, "")
       
      If bolTIPSamt = True Then 'Added by Lydia 2025/06/03 判斷TIPS請款
         If Val(m_CP156) > 0 Then
            'Modified by Lydia 2025/09/02 若該案只有一個請款階段，不顯示第Ｘ階段
            'strSubject = strSubject & "_TIPS輔導專案第" & PUB_ChgNumber2Chinese(m_CP156) & "階段請款事宜"
            strSubject = strSubject & "_TIPS輔導專案" & IIf(m_CP156 = "1" And m_CP156 = m_CP156Max, "", "第" & PUB_ChgNumber2Chinese(m_CP156) & "階段") & "請款事宜"
         Else  '代收代付706
            'Modified by Lydia 2024/07/31 +匯款證明及
            strSubject = strSubject & "_檢送" & Trim(lbeCusName) & " TIPS驗證費用之匯款證明及電子發票"
         End If
      'Added by Lydia 2025/06/03
      Else
         strSubject = m_LC01 & "-" & m_LC02 & IIf(m_LC03 & m_LC04 <> "000", "-" & m_LC03 & "-" & m_LC04, "") & "「" & textItem & "」請款事宜"
      End If
      'end 2025/06/03
      
      frm880019.txtSubject = strSubject
      '本文
      frm880019.txtContent = strContent
      
      '收件者
      frm880019.txtReceiver = lbeSale & " (" & lbeSaleName & "); "
      'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個
      If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
          strB1 = Pub_GetSpecMan("財務處應收處理人員")
      Else
          strB1 = Pub_GetSpecMan("財務處總帳人員") 'CC
      End If
      'Added by Lydia 2025/07/22
      If strEx01 = "Y" Then
         '世界先進(X85936000)、華邦電子(X36279010)因該兩家公司對於發票內容有特殊需求
      Else  '一般：財務改成正本收件者
         frm880019.txtReceiver = frm880019.txtReceiver & strB1 & " (" & PUB_ReadUserData(strB1) & "); "
      End If
      'end 2025/07/22
      
      'Modified by Lydia 2024/07/05 副本收受者之ACS01改為操作者+操作者部門主管A0924
      'strB1 = strB1 & ";acs01" 'Added by Lydia 2024/04/02 副本接收者，請預設有acs01@taie.com.tw
      'Modified by Lydia 2025/07/22 財務改成正本收件者
      'strB2 = GetDeptMan(PUB_GetST93(strUserNum), 2)
      'strB1 = strB1 & ";" & strUserNum & ";" & strB2
      strB1 = strUserNum & ";" & GetDeptMan(PUB_GetST93(strUserNum), 2)
      'end 2024/07/05
      
      'Added by Lydia 2024/11/05  加CC:ACS郵件通知主管 ---- by 教威 (已電話確認不管是顧服或創新業務都+CC)
      strB2 = Pub_GetSpecMan("ACS郵件通知主管")
      If strB2 <> "" Then strB1 = strB1 & ";" & strB2
      'end 2024/11/05
      
      If strB1 <> "" Then
         'Modified by Lydia 2024/07/05
         'frm880019.txtCopy = strB1 & " (" & GetStaffName(strB1, True) & "); "
         frm880019.txtCopy = Replace(PUB_ReadUserData(strB1, True, "1"), ",", ";")
      End If
      
      Screen.MousePointer = vbDefault

      frm880019.Show vbModal
      pbolDone2 = frm880019.m_bolDone
      Unload frm880019
   End If

   '對外Email
   If pContent <> "" Then
      '內文:增加收據抬頭,4/2 敬啟者後去掉2個換行
      strContent = Mid(pContent, InStr(pContent, "敬啟者：") + 6)
      'Added by Lydia 2024/04/02 只保留定稿內文
      strContent = Mid(strContent, 1, InStr(strContent, "如有任何問題請隨時不吝賜教，本所將竭誠提供服務。") + Len("如有任何問題請隨時不吝賜教，本所將竭誠提供服務。") - 1)
      strContent = Trim(lbeSaleName) & "您好，" & vbCrLf & vbCrLf & "惠請於收到電子發票後，轉寄下列訊息及附件請款通知函予客戶：" & vbCrLf & vbCrLf & strContent & vbCrLf
      strContent = strContent & vbCrLf & strLetterEnd '+信尾
      'end 2024/04/02
      '從卷宗區下載CUS.pdf
      strB1 = "select cpp02 from casepaperpdf where cpp01='" & m_CP09 & "' and upper(cpp02)=upper('" & m_LC01 & m_LC02 & IIf(m_LC03 & m_LC04 <> "000", "-" & m_LC03 & "-" & m_LC04, "") & "." & lbeProperty & ".CUS.PDF') "
      intB = 1
      Set rsBD = ClsLawReadRstMsg(intB, strB1)
      If intB = 1 Then
         If PUB_GetAttachFile_CPP(m_CP09, "" & rsBD.Fields("cpp02"), m_AttPath) = True Then
            strAttFile = "" & rsBD.Fields("cpp02")
         End If
      End If

      Screen.MousePointer = vbHourglass
      'E-Mail呼叫 frm880019:要將寄信的內容及寄信的成功時間儲存在資料庫中，便於事後查詢。
      frm880019.m_bolSaveMail = True
      frm880019.m_CP01 = m_LC01
      frm880019.m_CP02 = m_LC02
      frm880019.m_CP03 = m_LC03
      frm880019.m_CP04 = m_LC04
      frm880019.m_CP09 = m_CP09
      frm880019.m_CP10 = lbeProperty
      'Modified by Lydia 2024/07/05 預設寄件人strUserNum
      frm880019.SetParent Me, strUserNum
 
      strSubject = m_LC01 & "-" & m_LC02 & IIf(m_LC03 & m_LC04 <> "000", "-" & m_LC03 & "-" & m_LC04, "") & "號檢送TIPS"
      If bolTIPSamt = True Then 'Added by Lydia 2025/06/03 判斷TIPS請款
         If Val(m_CP156) > 0 Then
            'Modified by Lydia 2025/09/02 若該案只有一個請款階段，不顯示第Ｘ階段
            'strSubject = strSubject & "輔導專案第" & PUB_ChgNumber2Chinese(m_CP156) & "階段服務費用之電子發票"
            strSubject = strSubject & "輔導專案" & IIf(m_CP156 = "1" And m_CP156 = m_CP156Max, "", "第" & PUB_ChgNumber2Chinese(m_CP156) & "階段") & "服務費用之電子發票"
         Else  '代收代付706
            'Modified by Lydia 2024/07/31 +、匯款證明
            strSubject = strSubject & "驗證費用之請款單、匯款證明及電子發票"
         End If
      'Added by Lydia 2025/06/03
      Else
         strSubject = m_LC01 & "-" & m_LC02 & IIf(m_LC03 & m_LC04 <> "000", "-" & m_LC03 & "-" & m_LC04, "") & "檢送「" & textItem & "」服務費用之電子發票"
      End If
      'end 2025/06/03
      
      frm880019.txtSubject = strSubject
      '本文
      frm880019.txtContent = strContent
      '附件
      If strAttFile <> "" Then
         frm880019.SetAttach m_AttPath & "\" & strAttFile
      End If
      
      frm880019.txtReceiver = lbeSale & " (" & lbeSaleName & "); "
      'Added by Lydia 2024/04/02 副本接收者，請預設有acs01@taie.com.tw
      'Modified by Lydia 2024/07/05 副本收受者之ACS01改為操作者+操作者部門主管A0924
      'strB1 = "acs01"
      strB2 = GetDeptMan(PUB_GetST93(strUserNum), 2)
      strB1 = strUserNum & ";" & strB2
      'Added by Lydia 2024/11/05  加CC:ACS郵件通知主管 ---- by 教威 (已電話確認不管是顧服或創新業務都+CC)
      strB2 = Pub_GetSpecMan("ACS郵件通知主管")
      If strB2 <> "" Then strB1 = strB1 & ";" & strB2
      'end 2024/11/05
      
      If strB1 <> "" Then
         'Mark by Lydia 2024/07/05
         'frm880019.txtCopy = strB1 & " (" & GetStaffName(strB1, True) & "); "
         frm880019.txtCopy = Replace(PUB_ReadUserData(strB1, True, "1"), ",", ";")
      End If
      'end 2024/04/02
      
      Screen.MousePointer = vbDefault
      
      frm880019.Show vbModal
      pbolDone1 = frm880019.m_bolDone
      If strAttFile <> "" Then
         Kill m_AttPath & "\" & strAttFile
      End If
      Unload frm880019
   Else
      pbolDone1 = True
   End If

   Set rsBD = Nothing
End Sub


'Added by Lydia 2024/03/13 ACS-TIPS案請款作業通知Email
'Memo by Lydia 2024/03/22 保留; 原本定稿是開啟Word維護(frm1105_1)，再發Email
'Public Sub PUB_ShowMailACS(ByVal pCP09 As String, ByVal pContent As String, Optional ByRef pFromFrm As Form)
'Dim strSubject As String, strAttFile As String, strContent As String
'Dim strTitle1 As String, strTitle2 As String
'Dim pbolDone1 As Boolean, pbolDone2 As Boolean
'Dim pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String, pCP10 As String, pCP10name As String, pCP43cp10 As String
'Dim pCP27 As String, pCustName As String
'Dim pCP156 As String, pCP144 As String, pCP13 As String, pCP13n As String, pCP14 As String, pCP14n As String
'Dim intB As Integer, strB1 As String
'Dim rsBD As New ADODB.Recordset
'Dim pFileName As String, pET03 As String
'
'   strB1 = "select c1.cp27,c1.cp01,c1.cp02,c1.cp03,c1.cp04,c1.cp10,nvl(cpm03,cpm04) as cp10n," & _
'           " c1.cp156,c1.cp144,c1.cp13, s1.st02 as cp13n,c2.cp10 as cp43cp10 ,nvl(cu04,nvl(cu05,cu06)) as cname" & _
'           " from caseprogress c1, caseprogress c2 , casepropertymap,staff s1 , lawcase , customer" & _
'           " where c1.cp09='" & pCP09 & "' and c1.cp01=cpm01(+) and c1.cp10=cpm02(+) and c1.cp13=s1.st01(+) and c1.cp43=c2.cp09(+)" & _
'           " and c1.cp01=lc01(+) and c1.cp02=lc02(+) and c1.cp03=lc03(+) and c1.cp04=lc04(+) and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+)"
'   intB = 1
'   Set rsBD = ClsLawReadRstMsg(intB, strB1)
'   If intB = 1 Then
'      pCP01 = "" & rsBD.Fields("cp01")
'      pCP02 = "" & rsBD.Fields("cp02")
'      pCP03 = "" & rsBD.Fields("cp03")
'      pCP04 = "" & rsBD.Fields("cp04")
'      pCP10 = "" & rsBD.Fields("cp10")
'      pCP10name = "" & rsBD.Fields("cp10n")
'      pCP156 = "" & rsBD.Fields("cp156")
'      pCP144 = Pub_GetCP144Val("2", "" & rsBD.Fields("cp144"))
'      pCP13 = "" & rsBD.Fields("cp13")
'      pCP13n = "" & rsBD.Fields("cp13n")
'      pCP43cp10 = "" & rsBD.Fields("cp43cp10")
'      pCP27 = "" & rsBD.Fields("cp27")
'      pCustName = "" & rsBD.Fields("cname")
'   End If
'
'   If pContent = "" Then
'      '取得定稿內容: 參考frm071006
'      If Val(pCP156) = 0 Then
'         pET03 = "10"
'      Else
'         pET03 = Format(pCP156, "00")
'      End If
'      NowPrint pCP09, "01", pET03, False, strUserNum, , , True, pContent, , , , , False
'   End If
'   '對外Email
'   If pContent <> "" Then
'      '從卷宗區下載CUS.pdf
'      strB1 = "select cpp02 from casepaperpdf where cpp01='" & pCP09 & "' and upper(cpp02)=upper('" & pCP01 & pCP02 & IIf(pCP03 & pCP04 <> "000", "-" & pCP03 & "-" & pCP04, "") & "." & pCP10 & ".CUS.PDF') "
'      intB = 1
'      Set rsBD = ClsLawReadRstMsg(intB, strB1)
'      If intB = 1 Then
'         If PUB_GetAttachFile_CPP(pCP09, "" & rsBD.Fields("cpp02"), App.path & "\" & strUserNum) = True Then
'            strAttFile = "" & rsBD.Fields("cpp02")
'         End If
'      End If
'      Screen.MousePointer = vbHourglass
'      'E-Mail呼叫 frm880019:要將寄信的內容及寄信的成功時間儲存在資料庫中，便於事後查詢。
'      frm880019.m_bolSaveMail = True
'      frm880019.m_CP01 = pCP01
'      frm880019.m_CP02 = pCP02
'      frm880019.m_CP03 = pCP03
'      frm880019.m_CP04 = pCP04
'      frm880019.m_CP09 = pCP09
'      frm880019.m_CP10 = pCP10
'      If UCase(TypeName(pFromFrm)) <> "NOTHING" Then
'         frm880019.SetParent pFromFrm
'      End If
'      strSubject = pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 <> "000", "-" & pCP03 & "-" & pCP04, "") & "號檢送TIPS"
'      If Val(pCP156) > 0 Then
'         strSubject = strSubject & "輔導專案第" & PUB_ChgNumber2Chinese(pCP156) & "階段服務費用之電子發票"
'      Else  '代收代付706
'         strSubject = strSubject & "驗證費用之請款單及電子發票"
'      End If
'      strContent = Mid(pContent, InStr(pContent, "敬啟者：") + 4)
'
'      frm880019.txtSubject = strSubject
'      '本文
'      frm880019.txtContent = strContent
'      '附件
'      If strAttFile <> "" Then
'         frm880019.SetAttach App.path & "\" & strUserNum & "\" & strAttFile
'      End If
'      frm880019.txtReceiver = pCP13 & " (" & pCP13n & "); "
'
'      Screen.MousePointer = vbDefault
'
'      frm880019.Show vbModal
'      pbolDone1 = frm880019.m_bolDone
'      If strAttFile <> "" Then
'         Kill App.path & "\" & strUserNum & "\" & strAttFile
'      End If
'      Unload frm880019
'   Else
'      pbolDone1 = True
'   End If
'
'   If pCP43cp10 = "101" Then
'      strTitle1 = "TIPS首次驗證"
'   ElseIf pCP43cp10 = "1012" Then
'      strTitle1 = "TIPS再驗證"
'   ElseIf pCP43cp10 = "1013" Or pCP43cp10 = "1014" Then
'      strTitle1 = "TIPS抽驗"
'   End If
'   strContent = Trim(pCP13n) & "您好，" & vbCrLf & vbCrLf
'   If Val(pCP156) > 0 Then
'      strContent = strContent & "緣" & Trim(pCustName) & "業經" & Val(Mid(pCP27, 1, 4)) - 1911 & "年" & Mid(pCP27, 5, 2) & "月" & Mid(pCP27, 7, 2) & "日" & "已T2，"
'      strContent = strContent & "TIPS輔導專案第" & PUB_ChgNumber2Chinese(pCP156) & "階段工作項目已完成，惠請協助處理後續請款相關事宜，並請於收到電子發票後，將發票寄予客戶。" & vbCrLf & vbCrLf
'      strContent = strContent & "電子發票內容如下:" & vbCrLf
'      strContent = strContent & "品名：T1/T2" & vbCrLf
'      strContent = strContent & "金額：" & Format(pCP144, "##,##0") & vbCrLf
'      strContent = strContent & "備註：" & vbCrLf
'      If pCP156 < "3" Then
'         strTitle2 = "完成" & pCP10name
'      Else
'         strTitle2 = "提出" & pCP10name
'      End If
'      strContent = Replace(Replace(strContent, "T1", strTitle1), "T2", strTitle2)
'   Else  '代收代付706
'      strContent = strContent & "緣" & Trim(pCustName) & "TIPS驗證費用之電子發票如附件，" & vbCrLf
'      strContent = strContent & "惠請協助處理後續請款相關事宜，" & vbCrLf
'      strContent = strContent & "並請財務處協助開立本所" & pCP10name & "驗證費用之請款單，" & vbCrLf
'      strContent = strContent & "續行再請將本所請款單及資策會開立之電子發票一併寄予客戶請款。" & vbCrLf & vbCrLf
'      strContent = strContent & "電子發票內容如下:" & vbCrLf
'      strContent = strContent & "品名：" & pCP10name & strTitle1 & "/驗證申請費用" & vbCrLf
'      strContent = strContent & "金額：" & Format(pCP144, "##,##0") & vbCrLf
'      strContent = strContent & "備註：" & vbCrLf
'   End If
'   If strTitle1 <> "" Then
'      Screen.MousePointer = vbHourglass
'      'E-Mail呼叫 frm880019:要將寄信的內容及寄信的成功時間儲存在資料庫中，便於事後查詢。
'      frm880019.m_bolSaveMail = True
'      frm880019.m_CP01 = pCP01
'      frm880019.m_CP02 = pCP02
'      frm880019.m_CP03 = pCP03
'      frm880019.m_CP04 = pCP04
'      frm880019.m_CP09 = pCP09
'      frm880019.m_CP10 = pCP10
'      If UCase(TypeName(pFromFrm)) <> "NOTHING" Then
'         frm880019.SetParent pFromFrm
'      End If
'      strSubject = pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 <> "000", "-" & pCP03 & "-" & pCP04, "")
'
'      If Val(pCP156) > 0 Then
'         strSubject = strSubject & "_TIPS輔導專案第" & PUB_ChgNumber2Chinese(pCP156) & "階段請款事宜"
'      Else  '代收代付706
'         strSubject = strSubject & "_檢送" & pCustName & " TIPS驗證費用之電子發票"
'      End If
'      frm880019.txtSubject = strSubject
'      '本文
'      frm880019.txtContent = strContent
'
'      '收件者
'      frm880019.txtReceiver = pCP13 & " (" & pCP13n & "); "
       'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個-Lydia:不確定會不會再使用,故仍修改
'      If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
'          strB1 = Pub_GetSpecMan("財務處應收處理人員")
'      Else
'         strB1 = Pub_GetSpecMan("財務處總帳人員") 'CC
'      End If
'      If strB1 <> "" Then
'         frm880019.txtCopy = strB1 & " (" & GetStaffName(strB1, True) & "); "
'      End If
'
'      Screen.MousePointer = vbDefault
'
'      frm880019.Show vbModal
'      pbolDone2 = frm880019.m_bolDone
'      Unload frm880019
'   End If
'   Set rsBD = Nothing
'End Sub

'Added by Lydia 2024/05/10
Private Sub textItem_GotFocus()
   TextInverse textItem
End Sub

'Added by Lydia 2024/05/10
Private Sub textItem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      Forms(0).PopupMenu2 textItem
   End If
End Sub

'Added by Lydia 2025/06/03 取得請款金額
Private Function GetCP144Val_2(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pKind As String, ByVal pVAL01 As String) As String
Dim strA1 As String, intA As Integer
Dim strCTitle As String
Dim rsA1 As New ADODB.Recordset
Dim strA2 As String 'Added by Lydia 2025/10/03

   strCTitle = "請款金額："  'CP144內容=>請款金額：XXXXXX;
   GetCP144Val_2 = ""
   
   Select Case pKind
      Case "0" 'CP144內容=>請款金額：XXXXXX;
         GetCP144Val_2 = strCTitle & pVAL01 & ";"
      Case "1" '其他請款金額
         'Modified by Lydia 2025/10/03 增加請款發文定稿通知性質
         'strA1 = "select cp144 from caseprogress c1 where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp10 in (" & ACSforLetter & ") and cp158>0 and instr(cp144,'" & strCTitle & "') > 0 AAA"
         'strA1 = strA1 & " Union select c1.cp144 from caseprogress c1, caseprogress c2 where c1.cp01='" & pCP01 & "' and c1.cp02='" & pCP02 & "' and c1.cp03='" & pCP03 & "' and c1.cp04='" & pCP04 & "' and c1.cp43 is not null and c1.cp43=c2.cp09(+) and c2.cp10 in (" & ACSforLetter & ") and c1.cp158>0 and instr(c1.cp144,'" & strCTitle & "') > 0 AAA "
         strA2 = Mid(IIf(ACSforLetter <> "", ",", "") & ACSforLetter & IIf(ACSforTIPSAdd <> "", ",", "") & ACSforTIPSAdd, 2)
         strA1 = "select cp144 from caseprogress c1 where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp10 in (" & strA2 & ") and cp158>0 and instr(cp144,'" & strCTitle & "') > 0 AAA"
         strA1 = strA1 & " Union select c1.cp144 from caseprogress c1, caseprogress c2 where c1.cp01='" & pCP01 & "' and c1.cp02='" & pCP02 & "' and c1.cp03='" & pCP03 & "' and c1.cp04='" & pCP04 & "' and c1.cp43 is not null and c1.cp43=c2.cp09(+) and c2.cp10 in (" & strA2 & ") and c1.cp158>0 and instr(c1.cp144,'" & strCTitle & "') > 0 AAA "
         'end 2025/10/03
         If Len(pVAL01) > 9 And Mid(pVAL01, 1, 1) < "D" Then
            strA1 = Replace(strA1, "AAA", " and c1.cp09 not in (" & GetAddStr(pVAL01) & ")")
         Else
            strA1 = Replace(strA1, "AAA", "")
         End If
         intA = 1
         Set rsA1 = ClsLawReadRstMsg(intA, strA1)
         If intA = 1 Then
            rsA1.MoveFirst
            Do While Not rsA1.EOF
               If InStr("" & rsA1.Fields("cp144"), strCTitle) > 0 Then
                  strA1 = Mid("" & rsA1.Fields("cp144"), InStr("" & rsA1.Fields("cp144"), strCTitle) + Len(strCTitle))
                  If InStr(strA1, ";") > 0 Then
                     strA1 = Mid(strA1, 1, InStr(strA1, ";") - 1)
                  End If
                  GetCP144Val_2 = Val(GetCP144Val_2) + Val(strA1)
               End If
               rsA1.MoveNext
            Loop
         End If
      Case "2" 'CP144>>取得金額
         If InStr(pVAL01, strCTitle) > 0 Then
            strA1 = Mid(pVAL01, InStr(pVAL01, strCTitle) + Len(strCTitle))
            GetCP144Val_2 = Val(Mid(strA1, 1, InStr(strA1, ";") - 1))
         Else
            GetCP144Val_2 = pVAL01
         End If
   End Select
   Set rsA1 = Nothing
End Function
