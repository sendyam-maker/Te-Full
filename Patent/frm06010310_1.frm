VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010310_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-其他案件性質"
   ClientHeight    =   5772
   ClientLeft      =   -1248
   ClientTop       =   2316
   ClientWidth     =   7836
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   7836
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Height          =   1395
      Left            =   120
      TabIndex        =   57
      Top             =   4350
      Width           =   5415
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之姓名或名稱"
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   61
         Top             =   1050
         Width           =   2430
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之代表人"
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   60
         Top             =   810
         Width           =   1950
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之代理人"
         ForeColor       =   &H000000C0&
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   59
         Top             =   570
         Width           =   1950
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "變更申請人之地址"
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   58
         Top             =   300
         Width           =   1770
      End
      Begin VB.Label Label122 
         Caption         =   "+ 300"
         ForeColor       =   &H000000C0&
         Height          =   165
         Left            =   1260
         TabIndex        =   65
         Top             =   90
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label18 
         Caption         =   "(請先至客戶資料維護修改資料)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Index           =   2
         Left            =   2550
         TabIndex        =   64
         Top             =   330
         Width           =   2715
      End
      Begin VB.Label Label18 
         Caption         =   "(請先至客戶資料維護修改資料)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   165
         Index           =   3
         Left            =   2550
         TabIndex        =   63
         Top             =   1080
         Width           =   2715
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "同時辦理事項"
         Height          =   210
         Left            =   90
         TabIndex        =   62
         Top             =   90
         Width           =   1080
      End
      Begin VB.Shape Shape2 
         Height          =   1305
         Left            =   30
         Top             =   30
         Width           =   5310
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Height          =   315
      Left            =   150
      TabIndex        =   54
      Top             =   3990
      Width           =   2775
      Begin VB.TextBox TextPA178 
         Height          =   270
         Left            =   930
         MaxLength       =   1
         TabIndex        =   56
         Top             =   0
         Width           =   300
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "證書形式:        （1: 電子 2: 紙本）"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   55
         Top             =   30
         Width           =   2610
      End
   End
   Begin VB.Frame FraPA174 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   525
      Left            =   6900
      TabIndex        =   51
      Top             =   1230
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton CmdPA174 
         BackColor       =   &H00C0FFFF&
         Caption         =   "特殊字"
         Height          =   280
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   52
         Top             =   210
         Width           =   800
      End
      Begin VB.Label lblPA174 
         Caption         =   "有特殊字"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   35
         TabIndex        =   53
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   1185
      Left            =   3150
      TabIndex        =   37
      Top             =   3090
      Width           =   3015
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   1290
         MaxLength       =   1
         TabIndex        =   43
         Top             =   900
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   1410
         MaxLength       =   1
         TabIndex        =   40
         Top             =   330
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   990
         MaxLength       =   2
         TabIndex        =   38
         Text            =   "1"
         Top             =   30
         Width           =   495
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Index           =   1
         Left            =   1830
         MaxLength       =   2
         TabIndex        =   39
         Top             =   30
         Width           =   495
      End
      Begin VB.TextBox txtCP71 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   42
         Text            =   "6"
         Top             =   630
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CheckBox chk412 
         Enabled         =   0   'False
         Height          =   195
         Left            =   30
         TabIndex        =   41
         Top             =   675
         Value           =   1  '核取
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "是否逾期補繳:              (Y:是)"
         Height          =   180
         Left            =   30
         TabIndex        =   50
         Top             =   945
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "費用是否要雙倍:"
         Height          =   180
         Left            =   30
         TabIndex        =   49
         Top             =   375
         Width           =   1305
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "(Y:雙倍)"
         Height          =   180
         Left            =   1890
         TabIndex        =   48
         Top             =   375
         Width           =   645
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "繳納第:"
         Height          =   180
         Left            =   30
         TabIndex        =   47
         Top             =   75
         Width           =   585
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Left            =   1590
         TabIndex        =   46
         Top             =   75
         Width           =   180
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "年 年費"
         Height          =   180
         Left            =   2430
         TabIndex        =   45
         Top             =   75
         Width           =   585
      End
      Begin VB.Label lblCP71 
         AutoSize        =   -1  'True
         Caption         =   "延緩公告：延緩　　　個月"
         Height          =   180
         Left            =   300
         TabIndex        =   44
         Top             =   675
         Visible         =   0   'False
         Width           =   2160
      End
   End
   Begin VB.TextBox txtCP84 
      Height          =   270
      Left            =   3960
      MaxLength       =   7
      TabIndex        =   2
      Top             =   2490
      Width           =   990
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   795
      Left            =   180
      TabIndex        =   34
      Top             =   3150
      Width           =   1755
      Begin VB.TextBox txtCP136 
         Height          =   280
         Left            =   870
         TabIndex        =   6
         Top             =   360
         Width           =   420
      End
      Begin VB.TextBox txtCP135 
         Height          =   280
         Left            =   870
         TabIndex        =   5
         Top             =   60
         Width           =   420
      End
      Begin VB.Label lblCP136 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "總項數:"
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblCP135 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "總頁數:"
         Height          =   180
         Left            =   240
         TabIndex        =   35
         Top             =   90
         Width           =   585
      End
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2790
      Width           =   300
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6828
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4860
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5685
      TabIndex        =   8
      Top             =   70
      Width           =   1110
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2490
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   13
      Top             =   555
      Width           =   550
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1584
      MaxLength       =   6
      TabIndex        =   12
      Top             =   555
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2448
      MaxLength       =   1
      TabIndex        =   11
      Top             =   555
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   10
      Top             =   555
      Width           =   375
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "Y"
      Top             =   2820
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   180
      TabIndex        =   67
      Top             =   2535
      Width           =   945
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "繳費金額:"
      Height          =   180
      Left            =   3150
      TabIndex        =   66
      Top             =   2540
      Width           =   770
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1020
      TabIndex        =   14
      Top             =   1200
      Width           =   5805
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "13652;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   6270
      TabIndex        =   4
      Top             =   2820
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;556"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   5370
      TabIndex        =   33
      Top             =   2850
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   180
      X2              =   7620
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   180
      X2              =   7620
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "補／換發證書:           (1.補發證書 2.換發證書)"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   32
      Top             =   2840
      Width           =   3530
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   4020
      TabIndex        =   31
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   4020
      TabIndex        =   30
      Top             =   2010
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   180
      TabIndex        =   29
      Top             =   2010
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   0
      Left            =   4860
      TabIndex        =   28
      Top             =   600
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3016;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   4020
      TabIndex        =   27
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人　:"
      Height          =   180
      Left            =   180
      TabIndex        =   26
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   25
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   24
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   4020
      TabIndex        =   23
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   22
      Top             =   1260
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   1
      Left            =   1020
      TabIndex        =   21
      Top             =   930
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3016;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   2
      Left            =   4860
      TabIndex        =   20
      Top             =   930
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3016;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   4
      Left            =   1020
      TabIndex        =   19
      Top             =   1680
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3016;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   5
      Left            =   4860
      TabIndex        =   18
      Top             =   1680
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3016;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   6
      Left            =   1170
      TabIndex        =   17
      Top             =   2010
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3016;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   285
      Index           =   7
      Left            =   4860
      TabIndex        =   16
      Top             =   2010
      Width           =   2790
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4921;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   2250
      TabIndex        =   15
      Top             =   2850
      Width           =   2880
   End
End
Attribute VB_Name = "frm06010310_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/8 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Public strReceiveNo As String
'Modify by Morgan 2005/8/8 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, m_CP110 As String, m_AgentName As String
Dim cp() As String 'Modify By Sindy 2017/11/8
Dim intWhere As Integer
Dim m_CP43 As String
Dim m_CaseNo As String 'Add By Sindy 2017/11/8
Dim m_bolChkFee As Boolean 'Add By Sindy 2018/5/22
Dim bolDelay As Boolean '是否延期過
Dim m_strDelayCP09 As String '延期收文號
Dim m_strReExamCP27 As String '台灣再審發文日(若再審延期發文日)
Dim m_allPage As String, m_allItem As String '總頁數,總項數
Dim strCaseFee1 As String 'strCaseFee1 國家檔中繳費年度
Dim strCaseFee2 As String 'strCaseFee2 國家檔中起算日
'Add By Sindy 2018/11/28
Dim m_CP81 As String '可否減免
Dim m_lngDisc As Long '減免金額
Dim m_lngDisc1Year As Long '第一年減免金額 Add By Sindy 2020/3/31
Dim m_DiscType As String '減免身分
Dim m_strNP09_1 As String
Dim m_strOfficalFee As String '規費
Dim m_strServiceFee As String '服務費
Dim strCaseFee(1 To 2) As String 'strCaseFee(1):國家檔中繳費年度，strCaseFee(2):國家檔中起算日
'2018/11/28 END
Dim m_strNA81Appl As String 'Add By Sindy 2019/1/22
Dim m_SendDate As String, m_SendWord As String, m_SendNumber As String 'Add By Sindy 2019/7/23
Dim strCP72 As String, strCP50 As String, strCP51 As String, strCP52 As String
Dim strCP53 As String, strCP54 As String, strCP43 As String
Dim strOldReceiveNo As String, bolReadCP As Boolean
Dim m_str412CP09 As String 'Added by Morgan 2022/12/27
Dim m_PrevForm As Form '前一畫面 Add By Sindy 2023/2/16
Dim m_nFrm As Form 'Add By Sindy 2023/2/16


'Add By Sindy 2023/2/16
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Added by Morgan 2022/12/26
Private Sub chkAtt_Click(Index As Integer)
   If chkAtt(2).Value = 1 Or _
      chkAtt(4).Value = 1 Then
      If Label122.Visible = False Then
         txtCP84 = Val(txtCP84) + 300
         Label122.Visible = True
      End If
   Else
      If Label122.Visible = True Then
         txtCP84 = Val(txtCP84) - 300
         Label122.Visible = False
      End If
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean, strTmp As String
Dim strTmp1 As String, strLetter As String 'Add by Morgan 2008/1/29
Dim strFolder As String, strFileName As String 'Add By Sindy 2017/11/8
Dim varTmpNICK As Variant, TMPnick060104 As Integer, i As Integer, varTmp As Variant
Dim stET03 As String 'Added by Morgan 2022/12/26

   Select Case Index
      Case 0
         If cp(10) = 補換發證書 And Text6 = "" Then
            MsgBox "請選擇補換發類別 !", vbCritical
            Text6.SetFocus
            Text6_GotFocus
            Exit Sub
         End If
         If cp(10) = 自請撤回 And Text6 = "" Then
            MsgBox "請選擇自撤案件性質 !", vbCritical
            Text6_GotFocus
            Exit Sub
         End If
         
         'Add By Sindy 2018/11/28
         If cp(10) = 年費 Then
            '檢查繳納年費年數
            If Me.Text7(0).Text = "" Then
               MsgBox "年度不可空白 !", vbCritical
               Me.Text7(0).SetFocus
               Text7_GotFocus 0
               Exit Sub
            End If
            If pa(72) = "" Then
               If Text7(0) <> "1" Then
                  MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
                  Me.Text7(0).SetFocus
                  Text7_GotFocus 0
                  Exit Sub
               End If
            Else
               varTmpNICK = Split(pa(72), ",")
               For TMPnick060104 = UBound(varTmpNICK) To 0 Step -1
                  If Trim(varTmpNICK(TMPnick060104)) <> "" Then
                     Exit For
                  End If
               Next TMPnick060104
               If Text7(0) <> Val(varTmpNICK(TMPnick060104)) + 1 Then
                  MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
                  Me.Text7(0).SetFocus
                  Text7_GotFocus 0
                  Exit Sub
               End If
            End If
            If Me.Text7(1).Text = "" Then
               MsgBox "年度不可空白 !", vbCritical
               Me.Text7(1).SetFocus
               Text7_GotFocus 1
               Exit Sub
            End If
            If ChkRange(Text7(0), Text7(1), "繳費年度") = True Then
               For i = Text7(0) To Text7(1)
                  If InStr(pa(72), Format(i)) > 0 Then
                     bolChk = True
                     Exit For
                  End If
               Next
               If bolChk = True Then
                  MsgBox "繳費年度重覆，請查明後再輸入 !", vbCritical
                  Me.Text7(1).SetFocus
                  Text7_GotFocus 1
                  Exit Sub
               Else
                  varTmp = Split(strCaseFee(2), ",")
                  '改判斷繳費迄年是否繳超過專用期
                  strExc(0) = TransDate(CompDate(0, Text7(1) - 1, strCaseFee(1)), 1)
                  If Val(strExc(0)) > Val(pa(25)) Then
                     MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
                     Me.Text7(1).SetFocus
                     Text7_GotFocus 1
                     Exit Sub
'                  ElseIf Text7(1) = UBound(varTmp) + 1 Then
'                     Text7(7).Text = ""
'                  Else
'                     Text7(7).Text = TransDate(CompDate(0, Val(varTmp(Val(Text7(1).Text) - 1)), strCaseFee(1)), 1)
'                     '若計算出的下次繳費年度>=專用期止日, 則清空下次繳費日(存檔時不產生下一程序)
'                     If Me.Text7(7).Text <> "" Then
'                        If DBDATE(Me.Text7(7).Text) >= DBDATE(pa(25)) Then
'                           Me.Text7(7).Text = ""
'                        End If
'                     End If
                  End If
               End If
            Else
               Me.Text7(1).SetFocus
               Text7_GotFocus 1
               Exit Sub
            End If
         End If
         
         'Add by Morgan 2005/8/8
         If TxtValidate = False Then Exit Sub
         'Added by Lydia 2020/02/17 產生各式申請書時，若基本檔「名稱有特殊字」已勾選，彈訊息提醒，並一併開啟原始檔。
         If (pa(1) = "FCP" Or pa(1) = "P") And pa(174) = "Y" Then
             MsgBox MsgText(1111), vbInformation
             If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = False Then
                 Exit Sub
             End If
         End If
         'end 2020/02/17
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Modify By Sindy 2018/8/7
         If m_PrevForm.Text6 = "3" Then '電子送件
            m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
            If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Or Pub_StrUserSt03 = "M51" Then
               strFolder = PUB_Getdesktop
            Else
               strFolder = FCP電子送件檔案存放路徑
            End If
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            
            bolReadCP = False
            '1.基本資料
            'Modify By Sindy 2019/11/15 + 授權變更: Or (cp(10) = 變更 And GetCP10(cp(43)) = "704")
            If cp(10) = 授權 Or cp(10) = 終止授權 Or (cp(10) = 變更 And GetCP10(cp(43)) = "704") Then
               If cp(10) = 變更 And GetCP10(cp(43)) = "704" Then
                  strCP43 = cp(43)
                  strOldReceiveNo = strReceiveNo: strReceiveNo = strCP43
                  ReDim cp(TF_CP)
                  cp(9) = strCP43
                  Call PUB_ReadCaseProgressDatabase(cp(), intWhere)
                  'strCP72 = cp(72): strCP50 = cp(50): strCP51 = cp(51): strCP52 = cp(52): strCP53 = cp(53): strCP54 = cp(54)
                  bolReadCP = True
               End If
               StartLetterPA_EData "01", "05", strReceiveNo, pa, cp, False, , , , , m_strNA81Appl
               '被授權人--畫面上
               Call PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), IIf(cp(72) = "", cp(50) & cp(51) & cp(52), cp(72)), , , , , , , , cp(10), cp(50), cp(51), , , , , , , , , , , , , , , , , , , "E", "01", "05", strReceiveNo)
               NowPrint strReceiveNo, "01", "05", False, strUserNum, , , True, strExc(9)
               If bolReadCP = True Then
                  strReceiveNo = strOldReceiveNo
                  ReDim cp(TF_CP)
                  cp(9) = strReceiveNo
                  Call PUB_ReadCaseProgressDatabase(cp(), intWhere)
                  'cp(72) = strCP72: cp(50) = strCP50: cp(51) = strCP51: cp(52) = strCP52: cp(53) = strCP53: cp(54) = strCP54
               End If
            Else
            '2019/11/15 END
               StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False, , , , , m_strNA81Appl
               NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
            End If
            strFileName = strFolder & "\" & m_CaseNo & ".contact"
            Call PUB_MakeDoc(strExc(9), strFileName)
            '2.申請書
            If cp(10) = 申請優先權證明 Then
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "優先權證明文件申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            ElseIf cp(10) = 自請撤回 Then
               If StartLetter2("01", "10") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "10", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利申請案撤回申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            'Added by Morgan 2022/12/27
            ElseIf cp(10) = 補換發證書 And strSrvDate(1) >= "20230101" Then
               If StartLetter2("01", "26") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "26", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利證書補（換）發申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            'end 2022/12/27
            ElseIf cp(10) = 補換發證書 And Text6 = "1" Then '1.補發證書
               If StartLetter2("01", "00") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "00", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利證書補發申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            'Modify By Sindy 2018/11/2 + 領證及繳年費
            ElseIf cp(10) = 領證及繳年費 Then
               'Modified by Morgan 2022/12/26 112年起領證改用新申請書
               If strSrvDate(1) >= "20230101" Then
                  stET03 = "26"
                  strFileName = strFolder & "\" & "專利證書申請書"
               Else
                  stET03 = "04"
                  strFileName = strFolder & "\" & "申領專利證書及申請延緩公告申請書"
               End If
               If StartLetter2("01", stET03) = False Then Exit Sub
               NowPrint strReceiveNo, "01", stET03, False, strUserNum, , , True, strExc(9)
               'end 2022/12/26
               Call PUB_MakeDoc(strExc(9), strFileName)
               
               'Added by Morgan 2022/12/27 112年起延緩公告單獨新申請書
               If strSrvDate(1) >= "20230101" And chk412.Value = 1 And chk412.Visible = True Then
                  strFileName = strFolder & "\" & "延緩公告申請書"
                  If StartLetter3("01", stET03, m_str412CP09) = False Then Exit Sub
                  NowPrint m_str412CP09, "01", stET03, False, strUserNum, , , True, strExc(9)
                  Call PUB_MakeDoc(strExc(9), strFileName)
               End If
               'end 2022/12/27
               
            'Modify By Sindy 2018/11/28 + 年費
            ElseIf cp(10) = 年費 Then
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               'Modify By Sindy 2021/1/28 + m_CP81 = "Y"
               strFileName = strFolder & "\" & "專利年費" & _
                  IIf((m_DiscType <> "" Or m_lngDisc > 0) And m_CP81 = "Y", "減收", "") & "繳納申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            'Modify By Sindy 2019/1/3 + 申請英文證明
            ElseIf cp(10) = 申請英文證明 Then
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利證書英譯證明申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            ElseIf cp(10) = 繼承 Then
               If StartLetter2("01", "00") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "00", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利申請權繼承登記申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            ElseIf cp(10) = 提早公開 Then
               If StartLetter2("01", "00") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "00", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "發明專利提早公開申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            'Add By Sindy 2019/7/23
            'Add By Sindy 2022/6/2 + 439.專利權部分拋棄,440.申請權部分拋棄
            ElseIf cp(10) = 其他 Or cp(10) = "439" Or cp(10) = "440" Then
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "一般事項申復申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            '2019/7/23 END
            'Add By Sindy 2019/11/6 + 425.優先審查
            ElseIf cp(10) = "425" Then
               If StartLetter2("01", "00") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "00", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "發明專利優先審查申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            ElseIf cp(10) = 授權 Then '704.授權
               If StartLetter2("01", "00") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "00", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利權授權登記申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            'Modify By Sindy 2022/6/7 + 授權變更
            ElseIf cp(10) = 變更 And GetCP10(cp(43)) = "704" Then
               If StartLetter2("01", "12") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "12", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利權授權變更登記申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            '2022/6/7 END
            ElseIf cp(10) = 終止授權 Then '705.終止授權
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利權授權塗銷登記申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            '2019/11/6 END
            'Add By Sindy 2020/8/24 432.回復原狀,206.補充說明
            'Modify By Sindy 2025/2/18 修正:補充說明 是工程師操作, 產生專利補正文件申請書
            'ElseIf cp(10) = "432" Or cp(10) = 補充說明 Then
            ElseIf cp(10) = "432" Then
            '2025/2/18 END
               If StartLetter2("01", "02") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "02", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "專利補正文件申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            '2020/8/24 END
            'Add By Sindy 2022/6/2 + 245.延緩審查
            ElseIf cp(10) = "245" Then
               If pa(8) = "1" Then '發明
                  If StartLetter2("01", "01") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "發明專利申請延緩實體審查申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
               ElseIf pa(8) = "3" Then '設計
                  If StartLetter2("01", "02") = False Then Exit Sub
                  NowPrint strReceiveNo, "01", "02", False, strUserNum, , , True, strExc(9)
                  strFileName = strFolder & "\" & "設計專利申請延緩實體審查申請書"
                  Call PUB_MakeDoc(strExc(9), strFileName)
               End If
            'Add By Sindy 2022/12/2 124.回復優先權主張
            ElseIf cp(10) = "124" Then
               If StartLetter2("01", "01") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "01", False, strUserNum, , , True, strExc(9)
               strFileName = strFolder & "\" & "回復優先權主張申請書"
               Call PUB_MakeDoc(strExc(9), strFileName)
            '2022/12/2 END
            'Added by Morgan 2022/12/27 443 申請證書副本
            ElseIf cp(10) = "443" Then
               stET03 = "26"
               strFileName = strFolder & "\" & "專利證書副本申請書"
               If StartLetter2("01", stET03) = False Then Exit Sub
               NowPrint strReceiveNo, "01", stET03, False, strUserNum, , , True, strExc(9)
               Call PUB_MakeDoc(strExc(9), strFileName)
            'end 2022/12/27
            'Add By Sindy 2024/9/26 422 加速審查(再審查)
            'Modified by Morgan 2024/11/14  改為447再審查加速審查
            'ElseIf cp(10) = "422" Then
            ElseIf cp(10) = "447" Then
            'end 2024/11/14
               stET03 = "26"
               strFileName = strFolder & "\" & "發明專利再審查加速審查申請書"
               If StartLetter2("01", stET03) = False Then Exit Sub
               NowPrint strReceiveNo, "01", stET03, False, strUserNum, , , True, strExc(9)
               Call PUB_MakeDoc(strExc(9), strFileName)
            'end 2022/12/27
            End If
         Else
         '2018/8/7 END
            If Text8 = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            'Add By Sindy 2015/10/27
            If cp(10) = 實體審查 Then
   '            'Add By Sindy 2017/11/8
   '            If m_PrevForm.Text6 = "3" Then '電子送件
   '               m_CaseNo = pa(1) & IIf(Left(pa(2), 1) = "0", Mid(pa(2), 2), pa(2)) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "")
   '               'Ex.\\Typing2\電子送件暫存區\FCP57550\...
   '               If Pub_StrUserSt03 = "M51" Then
   '                  strFolder = PUB_Getdesktop
   '               Else
   '                  strFolder = FCP電子送件檔案存放路徑
   '               End If
   '               strFolder = strFolder & "\" & m_CaseNo
   '               If Dir(strFolder, vbDirectory) = "" Then
   '                  MkDir strFolder
   '               End If
   '
   '               StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False
   '               NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(9)
   '               strFileName = strFolder & "\" & m_CaseNo & ".contact"
   '               Call PUB_MakeDoc(strExc(9), strFileName)
   '
   '               If StartLetter2("01", "03") = False Then Exit Sub
   '               NowPrint strReceiveNo, "01", "03", False, strUserNum, , , True, strExc(9)
   '               strFileName = strFolder & "\" & "發明專利實體審查申請書"
   '               Call PUB_MakeDoc(strExc(9), strFileName)
   '               '2017/11/8 END
   '            Else
                  Call PUB_GetApplBook_FCP(pa(), Text1 & "-" & Text2 & "-" & Text3 & "-" & Text4, cp(10), , , , , , strReceiveNo)
   '            End If
            Else
            '2015/10/27 END
               Select Case cp(10)
                  Case 補換發證書
                     strTmp = "0" & Text6
                  Case 自請撤回
                     strTmp = "0" & Text6
                  Case Else
                     strTmp = "00"
               End Select
               strLetterDate = Text5.Text
               'Modify by Morgan 2008/1/31 訴願的自請撤回要印3份申請書
               'NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum
               StartLetter "01", strTmp
               If cp(10) = 自請撤回 And strTmp = "02" Then
                  NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum, , , , , 3
               Else
                  NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum
               End If
               'end 2008/1/31
            End If
         End If
         
         m_PrevForm.Show
         ' 90.08.27 modify by louis (回到原畫面要清除畫面)
         m_PrevForm.ClearForm
      Case 1
         m_PrevForm.Show
      Case 2
         Unload m_PrevForm
   End Select
   'Add By Sindy 2018/11/8
   'Modify By Sindy 2023/2/22
   'If Frame2.Visible = True Then
   If TypeName(m_nFrm) <> "Nothing" Then
   '2023/2/22 END
      Unload m_nFrm
   End If
   '2018/11/8 END
   Unload Me
End Sub
'Added by Morgan 2022/12/27
'延緩公告申請書
Private Function StartLetter3(ByVal ET01 As String, ByVal ET03 As String, ByVal ET02 As String) As Boolean
Dim strTxt(200) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim strOa02 As String

   ii = 0
   EndLetter ET01, ET02, ET03, strUserNum
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   Call PUB_GetApplPA_EData(ET01, ET03, ET02, pa())
   
   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      jj = 1
      Do While Not .EOF
         strOa02 = IIf(strOa02 <> "", strOa02 & "、", "") & PUB_ConvertNameFormat("" & .Fields("st02"))
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
         jj = jj + 1
         .MoveNext
      Loop
      End With
   End If
   
   If chk412.Value = 1 And chk412.Visible = True Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','延緩月數','" & txtCP71 & "')"
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter3 = True
   End If
End Function

'申請書
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(200) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP07Add2M As String
Dim strAD15 As String, strAD16 As String 'Add By Sindy 2019/11/15
Dim strOa02 As String 'Add By Sindy 2019/12/13
Dim strNote As String '備註內容
Dim strNote2 As String '申請內容
Dim idx As Integer
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
  
   'Add By Sindy 2019/11/15
   'Modify By Sindy 2022/6/16 + 授權變更: Or (cp(10) = 變更 And GetCP10(cp(43)) = "704")
   If cp(10) = 授權 Or cp(10) = 終止授權 Or (cp(10) = 變更 And GetCP10(cp(43)) = "704") Then
      If cp(10) = 變更 And GetCP10(cp(43)) = "704" Then
         strCP43 = cp(43)
         strOldReceiveNo = strReceiveNo ': strReceiveNo = strCP43
         ReDim cp(TF_CP)
         cp(9) = strCP43
         Call PUB_ReadCaseProgressDatabase(cp(), intWhere)
         strCP72 = cp(72): strCP50 = cp(50): strCP51 = cp(51): strCP52 = cp(52): strCP53 = cp(53): strCP54 = cp(54)
         bolReadCP = True
      End If
      '被授權人--畫面上
      Call PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), IIf(cp(72) = "", cp(50) & cp(51) & cp(52), cp(72)), , , , , , , , cp(10), cp(50), cp(51), , , , , , , , , , , , , , , , , , , "E", ET01, ET03, strReceiveNo)
      If bolReadCP = True Then
         'strReceiveNo = strOldReceiveNo
         ReDim cp(TF_CP)
         cp(9) = strReceiveNo
         Call PUB_ReadCaseProgressDatabase(cp(), intWhere)
         'cp(72) = strCP72: cp(50) = strCP50: cp(51) = strCP51: cp(52) = strCP52: cp(53) = strCP53: cp(54) = strCP54
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','授權期間(起)','" & strCP53 & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','授權期間(迄)','" & strCP54 & "')"
      End If
   End If
   '2019/11/15 END
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa())

   '出名代理人
   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      jj = 1
      Do While Not .EOF
         'Add By Sindy 2019/12/13
         strOa02 = IIf(strOa02 <> "", strOa02 & "、", "") & PUB_ConvertNameFormat("" & .Fields("st02"))
         '2019/12/13 END
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
         jj = jj + 1
         .MoveNext
      Loop
      End With
   End If
   
   '辦理依據
   If m_SendDate <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文日期','" & ChangeTStringToTDateString(m_SendDate) & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','智專字','" & m_SendWord & "')"

      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文號','" & m_SendNumber & "')"
   End If
   
   'Added by Morgan 2022/12/26
   If Text6 <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','補換發','" & IIf(Text6 = "1", "補發", "換發") & "')"
   End If
   If Frame4.Visible Then
      If chkAtt(1).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變地址','♀')"
      End If
      If chkAtt(2).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變代理人','♀')"
      End If
      If chkAtt(3).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變代表人','♀')"
      End If
      If chkAtt(4).Value = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','變名稱','♀')"
      End If
   End If
   'end 2022/12/26
   
   'Add By Sindy 2018/11/7 領證申請書 StartLetter2("01", "04")
   'Modify By Sindy 2018/11/28 + 年費 StartLetter2("01", "01")
   If cp(10) = 領證及繳年費 Or cp(10) = 年費 Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳納起年','" & Text7(0) & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳納迄年','" & Text7(1) & "')"
      
      If cp(10) = 領證及繳年費 Then
'cancel by sonia 2019/2/20 敏莉:取消加註
'         'Add By Sindy 2019/1/22 敏莉:領證申請書上加註: 請將證書上的專利權人繕打為: XX商XXXX公司
'         'Modify By Sindy 2019/1/23 敏莉:領證申請書上加註: 請將專利證書上的專利權人繕打為：「XX商XXXX公司」
'         If m_strNA81Appl <> "" Then
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','請將專利證書上的專利權人繕打為：「" & m_strNA81Appl & "」')"
'         End If
'         '2019/1/22 END
'end 2019/2/20
         
         m_DiscType = m_nFrm.m_DiscType '減免身分
         m_lngDisc = m_nFrm.m_lngDisc '減免金額
         m_lngDisc1Year = m_nFrm.m_lngDisc1Year '第一年減免金額
         
         '一般資格繳費項目
'         'Modify By Sindy 2019/10/24 無減免才要顯示
'         If Not (m_DiscType <> "" Or m_lngDisc > 0) Then
'         '2019/10/24 END
         'Modify By Sindy 2020/3/13
         If m_CP81 <> "Y" Then   '無減免
         '2020/3/13 END
            strTmp = ""
            'Modified by Morgan 2023/1/18 第1年年費加倍時也要考慮 Ex:FCP-065427
            If Text7(0) = 1 Then
               'strTmp = "繳納" & IIf(pa(8) = "1", "發明", IIf(pa(8) = "2", "新型", "設計")) & "專利證書費" & Format(m_nFrm.m_lngOfficalFee1) & "元及第1年年費" & Format(m_nFrm.m_lngOfficalFee1Year) & "元(共計" & Format(m_nFrm.m_lngFee1) & "元)"
               strExc(1) = Format(m_nFrm.m_lngOfficalFee1)
               strExc(2) = Format(IIf(Text9.Text = "Y", 2, 1) * (m_nFrm.m_lngOfficalFee1Year))
               strExc(3) = Format(Val(strExc(1)) + Val(strExc(2)))
               strTmp = "繳納" & IIf(pa(8) = "1", "發明", IIf(pa(8) = "2", "新型", "設計")) & "專利證書費" & strExc(1) & "元及第1年年費" & strExc(2) & "元(共計" & strExc(3) & "元)"
            End If
            If Text7(1) <> 1 Then
               If strTmp <> "" Then strTmp = strTmp & "，及"
               'strTmp = strTmp & "第" & IIf(Text7(0) = 1, "2", Text7(0)) & "年至第" & Text7(1) & "年年費計" & Format(m_nFrm.m_lngFee2) & "元，合計" & Format(m_nFrm.m_lngFee1 + m_nFrm.m_lngFee2) & "元。"
               strExc(4) = Format(m_nFrm.m_lngFee2)
               strExc(5) = Format(Val(strExc(3)) + Val(strExc(4)))
               strTmp = strTmp & "第" & IIf(Text7(0) = 1, "2", Text7(0)) & "年至第" & Text7(1) & "年年費計" & strExc(4) & "元，合計" & strExc(5) & "元。"
            End If
            'end 2023/1/18
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一般資格繳費項目','" & strTmp & "')"
         End If
      Else
         'Modify By Sindy 2019/11/15 年費第7年(含)以後自然人、學校、中小企業皆無減免，故應改成年費繳納申請書
         If Val(Text7(0)) >= 7 Then
            m_DiscType = ""
            m_lngDisc = 0
         End If
         '2019/11/15 END
      End If
      
      'Modify By Sindy 2020/3/13
      If m_CP81 = "Y" Then   '有減免
      '2020/3/13 END
         If m_DiscType <> "" Or m_lngDisc > 0 Then
'            If InStr(m_DiscType, "3") = 0 Then '無中小企業
'               strTmp = pa(26) '申請人1
'            Else
'               strTmp = pa(26 + (InStr(m_DiscType, "3") - 1)) '第幾個申請人
'            End If
            '**************************************************************************
            '申請人1~5
            '**************************************************************************
            For jj = 0 To 4
               If pa(26 + jj) <> "" Then
                  'Add By Sindy 2019/11/15
                  'Call PUB_GetAD03(pa(26), pa(9), m_DiscType, , strAD15, strAD16)
                  Call PUB_GetAD03(pa(26 + jj), pa(9), m_DiscType, , strAD15, strAD16) '申請人減免資料
                  '2019/11/15 END
                  
                  '符合年費減收資格
                  strTmp = ""
                  If m_DiscType = "1" Then
                     strTmp = "自然人"
                  ElseIf m_DiscType = "2" Then
                     strTmp = "學校"
                  ElseIf m_DiscType = "3" Then
                     strTmp = "中小企業"
                  End If
                  If jj = 0 Then
                     ii = ii + 1
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','符合年費減收資格','" & strTmp & "')"
                  Else
                     '備註
                     'If strNote <> "" Then strNote = strNote & vbCrLf '加折行
                     strNote = strNote & vbCrLf & "申請人" & jj + 1 & "減免資格身分：" & strTmp
                  End If
                  'Modify By Sindy 2019/1/11
                  strTmp = ""
                  '中小企業符合減收資格依據
                  If m_DiscType = "3" Then '中小企業
                     'Modify By Sindy 2019/11/15
         '               strTmp = vbCrLf & "　製造業、營造業、礦業及土石採取業實收資本額八千萬以下：           元" & vbCrLf & _
         '                        "　前項除外之其他行業前一年營業額一億元以下：           元" & vbCrLf & _
         '                        "　我國製造業、營造業、礦業及土石採取業實收資本額新台幣八千萬以上但經常僱用員工數未滿200人：員工數        人" & vbCrLf & _
         '                        "　我國前項除外之其他行業前一年營業額一億元以上者但經常僱用員工數未滿100人：員工數        人"
                     'Modify By Sindy 2020/7/24 傳入中小企業符合減收資格依據代碼 , 轉換中文
                     strTmp = PUB_AD15ToText(strAD15, strAD16)
                     '2020/7/24 END
'                     If strAD15 = "1" Then
'                        strTmp = "製造業、營造業、礦業及土石採取業實收資本額八千萬以下：" & strAD16 & "元"
'                     ElseIf strAD15 = "2" Then
'                        strTmp = "前項除外之其他行業前一年營業額一億元以下：" & strAD16 & "元"
'                     ElseIf strAD15 = "3" Then
'                        strTmp = "我國製造業、營造業、礦業及土石採取業實收資本額新台幣八千萬以上但經常僱用員工數未滿200人：員工數" & strAD16 & "人"
'                     Else '4
'                        strTmp = "我國前項除外之其他行業前一年營業額一億元以上者但經常僱用員工數未滿100人：員工數" & strAD16 & "人"
'                     End If
                     '2019/11/15 END
                  End If
         '            '中小企業符合減收資格依據
         '            strTmp = "第" & IIf(Text7(0) = 1, "2", Text7(0)) & "年至第" & Text7(1) & "年年費計" & Format(m_strOfficalFee) & "元，減免金額" & Format(m_lngDisc) & "元，合計" & Format(m_strOfficalFee - m_lngDisc) & "元。"
         '            strExc(1) = ""
         '            If InStr(m_DiscType, "1") > 0 And InStr(m_DiscType, "2") > 0 And InStr(m_DiscType, "3") > 0 Then
         '               strExc(1) = "為自然人、學校及中小企業且資格符合中小企業認定標準第2條第1項第1款之規定"
         '            Else
         '               If InStr(m_DiscType, "1") > 0 Then
         '                  strExc(1) = "為自然人"
         '                  If InStr(m_DiscType, "2") > 0 Then
         '                     strExc(1) = strExc(1) & "及學校"
         '                  ElseIf InStr(m_DiscType, "3") > 0 Then
         '                     strExc(1) = strExc(1) & "及中小企業且資格符合中小企業認定標準第2條第1項第1款之規定"
         '                  End If
         '               ElseIf InStr(m_DiscType, "2") > 0 Then
         '                  strExc(1) = "為學校"
         '                  If InStr(m_DiscType, "3") > 0 Then
         '                     strExc(1) = strExc(1) & "及中小企業且資格符合中小企業認定標準第2條第1項第1款之規定"
         '                  End If
         '               ElseIf InStr(m_DiscType, "3") > 0 Then
         '                  strExc(1) = "為中小企業且資格符合中小企業認定標準第2條第1項第1款之規定"
         '               End If
         '            End If
         '            If strExc(1) <> "" Then
         '               strExc(1) = strExc(1) & "，依據專利年費減免辦法規定"
         '               strTmp = strTmp & strExc(1)
         '            End If
                  If strTmp <> "" Then
                     If jj = 0 Then
                        ii = ii + 1
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','中小企業符合減收資格依據','" & strTmp & "')"
                     Else
                        '備註
                        If strNote <> "" Then strNote = strNote & vbCrLf '加折行
                        strNote = strNote & "中小企業符合減收資格依據：" & strTmp
                     End If
                  End If
               End If
            Next jj
            '**************************************************************************
            
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','減收','減收')"
            If cp(10) = 領證及繳年費 Then
               '減收資格繳費項目
               strTmp = ""
               'Modified by Morgan 2023/1/18 第1年年費加倍時也要考慮 Ex:FCP-065427
               If Text7(0) = 1 Then
                  'strTmp = "繳納" & IIf(pa(8) = "1", "發明", IIf(pa(8) = "2", "新型", "設計")) & "專利證書費" & Format(m_nFrm.m_lngOfficalFee1) & "元及第1年年費" & Format(m_nFrm.m_lngOfficalFee1Year - m_lngDisc1Year) & "元(共計" & Format(m_nFrm.m_lngFee1 - m_lngDisc1Year) & "元)"
                  strExc(1) = Format(m_nFrm.m_lngOfficalFee1)
                  strExc(2) = Format(IIf(Text9.Text = "Y", 2, 1) * (m_nFrm.m_lngOfficalFee1Year - m_lngDisc1Year))
                  strExc(3) = Format(Val(strExc(1)) + Val(strExc(2)))
                  strTmp = "繳納" & IIf(pa(8) = "1", "發明", IIf(pa(8) = "2", "新型", "設計")) & "專利證書費" & strExc(1) & "元及第1年年費" & strExc(2) & "元(共計" & strExc(3) & "元)"
               End If
               If Text7(1) <> 1 Then
                  If strTmp <> "" Then strTmp = strTmp & "，及"
                  'strTmp = strTmp & "第" & IIf(Text7(0) = 1, "2", Text7(0)) & "年至第" & Text7(1) & "年年費計" & Format(m_nFrm.m_lngFee2) - (m_lngDisc - m_lngDisc1Year) & "元，合計" & Format(m_nFrm.m_lngFee1 - m_lngDisc + m_nFrm.m_lngFee2) & "元。"
                  strExc(4) = Format(m_nFrm.m_lngFee2 - m_lngDisc + IIf(Text9.Text = "Y", 2, 1) * m_lngDisc1Year)
                  strExc(5) = Format(Val(strExc(3)) + Val(strExc(4)))
                  strTmp = strTmp & "第" & IIf(Text7(0) = 1, "2", Text7(0)) & "年至第" & Text7(1) & "年年費計" & strExc(4) & "元，合計" & strExc(5) & "元。"
               End If
               'end 2023/1/18
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','減收資格繳費項目','" & strTmp & "')"
            End If
         End If
      End If
      
      If chk412.Value = 1 And chk412.Visible = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','延緩月數','" & txtCP71 & "')"
      End If
      
      'Modify By Sindy 2020/5/13
      '申領專利證書及申請延緩公告申請書:
      '有備註
      If strNote <> "" And cp(10) = 領證及繳年費 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','" & strNote & "')"
      End If
      
      'Modify By Sindy 2022/11/11 ex:FCP-068268 二者都發生,申請內容均要帶出來
      '*********************************
      '申請內容
      '*********************************
      strNote2 = "": intI = 0
      'Add By Sindy 2019/1/30
      If Me.Text10.Text = "Y" Then '逾期補繳
         intI = intI + 1
         If intI > 1 Then strNote2 = strNote2 & vbCrLf
         '逾期補繳=>繳費金額含逾期費用。
         strNote2 = strNote2 & PUB_ChgNumber2Chinese(CStr(intI)) & "、繳費金額含逾期費用。"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','繳費金額含逾期費用。')"
      End If
      'Add By Sindy 2019/12/13
      If cp(10) = 年費 And pa(143) = "N" Then '年費申請人是否出名為"N"
         intI = intI + 1
         If intI > 1 Then strNote2 = strNote2 & vbCrLf
         strNote2 = strNote2 & PUB_ChgNumber2Chinese(CStr(intI)) & "、代理人" & strOa02 & "僅辦理年費繳納之事宜，本案後續相關程序之進行，均維持原代理人。"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','代理人" & strOa02 & "僅辦理年費繳納之事宜，本案後續相關程序之進行，均維持原代理人。')"
      '2019/1/30 END
      End If
      '年費繳納申請書:
      '有申請內容
      If strNote <> "" Then
         strNote2 = strNote2 & strNote
      End If
      If strNote2 <> "" And cp(10) = 年費 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','" & strNote2 & "')"
      End If
      '2020/5/13 END
      
      'Modify By Sindy 2019/1/30 有收414申請復權才需要顯示此段內容
      'If Text9 = "Y" Then '費用雙倍
      If PUB_ChkCPExist(cp, "414", 1) Then '1=未發文
      '2019/1/30 END
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請回復領證','♀')"
         
         'Added by Morgan 2022/10/7
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','電子稽核防止漏發文','♀')"
         'end 2022/10/7
      End If
   End If
   
   'Add By Sindy 2019/12/9
   If cp(10) = 自請撤回 Then
      'Modified by Lydia 2020/03/31 改模組A0802Query => CompNameQuery
      strTmp = "1.請准予撤回本申請案。" & vbCrLf & _
               "2.請退還本案已繳實體審查或再審查規費「7000」元。" & vbCrLf & _
               "3.檢還之國庫支票抬頭請開立：「" & CompNameQuery("2") & "」。" & vbCrLf & _
               "4.本案收據正本已遺失，檢附[收據無法檢還原因切結書]1份。"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','" & strTmp & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','文件描述','退費收據')"
   End If
   '2019/12/9 END
   
   'Add By Sindy 2022/6/2 + 439.專利權部分拋棄,440.申請權部分拋棄
   intI = 0
   For idx = 26 To 30
      If pa(idx) <> "" Then
         intI = intI + 1
      End If
   Next
   If cp(10) = "439" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','本案原有" & PUB_ChgNumber2Chinese(CStr(intI)) & "位申請人，今第申請人「　　　」欲主動放棄其專利權，故懇請　鈞局准予辦理，至感德便。')"
   ElseIf cp(10) = "440" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','本案原有" & PUB_ChgNumber2Chinese(CStr(intI)) & "位申請人，今第申請人「　　　」欲主動放棄其專利申請權，故懇請　鈞局准予辦理，至感德便。')"
   End If
   '2022/6/2 END
   
   'Add By Sindy 2022/6/2
   If cp(10) = "245" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','法定期限','" & ChangeWStringToTDateString(cp(7)) & "')"
   End If
   '2022/6/2 END
   
   'Add By Sindy 2020/8/24 432.回復原狀,206.補充說明
   If cp(10) = "432" Or cp(10) = 補充說明 Then
      If cp(10) = "432" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-申請書','" & m_CaseNo & ".ATT.DATA.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-委任書','" & m_CaseNo & ".POA.pdf')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-國際優先權證明文件','" & m_CaseNo & ".PRI.pdf')"
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註-其他申復事項','♀')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','備註','♀')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他','♀')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件描述','♀')"
      If cp(10) = 補充說明 Then
         strTmp = m_CaseNo & ".ADD.pdf"
      Else
         strTmp = "♀"
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','其他-文件檔名','" & strTmp & "')"
   End If
   '2020/8/24 END
   
   'Add By Sindy 2022/12/2 124.回復優先權主張
   If cp(10) = "124" Then
      '優先權資料
      strTmp = PUB_GetAppPridate(pa, ET01, strReceiveNo, ET03)
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','優先權資料','" & strTmp & "')"
   End If
   '2022/12/2 END
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & ".contact.pdf')"
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   If m_PrevForm.Text2 = "" Then Exit Sub
   With m_PrevForm
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/8/8
   ReDim pa(TF_PA)
   ReDim cp(TF_CP) 'Add By Sindy 2017/11/8
   
   Text5.Text = strSrvDate(2)
   ReadPatent
   'Add by Morgan 2005/8/8
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   'Modified by Morgan 2020/3/20 +cp10
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, cp(10), True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1100
   lstNameAgent.Width = 1300

   Combo1.ListIndex = 0
   If cp(10) = 補換發證書 Or cp(10) = 自請撤回 Then
      Text6.Visible = True
      Label18(0).Visible = True
      '92.9.28 add by sonia
      If cp(10) = 自請撤回 Then
         Label18(0).Caption = "自撤案件性質:           (1.申請 2.訴願)"
         strExc(0) = "SELECT CP10 FROM CASEPROGRESS WHERE CP09=(SELECT CP43 FROM CASEPROGRESS WHERE CP09='" & strReceiveNo & "')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 訴願 Then
               Text6 = "2"
            Else
               Text6 = "1"
            End If
         End If
      End If
      '92.9.28 end
   Else
      Text6.Visible = False
      Label18(0).Visible = False
   End If
   Select Case cp(10)
   Case 補換發證書, 申請優先權證明, 申請英文證明, 請求閱卷, 請求公告, 終止授權, "419"
      Text8 = ""
   Case Else
      Text8 = "Y"
   End Select
   
   FraPA174.BackColor = &H8000000F 'Added by Lydia 2020/02/21
End Sub

'讀取案件性質
Private Function GetCP10(p_CP09 As String, Optional strCol As String = "CP10") As String
   Dim stSQL As String, iRtn As Integer

   GetCP10 = ""
   If p_CP09 <> "" Then
      stSQL = "select " & strCol & " from caseprogress where cp09='" & p_CP09 & "'"
      iRtn = 1
      Set AdoRecordSet3 = ClsLawReadRstMsg(iRtn, stSQL)
      If iRtn = 1 Then
         GetCP10 = "" & AdoRecordSet3.Fields(0)
      End If
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing 'Add By Sindy 2023/2/16
   Set frm06010310_1 = Nothing
End Sub

Private Sub ReadPatent()
Dim rsTemp1 As New ADODB.Recordset, Lbl As Object
Dim strTmp1(0 To 5) As String, i As Integer
   
   For Each Lbl In Label12
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If pa(1) = "FCP" Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         'Text5 = pa(10) 'Removed by Morgan 2022/12/29
         Label12(1) = pa(11)
         Label12(2) = pa(22)
         AddCboName Combo1, pa(5), pa(6), pa(7)
      End If
   Else
      If pa(1) = "FG" Then
         If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
            'Text5 = pa(10) 'Removed by Morgan 2022/12/29
            Label12(1) = pa(11)
            AddCboName Combo1, pa(5), pa(6), pa(7)
         End If
      End If
   End If
   
   'Add By Sindy 2019/7/23
   '來函文號:
   strExc(0) = "select cp08,ed08 from caseprogress,edocument where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP09='" & cp(9) & "' AND ed11(+)=cp09 ORDER BY NVL(ED08,CP05) DESC"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp("ED08")) Then
         m_SendDate = RsTemp("ED08") - 19110000
         If Not IsNull(RsTemp("cp08")) Then
            strExc(0) = Mid(RsTemp("cp08"), InStr(RsTemp("cp08"), "智專") + 2, Len(RsTemp("cp08")))
            m_SendWord = Mid(strExc(0), 1, InStr(strExc(0), "字") - 1)
            strExc(0) = Replace(strExc(0), m_SendWord & "字第", "")
            m_SendNumber = Mid(strExc(0), 1, InStr(strExc(0), "號") - 1)
         End If
      End If
   End If
   
   'Add By Sindy 2017/11/8
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If Val(cp(84)) > 0 Then
         txtCP84 = cp(84) '發文規費
      ElseIf Val(cp(17)) > 0 Then
         txtCP84 = cp(17) '規費
      Else
         txtCP84 = 0
      End If
   End If
   '2017/11/8 END
   
   'Added by Morgan 2022/12/26 領證預設其他則需人工指定
   If cp(10) = 領證及繳年費 Then
      TextPA178 = PUB_GetCertType(pa(1), pa(2), pa(3), pa(4))
   Else
      TextPA178 = ""
   End If
   'end 2022/12/26
   
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2,cp43,cp10,CP110,CP43,CP135,CP136,CP84 from caseprogress,casepropertymap,staff," & _
      "staff staff1 where cp09='" & strReceiveNo & "' AND cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and " & _
      "cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP43 = "" & .Fields("CP43") 'Add by Morgan 2008/1/29
      m_CP110 = "" & .Fields("CP110")
      'Add By Cheng 2002/07/17
      If Not IsNull(.Fields(0)) Then Label12(0) = .Fields(0)
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1)
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2)
      If Not IsNull(.Fields(3)) Then
         strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields(0)) Then Label12(6) = TransDate(rsTemp1.Fields(0), 1)
            If Not IsNull(rsTemp1.Fields(1)) Then Label12(7) = rsTemp1.Fields(1)
         End If
      End If
   End If
   End With
   
   Me.Height = 4550 'Added by Morgan 2022/12/28
   Frame4.Visible = False 'Added by Morgan 2022/12/27
   Frame3.Visible = False 'Added by Morgan 2022/12/26
   Frame2.Visible = False 'Add By Sindy 2018/11/8
   'Add By Sindy 2018/5/22
   Frame1.Visible = False
   If cp(10) = 實體審查 Then
      bolDelay = False
'      m_bol107NewFee = True
      'Modify by Morgan 2006/8/18 加判斷107(再審),803(舉發),301,302,303,305(改請)才要
      'Modified by Morgan 2013/8/26 +507 -- FCP032929
      If InStr("107,803,301,302,303,305,507", cp(10)) > 0 Then
         'Add by Morgan 2004/9/8 檢查是否有延期，若有則規費預設0
         bolDelay = PUB_ChkDelay(strReceiveNo, m_strDelayCP09, strExc(1))
'         If bolDelay = True Then
'            If strExc(1) < "20130101" Then m_bol107NewFee = False 'Added by Morgan 2013/1/9
'            cp(17) = "0"
'         End If
      End If
      
      m_strReExamCP27 = "" 'Added by Morgan 2013/1/10
      'Added by Morgan 2013/1/10
      If (cp(10) = "210" Or cp(10) = "203" Or cp(10) = "204" Or cp(10) = "205" Or cp(10) = "206") Then
         m_strReExamCP27 = PUB_GetReExamDate(cp)
'         If m_strReExamCP27 > "20130000" Then
'            m_bolFixNewFee = True
'         End If
      End If
      'end 2013/1/10
      
      Frame1.Visible = True
      m_bolChkFee = True
      '讀取總頁數和總項數(統計已發文)
      m_allPage = 0: m_allItem = 0
      '總頁數:最近一筆進度的頁數
      'Modify By Sindy 2018/5/21 取消 and cp158>0,改為 and cp159=0
      If Val(cp(135)) > 0 Then
         m_allPage = Val(cp(135))
      Else
         strExc(0) = "select cp09,cp10,nvl(cp135,0) from caseprogress" & _
                     " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                     " and cp159=0" & _
                     " and nvl(cp135,0)>0" & _
                     " ORDER BY CP69 DESC,CP70 DESC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_allPage = Val("" & RsTemp.Fields(2))
         End If
      End If
      txtCP135.Text = m_allPage '總頁數
      '總項數:增加項數-刪除未審項數-刪除已審項數
      If Val(cp(136)) > 0 Then
         m_allItem = Val(cp(136))
      Else
         strExc(0) = "select sum(nvl(cp136,0)),sum(nvl(cp137,0)),sum(nvl(cp138,0)) from caseprogress" & _
                     " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                     " and cp159=0" & _
                     " and (nvl(cp136,0)>0 or nvl(cp137,0)>0 or nvl(cp138,0)>0)" & _
                     " ORDER BY CP69 DESC,CP70 DESC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_allItem = Val("" & RsTemp.Fields(0)) - Val("" & RsTemp.Fields(1)) - Val("" & RsTemp.Fields(2))
         End If
      End If
      txtCP136.Text = m_allItem '總項數
      Call txtCP135_Validate(False)
   '2018/5/22 END
   ElseIf cp(10) = 領證及繳年費 Then
      txtCP84.Enabled = False 'Added by Morgan 2022/12/27
      Set m_nFrm = Forms(0).GetForm("frm060104_7") 'Add By Sindy 2023/2/16
      m_nFrm.SetParent Me
      m_nFrm.Hide
      Frame2.Top = 2480: Frame2.Left = 2270
      Frame2.Visible = True
      'Added by Morgan 2022/12/26 112年起領證、延緩公告需分開自申請
      If strSrvDate(1) >= "20230101" Then
         Me.Height = 5500
         Frame3.Top = 3400: Frame3.Left = 220
         Frame3.Visible = True
         Frame4.Top = Frame3.Top + Frame3.Height: Frame4.Left = 220
         Frame4.Visible = True
      End If
      'end 2022/12/26
      'Modified by Morgan 2022/12/27 +m_str412CP09
      If PUB_ChkCPExist(cp, "412", 1, m_str412CP09) Then '有延緩公告  '2020/8/18 modify by sonia 延緩公告預設月數由3個月改6個月(何淑華)
         Me.chk412.Visible = True
         Me.chk412.Enabled = True
         Me.lblCP71.Visible = True
         Me.lblCP71.Enabled = True
         Me.txtCP71.Visible = True
         Me.txtCP71.Enabled = True
      End If
      
      Me.Text7(0).Text = m_nFrm.Text7(0).Text
      Me.Text7(1).Text = m_nFrm.Text7(1).Text
      m_nFrm.ChkDouble 'Added by Morgan 2022/10/7
      Me.Text9.Text = m_nFrm.Text6.Text
      strCaseFee1 = m_nFrm.strCaseFee1
      strCaseFee2 = m_nFrm.strCaseFee2
      m_CP81 = m_nFrm.m_CP81 '可否減免
      
   'Added by Morgan 2022/12/27
   '443 申請證書副本
   ElseIf cp(10) = "443" Then
      txtCP84.Enabled = False
   ElseIf cp(10) = 補換發證書 Then
      If strSrvDate(1) >= "20230101" Then
         txtCP84.Enabled = False
         Frame3.Top = 3380: Frame3.Left = 220
         Frame3.Visible = True
      End If
   'end 2022/12/27
   'Add By Sindy 2018/11/28
   ElseIf cp(10) = 年費 Then
      txtCP84.Enabled = False 'Added by Morgan 2022/12/27
      Frame2.Top = 2480: Frame2.Left = 2270
      Frame2.Visible = True
      Text7(0).Text = ""
      Text7(0).Enabled = True
      Label4.Visible = True
      Text10.Visible = True
      
      m_CP81 = "" '設定案件是否可減免
      If PUB_GetFCPCaseDiscState(pa(1) & pa(2) & pa(3) & pa(4), m_DiscType) Then
         m_CP81 = "Y"
      Else
         m_CP81 = "N"
      End If
      
      m_strNP09_1 = PUB_GetNextFeeDate(pa)
      '若法定期限為假日時, 抓大於法定期限最近的工作天
      If m_strNP09_1 <> "" Then
         m_strNP09_1 = DBDATE(PUB_GetLawDay(DBDATE(m_strNP09_1)))
      End If
      If strSrvDate(1) > m_strNP09_1 Then
         '費用雙倍
         Me.Text9.Text = "Y"
         Me.Text10.Text = "Y" '逾期補繳
      Else
         '取消費用雙倍
         '2005/3/1 已設雙倍不可清除
         'Me.Text9.Text = ""
      End If
      
      strTmp1(0) = strReceiveNo
      For i = 1 To 4
         strTmp1(i) = pa(i)
      Next
      If GetMoneyDate(Val(pa(8)), pa(9), strTmp1, strCaseFee(1), strCaseFee(2)) = True Then
      End If
      '若尚未發證則依公式計算專用期止日
      If pa(25) = "" Then
         If GetMoneyDate(Val(pa(8)) + 10, pa(9), strTmp1, strCaseFee(1), strCaseFee(2), pa(25)) Then '抓專用期起止日
             pa(25) = ChangeWStringToTString(pa(25))
             If GetMoneyDate(pa(8), pa(9), strTmp1, strCaseFee(1), strCaseFee(2)) = True Then
                '舊法新型專用期12年
                If pa(9) = "000" And pa(8) = "2" And Val(pa(14)) > 0 And Val(pa(14)) < 930701 Then
                   strCaseFee(2) = "1,2,3,4,5,6,7,8,9,10,11,12"
                End If
             End If
         End If
      End If
   End If
   
   'Add By Sindy 2018/8/7
   If m_PrevForm.Text6 = "3" Then '電子送件
      Label18(1).Visible = False
      Text8.Visible = False
   End If
   
   'Added by Lydia 2020/02/21 預設「名稱有特殊字」
   FraPA174.Visible = False
   If pa(1) = "FCP" Or pa(1) = "P" Then
       If pa(174) = "Y" Then
          FraPA174.Visible = True
       End If
   End If
   'end 2020/02/21
      
   'Modified by Morgan 2022/12/27 112年起領證可同時辦變更規費可能+300，且原來存檔前都會重算，故規費改為不可修改一律系統計算
   'If cp(10) = "601" And Val(cp(84)) = 0 Then Call CountYearFee 'Added by Morgan 2022/7/13 領證發文規費預設可能不正確,都重算一下
   Call CountYearFee
   'end 2022/12/27
End Sub

Private Sub Text10_GotFocus()
   InverseTextBox Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub TextPA178_GotFocus()
   TextInverse TextPA178
End Sub

Private Sub TextPA178_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "1" And Chr(KeyAscii) <> "2" Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text7(0) <> "" And Text7(1) <> "" Then
      Call CountYearFee '計算年費
   End If
End Sub

Private Sub txtCP135_GotFocus()
   TextInverse txtCP135
   CloseIme
End Sub

Private Sub txtCP135_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP135_Validate(Cancel As Boolean)
   If m_bolChkFee Then
      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                                txtCP135, txtCP136, , txtCP84)
   End If
End Sub

Private Sub txtCP136_GotFocus()
   TextInverse txtCP136
   CloseIme
End Sub

Private Sub txtCP136_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP136_Validate(Cancel As Boolean)
   If m_bolChkFee Then
      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                                txtCP135, txtCP136, , txtCP84)
   End If
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text8_GotFocus()
  TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add by Morgan 2005/8/8
Private Function TxtValidate() As Boolean
   Dim Cancel As Boolean
   
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   If Frame2.Visible = True Then 'Added by Lydia 2020/02/24 有顯示才檢查
        'Add By Sindy 2020/2/7
        Text7_Validate 0, Cancel
        If Cancel = True Then Exit Function
        Text7_Validate 1, Cancel
        If Cancel = True Then Exit Function
        '2020/2/7 END
   End If 'Added by Lydia 2020/02/24
   
   'Added by Morgan 2022/12/26
   If Frame3.Visible Then
      If TextPA178 = "" Then
         MsgBox "請輸入證書形式！", vbExclamation
         TextPA178.SetFocus
         Exit Function
      End If
   End If
   'end 2022/12/26
   
   'Added by Morgan 2025/7/23
   If cp(10) = 年費 Then
      If MsgBox("出申請書同時設定暫不繳Y，若需要改回大批CSV繳納，請將[年費維護] 暫不繳Y清除！" & vbCrLf & vbCrLf & "是否確定要繼續？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   'end 2025/7/23
   
   TxtValidate = True
   
   Call CountYearFee '計算年費
End Function

Private Sub CountYearFee()
   'Add By Sindy 2018/11/28
   If cp(10) = 年費 Then
      '取得領證及繳年費相關費用
      PUB_GetPatentYearFee pa(9), pa(8), "Y00000000", cp(10), Me.Text7(0).Text, Me.Text7(1).Text, IIf(Me.Text9.Text = "Y", True, False), m_CP81, pa(14), strSrvDate(1), m_strOfficalFee, m_strServiceFee, m_lngDisc
      txtCP84 = m_strOfficalFee
   '2018/11/28 END
   ElseIf cp(10) = 領證及繳年費 Then
      m_nFrm.m_lngOfficalFee1 = 0
      m_nFrm.m_lngOfficalFee1Year = 0
      m_nFrm.m_lngFee1 = 0
      m_nFrm.m_lngFee2 = 0
      m_nFrm.m_strOfficalFee = 0
      m_nFrm.m_lngDisc = 0
      m_nFrm.m_lngDisc1Year = 0
      '取得領證及繳年費相關費用
      m_nFrm.GetPatentYearFee pa(9), pa(8), "Y00000000", cp(10), Me.Text7(0).Text, Me.Text7(1).Text, IIf(Me.Text9.Text = "Y", True, False)
      txtCP84 = m_nFrm.m_strOfficalFee 'Add By Sindy 2018/11/15
      If Label122.Visible Then txtCP84 = Val(txtCP84) + 300 'Added by Morgan 2022/12/26
      txtCP84.Enabled = False 'Added by Morgan 2022/12/27
   End If
End Sub

'Add by Morgan 2005/8/8
Private Function FormSave() As Boolean
Dim strConSql As String
Dim stUpdate As String
   
On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   
   'Add By Sindy 2018/5/22
   If Frame1.Visible = True Then
      '檢查此文號是否未發文未取消收文,才需要儲存資料
      strSql = "select cp09" & _
               " From CASEPROGRESS" & _
               " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         '先清除此案號總頁/項數,後面SQL會將總頁/項數儲存在此筆文號中
         strSql = "UPDATE CASEPROGRESS SET cp135=null,cp136=null,cp137=null,cp138=null" & _
                  " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'"
         cnnConnection.Execute strSql
      End If
      stUpdate = stUpdate & ",cp135=" & Val(txtCP135) & _
                  ",cp136=" & Val(txtCP136)
      '2018/5/22 END
   'Add By Sindy 2018/11/8
   ElseIf Frame2.Visible = True Then
      'Add By Sindy 2019/12/13 當601領證及605年費key繳費年度而產生電子送件申請書時，將key的年度自動帶到發文作業的年度。
      If Val(Text7(0)) > 0 And Val(Text7(1)) > 0 Then
         stUpdate = stUpdate & ",cp53=" & Val(Text7(0)) & ",cp54=" & Val(Text7(1))
      End If
      '2019/12/13 END
      
      'Modified by Lydia 2019/01/14 年費不需回寫CP71 (ex.FCP-55649,FCP-55650,FCP-55654)
      'stUpdate = stUpdate & ",cp81=" & CNULL(m_CP81) & _
                  ",cp71=" & CNULL(Val(txtCP71))
      stUpdate = stUpdate & ",cp81=" & CNULL(m_CP81) & _
                     IIf(lblCP71.Visible = True And txtCP71.Visible = True, ", cp71=" & CNULL(Val(txtCP71)), " ")
   
      'Added by Morgan 2020/3/9 "暫不繳納"自動上 "Y"，以防重覆繳納(CSV整批發文)--敏莉
      If cp(10) = 年費 Then
         'Modify By Sindy 2024/5/28 "暫不繳納" 改用獨立欄位存放
         'stUpdate = stUpdate & ",cp141='4'"
         'Modified by Morgan 2025/7/23 還原(此設定是控制不重複繳納並非不送件)--Winfrey
         'stUpdate = stUpdate & ",cp176='Y'"
         stUpdate = stUpdate & ",cp141='4'"
         'end 2025/7/23
      End If
      'end 2020/3/9
   End If
   
   'Modify By Sindy 2018/8/7
   If lstNameAgent.Visible = True Then
      cp(110) = m_CP110
      stUpdate = stUpdate & ",cp110=" & CNULL(m_CP110)
   End If
   stUpdate = stUpdate & ",cp84=" & Val(txtCP84) '發文規費
   If m_PrevForm.Text6 = "3" Then '電子送件
      stUpdate = stUpdate & ",cp118='A'"
   Else
      stUpdate = stUpdate & ",cp118=null"
   End If
   If stUpdate <> "" Then
      stUpdate = Mid(stUpdate, 2)
      'Modify By Sindy 2018/6/19 + and cp158=0 and cp159=0
      strSql = " UPDATE CASEPROGRESS SET " & stUpdate & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql, intI
   End If
   
   'Added by Morgan 2022/12/26
   '證書形式
   If Frame3.Visible = True Then
      strSql = "Update patent Set pa178='" & TextPA178 & "' " & _
            "WHERE pa01='" & pa(1) & "' and pa02='" & pa(2) & "'" & _
             " and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
      cnnConnection.Execute strSql
   End If
   'end 2022/12/26
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrorHandler:
   If Err.Number <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function
'Add by Morgan 2005/8/8
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(5) As String
   Dim ii As Integer

   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   If cp(10) = 自請撤回 And ET03 = "02" And m_CP43 <> "" Then
      strExc(0) = "select cp08 from caseprogress a where cp43 in (select b.cp09 from caseprogress b where b.cp43='" & m_CP43 & "' and b.cp10='404') and cp10='1004' and cp08 is not null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','延期受理機關文號','" & RsTemp(0) & "')"
      End If
      
      strExc(0) = "select 1 from caseprogress a where cp09 in (select b.cp43 from caseprogress b where b.cp09='" & m_CP43 & "' and b.cp10='" & 訴願 & "') and cp10='1002'" & _
         " and exists(select * from caseprogress c where c.cp09=a.cp43 and c.cp10='803')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','舉發','舉發')"
      End If
      
      strExc(0) = "select cp27 from caseprogress a where cp43='" & m_CP43 & "' and cp10='404' and cp27>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','延期發文日','" & RsTemp(0) & "')"
      End If
      
      If Not ClsLawExecSQL(ii, strTxt) Then
         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
      End If
   End If
End Sub

''Add By Sindy 2017/11/8
''申請書
'Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
'Dim strTxt(200) As String, strTmp As String
'Dim ii As Integer, jj As Integer
'
'   EndLetter ET01, strReceiveNo, ET03, strUserNum
'
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
'
'   'Modify By Sindy 2017/11/15
'   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa())
''   For jj = 1 To 5
''      If pa(25 + jj) <> "" Then
''         '申請人
''         strExc(0) = " SELECT C.*,N1.NA72 X1,N2.NA72 X2" & _
''            " FROM CUSTOMER C,NATION N1,NATION N2 WHERE CU01='" & Left(ChangeCustomerL(pa(25 + jj)), 8) & "'" & _
''            " and cu02='" & Mid(ChangeCustomerL(pa(25 + jj)), 9) & "' AND N1.NA01(+)=CU10 AND N2.NA01(+)=CU87"
''         intI = 1
''         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''         If intI = 1 Then
''            ii = ii + 1
''            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-國籍','" & RsTemp("X1") & "')"
''
''            ii = ii + 1
''            If RsTemp("CU10") < "011" Then
''               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-ID','" & RsTemp("CU11") & "')"
''            End If
''
''            ii = ii + 1
''            If RsTemp("CU15") = "0" Then
''               strTmp = "申請人" & jj & "-中文姓名"
''            Else
''               strTmp = "申請人" & jj & "-中文名稱"
''            End If
''            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL("" & RsTemp("CU04")) & "')"
''
''            ii = ii + 1
''            If RsTemp("CU15") = "0" Then
''               strTmp = "申請人" & jj & "-英文姓名"
''            Else
''               strTmp = "申請人" & jj & "-英文名稱"
''            End If
''            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL(RTrim(Trim("" & RsTemp("CU05")) & " " & Trim("" & RsTemp("CU88")) & " " & Trim("" & RsTemp("CU89")) & " " & Trim("" & RsTemp("CU90")))) & "')"
''         End If
''      End If
''   Next
'
'   '出名代理人
'   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      jj = 1
'      Do While Not .EOF
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'         jj = jj + 1
'         .MoveNext
'      Loop
'      End With
'   End If
'
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','頁數','" & IIf(cp(135) = "", "", cp(135)) & "')"
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','項數','" & IIf(cp(136) = "", "", cp(136)) & "')"
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & IIf(cp(84) = "", "", cp(84)) & "')"
'
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件-基本資料表','" & m_CaseNo & ".contact.pdf')"
'
'   'Add By Sindy 2018/1/18
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請修正','否')"
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','一併申請誤譯訂正','否')"
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','同時辦理事項','♀')"
'   '2018/1/18 END
'
'   If Not ClsLawExecSQL(ii, strTxt) Then
'      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
'   Else
'      StartLetter2 = True
'   End If
'End Function

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
   CloseIme
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add By Sindy 2018/11/8
Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub Text7_GotFocus(Index As Integer)
  TextInverse Text7(Index)
End Sub
Private Sub Text7_Validate(Index As Integer, Cancel As Boolean)
Dim i As Integer, bolChk As Boolean, varTmp As Variant
Dim varTmpNICK As Variant, TMPnick060104 As Integer
Dim strNextFeeDate As String '下次繳費日
   
   If Text7(Index) <> "" Then
      'Add By Sindy 2018/11/28
      If cp(10) = 年費 Then
         If Index = 0 Then
            If pa(72) = "" Then
               If Text7(0) <> "1" Then
                  MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
                  Cancel = True
               End If
            Else
               varTmpNICK = Split(pa(72), ",")
               For TMPnick060104 = UBound(varTmpNICK) To 0 Step -1
                  If Trim(varTmpNICK(TMPnick060104)) <> "" Then
                     Exit For
                  End If
               Next TMPnick060104
               If Text7(0) <> Val(varTmpNICK(TMPnick060104)) + 1 Then
                  MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
                  Cancel = True
               End If
            End If
         ElseIf Index = 1 Then
            If ChkRange(Text7(0), Text7(1), "繳費年度") = True Then
               For i = Text7(0) To Text7(1)
                  If InStr(pa(72), Format(i)) > 0 Then
                     bolChk = True
                     Exit For
                  End If
               Next
               If bolChk = True Then
                  MsgBox "繳費年度重覆，請查明後再輸入 !", vbCritical
                  Cancel = True
               Else
                  varTmp = Split(strCaseFee(2), ",")
                  '改判斷繳費迄年是否繳超過專用期
                  strExc(0) = TransDate(CompDate(0, Text7(1) - 1, strCaseFee(1)), 1)
                  If Val(strExc(0)) > Val(pa(25)) Then
                     MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
                     Cancel = True
                  ElseIf Text7(1) = UBound(varTmp) + 1 Then
                     'Text7(7).Text = "" 'Mark by Lydia 2024/07/05 debug; 另外已到最大應繳年度，起迄相同
'                  Else
'                     '原算出的下次繳費日多一天
'                     strNextFeeDate = CompDate(0, Val(varTmp(Val(Text7(1).Text) - 1)), strCaseFee(1))
'                     '避免計算下次繳費日時出錯
'                     If strNextFeeDate <> "" Then
'                        Text7(7).Text = ChangeWDateStringToTString(DateSerial(Left(strNextFeeDate, 4), Mid(strNextFeeDate, 5, 2), Right(strNextFeeDate, 2) - 1))
'                     Else
'                        Text7(7).Text = ""
'                     End If
'                     '若計算出的下次繳費年度>=專用期止日, 則清空下次繳費日(存檔時不產生下一程序)
'                     If Me.Text7(7).Text <> "" Then
'                        If DBDATE(Me.Text7(7).Text) >= DBDATE(pa(25)) Then
'                           Me.Text7(7).Text = ""
'                        End If
'                     End If
                  End If
               End If
            Else
               Cancel = True
            End If
         End If
         
      '領證及繳年費
      Else
         If Index = 1 Then
            If ChkRange(Text7(0), Text7(1), "繳費年度") = True Then
               For i = Text7(0) To Text7(1)
                  If InStr(pa(72), Format(i)) > 0 Then
                     bolChk = True
                     Exit For
                  End If
               Next
               If bolChk = True Then
                  MsgBox "繳費年度錯誤，請查明後再輸入 !", vbCritical
                  Cancel = True
               Else
                  varTmp = Split(strCaseFee2, ",")
                  If Text7(1) > UBound(varTmp) + 1 Then
                     MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
                     Cancel = True
                  ElseIf Text7(1) = UBound(varTmp) + 1 Then
   '                  Text7(7).Text = ""
                  Else
                     If m_CP81 = "Y" And pa(8) = "3" And Val(Text7(1)) < 3 And Val(Text7(1)) <> UBound(varTmp) + 1 Then
                        If UBound(varTmp) + 1 < 3 Then
                           strExc(1) = UBound(varTmp) + 1
                        Else
                           strExc(1) = 3
                        End If
                        'Modified by Morgan 2022/7/12 Ex:FCP-066150--何淑華
                        'MsgBox "繳費年度請輸入 " & strExc(1) & " 以上(可減免客戶1~3年免繳年費)!!"
                        'Cancel = True
                        If MsgBox("申請人為個人或中小企業1~3年可免繳年費，確定只繳" & Val(Text7(1)) & "年？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                           Cancel = True
                        End If
                        'end 2022/7/12
                     End If
                  End If
               End If
            Else
               Cancel = True
            End If
         End If
      End If
      
      Call CountYearFee '計算年費
   Else
      MsgBox "年度不可空白 !", vbCritical
      TextInverse Text7(Index)
   End If
End Sub
'2018/11/8 END

'Added by Lydia 2020/02/21 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()

    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub
