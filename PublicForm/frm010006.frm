VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010006 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   6180
   ClientLeft      =   5550
   ClientTop       =   1545
   ClientWidth     =   9045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9045
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6525
      TabIndex        =   62
      Top             =   60
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5700
      TabIndex        =   61
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7650
      TabIndex        =   60
      Top             =   60
      Width           =   800
   End
   Begin VB.Frame fraWindow1 
      BorderStyle     =   0  '沒有框線
      Height          =   5475
      Left            =   60
      TabIndex        =   21
      Top             =   600
      Width           =   8895
      Begin VB.Frame fraPromoter 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame1"
         Height          =   345
         Left            =   120
         TabIndex        =   56
         Top             =   4770
         Width           =   4035
         Begin MSForms.TextBox txtAdviser 
            Height          =   300
            Index           =   11
            Left            =   990
            TabIndex        =   20
            Top             =   0
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   6
            Size            =   "1931;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPromoter 
            Height          =   300
            Left            =   2130
            TabIndex        =   58
            Top             =   0
            Width           =   1665
            VariousPropertyBits=   27
            Size            =   "2937;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label25 
            Caption         =   "承辦人："
            Height          =   255
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.TextBox txtRecieveCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1050
         TabIndex        =   55
         Top             =   90
         Width           =   1452
      End
      Begin VB.Frame fraWindow2 
         Height          =   2475
         Left            =   60
         TabIndex        =   37
         Top             =   780
         Width           =   8745
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   41
            Top             =   210
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   3175
            MaxLength       =   1
            TabIndex        =   40
            Top             =   210
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1910
            MaxLength       =   6
            TabIndex        =   39
            Top             =   210
            Width           =   1212
         End
         Begin VB.TextBox txtSystem 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1125
            MaxLength       =   3
            TabIndex        =   38
            Top             =   210
            Width           =   732
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "接洽人："
            Height          =   180
            Left            =   5790
            TabIndex        =   54
            Top             =   600
            Width           =   720
         End
         Begin VB.Label Label15 
            Caption         =   "客戶編號5："
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1770
            Width           =   1005
         End
         Begin VB.Label Label12 
            Caption         =   "客戶編號4："
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1485
            Width           =   1005
         End
         Begin VB.Label Label10 
            Caption         =   "客戶編號3："
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1185
            Width           =   1005
         End
         Begin VB.Label Label8 
            Caption         =   "客戶編號2："
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   885
            Width           =   1005
         End
         Begin VB.Label Label4 
            Caption         =   "案件名稱（40）："
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   2130
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "本所案號："
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "客戶編號1："
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   585
            Width           =   1005
         End
         Begin MSForms.TextBox txtAdviser 
            Height          =   300
            Index           =   17
            Left            =   1140
            TabIndex        =   8
            Top             =   1740
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAdviser 
            Height          =   300
            Index           =   16
            Left            =   1140
            TabIndex        =   7
            Top             =   1440
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAdviser 
            Height          =   300
            Index           =   15
            Left            =   1140
            TabIndex        =   6
            Top             =   1140
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAdviser 
            Height          =   300
            Index           =   14
            Left            =   1140
            TabIndex        =   5
            Top             =   840
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAdviser 
            Height          =   300
            Index           =   3
            Left            =   1620
            TabIndex        =   9
            Top             =   2085
            Width           =   6615
            VariousPropertyBits=   671105051
            Size            =   "11663;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox txtAdviser 
            Height          =   300
            Index           =   4
            Left            =   1140
            TabIndex        =   3
            Top             =   540
            Width           =   1095
            VariousPropertyBits=   671105051
            MaxLength       =   9
            Size            =   "1926;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ComboBox cboContact 
            Height          =   315
            Left            =   6570
            TabIndex        =   4
            Top             =   533
            Width           =   1770
            VariousPropertyBits=   679495707
            DisplayStyle    =   7
            Size            =   "3122;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetition 
            Height          =   300
            Index           =   4
            Left            =   2280
            TabIndex        =   46
            Top             =   1755
            Width           =   3555
            VariousPropertyBits=   27
            Size            =   "6271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetition 
            Height          =   300
            Index           =   3
            Left            =   2280
            TabIndex        =   45
            Top             =   1455
            Width           =   3555
            VariousPropertyBits=   27
            Size            =   "6271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetition 
            Height          =   300
            Index           =   2
            Left            =   2280
            TabIndex        =   44
            Top             =   1155
            Width           =   3555
            VariousPropertyBits=   27
            Size            =   "6271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetition 
            Height          =   300
            Index           =   1
            Left            =   2280
            TabIndex        =   43
            Top             =   855
            Width           =   3555
            VariousPropertyBits=   27
            Size            =   "6271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetition 
            Height          =   300
            Index           =   0
            Left            =   2280
            TabIndex        =   42
            Top             =   555
            Width           =   3075
            VariousPropertyBits=   27
            Size            =   "5424;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "有★★的應收帳款簽核控管"
         Height          =   285
         Left            =   4440
         TabIndex        =   19
         Top             =   5100
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "現金或支票"
         Height          =   285
         Left            =   6810
         TabIndex        =   18
         Top             =   4800
         Width           =   1215
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   13
         Left            =   5610
         TabIndex        =   17
         Top             =   4770
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   12
         Left            =   5430
         TabIndex        =   16
         Top             =   4410
         Width           =   2940
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "5186;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   10
         Left            =   5940
         TabIndex        =   14
         Top             =   4050
         Width           =   495
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblSales 
         Height          =   300
         Left            =   2220
         TabIndex        =   59
         Top             =   3690
         Width           =   1935
         VariousPropertyBits=   27
         Size            =   "3413;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   6
         Left            =   2490
         TabIndex        =   11
         Top             =   3330
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   5
         Left            =   1110
         TabIndex        =   10
         Top             =   3330
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   8
         Left            =   1110
         TabIndex        =   13
         Top             =   4050
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   5
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   9
         Left            =   1110
         TabIndex        =   15
         Top             =   4410
         Width           =   1095
         VariousPropertyBits=   671105051
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   7
         Left            =   1110
         TabIndex        =   12
         Top             =   3690
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   0
         Left            =   5115
         TabIndex        =   0
         Top             =   90
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1926;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   1
         Left            =   1050
         TabIndex        =   1
         Top             =   450
         Width           =   600
         VariousPropertyBits=   671105049
         MaxLength       =   4
         Size            =   "1058;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   2
         Left            =   5115
         TabIndex        =   2
         Top             =   450
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "656;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "預定收款日："
         Height          =   180
         Left            =   4440
         TabIndex        =   36
         Top             =   4860
         Width           =   1080
      End
      Begin VB.Label Label30 
         Caption         =   "分所案號："
         Height          =   255
         Left            =   4440
         TabIndex        =   35
         Top             =   4433
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   2280
         X2              =   2400
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label6 
         Caption         =   "聘任期間："
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "是否開電腦收據：           （N：不開)"
         Height          =   255
         Left            =   4440
         TabIndex        =   33
         Top             =   4073
         Width           =   3015
      End
      Begin VB.Label lblDepartment 
         Height          =   255
         Left            =   5220
         TabIndex        =   32
         Top             =   3713
         Width           =   3135
      End
      Begin VB.Label Label18 
         Caption         =   "業務區："
         Height          =   255
         Left            =   4440
         TabIndex        =   31
         Top             =   3713
         Width           =   855
      End
      Begin VB.Label lblCaseProperty 
         Height          =   252
         Left            =   1812
         TabIndex        =   30
         Top             =   552
         Width           =   2172
      End
      Begin VB.Label lblCaseSource 
         Height          =   252
         Left            =   5640
         TabIndex        =   29
         Top             =   528
         Width           =   2772
      End
      Begin VB.Label Label5 
         Caption         =   "案件來源："
         Height          =   255
         Left            =   4110
         TabIndex        =   28
         Top             =   510
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "案件性質："
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "智權人員："
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3713
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "費用："
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   4433
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "郵遞區號："
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   4073
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "收文日："
         Height          =   255
         Left            =   4110
         TabIndex        =   23
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "收文號："
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   150
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm010006"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/04/28 Form2.0已修改(txtAdviser(index)、lblPetition(index)、lblSales、lblPromoter、cboContact
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

'bolLeave判斷離開時，是否要彈出詢問視窗
'LastData上一次存檔時，所輸入之收文日
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim bolLeave As Boolean, LastDate As String, intLeaveKind As Integer
Dim strNation As String
'Add by Morgan 2004/4/15
'是否已觸發 Form Active 事件
Dim bolActive As Boolean
'add by nickc 2007/12/12
Dim IsSaveData As Boolean
Dim strAppNo1 As String '申請人1編號
Dim dblAmt As Double, dblPFee As Double, dblTFee As Double, m_CP150 As String 'Add By Sindy 2012/11/06
Dim dblChkAmt As Double 'Add By Sindy 2012/12/10
'Added by Lydia 2020/02/03
Dim dblCu183 As Double '個人之應收帳款上限
Dim dblAmtR As Double, dblPFeeR As Double, dblTFeeR As Double '關係企業之應收帳款金額
'end 2020/02/03

'Added by Lydia 2019/02/14
Dim m_SalesST15 As String '畫面上智權人員的收文部門
Dim m_Tuser As String '創新業務部預設收文人員
'Added by Lydia 2019/09/16
Dim m_SalesST06 As String '智權人員的所別
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_LOS01 As String '案源總收文號
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '案源總收文號之本所案號
Dim m_LOS02 As String '案源案件類型
Dim m_LOS15 As String '案源單號
Dim m_LOS04 As String  '介紹人
Dim m_LOS04_1 As String, m_LOS04_1st15 As String, m_LOS04_1st06 As String '介紹人(第一位)、收文部門、所別
Dim m_LOS05 As String  '介紹客戶
Dim m_LOS12 As String  '介紹日
Dim m_Los04_N1 As String, m_Los05_N As String  'Added by Lydia 2020/10/05 LA補案源之介紹人(第一位), 介紹客戶
'Mark by Lydia 2022/09/06 改抓特殊設定
'Private Const cnt應收帳款檢查排除 As String = "74018,70005" 'Added by Lydia 2022/06/15 應收帳款上限檢查排除特定人員: 如果人員有異動, 請一併修改接洽單frm090801和收文frm010004~frm010007

'Added by Lydia 2022/09/14 櫃台收文模組化
Private Const 收文存檔模組化啟用日 = 20220928 '完成後先開始使用
Dim modCP() As String, modBase() As String ' 收文 和 基本檔
Dim mType As String, mCaseNo As String  '特殊管制

'Added by Lydia 2022/09/14 設定陣列
Private Sub SetDBArray(ByVal bolReset As Boolean, ByVal pSNo As String, ByVal pCD01 As String, Optional ByVal pCD02 As String, Optional ByVal pCD03 As String, Optional ByVal pCD04 As String)
'pSNo: 現在的收文號 (1碼=新增)
'pCD01~pCD04: 本所案號
Dim intKind As Integer, intWhere As Integer
Dim strTmpA As String

   If bolReset = True Then
      If ClsPDGetSystemKind(pCD01, intKind) = True Then
        Select Case intKind
           Case 專利
              ReDim Preserve modBase(TF_PA) As String
           Case 商標
              ReDim Preserve modBase(TF_TM) As String
           Case 法務
              ReDim Preserve modBase(TF_LC) As String
           Case 顧問
              ReDim Preserve modBase(TF_HC) As String
           Case Else
              ReDim Preserve modBase(tf_SP) As String
        End Select
      End If
      ReDim Preserve modCP(TF_CP) As String
      modBase(1) = pCD01
      modBase(2) = pCD02
      'Added by Lydia 2022/11/11  debug: CFP-029190-0-40收文子案存成母案
      modBase(3) = pCD03
      modBase(4) = pCD04
      'end 2022/11/11
      If pCD01 <> "" And pCD02 <> "" Then
         If modBase(3) = "" Then modBase(3) = "0"
         If modBase(4) = "" Then modBase(4) = "00"
         'Modified by Lydia 2023/05/12 + false
         If PUB_ReadCaseData(modBase, intKind, intWhere, False) = True Then
         End If
      End If
      modCP(1) = modBase(1)
      modCP(2) = modBase(2)
      modCP(3) = modBase(3)
      modCP(4) = modBase(4)
   Else
      If modBase(3) = "" Then
          modBase(3) = "0"
          modCP(3) = "0"
      End If
      If modBase(4) = "" Then
          modBase(4) = "00"
          modCP(4) = "00"
      End If
      '考慮多案收文,再設定一次
      modBase(1) = pCD01
      modBase(2) = pCD02
      modCP(1) = modBase(1)
      modCP(2) = modBase(2)
      modCP(3) = modBase(3)
      modCP(4) = modBase(4)
      '---------------
      modBase(6) = txtAdviser(3)  '案件名稱(中)
      modBase(7) = txtAdviser(12)  '分所案號
      '當事人1~5
      modBase(5) = ChangeCustomerL(txtAdviser(4))
      modBase(24) = ChangeCustomerL(txtAdviser(14))
      modBase(25) = ChangeCustomerL(txtAdviser(15))
      modBase(26) = ChangeCustomerL(txtAdviser(16))
      modBase(27) = ChangeCustomerL(txtAdviser(17))
      
      '申請人聯絡人編號
      If cboContact.Locked = False Then
         If cboContact.ListIndex >= 0 Then
            modBase(23) = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
            If Val(modBase(23)) > 0 Then
            'Add by Morgan 2008/8/7 若個案接洽人與客戶檔的預設接洽人相同時不必設定
               PUB_GetContact modBase(5), strTmpA, True
               If modBase(23) = strTmpA Then
                  modBase(23) = ""
               End If
            '排除空白=00
            ElseIf modBase(23) = "00" And Trim(cboContact.Text) = "" Then
               modBase(23) = ""
            End If
         End If
      End If
      
      modCP(9) = txtRecieveCode  '收文號
      modCP(5) = ChangeTStringToWString(txtAdviser(0)) '收文日
      modCP(10) = Trim(txtAdviser(1)) '案件性質
      modCP(11) = Trim(txtAdviser(2)) '案件來源
      modCP(12) = GetST15(txtAdviser(7))
      modCP(13) = Trim(txtAdviser(7))       '智權人員
      If fraPromoter.Visible = True Then
         modCP(14) = Trim(txtAdviser(11))    '承辦人
      End If
      modCP(16) = txtAdviser(9)    '費用
      'modCP(18) = Val(modCP(16)) / 1000   '點數 ---存檔時計算
      modCP(32) = txtAdviser(10) '是否開電腦收據

       '聘任期間
       modCP(53) = ChangeTStringToWString(txtAdviser(5))
       modCP(54) = ChangeTStringToWString(txtAdviser(6))
      '有★★的應收帳款簽核控管
      If Check2.Visible = True Then
         modCP(150) = IIf(Check2.Value = 1, "Y", "")
      End If
      
      '特殊管制
      mType = "": mCaseNo = ""
      If m_LOS02 <> "" And m_LOS15 <> "" Then
          mType = "LOS案源收文"
          mCaseNo = m_LOS02 & "," & m_LOS15
      End If
   End If
   
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      '2011/4/22 MODIFY BY SONIA 分所智權人員則多一天
      'txtAdviser(13) = PUB_GetWorkDayAfterSysDate(CDbl(txtAdviser(0)) + 19110000, 5)
      If PUB_GetST06(txtAdviser(7)) <> "1" Then
         txtAdviser(13) = PUB_GetWorkDayAfterSysDate(CDbl(txtAdviser(0)) + 19110000, 6)
      Else
         txtAdviser(13) = PUB_GetWorkDayAfterSysDate(CDbl(txtAdviser(0)) + 19110000, 5)
      End If
      '2011/4/22 END
      txtAdviser(13).Locked = True
   Else
      txtAdviser(13).Locked = False
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim varSaveCursor, strAuto1 As String, strAuto2 As String, i As Integer
Dim mBillNo As String, mMemo As String 'Added by Lydia 2019/05/13
Dim bolSaveOK As Boolean  'Added by Lydia 2022/09/14

If Index = 0 Then
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
   End If
   
   m_SalesST15 = GetST15(txtAdviser(7)) 'Added by Lydia 2019/02/14
   
   'Added by Lydia 2020/04/08 檢查案件或智權人員是否為法務部
   If PUB_ChkSalesL(txtSystem, txtAdviser(7).Text) = False Then
        txtAdviser(7).SetFocus
        Call txtAdviser_GotFocus(7)
        Screen.MousePointer = vbDefault
        Exit Sub
   End If
   'end 2020/04/08
   
   'Added by Lydia 2021/09/10 修正畫面所有含跳行符號的文字框; 9/10 FCT-47909收文申請,彼所案號中間有換行
   PUB_FilterFormText Me
      
   strAuto1 = txtRecieveCode
   For i = 0 To 11
          If txtAdviser(i).Enabled And txtAdviser(i).Visible Then
             If CheckKeyIn(i) <> 1 Then
                txtAdviser(i).SetFocus
                txtAdviser_GotFocus (i)
                Exit For
             End If
          End If
   Next
   If i = 12 Then
      strAuto1 = txtRecieveCode
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
      
      'Add By Sindy 2011/7/26 L或LA新案時, 收文程式檢查
      If (txtSystem = "L" Or txtSystem = "LA") And txtCode(0) = "" Then
         '申請人1為X65299謝律師, 未輸入申請人2時
         If Left(Trim(txtAdviser(4)), 6) = "X65299" And Trim(txtAdviser(14)) = "" Then
            Screen.MousePointer = varSaveCursor
            MsgBox "請輸入客戶編號2, 若接洽單未填寫請智權人員補填實際客戶資料！", vbExclamation + vbOKOnly
            txtAdviser(14).SetFocus
            Call txtAdviser_GotFocus(14)
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         '申請人2~5為X65299謝律師時
         If Left(Trim(txtAdviser(14)), 6) = "X65299" Or Left(Trim(txtAdviser(15)), 6) = "X65299" Or _
            Left(Trim(txtAdviser(16)), 6) = "X65299" Or Left(Trim(txtAdviser(17)), 6) = "X65299" Then
            Screen.MousePointer = varSaveCursor
            MsgBox "與謝律師合作案件請於客戶編號1輸入X65299謝智硯律師事務所, 客戶編號2欄填實際客戶資料！", vbExclamation + vbOKOnly
            If Left(Trim(txtAdviser(14)), 6) = "X65299" Then txtAdviser(14).SetFocus: Call txtAdviser_GotFocus(14)
            If Left(Trim(txtAdviser(15)), 6) = "X65299" Then txtAdviser(15).SetFocus: Call txtAdviser_GotFocus(15)
            If Left(Trim(txtAdviser(16)), 6) = "X65299" Then txtAdviser(16).SetFocus: Call txtAdviser_GotFocus(16)
            If Left(Trim(txtAdviser(17)), 6) = "X65299" Then txtAdviser(17).SetFocus: Call txtAdviser_GotFocus(17)
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
      '2011/7/26 End
      
    'add by nickc 2007/11/12 加入檢查特殊客戶
    Dim IsSpecCu As Boolean
    IsSpecCu = False
    If IsSpecCu = False And txtAdviser(4) <> "" Then
        strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtAdviser(4)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtAdviser(4)), 9, 1) & "' "
        CheckOC3
        AdoRecordSet3.CursorLocation = adUseClient
        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If AdoRecordSet3.RecordCount <> 0 Then
            If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                IsSpecCu = True
            End If
        End If
    End If
    'Add By Sindy 2011/1/18
    If IsSpecCu = False And txtAdviser(14) <> "" Then
        strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtAdviser(14)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtAdviser(14)), 9, 1) & "' "
        CheckOC3
        AdoRecordSet3.CursorLocation = adUseClient
        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If AdoRecordSet3.RecordCount <> 0 Then
            If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                IsSpecCu = True
            End If
        End If
    End If
    If IsSpecCu = False And txtAdviser(15) <> "" Then
        strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtAdviser(15)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtAdviser(15)), 9, 1) & "' "
        CheckOC3
        AdoRecordSet3.CursorLocation = adUseClient
        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If AdoRecordSet3.RecordCount <> 0 Then
            If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                IsSpecCu = True
            End If
        End If
    End If
    If IsSpecCu = False And txtAdviser(16) <> "" Then
        strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtAdviser(16)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtAdviser(16)), 9, 1) & "' "
        CheckOC3
        AdoRecordSet3.CursorLocation = adUseClient
        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If AdoRecordSet3.RecordCount <> 0 Then
            If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                IsSpecCu = True
            End If
        End If
    End If
    If IsSpecCu = False And txtAdviser(17) <> "" Then
        strSql = "select cu01,cu02,cu121  from customer Where CU01='" & Mid(ChangeCustomerL(txtAdviser(17)), 1, 8) & "' And CU02='" & Mid(ChangeCustomerL(txtAdviser(17)), 9, 1) & "' "
        CheckOC3
        AdoRecordSet3.CursorLocation = adUseClient
        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If AdoRecordSet3.RecordCount <> 0 Then
            If CheckStr(AdoRecordSet3.Fields("cu121")) = "Y" Then
                IsSpecCu = True
            End If
        End If
    End If
    '2011/1/18 End
    If IsSpecCu Then
      'Modify By Sindy 2023/1/30 排除有輸入案源編號者,已有Flow簽核不需檢查
      If m_LOS15 = "" Then
      '2023/1/30 END
          If MsgBox("請確認此客戶接洽單主管是否核示??", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
              Screen.MousePointer = vbDefault
              Exit Sub
          End If
      End If
    End If
      
      'Add By Sindy 2010/12/31 費用檢查提到存檔前檢查
      '郭 請作單 X14843050 不管
      'Modify By Sindy 2011/1/18 增加客戶編號2,3,4,5檢查
      'modify by sonia 2013/11/19 加X3928904,X69514 葉經理
      'modify by sonia 2014/9/11 取消X69514,已轉外專
      If Mid(txtAdviser(4), 1, 8) <> "X1484305" And Mid(txtAdviser(14), 1, 8) <> "X1484305" And Mid(txtAdviser(15), 1, 8) <> "X1484305" And Mid(txtAdviser(16), 1, 8) <> "X1484305" And Mid(txtAdviser(17), 1, 8) <> "X1484305" And _
         Mid(txtAdviser(4), 1, 8) <> "X3928904" And Mid(txtAdviser(14), 1, 8) <> "X3928904" And Mid(txtAdviser(15), 1, 8) <> "X3928904" And Mid(txtAdviser(16), 1, 8) <> "X3928904" And Mid(txtAdviser(17), 1, 8) <> "X3928904" Then
         'MODIFY BY SONIA 2014/7/17 +傳規費 CFP-027024
         If ClsPDGetCaseFee(txtSystem, txtAdviser(4), txtAdviser(1), Val(txtAdviser(9)), 0) = 0 Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
      
      'Add By Sindy 2011/1/18 檢查客戶編號的輸入順序
      If (Trim(txtAdviser(14)) <> "" And Trim(txtAdviser(4)) = "") Or _
         (Trim(txtAdviser(15)) <> "" And Trim(txtAdviser(14)) = "") Or _
         (Trim(txtAdviser(16)) <> "" And Trim(txtAdviser(15)) = "") Or _
         (Trim(txtAdviser(17)) <> "" And Trim(txtAdviser(16)) = "") Then
         ShowMsg "請依序輸入客戶編號!"
         If Trim(txtAdviser(14)) <> "" And Trim(txtAdviser(4)) = "" Then txtAdviser(14).SetFocus: Call txtAdviser_GotFocus(14)
         If Trim(txtAdviser(15)) <> "" And Trim(txtAdviser(14)) = "" Then txtAdviser(15).SetFocus: Call txtAdviser_GotFocus(15)
         If Trim(txtAdviser(16)) <> "" And Trim(txtAdviser(15)) = "" Then txtAdviser(16).SetFocus: Call txtAdviser_GotFocus(16)
         If Trim(txtAdviser(17)) <> "" And Trim(txtAdviser(16)) = "" Then txtAdviser(17).SetFocus: Call txtAdviser_GotFocus(17)
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      '2011/1/18 End
      
      '2011/4/21 add by sonia
Dim strHC23 As String, strContact As String

      If cboContact.Locked = False Then
         strContact = ""
         If cboContact.ListCount > 2 Then
            'Modified by Lydia 2021/04/28 改成Form 2.0;
            'strHC23 = Format(cboContact.ItemData(cboContact.ListIndex), "00")
            strHC23 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
            PUB_GetContact strAppNo1, strContact, True
            If strHC23 = strContact Or strHC23 = "00" Then
               If MsgBox("請確定接洽人欄是否有為★, 是否要選擇其他接洽人!!", vbYesNo, "警告！") = vbYes Then
                   Screen.MousePointer = varSaveCursor
                   cboContact.SetFocus
                   Exit Sub
               End If
            End If
         End If
      End If
      '2011/4/21 end
      
      'Added by Lydia 2019/05/13 改模組(一併取得)
      If Left(m_SalesST15, 1) <> "F" And txtAdviser(4).Text <> "" And Val(txtAdviser(9).Text) > 0 Then
          'Modified by Lydia 2022/06/13 傳入收文之本所案號,案件性質(可用,串接)
          'Call PUB_GetBillDataAll("3", txtAdviser(4), dblAmt, dblPFee, dblTFee, , , TransDate(txtAdviser(0), 2), mBillNo, mMemo)
          'Modified by Lydia 2022/06/15 傳入收文之智權人員
          Call PUB_GetBillDataAll("3", txtAdviser(4), txtSystem & IIf(txtCode(0) <> "", txtCode(0) & Left(txtCode(1) & "0", 1) & Left(txtCode(2) & "00", 2), ""), txtAdviser(1), Trim(txtAdviser(7)), dblAmt, dblPFee, dblTFee, , , TransDate(txtAdviser(0), 2), mBillNo, mMemo)
      End If
      
      'Add By Sindy 2012/11/06 非T*案件(TF要含)若已送件之應收款超過15萬以上,智權人員非國外部且有費用者須做下列控管
      'Modified by Lydia 2017/06/19 +判斷有申請人編號
      'If (Left(Trim(txtSystem), 1) <> "T" Or Trim(txtSystem) = "TF") And _
         Left(PUB_GetStaffST15(Trim(txtAdviser(7)), "1"), 1) <> "F" And _
         Val(txtAdviser(9)) > 0 And _
         Check2.Value = 0 Then
      'Modified by Lydia 2019/04/08 PUB_GetStaffST15(Trim(txtAdviser(7)), "1") => m_SalesST15
      If (Left(Trim(txtSystem), 1) <> "T" Or Trim(txtSystem) = "TF") And _
         Left(m_SalesST15, 1) <> "F" And _
         Val(txtAdviser(9)) > 0 And _
         Check2.Value = 0 And Trim(txtAdviser(4)) <> "" Then
      'end 2017/06/19
         'Mark by Lydia 2019/05/13 改模組(一併取得)
         'GetBillData txtAdviser(4), dblAmt, dblPFee, dblTFee
         
         'Add By Sindy 2012/12/10 取得客戶應收帳款收文檢查上限
         'Modified by Lydia 2020/02/03 應收帳款上限分開管制為個人"應收帳款上限"和"集團應收帳款上限"
         'dblChkAmt = PUB_GetCustRecAmtLmt(txtAdviser(4))
         '2012/12/10 End
         dblCu183 = PUB_GetCustRecAmtLmt(txtAdviser(4), dblChkAmt)
         'Added by Lydia 2020/02/03 判斷是否有集團上限
         If dblChkAmt = 0 Then
             dblAmtR = 0: dblPFeeR = 0: dblTFeeR = 0
         Else   '有集團上限才抓關係企業的應收帳款金額
             GetBillData txtAdviser(4), dblAmtR, dblPFeeR, dblTFeeR
         End If
         'end 2020/02/03
         
         '已送件之應收款超過30萬以上(不含T*案件應收款),提醒
         'Modify By Sindy 2012/12/10 檢查的30萬改不要固定金額,抓CustRecAmtLmt
         'If dblAmt >= 300000 Then
         'Modified by Lydia 2020/02/03 應收帳款上限分開管制為個人"應收帳款上限"和"集團應收帳款上限"
         'If dblAmt >= dblChkAmt Then
         ''2012/12/10 End
         'Modified by Lydia 2022/06/15 排除特定人員
         'Modified by Lydia 2022/09/06 改抓特殊設定
         'If InStr(cnt應收帳款檢查排除, Trim(txtAdviser(7))) = 0 And dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
         'Modified by Lydia 2022/09/21 案源要判斷介紹人是否在應收帳款上限檢查排除名單內
         'If InStr(Pub_GetSpecMan("應收帳款上限檢查排除"), Trim(txtAdviser(7))) = 0 And dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
         strExc(7) = InStr(Pub_GetSpecMan("應收帳款上限檢查排除"), IIf(m_LOS04_1 <> "", m_LOS04_1, Trim(txtAdviser(7))))
         If Val(strExc(7)) = 0 And dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
         'end 2022/09/21
            'Modified by Lydia 2018/09/20 預設按鈕改成"否" vbDefaultButton1=>vbDefaultButton2
            'Modify By Sindy 2023/1/30 排除有輸入案源編號者,已有Flow簽核不需檢查
            If m_LOS15 = "" Then
            '2023/1/30 END
               If MsgBox("請注意接洽單上是否有註明應收帳款超額，需主管簽核才可收文！是否可收文？" & vbCrLf & _
                         "（接洽單上若有★★的應收帳款簽核控管，是否已勾選畫面上的註記欄位了？）", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
            End If
'         '已送件之應收款超過15萬以上(不含T*案件應收款),提醒
'         ElseIf dblAmt >= 150000 Then
'            If MsgBox("請注意接洽單上是否有註明應收帳款超額，需主管簽核才可收文！是否可收文？" & vbCrLf & _
'                      "（接洽單上若有★★的應收帳款簽核控管，是否已勾選畫面上的註記欄位了？）", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
'               Screen.MousePointer = varSaveCursor
'               Exit Sub
'            End If
         End If
      End If
      '2012/11/06 End
      
      'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
      'Modified by Lydia 2019/04/08 智權人員非國外部
      'If txtAdviser(4).Text <> "" And Val(txtAdviser(9).Text) > 0 Then
      If Left(m_SalesST15, 1) <> "F" And txtAdviser(4).Text <> "" And Val(txtAdviser(9).Text) > 0 Then
         'Modified by Lydia 2019/05/13 改模組(一併取得)
         'If GetBillDate(txtAdviser(4), TransDate(txtAdviser(0), 2), strExc(1), strExc(2)) = True Then
         If mMemo <> "" Then
            'Modified by Lydia 2018/10/29 改訊息
            'If MsgBox("請注意接洽單上是否有註明" & vbCrLf & strExc(2) & vbCrLf & "，請交主管簽核並且有主管簽核。" & vbCrLf & "是否可收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
            'Modify By Sindy 2023/1/30 排除有輸入案源編號者,已有Flow簽核不需檢查
            If m_LOS15 = "" Then
            '2023/1/30 END
               If MsgBox("請注意接洽單上是否有註明" & vbCrLf & mMemo & "，請交主管簽核。" & vbCrLf & "並且有主管簽核，是否可收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Screen.MousePointer = varSaveCursor
                  Exit Sub
               End If
            End If
         End If
      End If
      'end 2018/08/22
      
      'Added by Lydia 2021/02/22 增加檢查重複聘任期間，彈訊息與智權人員確認後，方可收文。 ex. LA-003219於109/1/9已有顧問聘任期間(輸錯109/2/1~111/1/31)，又於110/1/20重複顧問聘任期間
      If txtSystem = "LA" And Len(txtCode(0)) = 6 And txtAdviser(1) = 顧問聘任 Then
        strSql = "select cp53,cp54 from caseprogress Where Cp09 In (" & _
                     "Select Substr(Max(Cp05||Cp09),9,9) Mno From Caseprogress Where Cp01='" & txtSystem & "' And Cp02='" & txtCode(0) & "' And Cp03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' And Cp04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' And Cp10='0' And Cp158=0 And Cp159=0) "
        CheckOC3
        AdoRecordSet3.CursorLocation = adUseClient
        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If AdoRecordSet3.RecordCount <> 0 Then
            If Val("" & AdoRecordSet3.Fields("cp54")) >= Val(DBDATE(txtAdviser(5))) Then
                strExc(1) = "已有聘任期間：" & ChangeWStringToTDateString("" & AdoRecordSet3.Fields("cp53")) & "-" & ChangeWStringToTDateString("" & AdoRecordSet3.Fields("cp54")) & vbCrLf & "請與智權人員聯繫，確認是否繼續收文？"
                If MsgBox(strExc(1), vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                   Screen.MousePointer = varSaveCursor
                   Exit Sub
                End If
            End If
        End If
      End If
      'end 2021/02/22

      'Modified by Lydia 2022/09/14 判斷啟用日
      'If SaveDatabase(strAuto1, strAuto2) Then
      bolSaveOK = False
      If strSrvDate(1) < 收文存檔模組化啟用日 Then
         bolSaveOK = SaveDatabase(strAuto1, strAuto2)
      Else
         Call SetDBArray(False, txtRecieveCode, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
         bolSaveOK = PUB_SaveFrm010006(Me.Name, frm010001.intSaveMode, frm010001.intModifyKind, frm010001.intChoose, modBase, modCP, txtAdviser(8), IsSaveData, mType, mCaseNo)
         
         If frm010001.intModifyKind = 0 And bolSaveOK = True Then
             txtCode(0) = modBase(2)
             strAuto1 = modCP(9)
             strAuto2 = modBase(2)
         End If
      End If
      If bolSaveOK = True Then
      'end 2022/09/14
         Me.Enabled = False
         PUB_SendMailCache 'Add by Sindy 2022/9/29
         frm010001.ClearForm strAuto1, strAuto2
         bolLeave = True
         intLeaveKind = 1
         If frm010001.intModifyKind = 0 Then LastDate = txtAdviser(0).Text
         Unload Me
      End If
   End If
   Screen.MousePointer = varSaveCursor
Else
   If Index = 2 Then
      intLeaveKind = 0
   Else
      intLeaveKind = 1
   End If
   Unload Me
End If
End Sub

Private Sub ReadAdviserDatabaseR()
'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
Dim hc01 As String, hc02 As String, hc03 As String, hc04 As String, hc05 As String, _
              hc06 As String, hc07 As String, cp05 As String, CP10 As String, cp11 As String, cp53 As String, cp54 As String, _
              cp13 As String, cp14 As String, cp16 As String, cp32 As String, cu30 As String, i As Integer, rt As Boolean, _
              hc24 As String, hc25 As String, hc26 As String, hc27 As String
Dim strTemp As String
Dim CP150 As String 'Add By Sindy 2012/11/08

'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
rt = ReadHireDatabase(frm010001.intModifyKind, txtSystem, txtCode(0), _
          IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), hc05, hc06, _
          cp05, txtRecieveCode, CP10, cp11, cp53, cp54, cp13, cp16, cp32, cu30, cp14, hc07, _
          hc24, hc25, hc26, hc27, CP150)
If rt Then
   If frm010001.intModifyKind <> 0 Then
      txtAdviser(0) = cp05
      txtAdviser(2) = cp11
      txtAdviser(5) = cp53
      txtAdviser(6) = cp54
      txtAdviser(7) = cp13
      txtAdviser(8) = cu30
      txtAdviser(9) = cp16
      txtAdviser(10) = cp32
      txtAdviser(11) = cp14
      'txtAdviser(12) = hc07 'Mark by Lydia 2020/03/24 分所案號:後面處理
      CheckKeyIn 7
      CheckKeyIn 2
      CheckKeyIn 11
      txtAdviser(1) = 顧問聘任
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCaseProperty(txtSystem, txtAdviser(1), strTemp) Then
      If ClsPDGetCaseProperty(txtSystem, txtAdviser(1), strTemp) Then
         lblCaseProperty.Caption = strTemp
      End If
      'Add By Sindy 2012/11/08
      If CP150 = "Y" Then
         Me.Check2.Value = 1
      End If
      '2012/11/08 End
   End If
   txtAdviser(3) = hc06
   txtAdviser(4) = hc05
   CheckKeyIn 4
   'Add By Sindy 2011/1/18
   txtAdviser(14) = hc24
   txtAdviser(15) = hc25
   txtAdviser(16) = hc26
   txtAdviser(17) = hc27
   CheckKeyIn 14
   CheckKeyIn 15
   CheckKeyIn 16
   CheckKeyIn 17
   '2011/1/18 End
Else
   If frm010001.intModifyKind <> 0 Then
      MsgBox "讀取資料時發生錯誤!!", vbCritical
      bolLeave = True
      Unload Me
   Else
      txtAdviser(0) = cp05
      txtAdviser(2) = cp11
      txtAdviser(5) = cp53
      txtAdviser(6) = cp54
      'txtAdviser(7) = cp13  '2011/5/11 cancel by sonia 偶而改智權人員收文會忘記打所以不自動帶
      txtAdviser(8) = cu30
      txtAdviser(9) = cp16
      txtAdviser(10) = cp32
      txtAdviser(11) = cp14
      CheckKeyIn 7
      CheckKeyIn 2
      CheckKeyIn 11
      txtAdviser(1) = 顧問聘任
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCaseProperty(txtSystem, txtAdviser(1), strTemp) Then
      If ClsPDGetCaseProperty(txtSystem, txtAdviser(1), strTemp) Then
         lblCaseProperty.Caption = strTemp
      End If
   End If
End If

    'Added by Lydia 2020/03/24 分所案號: 舊案若有分所案號則帶出並鎖住，若無分所案號則開放可輸入並更新回基本檔。
    txtAdviser(12).Locked = False
    txtAdviser(12) = hc07
    If Trim(hc07) <> "" Then
        txtAdviser(12).Locked = True
    End If
    'end 2020/03/24

'NICK 900803 **********************
If frm010001.intChoose = 1 Then
   txtAdviser(2) = "90"
   CheckKeyIn (2)
End If
' **********************
End Sub

Private Function SaveDatabase(ByRef strRecieveAuto As String, ByRef strCaseAuto As String) As Boolean
Dim adoquery As New ADODB.Recordset
Dim strHC23 As String, strContact As String
   
   'Add by Morgan 2008/8/5
   If cboContact.Locked = False Then
      If cboContact.ListIndex >= 0 Then
         'Modified by Lydia 2021/04/28 改成Form 2.0
         'If Val(cboContact.ItemData(cboContact.ListIndex)) > 0 Then
         '   strHC23 = Format(cboContact.ItemData(cboContact.ListIndex), "00")
         strHC23 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
         If Val(strHC23) > 0 Then
         'end 2021/04/28
            'Add by Morgan 2008/8/7 若個案接洽人與客戶檔的預設接洽人相同時不必設定
            PUB_GetContact strAppNo1, strContact, True
            If strHC23 = strContact Then
               strHC23 = ""
            End If
         'Added by Lydia 2022/09/16 排除空白=00
         ElseIf strHC23 = "00" And Trim(cboContact.Text) = "" Then
             strHC23 = ""
         'end 2022/09/16
         End If
      End If
   Else
      strHC23 = "HC23"
   End If
      
   m_SalesST15 = GetST15(txtAdviser(7).Text) 'Added by Lydia 2019/02/14
   
   If frm010001.intModifyKind = 0 Then
      If strHC23 = "HC23" Then strHC23 = "" 'Add by Morgan 2008/8/7
      'Modify By Sindy 2011/1/18 +客戶編號2,3,4,5
      SaveDatabase = InsertHireDatabase(frm010001.intSaveMode, txtSystem, txtCode(0), _
               IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtAdviser(4), txtAdviser(3), Me.txtAdviser(12).Text, txtAdviser(0), txtAdviser(1), _
               txtAdviser(2), txtAdviser(5), txtAdviser(6), txtAdviser(7), txtAdviser(9), txtAdviser(10), txtAdviser(8), txtAdviser(11), strRecieveAuto, strCaseAuto, strHC23, _
               txtAdviser(14), txtAdviser(15), txtAdviser(16), txtAdviser(17))
   Else
      'Modify By Sindy 2011/1/18 +客戶編號2,3,4,5
      SaveDatabase = UpdateHireDatabase(txtSystem, txtCode(0), _
               IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtAdviser(4), txtAdviser(3), Me.txtAdviser(12).Text, txtRecieveCode, txtAdviser(0), txtAdviser(1), _
               txtAdviser(2), txtAdviser(5), txtAdviser(6), txtAdviser(7), txtAdviser(9), txtAdviser(10), txtAdviser(8), txtAdviser(11), strHC23, _
               txtAdviser(14), txtAdviser(15), txtAdviser(16), txtAdviser(17))
   End If
   
   'add by nickc 2007/11/09 測試解決mail 發不到的時候會存兩筆的錯誤
   On Error GoTo 0    '歸零
   'add by nickc 2005/09/05
   If frm010001.intModifyKind = 0 Then
      'Modify By Sindy 2011/1/18
            '當收文業務區與客戶檔業務區不同時發 mail  及提示
            Dim oStrCuSales1 As String
            Dim oStrCuSales2 As String
            Dim oStrCuSales3 As String
            Dim oStrCuSales4 As String
            Dim oStrCuSales5 As String
            Dim oContext As String
            Dim oMailCount As String
            '秀玲說，其中一個符合就不發了
            Dim IsMail As Boolean
            IsMail = True
            
            oStrCuSales1 = ""
            oStrCuSales2 = ""
            oStrCuSales3 = ""
            oStrCuSales4 = ""
            oStrCuSales5 = ""
            
            'Added by Lydia 2020/10/05 (9/30) 若該收文號點數>0但無案源(自行收文者)時，若案件的客戶為非法律所的客戶時則為A3類案源，不論新舊案，系統自動新增TT-999999案進度(B類收文)及法律所案源資料。若為新案業務區不同的Email照舊通知。
            If m_Los05_N <> "" Then  '因為櫃台無法處理,所以只發email
                m_LOS05 = m_Los05_N
                m_LOS04_1 = m_Los04_N1
                m_LOS04_1st15 = GetST15(m_LOS04_1)
            End If
            'end 2020/10/05
            
        'Modified by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
        'If strSrvDate(1) < 智慧所更名日 Then 'Added by Lydia 2020/03/24 智慧所更名日:並請取消智權人員與客戶檔智權人員檢查的控制。
        'Modified by Lydia 2022/11/03 debug:非案源(補案源不算)改用畫面判斷
        'If strSrvDate(1) >= 法律所案源收文啟用日 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
        If Trim(frm010001.txtLOS15) = "" Then
            oContext = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtAdviser(3) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtAdviser(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
            
            oMailCount = ""
            'Modified by Lydia 2019/02/14
            'If GetST15(txtAdviser(7).Text) <> GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(4).Text) <> "" Then
            '   If Left(Trim(GetST15(txtAdviser(7).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1)), 1) = "F" Then
            If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(4).Text) <> "" Then
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1)), 1) = "F" Then
            'end 2019/02/14
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales1 & ";"
                  oContext = oContext & vbCrLf + "客戶編號1： " + GetCustomerName(ChangeCustomerL(txtAdviser(4).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales1)
               End If
             '秀玲說，其中一個符合就不發了
             Else
                   If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(4).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtAdviser(4)) <> "" And Trim(txtAdviser(7)) <> "" Then
                If PUB_ChkOldCustomer(True, txtAdviser(4), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If

            'Modified by Lydia 2019/02/14
            'If GetST15(txtAdviser(7).Text) <> GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales2) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(14).Text) <> "" Then
            '   If Left(Trim(GetST15(txtAdviser(7).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales2)), 1) = "F" Then
            If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales2) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(14).Text) <> "" Then
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales2)), 1) = "F" Then
            'end 2019/02/14
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales2 & ";"
                  oContext = oContext & vbCrLf + "客戶編號2： " + GetCustomerName(ChangeCustomerL(txtAdviser(14).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales2)
               End If
             '秀玲說，其中一個符合就不發了
             Else
                   If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(14).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtAdviser(14)) <> "" And Trim(txtAdviser(7)) <> "" Then
                If PUB_ChkOldCustomer(True, txtAdviser(14), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If

            'Modified by Lydia 2019/02/14
            'If GetST15(txtAdviser(7).Text) <> GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales3) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(15).Text) <> "" Then
            '   If Left(Trim(GetST15(txtAdviser(7).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales3)), 1) = "F" Then
            If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales3) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(15).Text) <> "" Then
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales3)), 1) = "F" Then
            'end 2019/02/14
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales3 & ";"
                  oContext = oContext & vbCrLf + "客戶編號3： " + GetCustomerName(ChangeCustomerL(txtAdviser(15).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales3)
               End If
             '秀玲說，其中一個符合就不發了
             Else
                   If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(15).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtAdviser(15)) <> "" And Trim(txtAdviser(7)) <> "" Then
                If PUB_ChkOldCustomer(True, txtAdviser(15), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If

            'Modified by Lydia 2019/02/14
            'If GetST15(txtAdviser(7).Text) <> GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales4) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(16).Text) <> "" Then
            '   If Left(Trim(GetST15(txtAdviser(7).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales4)), 1) = "F" Then
            If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales4) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(16).Text) <> "" Then
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales4)), 1) = "F" Then
            'end 2019/02/14
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales4 & ";"
                  oContext = oContext & vbCrLf + "客戶編號4： " + GetCustomerName(ChangeCustomerL(txtAdviser(16).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales4)
               End If
             Else
                   If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(16).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtAdviser(16)) <> "" And Trim(txtAdviser(7)) <> "" Then
                If PUB_ChkOldCustomer(True, txtAdviser(16), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If

            'Modified by Lydia 2019/02/14
            'If GetST15(txtAdviser(7).Text) <> GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales5) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(17).Text) <> "" Then
            '   If Left(Trim(GetST15(txtAdviser(7).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales5)), 1) = "F" Then
            If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales5) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(17).Text) <> "" Then
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales5)), 1) = "F" Then
            'end 2019/02/14
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales5 & ";"
                  oContext = oContext & vbCrLf + "客戶編號5： " + GetCustomerName(ChangeCustomerL(txtAdviser(17).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales5)
               End If
             Else
                   If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(17).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            'Added by Lydia 2019/09/16 檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtAdviser(17)) <> "" And Trim(txtAdviser(7)) <> "" Then
                If PUB_ChkOldCustomer(True, txtAdviser(17), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If
        End If 'Added by Lydia 2022/11/03
        
        'Added by Lydia 2022/11/03 debug:區分案源要用LOS04, 補案源不算
        If strSrvDate(1) >= 法律所案源收文啟用日 And Trim(frm010001.txtLOS15) <> "" And m_LOS05 <> "" And m_LOS04_1 <> "" Then
            oContext = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtAdviser(3) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtAdviser(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
            oMailCount = ""
        'end 2022/11/03
            If txtAdviser(4) <> "" Then
                If ChkSameCuArea(Trim(txtAdviser(4)), m_LOS04_1) = False Then
                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1)), 1) = "F" Then
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
                    Else
                        oMailCount = oMailCount & oStrCuSales1 & ";"
                        oContext = oContext & vbCrLf + "客戶編號1： " + GetCustomerName(ChangeCustomerL(txtAdviser(4).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales1)
                    End If
                Else
                       IsMail = False
                End If
                '檢查是否為待活化客戶
                If PUB_ChkOldCustomer(True, txtAdviser(4), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                    IsMail = False
                End If
            End If

            If txtAdviser(14) <> "" Then
                If ChkSameCuArea(Trim(txtAdviser(14)), m_LOS04_1) = False Then
                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales2)), 1) = "F" Then
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
                    Else
                        oMailCount = oMailCount & oStrCuSales2 & ";"
                        oContext = oContext & vbCrLf + "客戶編號2 " + GetCustomerName(ChangeCustomerL(txtAdviser(14).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales2)
                    End If
                Else
                       IsMail = False
                End If
                '檢查是否為待活化客戶
                If PUB_ChkOldCustomer(True, txtAdviser(14), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                    IsMail = False
                End If
            End If

            If txtAdviser(15) <> "" Then
                If ChkSameCuArea(Trim(txtAdviser(15)), m_LOS04_1) = False Then
                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales3)), 1) = "F" Then
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
                    Else
                        oMailCount = oMailCount & oStrCuSales3 & ";"
                        oContext = oContext & vbCrLf + "客戶編號3： " + GetCustomerName(ChangeCustomerL(txtAdviser(15).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales3)
                    End If
                Else
                       IsMail = False
                End If
                '檢查是否為待活化客戶
                If PUB_ChkOldCustomer(True, txtAdviser(15), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                    IsMail = False
                End If
            End If

            If txtAdviser(16) <> "" Then
                If ChkSameCuArea(Trim(txtAdviser(16)), m_LOS04_1) = False Then
                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales4)), 1) = "F" Then
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
                    Else
                        oMailCount = oMailCount & oStrCuSales4 & ";"
                        oContext = oContext & vbCrLf + "客戶編號4： " + GetCustomerName(ChangeCustomerL(txtAdviser(16).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales4)
                    End If
                Else
                       IsMail = False
                End If
                '檢查是否為待活化客戶
                If PUB_ChkOldCustomer(True, txtAdviser(16), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                    IsMail = False
                End If
            End If
            
            If txtAdviser(17) <> "" Then
                If ChkSameCuArea(Trim(txtAdviser(17)), m_LOS04_1) = False Then
                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales5)), 1) = "F" Then
                        '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
                    Else
                        oMailCount = oMailCount & oStrCuSales5 & ";"
                        oContext = oContext & vbCrLf + "客戶編號5： " + GetCustomerName(ChangeCustomerL(txtAdviser(17).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales5)
                    End If
                Else
                       IsMail = False
                End If
                '檢查是否為待活化客戶
                If PUB_ChkOldCustomer(True, txtAdviser(17), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                    IsMail = False
                End If
            End If
            'end 2020/05/20
            
            '若申請人全空白，不發
            If IsMail = False Or (Trim(txtAdviser(4)) = "" And Trim(txtAdviser(14)) = "" And Trim(txtAdviser(15)) = "" And Trim(txtAdviser(16)) = "" And Trim(txtAdviser(17)) = "") Then
                 oMailCount = ""
            End If
        End If 'Added by Lydia 2020/03/24
        
            'TXTSYSTEM只判斷1碼,因為FG
            If UCase(Mid(txtSystem, 1, 1)) <> "F" And oMailCount <> "" Then
               '申請人為 X65299 或 X03072 的所有關係企業都不檢查業務區
               If Left(Trim(txtAdviser(4)), 6) <> "X65299" And Left(Trim(txtAdviser(4)), 6) <> "X03072" And _
                  Left(Trim(txtAdviser(14)), 6) <> "X65299" And Left(Trim(txtAdviser(14)), 6) <> "X03072" And _
                  Left(Trim(txtAdviser(15)), 6) <> "X65299" And Left(Trim(txtAdviser(15)), 6) <> "X03072" And _
                  Left(Trim(txtAdviser(16)), 6) <> "X65299" And Left(Trim(txtAdviser(16)), 6) <> "X03072" And _
                  Left(Trim(txtAdviser(17)), 6) <> "X65299" And Left(Trim(txtAdviser(17)), 6) <> "X03072" Then
                  
                   '加發秀玲
                  'Modified by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
                  'Added by Lydia 2022/11/03
                  If m_LOS05 = "" Or m_LOS04_1 = "" Then
                       MsgBox "收文智權人員與客戶智權人員不同業務區，準備發 mail ！", , "注意！"
                       oMailCount = oMailCount & Trim(txtAdviser(7).Text) & ";83002"
                       oMailCount = oMailCount & PUB_ChkForLawMan(Trim(txtAdviser(4)), txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
                       oContext = oContext & vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！"
                  Else
                  'end 2022/11/03
                      MsgBox "案源介紹人員與客戶智權人員不同業務區，準備發 mail ！", , "注意！"
                      'Modified by Lydia 2022/07/15 通知法律所的智權人員沒有意義，應該要改為案源介紹人員. ex.L-006547
                      oMailCount = oMailCount & m_LOS04_1 & ";83002"
                      oContext = oContext & vbCrLf + "案源介紹人員： " + GetStaffName(m_LOS04_1) + vbCrLf + vbCrLf + "智權人員(區)不同！"
                  End If 'Added by Lydia 2022/11/03
                  PUB_SendMail strUserNum, oMailCount, "", "案件收文通知--此案收文非原智權人員(區)！", oContext
               End If
            End If
   End If

End Function

Private Sub Form_Activate()
   'Add by Morgan 2004/4/15
   If bolActive Then
      Exit Sub
   Else
      bolActive = True
   End If
   
Dim strKindName As String

Me.Refresh

'根據intModifyMode來調整fraWindow1 , fraWindow2
Select Case frm010001.intModifyKind
             Case 0
                        '新增：所有欄位皆可輸入
                        fraWindow1.Enabled = True
                        Select Case frm010001.intSaveMode
                                     Case 0
                                                fraWindow2.Enabled = False
                                     Case 1
                                                fraWindow2.Enabled = True
                        End Select
                        If LastDate = "" Then
                           txtAdviser(0).Text = GetTaiwanTodayDate
                        Else
                           txtAdviser(0).Text = LastDate
                        End If
                        txtAdviser_GotFocus 0
             Case 1
                        '修改：中間欄位不可輸入
                        fraWindow1.Enabled = True
                        Dim bolNew As Boolean
                        'edit by nickc 2007/02/06 不用 dll 了
                        'If obj001.IsNewCase(txtRecieveCode, bolNew) Then
                        If Cls001IsNewCase(txtRecieveCode, bolNew) Then
                           If bolNew Then
                              fraWindow2.Enabled = True
                           Else
                              fraWindow2.Enabled = False
                           End If
                        Else
                           bolLeave = True
                           Unload Me
                           Exit Sub
                        End If
             Case 2
                        '刪除：所有欄位皆不可輸入
                        cmdOK(0).Visible = False
                        fraWindow1.Enabled = False
                        fraWindow2.Enabled = False
End Select
If frm010001.intModifyKind <> 0 Or frm010001.intSaveMode <> 1 Then
   ReadAdviserDatabaseR
End If

Call ReadLOS 'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔

   'Added by Lydia 2022/09/14
   If strSrvDate(1) >= 收文存檔模組化啟用日 Then
       Call SetDBArray(True, txtRecieveCode, txtSystem, txtCode(0), txtCode(1), txtCode(2))
   End If
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   If frm010001.intChoose = 1 Then
      fraPromoter.Visible = True
   Else
      fraPromoter.Visible = False
   End If
   'add by nickc 2007/12/12
   IsSaveData = False
   'Add by Morgan 2008/8/5
   If frm010001.m_blnNewCase = True Then
      cboContact.Locked = False
   Else
      cboContact.Locked = True
   End If
   'end 2008/8/5
   
   'Added by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
   Label32.Visible = False
   txtAdviser(13).Visible = False
   Check1.Visible = False
   
   fraPromoter.BackColor = &H8000000F 'Added by Lydia 2021/06/09
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If frm010001.intModifyKind = 0 Or frm010001.intModifyKind = 1 Then
      If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Cancel = 1
      End If
   End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

PUB_SendMailCache 'Added by Lydia 2020/05/20

Where01ToGo intLeaveKind
intLeaveKind = 0
'Add By Cheng 2002/07/18
'Set frm010006 = Nothing 'Remove by Lydia 2021/12/13 Form2.0會有問題，改在呼叫時清除記憶體變數
stChkForm = Me.Name 'Add by Amy 2021/12/21
End Sub

'Remove by Lydia 2021/04/28 Form 2.0的Label沒有Change模組
'Private Sub lblPetition_Change(Index As Integer)
''Add By Cheng 2002/01/24
'If Me.txtSystem.Text = "LA" Then
'   If frm010001.intModifyKind = 0 Then '新增狀態
'      Me.txtAdviser(3).Text = Me.lblPetition(0).Caption
'   End If
'End If
'End Sub
'end 2021/04/28

Private Sub txtAdviser_Change(Index As Integer)
Select Case Index
             Case 2
                        lblCaseSource.Caption = ""
             Case 4 '客戶編號1
                        lblPetition(0).Caption = ""
                        txtAdviser(8).Text = ""
             Case 7
                        lblSales.Caption = ""
                        lblDepartment = ""
                        m_SalesST15 = "" 'Added by Lydia 2019/02/14
             'Add By Sindy 2011/1/18
             Case 14, 15, 16, 17 '客戶編號2,3,4,5
                        lblPetition(Index - 13).Caption = ""
End Select
End Sub

Private Sub txtAdviser_Validate(Index As Integer, Cancel As Boolean)

'add by nick 2005/01/04 智權人員
If Index = 7 Then
   If txtAdviser(Index).Text <> "" And txtAdviser(Index) < "63001" Then
      MsgBox "智權人員不可小於 63001！", , "注意！"
      Cancel = True
      Exit Sub
   End If
   'Modify By Sindy 2011/1/18
   '因為之前的 智權人員並沒有抓
   Dim strTemp As String, strTemp1 As String
   If Not ClsPDGetStaff(txtAdviser(Index).Text, strTemp, strTemp1) Then
       Cancel = True
       Exit Sub
   End If
   'Modified by Lydia 2019/02/14
   'GetST15 txtAdviser(Index).Text, strTemp1
   m_SalesST15 = GetST15(txtAdviser(Index).Text, strTemp1)
   lblSales.Caption = strTemp
   lblDepartment = strTemp1
   
   'Added by Lydia 2020/04/08 檢查案件或智權人員是否為法務部
   If PUB_ChkSalesL(txtSystem, txtAdviser(Index).Text) = False Then
   End If
   'end 2020/04/08
   
   'Added by Lydia 2019/02/14 創新業務部人員收文控管
   If PUB_ChkIsT10T20("2", txtAdviser(Index).Text, m_Tuser, strTemp) = True Then
        txtAdviser(Index) = m_Tuser
        lblSales.Caption = strTemp
        txtAdviser(Index).SetFocus
        Call txtAdviser_GotFocus(Index)
        Cancel = True
        Exit Sub
   End If
   'end 2019/02/14
   
   'Added by Lydia 2020/03/24 智慧所更名日起檢查智權人員非法律所人員不可收文
   'Remove by Lydia 2020/05/29 重複判斷
   'If strSrvDate(1) >= 智慧所更名日 And Me.txtAdviser(Index).Text <> "" Then
   '    If PUB_ChkLCompStaff(Me.txtAdviser(Index).Text) = False Then
   '         MsgBox "智權人員非法律所人員不可收文！", , "注意！"
   '         Cancel = True
   '         Exit Sub
   '    End If
   'End If
   'end 2020/03/24
   
        '當收文業務區與客戶檔業務區不同時發 mail  及提示
        Dim oStrCuSales1 As String
        Dim oStrCuSales2 As String
        Dim oStrCuSales3 As String
        Dim oStrCuSales4 As String
        Dim oStrCuSales5 As String
        Dim oMailCount As String
        '秀玲說，其中一個符合就不發了
        Dim IsMail As Boolean
        IsMail = True
        oStrCuSales1 = ""
        oStrCuSales2 = ""
        oStrCuSales3 = ""
        oStrCuSales4 = ""
        oStrCuSales5 = ""
        oMailCount = ""
   'Modified by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
   'If strSrvDate(1) < 智慧所更名日 Then 'Added by Lydia 2020/03/24 智慧所更名日:並請取消智權人員與客戶檔智權人員檢查的控制。
'        'Modified by Lydia 2019/02/14
'        'If GetST15(txtAdviser(7).Text) <> GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(4).Text) <> "" Then
'        If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(4).Text) <> "" Then
'        '秀玲說，其中一個符合就不發了
'        Else
'              If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(4).Text) <> "" Then
'                  IsMail = False
'              End If
'        End If
'        'Added by Lydia 2019/09/16 檢查是否為待活化客戶
'        If m_SalesST06 <> "" And Trim(txtAdviser(4)) <> "" And Trim(txtAdviser(7)) <> "" Then
'            If PUB_ChkOldCustomer(False, txtAdviser(4), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
'               IsMail = False
'            End If
'        End If
'
'        'Modified by Lydia 2019/02/14
'        'If GetST15(txtAdviser(7).Text) <> GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales2) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(14).Text) <> "" Then
'        If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales2) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(14).Text) <> "" Then
'        '秀玲說，其中一個符合就不發了
'        Else
'              If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(14).Text) <> "" Then
'                  IsMail = False
'              End If
'        End If
'        'Added by Lydia 2019/09/16 檢查是否為待活化客戶
'        If m_SalesST06 <> "" And Trim(txtAdviser(14)) <> "" And Trim(txtAdviser(7)) <> "" Then
'            If PUB_ChkOldCustomer(False, txtAdviser(14), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
'               IsMail = False
'            End If
'        End If
'
'        'Modified by Lydia 2019/02/14
'        'If GetST15(txtAdviser(7).Text) <> GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales3) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(15).Text) <> "" Then
'        If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales3) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(15).Text) <> "" Then
'        '秀玲說，其中一個符合就不發了
'        Else
'              If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(15).Text) <> "" Then
'                  IsMail = False
'              End If
'        End If
'        'Added by Lydia 2019/09/16 檢查是否為待活化客戶
'        If m_SalesST06 <> "" And Trim(txtAdviser(15)) <> "" And Trim(txtAdviser(7)) <> "" Then
'            If PUB_ChkOldCustomer(False, txtAdviser(15), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
'               IsMail = False
'            End If
'        End If
'
'        'Modified by Lydia 2019/02/14
'        'If GetST15(txtAdviser(7).Text) <> GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales4) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(16).Text) <> "" Then
'        If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales4) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(16).Text) <> "" Then
'        '秀玲說，其中一個符合就不發了
'        Else
'              If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(16).Text) <> "" Then
'                  IsMail = False
'              End If
'        End If
'        'Added by Lydia 2019/09/16 檢查是否為待活化客戶
'        If m_SalesST06 <> "" And Trim(txtAdviser(16)) <> "" And Trim(txtAdviser(7)) <> "" Then
'            If PUB_ChkOldCustomer(False, txtAdviser(16), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
'               IsMail = False
'            End If
'        End If
'
'        'Modified by Lydia 2019/02/14
'        'If GetST15(txtAdviser(7).Text) <> GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales5) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(17).Text) <> "" Then
'        If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales5) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(17).Text) <> "" Then
'        '秀玲說，其中一個符合就不發了
'        Else
'              If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(17).Text) <> "" Then
'                  IsMail = False
'              End If
'        End If
'        'Added by Lydia 2019/09/16 檢查是否為待活化客戶
'        If m_SalesST06 <> "" And Trim(txtAdviser(17)) <> "" And Trim(txtAdviser(7)) <> "" Then
'            If PUB_ChkOldCustomer(False, txtAdviser(17), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
'               IsMail = False
'            End If
'        End If
   If strSrvDate(1) >= 法律所案源收文啟用日 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
        If txtAdviser(4) <> "" Then
            If ChkSameCuArea(Trim(txtAdviser(4)), m_LOS04_1) = False Then
            Else
                   IsMail = False
            End If
            '檢查是否為待活化客戶
            If PUB_ChkOldCustomer(True, txtAdviser(4), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                IsMail = False
            End If
        End If
        If txtAdviser(14) <> "" Then
            If ChkSameCuArea(Trim(txtAdviser(14)), m_LOS04_1) = False Then
            Else
                   IsMail = False
            End If
            '檢查是否為待活化客戶
            If PUB_ChkOldCustomer(True, txtAdviser(14), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                IsMail = False
            End If
        End If
        If txtAdviser(15) <> "" Then
            If ChkSameCuArea(Trim(txtAdviser(15)), m_LOS04_1) = False Then
            Else
                   IsMail = False
            End If
            '檢查是否為待活化客戶
            If PUB_ChkOldCustomer(True, txtAdviser(15), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                IsMail = False
            End If
        End If
        If txtAdviser(16) <> "" Then
            If ChkSameCuArea(Trim(txtAdviser(16)), m_LOS04_1) = False Then
            Else
                   IsMail = False
            End If
            '檢查是否為待活化客戶
            If PUB_ChkOldCustomer(True, txtAdviser(16), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                IsMail = False
            End If
        End If
        If txtAdviser(17) <> "" Then
            If ChkSameCuArea(Trim(txtAdviser(17)), m_LOS04_1) = False Then
            Else
                   IsMail = False
            End If
            '檢查是否為待活化客戶
            If PUB_ChkOldCustomer(True, txtAdviser(17), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06) = True Then
                IsMail = False
            End If
        End If
        'end 2020/05/20
        If UCase(Mid(txtSystem, 1, 1)) <> "F" And IsMail = True And (txtAdviser(4) <> "" Or txtAdviser(14) <> "" Or txtAdviser(15) <> "" Or txtAdviser(16) <> "" Or txtAdviser(17) <> "") Then
             '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail，不顯示訊息
             oMailCount = ""
             'Modified by Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
'             If txtAdviser(4) <> "" Then
'                'Modified by Lydia 2019/02/14
'                'If Left(Trim(GetST15(txtAdviser(7).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1)), 1) = "F" Then
'                If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1)), 1) = "F" Then
'                Else
'                   oMailCount = "Y"
'                End If
'             End If
'             If txtAdviser(14) <> "" Then
'                'Modified by Lydia 2019/02/14
'                'If Left(Trim(GetST15(txtAdviser(7).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales1)), 1) = "F" Then
'                If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales1)), 1) = "F" Then
'                Else
'                   oMailCount = "Y"
'                End If
'             End If
'             If txtAdviser(15) <> "" Then
'                'Modified by Lydia 2019/02/14
'                'If Left(Trim(GetST15(txtAdviser(7).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales1)), 1) = "F" Then
'                If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales1)), 1) = "F" Then
'                Else
'                   oMailCount = "Y"
'                End If
'             End If
'             If txtAdviser(16) <> "" Then
'                'Modified by Lydia 2019/02/14
'                'If Left(Trim(GetST15(txtAdviser(7).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales1)), 1) = "F" Then
'                If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales1)), 1) = "F" Then
'                Else
'                   oMailCount = "Y"
'                End If
'             End If
'             If txtAdviser(17) <> "" Then
'                'Modified by Lydia 2019/02/14
'                'If Left(Trim(GetST15(txtAdviser(7).Text)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales1)), 1) = "F" Then
'                If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales1)), 1) = "F" Then
'                Else
'                   oMailCount = "Y"
'                End If
'             End If
            If txtAdviser(4) <> "" Then
               If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1)), 1) = "F" Then
               Else
                  oMailCount = "Y"
               End If
            End If
            If txtAdviser(14) <> "" Then
               If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales1)), 1) = "F" Then
               Else
                  oMailCount = "Y"
               End If
            End If
            If txtAdviser(15) <> "" Then
               If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales1)), 1) = "F" Then
               Else
                  oMailCount = "Y"
               End If
            End If
            If txtAdviser(16) <> "" Then
               If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales1)), 1) = "F" Then
               Else
                  oMailCount = "Y"
               End If
            End If
            If txtAdviser(17) <> "" Then
               If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales1)), 1) = "F" Then
               Else
                  oMailCount = "Y"
               End If
            End If
             'end 2020/05/20
             
             If Trim(oMailCount) <> "" Then
                '申請人為 X65299 或 X03072 的所有關係企業都不檢查業務區
                If Left(Trim(txtAdviser(4)), 6) <> "X65299" And Left(Trim(txtAdviser(4)), 6) <> "X03072" And _
                   Left(Trim(txtAdviser(14)), 6) <> "X65299" And Left(Trim(txtAdviser(14)), 6) <> "X03072" And _
                   Left(Trim(txtAdviser(15)), 6) <> "X65299" And Left(Trim(txtAdviser(15)), 6) <> "X03072" And _
                   Left(Trim(txtAdviser(16)), 6) <> "X65299" And Left(Trim(txtAdviser(16)), 6) <> "X03072" And _
                   Left(Trim(txtAdviser(17)), 6) <> "X65299" And Left(Trim(txtAdviser(17)), 6) <> "X03072" Then
                   'Modified by Lydia 2020/05/20 Lydia 2020/05/20 法律所案源收文：若介紹客戶為舊客戶但與介紹人不同區時發Mail通知相關人員
                   'MsgBox "收文智權人員與客戶智權人員不同業務區！", , "注意！"
                   MsgBox "案源介紹人員與客戶智權人員不同業務區！", , "注意！"
                End If
             End If
        End If
        '2011/1/18 End
   End If 'Added by Lydia 2020/03/24
End If

'add by nick 2005/01/18 客戶編號1
'Modify By Sindy 2011/1/18 客戶編號2,3,4,5
Dim strCol As String
If Index = 4 Or Index = 14 Or Index = 15 Or Index = 16 Or Index = 17 Then
   If Index = 4 Then strCol = "hc05"
   If Index = 14 Then strCol = "hc24"
   If Index = 15 Then strCol = "hc25"
   If Index = 16 Then strCol = "hc26"
   If Index = 17 Then strCol = "hc27"
   '檢查客戶是否有案子
   strSql = "select max(hc02) from hirecase where " & strCol & "='" & txtAdviser(Index).Text & "' "
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .MaxRecords <> 0 Then
         If txtCode(0).Text <> CheckStr(.Fields(0).Value) Then
            MsgBox "此客戶顧問號錯誤，應為 LA-" & CheckStr(.Fields(0).Value) & " "
            Cancel = True
         End If
      End If
   End With
End If

If CheckKeyIn(Index) = -1 Then
   Cancel = True
   txtAdviser(Index).SetFocus 'Added by Lydia 2021/06/09
   txtAdviser_GotFocus (Index)
'Added by Lydia 2021/06/09
Else
   If Index = 2 Then
       
   End If
End If
End Sub

Private Function CheckKeyIn(ByRef intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, strCusTemp As String
Static strLastCus As String

CheckKeyIn = -1
Select Case intIndex
             Case 0, 5
                        If CheckIsTaiwanDate(txtAdviser(intIndex).Text) Then
                            CheckKeyIn = 1
                        End If
             Case 2
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCaseSource(txtAdviser(intIndex).Text, strTemp) Then
                        If ClsPDGetCaseSource(txtAdviser(intIndex).Text, strTemp) Then
                           lblCaseSource.Caption = strTemp
                           CheckKeyIn = 1
                        End If
             Case 3
                        If txtAdviser(intIndex) = "" Then
                           ShowMsg MsgText(1041)
                        ElseIf CheckLengthIsOK(txtAdviser(intIndex), 40) Then
                           CheckKeyIn = 1
                        End If
             Case 4 '客戶編號1
                        'Add By Sindy 2011/1/18
                        If txtAdviser(intIndex) = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        If intIndex = 4 Then
                           If txtAdviser(intIndex) = txtAdviser(14) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(15) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(16) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(17) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        '2011/1/18 End
                        
                        'Added by Lydia 2020/05/20 法律所案源收文：檢查介紹客戶是否為申請人1
                        If m_LOS05 <> "" And txtAdviser(4) <> "" And ChangeCustomerS(m_LOS05) <> ChangeCustomerS(txtAdviser(4)) Then
                            MsgBox "客戶編號1請輸入 " & ChangeCustomerS(m_LOS05), vbExclamation, "檢查介紹客戶"
                            txtAdviser(4).SetFocus
                            txtAdviser_GotFocus 4
                            CheckKeyIn = -1
                            Exit Function
                        End If
                        'end 2020/05/20
                        
                        strCusTemp = txtAdviser(intIndex)
                        'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
                        'If objPublicData.GetCustomer(strCusTemp, strTemp, strTemp1) Then
                        'Modify By Sindy 2015/8/27 +txtSystem
                        'Modified by Lydia 2023/03/06 傳入本所案號 , , , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
                        If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, , , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) Then
                           txtAdviser(intIndex) = strCusTemp
                           lblPetition(0).Caption = strTemp
                           If strLastCus <> strCusTemp Or txtAdviser(8) = "" Then
                              txtAdviser(8).Text = strTemp1
                              strLastCus = strCusTemp
                           End If
                           CheckKeyIn = 1
                           'Add by Morgan 2008/8/5
                           If ChangeCustomerL(strCusTemp) <> strAppNo1 Then
                              strAppNo1 = ChangeCustomerL(strCusTemp)
                              'Modified by Lydia 2021/04/28 改成Form 2.0
                              'PUB_AddContact strAppNo1, cboContact, , True
                              strExc(10) = cboContact.Tag
                              'Added by Lydia 2022/11/25 區分有無輸入接洽人; ex.P-130652接洽人不是客戶預設接洽人
                              If cboContact.Text <> "" Then
                                 strExc(9) = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
                                 PUB_AddContact strAppNo1, cboContact, strExc(9), True, True, strExc(10)
                              Else
                              'end 2022/22/25
                                  PUB_AddContact strAppNo1, cboContact, , True, True, strExc(10)
                              End If  'Added by Lydia 2022/11/25
                              cboContact.Tag = strExc(10)
                              'end 2021/04/28
                           End If
                        End If
                        If CheckKeyIn <> -1 Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetCustomerNation(strCusTemp, strNation) Then
                           If ClsPDGetCustomerNation(strCusTemp, strNation) Then
                              'If strNation >= "010" Then
                              '   txtAdviser(10) = "N"
                              'Else
                              '   txtAdviser(10) = ""
                              'End If
                           End If
                        End If
                        'Add By Cheng 2003/09/08
                        If CheckKeyIn = 1 Then
                            '2010/9/30 modify by sonia 新增時才要檢查
                            'If frm010001.m_blnNewCase = True Then
                            If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 Then
                                '若輸入9碼且最後一碼不為"0"
                                If Len(Me.txtAdviser(intIndex).Text) = 9 And Right(Me.txtAdviser(intIndex).Text, 1) <> "0" Then
                                    MsgBox "此客戶已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                                    CheckKeyIn = -1
                                End If
                            End If
                        End If
             Case 6
                        If CheckIsTaiwanDate(txtAdviser(intIndex).Text) Then
                           If Val(txtAdviser(intIndex - 1)) < Val(txtAdviser(intIndex)) Then
                              CheckKeyIn = 1
                           Else
                              ShowMsg MsgText(1042)
                           End If
                        End If
             Case 7
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetStaff(txtAdviser(intIndex).Text, strTemp, strTemp1) Then
                        If ClsPDGetStaff(txtAdviser(intIndex).Text, strTemp, strTemp1) Then
                           CheckKeyIn = 1
                        End If
                        lblSales.Caption = strTemp
                        
                        'Modified by Lydia 2019/02/14
                        'strTemp = GetST15(txtAdviser(intIndex).Text, strTemp1)
                        m_SalesST15 = GetST15(txtAdviser(intIndex).Text, strTemp1)
                        lblDepartment = strTemp1
'             Case 9
'                    'edit by nickc 2008/05/30 郭 請作單 X14843050 不管
'                    If Mid(txtAdviser(4), 1, 8) <> "X1484305" Then
'                        'If txtAdviser(intIndex) <> "" Then
'                           'edit by nickc 2007/02/02 不用 dll 了
'                           'If objPublicData.GetCaseFee(txtSystem, txtAdviser(4), txtAdviser(1), Val(txtAdviser(intIndex))) = 1 Then
'                           If ClsPDGetCaseFee(txtSystem, txtAdviser(4), txtAdviser(1), Val(txtAdviser(intIndex))) = 1 Then
'                              CheckKeyIn = 1
'                           End If
'                        'Else
'                        '   CheckKeyIn = 1
'                        'End If
'                    Else
'                        CheckKeyIn = 1
'                    End If
             Case 10
                        'If strNation >= "010" Then
                        '   If txtAdviser(10) <> "N" Then
                        '      ShowMsg "申請人國籍非台灣時, 是否開電腦收據必須為 N"
                        '      CheckKeyIn = -1
                        '      Exit Function
                        '   End If
                        'End If
                        If txtAdviser(intIndex) = "" Or txtAdviser(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 11
                        If txtAdviser(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetStaff(txtAdviser(intIndex), strTemp) Then
                        If ClsPDGetStaff(txtAdviser(intIndex), strTemp) Then
                           lblPromoter = strTemp
                           CheckKeyIn = 1
                        End If
                        End If
             'add by nickc 2005/10/06 加長分所號
             Case 12
                        If CheckLengthIsOK(txtAdviser(intIndex), 50) Then
                            CheckKeyIn = 1
                        End If
            'add by nickc 2008/05/02 加預定收款日
             Case 13
                        If txtAdviser(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtAdviser(intIndex).Text) Then
                                'edit by nickc 2008/05/21 原先判斷系統日，改成判斷收文日，因為分所會去補分所號
                                'If DBDATE(txtAdviser(intIndex).Text) >= strSrvDate(1) Then
                                If DBDATE(txtAdviser(intIndex).Text) >= DBDATE(txtAdviser(0).Text) Then
                                   CheckKeyIn = 1
                                Else
                                    'edit by nickc 2008/05/21 原先判斷系統日，改成判斷收文日，因為分所會去補分所號
                                    'MsgBox "預定收款日必須>= 系統日", vbOKOnly + vbCritical, "輸入錯誤！"
                                    MsgBox "預定收款日必須>= 收文日", vbOKOnly + vbCritical, "輸入錯誤！"
                                End If
                           End If
                        End If
             'Add By Sindy 2011/1/18
             Case 14, 15, 16, 17 '客戶編號2,3,4,5
                        If txtAdviser(intIndex) = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        If intIndex = 14 Then
                           If txtAdviser(intIndex) = txtAdviser(4) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(15) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(16) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(17) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        If intIndex = 15 Then
                           If txtAdviser(intIndex) = txtAdviser(4) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(14) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(16) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(17) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        If intIndex = 16 Then
                           If txtAdviser(intIndex) = txtAdviser(4) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(14) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(15) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(17) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        If intIndex = 17 Then
                           If txtAdviser(intIndex) = txtAdviser(4) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(14) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(15) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(16) Then
                              MsgBox "客戶編號不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        strCusTemp = txtAdviser(intIndex)
                        '檢查該申請人或代理人狀態，若為不再使用則停在原地
                        'Modify By Sindy 2015/8/27 +txtSystem
                        'Modified by Lydia 2023/03/06 傳入本所案號 , , , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
                        If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, , , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) Then
                           txtAdviser(intIndex) = strCusTemp
                           lblPetition(intIndex - 13).Caption = strTemp
                           CheckKeyIn = 1
                        End If
                        If CheckKeyIn = 1 Then
                            '新增時才要檢查
                            If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 Then
                                '若輸入9碼且最後一碼不為"0"
                                If Len(Me.txtAdviser(intIndex).Text) = 9 And Right(Me.txtAdviser(intIndex).Text, 1) <> "0" Then
                                    MsgBox "此客戶已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                                    CheckKeyIn = -1
                                End If
                            End If
                        End If
             Case Else
                        CheckKeyIn = 1
End Select
End Function

'Modified by Lydia 2021/04/28
'Private Sub txtAdviser_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtAdviser_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
             'Modify By Sindy 2011/1/18 +14,15,16,17
             Case 4, 7, 10, 11, 14, 15, 16, 17
                       KeyAscii = UpperCase(KeyAscii)
             Case 8
                       'Modified by Lydia 2021/12/14 +物件名稱
                       KeyAscii = ChangeZIP(KeyAscii, txtAdviser(Index))
End Select
End Sub

Private Sub txtAdviser_GotFocus(Index As Integer)

txtAdviser(Index).SelStart = 0
txtAdviser(Index).SelLength = Len(txtAdviser(Index).Text)
'Remove by Lydia 2021/06/09
''切換輸入法
'Select Case Index
'             Case 3
'                        'edit by nickc 2007/06/06 切換輸入法改用API
'                        'txtAdviser(Index).IMEMode = 1
'                        OpenIme
'             Case Else
'                        'edit by nickc 2007/06/06 切換輸入法改用API
'                        'txtAdviser(Index).IMEMode = 2
'                        CloseIme
'End Select
End Sub

Private Sub txtAdviser_LostFocus(Index As Integer)
'關閉輸入法
'edit by nickc 2007/06/06 切換輸入法改用API
'txtAdviser(Index).IMEMode = 2
'CloseIme 'Removed by Morgan 2016/10/20 會造成 Win7 的切換錯誤
End Sub

'讀取顧問聘任資料庫
'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
Private Function ReadHireDatabase(ByRef intModifyKind As Integer, ByRef hc01 As String, _
             ByRef hc02 As String, ByRef hc03 As String, ByRef hc04 As String, ByRef hc05 As String, _
             ByRef hc06 As String, ByRef cp05 As String, ByRef CP09 As String, ByRef CP10 As String, ByRef cp11 As String, ByRef cp53 As String, ByRef cp54 As String, _
             ByRef cp13 As String, ByRef cp16 As String, ByRef cp32 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef hc07 As String, _
             ByRef hc24 As String, ByRef hc25 As String, ByRef hc26 As String, ByRef hc27 As String, ByRef CP150 As String) As Boolean
   Dim strSql As String, rsRecordset As New ADODB.Recordset, strTemp As String
   'Add by Morgan 2004/4/15
   '收據號碼
   Dim stCP60 As String
   
On Error GoTo ErrHand
If intModifyKind <> 0 Then
   'Add by Morgan 2004/4/15
   '收據號碼
   'strSQL = "select cp05,cp09,cp10,cp11,cp53,cp54,cp13,cp16,cp32,cp14 from caseprogress where cp09=" + CNULL(cp09)
   strSql = "select cp05,cp09,cp10,cp11,cp53,cp54,cp13,cp16,cp32,cp14,cp60,cp150 from caseprogress where cp09=" + CNULL(CP09)
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      cp05 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
      If cp05 <> "" Then cp05 = ChangeWStringToTString(cp05)
      CP09 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
      CP10 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
      cp11 = IIf(IsNull(rsRecordset.Fields(3)), "", rsRecordset.Fields(3))
      cp53 = IIf(IsNull(rsRecordset.Fields(4)), "", rsRecordset.Fields(4))
      If cp53 <> "" Then cp53 = ChangeWStringToTString(cp53)
      cp54 = IIf(IsNull(rsRecordset.Fields(5)), "", rsRecordset.Fields(5))
      If cp54 <> "" Then cp54 = ChangeWStringToTString(cp54)
      cp13 = IIf(IsNull(rsRecordset.Fields(6)), "", rsRecordset.Fields(6))
      cp16 = IIf(IsNull(rsRecordset.Fields(7)), "", rsRecordset.Fields(7))
      cp32 = IIf(IsNull(rsRecordset.Fields(8)), "", rsRecordset.Fields(8))
      cp14 = IIf(IsNull(rsRecordset.Fields(9)), "", rsRecordset.Fields(9))
      'Add by Morgan 2004/4/15
      stCP60 = "" & rsRecordset.Fields("cp60")
      If stCP60 <> "" Then
         txtAdviser(9).Enabled = False
         'add by nickc 2006/12/25 加鎖智權人員
         txtAdviser(7).Enabled = False
      End If
      CP150 = "" & rsRecordset.Fields("cp150") 'Add By Sindy 2012/11/08
   Else
      ShowMsg MsgText(1502)
      Exit Function
   End If
   rsRecordset.Close
End If
'Modify By Cheng 2003/12/22
'strSQL = "select hc05,hc06 from hirecase where hc01=" + CNULL(hc01) + " and hc02=" + CNULL(hc02) + " and hc03=" + CNULL(hc03) + " and hc04=" + CNULL(hc04)
'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
strSql = "select hc05,hc06, hc07,hc23,hc24,hc25,hc26,hc27 from hirecase where hc01=" + CNULL(hc01) + " and hc02=" + CNULL(hc02) + " and hc03=" + CNULL(hc03) + " and hc04=" + CNULL(hc04)
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSql, cnnConnection
If rsRecordset.RecordCount > 0 Then
   hc05 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
   hc06 = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
   hc07 = IIf(IsNull(rsRecordset.Fields(2)), "", rsRecordset.Fields(2))
   'Add By Sindy 2011/1/18
   hc24 = IIf(IsNull(rsRecordset.Fields("hc24")), "", rsRecordset.Fields("hc24"))
   hc25 = IIf(IsNull(rsRecordset.Fields("hc25")), "", rsRecordset.Fields("hc25"))
   hc26 = IIf(IsNull(rsRecordset.Fields("hc26")), "", rsRecordset.Fields("hc26"))
   hc27 = IIf(IsNull(rsRecordset.Fields("hc27")), "", rsRecordset.Fields("hc27"))
   '2011/1/18 End
   'Add by Morgan 2008/8/5
   strAppNo1 = "" & rsRecordset("hc05")
   'Modified by Lydia 2021/04/28 改成Form 2.0
   'PUB_AddContact strAppNo1, cboContact, "" & rsRecordset("hc23"), True
   ''end 2008/8/5
   strExc(10) = cboContact.Tag
   PUB_AddContact strAppNo1, cboContact, "" & rsRecordset("hc23"), True, True, strExc(10)
   cboContact.Tag = strExc(10)
   'end 2021/04/28
   
   rsRecordset.Close
   'strSQL = "select cu30 from customer where cu01||cu02='" + hc05 + "'"
   strSql = "select cu30 from customer where cu01=" + CNULL(Mid(hc05, 1, 8)) + " AND cu02=" + CNULL(Mid(hc05, 9, 1))
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      cu30 = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
      hc05 = ChangeCustomerS(hc05)
      'Add By Sindy 2011/1/18
      hc24 = ChangeCustomerS(hc24)
      hc25 = ChangeCustomerS(hc25)
      hc26 = ChangeCustomerS(hc26)
      hc27 = ChangeCustomerS(hc27)
      '2011/1/18 End
      ReadHireDatabase = True
   Else
      ShowMsg MsgText(1503)
      Exit Function
   End If
Else
   If intModifyKind <> 0 Then
      ShowMsg "找不到此本所案號在基本檔之資料"
   End If
End If
rsRecordset.Close

'add by nickc 2008/05/02 抓預定收款日
''Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'strSql = "select rd05 from ReceivablesDay where (rd01,rd02,rd03) in (select rd01,rd02,max(rd03) from ReceivablesDay where (rd02) in (select max(rd02) from ReceivablesDay where rd01='" & CP09 & "' ) and rd01='" & CP09 & "' group by rd01,rd02) "
'rsRecordset.CursorLocation = adUseClient
'rsRecordset.Open strSql, cnnConnection
'If rsRecordset.RecordCount > 0 Then
'   txtAdviser(13) = IIf(IsNull(rsRecordset.Fields(0)), "", TAIWANDATE(rsRecordset.Fields(0)))
'Else
'   txtAdviser(13) = ""
'End If
'txtAdviser(13).Tag = txtAdviser(13) 'Add by Morgan 2010/12/9
'rsRecordset.Close
'end 2018/08/22

Exit Function
ErrHand:
   ShowMsg "資料讀取失敗,請洽系統管理者!"  '2010/8/18 add by sonia
End Function

'修改顧問聘任資料庫
'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
Private Function UpdateHireDatabase(ByRef hc01 As String, _
      ByRef hc02 As String, ByRef hc03 As String, ByRef hc04 As String, ByRef hc05 As String, _
      ByRef hc06 As String, ByRef hc07 As String, ByRef CP09 As String, _
      ByRef cp05 As String, ByRef CP10 As String, ByRef cp11 As String, ByRef cp53 As String, _
      ByRef cp54 As String, ByRef cp13 As String, ByRef cp16 As String, _
      ByRef cp32 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef HC23 As String, _
      ByRef hc24 As String, ByRef hc25 As String, ByRef hc26 As String, ByRef hc27 As String) As Boolean
Dim strSql As String
Dim adoquery As New ADODB.Recordset

'add by nickc 2007/12/12
If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand
cp05 = ChangeTStringToWString(cp05)
cp53 = ChangeTStringToWString(cp53)
cp54 = ChangeTStringToWString(cp54)
hc05 = ChangeCustomerL(hc05)
'Add By Sindy 2011/1/18
hc24 = ChangeCustomerL(hc24)
hc25 = ChangeCustomerL(hc25)
hc26 = ChangeCustomerL(hc26)
hc27 = ChangeCustomerL(hc27)
'2011/1/18 End
cnnConnection.BeginTrans
'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
strSql = "update hirecase set hc05=" + CNULL(hc05) + ",hc06=" + CNULL(hc06) + ",hc07=" + CNULL(hc07) + ",hc24=" + CNULL(hc24) + ",hc25=" + CNULL(hc25) + ",hc26=" + CNULL(hc26) + ",hc27=" + CNULL(hc27)
'Add by Morgan 2008/8/5 +HC23
If UCase(HC23) <> "HC23" Then
   strSql = strSql + ",HC23=" + CNULL(HC23)
End If
strSql = strSql + " where hc01=" + CNULL(hc01) + " and hc02=" + CNULL(hc02) + " and hc03=" + CNULL(hc03) + " and hc04=" + CNULL(hc04)
cnnConnection.Execute strSql

'Add By Sindy 2012/11/06 有★★的應收帳款簽核控管
m_CP150 = ""
If Check2.Value = 1 Then m_CP150 = "Y"
'2012/11/06 End

'Modify By Sindy 2012/11/06 +CP150
strSql = "update caseprogress set cp05=" + CNULL(cp05) + ",cp10=" + CNULL(CP10) + ",cp11=" + CNULL(cp11) + ",cp53=" + CNULL(cp53) + ",cp54=" + CNULL(cp54) + ",cp13=" + CNULL(cp13) + _
   ",cp14=" + CNULL(cp14) + ",cp16=" + CNULL(cp16) + ",cp32=" + CNULL(cp32) + ",cp18=" & CNULL(IIf(Val(cp16) / 1000 = 0, "", Val(cp16) / 1000)) & ",cp150=" & CNULL(m_CP150) & " where cp09=" + CNULL(CP09)
cnnConnection.Execute strSql
strSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(cp13) + ") where cp09=" + CNULL(CP09)
cnnConnection.Execute strSql
        'Add By nickc 2007/08/21
        '若為接洽記錄單(櫃台收文)
        'Modify by Morgan 2007/10/26 費用可改時才做，否則已收款資料會被還原
        'If frm010001.intChoose = 0 Then
        If frm010001.intChoose = 0 And txtAdviser(9).Enabled = True Then
        'end 2007/10/26
            '未收金額 = 費用
            strSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(CP09)
            cnnConnection.Execute strSql
        End If
        
'Added by Lydia 2022/11/29 非內部收文並且有費用，先統一設定CP20=Null ;
If frm010001.intChoose = 0 And Val(cp16) > 0 Then
    strSql = "update caseprogress set cp20=null where cp09=" + CNULL(CP09)
    cnnConnection.Execute strSql
End If
'end 2022/11/29
'Add By Cheng 2002/05/10
'若為內部收文作業時, 案件進度檔的是否向客戶收款設定為"N"
If frm010001.intChoose = 1 Then
   strSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(CP09)
   cnnConnection.Execute strSql
End If

strSql = "update customer set cu30=" + CNULL(cu30) + " where cu01=" + CNULL(Mid(hc05, 1, 8)) + " and cu02=" + CNULL(Mid(hc05, 9, 1))
cnnConnection.Execute strSql

adoquery.CursorLocation = adUseClient
'adoquery.Open "select np01 from nextprogress where np02 = '" & hc01 & "' and np03 = '" & hc02 & "' and np04 = '" & hc03 & "' and np05 = '" & hc04 & "' and np07 = '" & cp10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
adoquery.Open "select np01 from nextprogress where np02 = '" & hc01 & "' and np03 = '" & hc02 & "' and np04 = '" & hc03 & "' and np05 = '" & hc04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
'Modify By Cheng 2002/05/10
'若在下一程序檔只抓到一筆資料時, 才要抓下一程序檔的總收文號更新案件進度檔的相關總收文號
'If adoquery.RecordCount <> 0 Then
If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
   If IsNull(adoquery.Fields(0).Value) = False Then
      cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
   End If
End If
adoquery.Close

'add by nickc 2008/05/02 儲存預定收款日
'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'Dim rtCnt As Integer
''Modify by Morgan 2010/12/9
''If txtAdviser(13) <> "" Then
''    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " ", rtCnt
'If txtAdviser(13) <> "" And txtAdviser(13) <> txtAdviser(13).Tag Then
'    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0) + 1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
''end 2010/12/9
'    If rtCnt = 0 Then
'        cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from dual "
'    End If
'End If
'end 2018/08/22

If m_LOS15 <> "" Then PUB_UpdateTTFee m_LOS15 'Added by Morgan 2022/4/14

cnnConnection.CommitTrans
UpdateHireDatabase = True
Exit Function
ErrHand:
cnnConnection.RollbackTrans
ShowMsg MsgText(9004)
'add by nickc 2007/12/12
IsSaveData = False
End Function

'新增顧問聘任至資料庫
'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
Private Function InsertHireDatabase(ByRef intSaveMode As Integer, ByRef hc01 As String, _
             ByRef hc02 As String, ByRef hc03 As String, ByRef hc04 As String, ByRef hc05 As String, _
             ByRef hc06 As String, ByRef hc07 As String, _
             ByRef cp05 As String, ByRef CP10 As String, ByRef cp11 As String, ByRef cp53 As String, _
             ByRef cp54 As String, ByRef cp13 As String, ByRef cp16 As String, _
             ByRef cp32 As String, ByRef cu30 As String, ByRef cp14 As String, ByRef CP09 As String, ByRef cp02 As String, ByRef HC23 As String, _
             ByRef hc24 As String, ByRef hc25 As String, ByRef hc26 As String, ByRef hc27 As String) As Boolean
Dim strSql As String, strAutoNumber As String, bolError As Boolean, cp31 As String
Dim adoquery As New ADODB.Recordset
'add by nickc 2007/12/12
If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand
'傳入0為重複之本所案號(新增舊案)，1為正確之本所案號(新增新案)
cp05 = ChangeTStringToWString(cp05)
cp53 = ChangeTStringToWString(cp53)
cp54 = ChangeTStringToWString(cp54)
hc05 = ChangeCustomerL(hc05)
'Add By Sindy 2011/1/18
hc24 = ChangeCustomerL(hc24)
hc25 = ChangeCustomerL(hc25)
hc26 = ChangeCustomerL(hc26)
hc27 = ChangeCustomerL(hc27)
'2011/1/18 End
'edit by nickc 2007/02/06 不用 dll 了
'Dim objPublicData As Object
'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
cnnConnection.BeginTrans
If intSaveMode = 1 Then
   If hc02 = "" Then
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetAutoNumber(hc01, strAutoNumber, True, False) Then
      If ClsPDGetAutoNumber(hc01, strAutoNumber, True, False) Then
         hc02 = strAutoNumber
      Else
         bolError = True
      End If
   End If
   If bolError = False Then
      cp02 = hc02
      'Modify by Morgan 2008/8/5 +HC23
      'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
      strSql = "insert into hirecase (hc01,hc02,hc03,hc04,hc05,hc06, hc07,hc23,hc24,hc25,hc26,hc27) values (" + _
          CNULL(hc01) + "," + CNULL(hc02) + "," + CNULL(hc03) + "," + CNULL(hc04) + "," + _
          CNULL(hc05) + "," + CNULL(hc06) + "," + CNULL(hc07) + "," + CNULL(HC23) + "," + _
          CNULL(hc24) + "," + CNULL(hc25) + "," + CNULL(hc26) + "," + CNULL(hc27) + ")"
      cnnConnection.Execute strSql
      cp31 = "Y"
   Else
      bolError = True
   End If
End If
If bolError = False Then
   'edit by nickc 2007/02/06 不用 dll 了
   'If objPublicData.GetAutoNumber(Left(CP09, 1), strAutoNumber, True, True) Then
   If ClsPDGetAutoNumber(Left(CP09, 1), strAutoNumber, True, True) Then
      'add by nick 2005/01/07
      If Me.txtAdviser(12).Text <> "" Then
           strSql = "Update hirecase Set hc07='" & ChgSQL(Me.txtAdviser(12).Text) & "' Where hc01=" + CNULL(hc01) + " and hc02=" + CNULL(hc02) + " and hc03=" + CNULL(hc03) + " and hc04=" + CNULL(hc04)
           cnnConnection.Execute strSql
      End If
       'add by nick 2004/10/29
       If CP10 = 顧問聘任 Then
            strSql = "update caseprogress set cp27=" & Trim(ServerDate) & " where cp01=" & CNULL(hc01) & " and cp02=" & CNULL(hc02) & " and cp03=" & CNULL(hc03) & " and cp04=" & CNULL(hc04) & " and cp27 is null and cp57 is null "
            cnnConnection.Execute strSql
        End If
      CP09 = CP09 + strAutoNumber
      
      'Add By Sindy 2012/11/06 有★★的應收帳款簽核控管
      m_CP150 = ""
      If Check2.Value = 1 Then m_CP150 = "Y"
      '2012/11/06 End
      
      'Modify By Sindy 2012/11/06 +CP150
      strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp14,cp53,cp54,cp13,cp16," + _
           "cp31,cp32,cp18,cp150) values (" + CNULL(hc01) + "," + CNULL(hc02) + "," + CNULL(hc03) + "," + CNULL(hc04) + "," + CNULL(cp05) + "," + _
           CNULL(CP09) + "," + CNULL(CP10) + "," + CNULL(cp11) + "," + CNULL(cp14) + "," + CNULL(cp53) + "," + CNULL(cp54) + "," + CNULL(cp13) + "," + CNULL(cp16) + "," + _
           CNULL(cp31) + "," + CNULL(cp32) + "," + CNULL(IIf(Val(cp16) / 1000 = 0, "", Val(cp16) / 1000)) & "," + CNULL(m_CP150) + ")"
      cnnConnection.Execute strSql
      strSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(cp13) + ") where cp09=" + CNULL(CP09)
      cnnConnection.Execute strSql
        
        '若為接洽記錄單(櫃台收文)
        'Modify by Morgan 2007/10/26 費用可改時才做，否則已收款資料會被還原
        'If frm010001.intChoose = 0 Then
        If frm010001.intChoose = 0 And txtAdviser(9).Enabled = True Then
        'end 2007/10/26
            '未收金額 = 費用
            strSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(CP09)
            cnnConnection.Execute strSql
        End If
      'Added by Lydia 2022/11/29 非內部收文並且有費用，先統一設定CP20=Null ;
      If frm010001.intChoose = 0 And Val(cp16) > 0 Then
          strSql = "update caseprogress set cp20=null where cp09=" + CNULL(CP09)
          cnnConnection.Execute strSql
      End If
      'end 2022/11/29
      'Add By Cheng 2002/05/10
      '若為內部收文作業時, 案件進度檔的是否向客戶收款設定為"N"
      If frm010001.intChoose = 1 Then
         strSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(CP09)
         cnnConnection.Execute strSql
      End If
      
      strSql = "update customer set cu30=" + CNULL(cu30) + " where cu01=" + CNULL(Mid(hc05, 1, 8)) + " and cu02=" + CNULL(Mid(hc05, 9, 1))
      cnnConnection.Execute strSql
      
      'Add By Sindy 2011/3/17 為顧問聘任(cp10=0)且聘任期間>系統日時,若cu153為null時則更新為Y,為N者不可更新
      If CP10 = 顧問聘任 And Val(txtAdviser(6)) > Val(strSrvDate(2)) Then
         strSql = "update customer set cu153='Y' where cu01='" + Mid(hc05, 1, 8) + "' and cu02='" + Mid(hc05, 9, 1) + "' and cu153 is null"
         cnnConnection.Execute strSql
         strSql = "update potcustcont set pcc23='Y' where pcc01='" + Mid(hc05, 1, 8) + "' and pcc23 is null"
         cnnConnection.Execute strSql
      End If
      
      'edit by nickc 2007/02/06 不用 dll 了
      'If obj001.SetCaseProgressFee(hc01, 台灣國家代號, 顧問聘任, CP09) = False Then bolError = True
      If Cls001SetCaseProgressFee(hc01, 台灣國家代號, 顧問聘任, CP09) = False Then bolError = True
   Else
      bolError = True
   End If
End If
adoquery.CursorLocation = adUseClient
'adoquery.Open "select np01 from nextprogress where np02 = '" & hc01 & "' and np03 = '" & hc02 & "' and np04 = '" & hc03 & "' and np05 = '" & hc04 & "' and np07 = '" & cp10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
adoquery.Open "select np01 from nextprogress where np02 = '" & hc01 & "' and np03 = '" & hc02 & "' and np04 = '" & hc03 & "' and np05 = '" & hc04 & "' and np06 is null and np07 = '" & CP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
'Modify By Cheng 2002/05/10
'若在下一程序檔只抓到一筆資料時, 才要抓下一程序檔的總收文號更新案件進度檔的相關總收文號
'If adoquery.RecordCount <> 0 Then
If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
   If IsNull(adoquery.Fields(0).Value) = False Then
      cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & CP09 & "'"
   End If
End If
adoquery.Close
'add by nickc 2008/05/02 儲存預定收款日
'Remove by Lydia 2018/08/22 (應收帳款管控)取消預定收款日,改成付款週期
'If bolError = False Then
'   Dim rtCnt As Integer
'   'Modify by Morgan 2010/12/9
'   'If txtAdviser(13) <> "" Then
'   '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " ", rtCnt
'   If txtAdviser(13) <> "" And txtAdviser(13) <> txtAdviser(13).Tag Then
'         cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0) + 1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from receivablesday where rd01='" & CP09 & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
'   'end 2010/12/9
'         If rtCnt = 0 Then
'             cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & CP09 & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from dual "
'         End If
'   End If
'End If
'end 2018/08/22

    'Added by Lydia 2020/05/20 法律所案源收文：存檔時案源單號存CP162、案源總收文號(LOS01)存CP64欄"案源：本所案號(總收文號)
    If strSrvDate(1) >= 法律所案源收文啟用日 And txtSystem = "LA" And m_LOS15 <> "" Then
        strSql = ""
        If m_LOS01 <> "" Then strSql = ",cp64=" & CNULL("案源：" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 <> "0", "-" & m_LOS01cp03, "") & IIf(m_LOS01cp04 <> "00", "-" & m_LOS01cp04, "") & "(" & m_LOS01 & ");")
        strSql = "update caseprogress set CP162=" & CNULL(m_LOS15) & strSql & " where cp09=" & CNULL(CP09)
        cnnConnection.Execute strSql, intI
       
        '並回寫收文號至案源檔的法律所總收文號欄。
        '5/26 若輸入之案源單號已有法律所總收文號且為同案號同日收文者，則為同一接洽單之其他性質。
        strSql = "update LawOfficeSource set los06='" & CP09 & "' where los06 is null and los15=" & CNULL(frm010001.txtLOS15)
        cnnConnection.Execute strSql, intI
        
        '若案源資料的介紹客戶LOS05為空時表示新客戶要回寫並更新(收文時輸入的)客戶智權人員(CU12CU13)為介紹人(LOS04第一人)
        If m_LOS05 = "" And Trim(txtAdviser(4) & txtAdviser(14) & txtAdviser(15) & txtAdviser(16) & txtAdviser(17)) <> "" Then
            '並且回寫案源介紹客戶編號LOS05
            strSql = "update LawOfficeSource set los05='" & ChangeCustomerL(txtAdviser(4)) & "' where los05 is null and los15=" & CNULL(m_LOS15)
            cnnConnection.Execute strSql, intI
            If intI > 0 Then
               strExc(1) = "4": strExc(2) = "14": strExc(3) = "15": strExc(4) = "16": strExc(5) = "17"
               For intI = 1 To 5
                   If Trim(txtAdviser(Val(strExc(intI)))) <> "" Then
                        strSql = "update customer set cu12='" & m_LOS04_1st15 & "',cu13='" & m_LOS04_1 & "' where cu01='" & Left(ChangeCustomerL(txtAdviser(Val(strExc(intI)))), 8) & "' and cu02='" & Right(ChangeCustomerL(txtAdviser(Val(strExc(intI)))), 1) & "'"
                        Pub_SeekTbLog strSql
                        cnnConnection.Execute strSql
                   End If
               Next intI
               m_Los05_N = ChangeCustomerL(txtAdviser(4))   'Added by Lydia 2022/11/10 客戶編號後建=m_LOS05=空白
            End If
        End If
        '最後才做-->客戶編號回寫後，案源案件類型A，若無點數則保留類型A，若有點數則判斷同一客戶編號介紹日前若有A1則此筆設為A2，若無則設為A1。
                            '計算案源之費用及點數，更新回案源總收文號LOS01之費用及點數，以利智慧所開立收據。
                            '案源為TT-999999時同時上發文日CP27為系統日(為無發文日者才更新)。
                            '5/6跟楊監察人確認國外部介紹案源以相同分潤方式計算，不管國外代理人仍以客戶為介紹基準。
        If m_LOS02 = "A" And Val(txtAdviser(9)) > 0 Then  '費用改為點數
           strSql = "select los02 from LawOfficeSource where los12<'" & m_LOS12 & "' and los02='A1' and los05='" & IIf(m_LOS05 <> "", m_LOS05, ChangeCustomerL(txtAdviser(4))) & "' "
           intI = 1
           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
           If intI = 1 Then
               strSql = "update LawOfficeSource set los02='A2' where los15='" & m_LOS15 & "' "
               cnnConnection.Execute strSql
           Else
               strSql = "update LawOfficeSource set los02='A1' where los15='" & m_LOS15 & "' "
               cnnConnection.Execute strSql
           End If
           '案源為TT-999999時同時上發文日CP27為系統日(為無發文日者才更新)。
           If m_LOS01cp01 & m_LOS01cp02 = "TT999999" Then
               strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_LOS01 & "' and nvl(cp27,0)=0 "
               cnnConnection.Execute strSql
           End If
        End If
        
       '計算案源之費用及點數，更新回案源總收文號LOS01之費用及點數，以利智慧所開立收據。
       PUB_UpdateTTFee m_LOS15 'Added by Morgan 2020/9/29 同案源單號的每個收文性質都要(費用加總)
    End If
    'end 2020/05/20
    
    'Added by Lydia 2020/05/20 法律所案源收文：若該收文號點數>0但無案源(自行收文者)時，若案件的客戶為非法律所的客戶時則仍算A類案源(另寫函數參照作帳規則設定為A1~A4)。
                                       '系統自動新增TT-999999案進度(B類收文)及法律所案源資料(同最後一筆案源的資料)。
    'Memo by Lydia 2020/10/05 (9/30) 若該收文號點數>0但無案源(自行收文者)時，若案件的客戶為非法律所的客戶時則為A3類案源，不論新舊案，系統自動新增TT-999999案進度(B類收文)及法律所案源資料。
    'Modified by Morgan 2021/1/8 台一關係企業 X03072 除外
    If strSrvDate(1) >= 法律所案源收文啟用日 And txtSystem = "LA" And m_LOS15 = "" And Val(txtAdviser(9)) > 0 And Left(txtAdviser(4), 6) <> "X03072" Then
        'Modified by Lydia 2020/10/05 + st01
        strSql = "select cu01,cu02,st15,st01 from customer,staff where cu01='" & Mid(ChangeCustomerL(txtAdviser(4)), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(txtAdviser(4)), 9, 1) & "'  and cu13=st01(+) "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
            strExc(1) = Left("" & RsTemp.Fields("st15"), 1)
            If strExc(1) <> "L" Then
               '非法律所的舊客戶時則仍算A類案源(另寫函數參照作帳規則設定為A1~A4)
               'Modified by Lydia 2020/10/05 (9/30) 若該收文號點數>0但無案源(自行收文者)時，若案件的客戶為非法律所的客戶時則為A3類案源，不論新舊案，系統自動新增TT-999999案進度(B類收文)及法律所案源資料。
               'strSql = "select * from LawOfficeSource where los12<'" & strSrvDate(1) & "' and los02 like 'A%' and los05='" & ChangeCustomerL(txtAdviser(4)) & "' " & _
                           "order by los12 desc, los13 desc "
               'intI = 1
               'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               'If intI = 1 Then
               '   RsTemp.MoveFirst
               '   strExc(1) = AutoNo("B", 6) 'TT收文號
               '   '案源類別
               '    If "" & RsTemp.Fields("los02") = "A" Then
               '        strExc(2) = "A1"
               '    Else
               '        strExc(2) = "A2"
               '    End If
              '
              '     'TT新增B類收文
              '     strExc(3) = txtAdviser(7)
              '     If "" & RsTemp.Fields("LOS04") <> "" Then '抓介紹人1
              '        If InStr("" & RsTemp.Fields("LOS04"), ",") = 0 Then
              '            strExc(3) = "" & RsTemp.Fields("LOS04")
              '        Else
              '            strExc(3) = Mid(RsTemp.Fields("LOS04"), 1, InStr("" & RsTemp.Fields("LOS04"), ",") - 1)
              '        End If
              '     End If
                   strExc(2) = "A3"   '案源類型
                   strExc(3) = "" & RsTemp.Fields("st01")  'B類收文之智權人員: 介紹人第一人
                   strExc(5) = strExc(3)  '介紹人
                   If txtCode(0) <> "" Then
                       '如為舊案且曾有A3類案源時，介紹人員同最後一筆案源，否則以客戶目前的智權人員為介紹人。
                       strSql = "Select * From Lawofficesource Where los02='A3' and los07||los08 is null and Los15 In " & _
                                   "(select max(cp162) from caseprogress where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and cp04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' and cp162 is not null) "
                       intI = 1
                       Set adoquery = ClsLawReadRstMsg(intI, strSql)
                       If intI = 1 Then
                            '用於案源之接洽人取得在職員工編號和介紹人第一人
                            strExc(5) = PUB_GetNowStaff("" & adoquery.Fields("los04"), strExc(3))
                       End If
                   End If
                   m_Los04_N1 = strExc(3)
                   If strExc(3) <> "" Then
                        strExc(1) = AutoNo("B", 6) 'TT收文號
              ' end 2020/10/05
                        'Modified by Morgan 2021/1/8 +cp20,cp27,cp32
                        'Modified by Lydia 2022/11/09 +CP27=系統日
                        strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp11,cp12,cp13,cp20,cp32,CP162,CP27)" & _
                           " values('TT','999999','0','00'," & strSrvDate(1) & ",null ,'" & strExc(1) & "'" & _
                           ",'735','07','" & GetST15(strExc(3)) & "','" & strExc(3) & "','N','N',null, " & strSrvDate(1) & " )"
                        cnnConnection.Execute strSql
                        '法律所案源資料(同最後一筆案源的資料), 案源單號=TT總收文號
                        strExc(4) = AutoNo("LOS", 5, , True)
                        'Modified by Lydia 2020/10/05
                        'strSql = "insert into LawOfficeSource(LOS01,LOS02,LOS03,LOS04,LOS05,LOS06,LOS10,LOS11,LOS12,LOS13,LOS15)" & _
                           " values ('" & strExc(1) & "','" & strExc(2) & "' ,'" & RsTemp.Fields("los03") & "'" & _
                           ",'" & RsTemp.Fields("los04") & "','" & RsTemp.Fields("los05") & "','" & CP09 & "','" & strExc(1) & "'" & _
                           ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss'),'" & strExc(4) & "')"
                        strSql = "insert into LawOfficeSource(LOS01,LOS02,LOS03,LOS04,LOS05,LOS06,LOS10,LOS11,LOS12,LOS13,LOS15)" & _
                           " values ('" & strExc(1) & "','" & strExc(2) & "' ,'" & txtAdviser(7) & "'" & _
                           ",'" & strExc(5) & "','" & ChangeCustomerL(txtAdviser(4)) & "','" & CP09 & "','" & strExc(1) & "'" & _
                           ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss'),'" & strExc(4) & "')"
                        'Modified by Lydia 2022/09/07 debug
                        'm_Los05_N = strExc(4)
                        m_Los05_N = ChangeCustomerL(txtAdviser(4))
                        'end 2020/10/05
                        cnnConnection.Execute strSql
                        'Added by Lydia 2020/10/05 收文之進度加註案源
                        'Modified by Lydia 2022/09/07 +cp162
                        strSql = "Update CaseProgress Set cp64=" & CNULL("案源：TT-999999(" & strExc(1) & ");") & "||cp64, cp162='" & strExc(4) & "' where cp09=" & CNULL(CP09)
                        cnnConnection.Execute strSql
                        
                        '計算案源之費用及點數，更新回案源總收文號LOS01之費用及點數，以利智慧所開立收據。
                        PUB_UpdateTTFee strExc(4) 'Added by Morgan 2020/9/29
                   End If 'Added by Lydia 2020/10/05
               'End If 'Remove by Lydia 2020/10/05
            End If
        End If
    End If
    'end 2020/05/20
         
If bolError Then
   cnnConnection.RollbackTrans
   ShowMsg MsgText(9004)
'add by nickc 2007/12/12
IsSaveData = False
Else
   cnnConnection.CommitTrans
   InsertHireDatabase = True
   'add by nickc 2006/03/27
   txtCode(0) = hc02
End If
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = Nothing
'add by nickc 2005/08/12
txtCode(0) = hc02
Exit Function
ErrHand:
'edit by nickc 2007/02/06 不用 dll 了
'Set objPublicData = Nothing
cnnConnection.RollbackTrans
'edit by nickc 2006/03/07 解決 cp02=null 的問題
'add by nickc 2005/08/25
'txtCode(0) = ""
ShowMsg MsgText(9004)
'add by nickc 2007/12/12
IsSaveData = False
Resume
End Function

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Me.txtAdviser
   If objTxt.Enabled = True Then
      Cancel = False
      txtAdviser_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Added Lydia 2020/07/06 顧問聘任需輸入聘任期間
If txtSystem = "LA" And txtAdviser(1) = 顧問聘任 Then
    If (Trim(txtAdviser(4)) = "" Or Trim(lblPetition(0).Caption) = "") And txtAdviser(5).Enabled = True Then
        MsgBox "客戶編號1不可空白！", vbCritical, "檢核資料"
        txtAdviser(4).SetFocus
        Call txtAdviser_GotFocus(4)
        Exit Function
    End If
End If
'end 2020/07/06

TxtValidate = True
End Function

'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔
Private Sub ReadLOS()
Dim strR As String, intR As Integer
Dim rsRd As New ADODB.Recordset
     
    m_LOS15 = ""
    m_Los05_N = "": m_Los04_N1 = ""   'Added by Lydia 2020/10/05
    If frm010001.intModifyKind = 0 And strSrvDate(1) >= 法律所案源收文啟用日 Then
        If frm010001.txtLOS15 <> "" Then
            strR = "select X.*,cp01,cp02,cp03,cp04 from LawOfficeSource X,caseprogress where los15=" & CNULL(frm010001.txtLOS15) & " and los01=cp09(+) "
            intR = 1
            Set rsRd = ClsLawReadRstMsg(intR, strR)
            If intR = 1 Then
                '案源總收文號
                m_LOS01 = "" & rsRd.Fields("LOS01")
                '案源總收文號之本所案號
                m_LOS01cp01 = "" & rsRd.Fields("cp01")
                m_LOS01cp02 = "" & rsRd.Fields("cp02")
                m_LOS01cp03 = "" & rsRd.Fields("cp03")
                m_LOS01cp04 = "" & rsRd.Fields("cp04")
                '(原)案源案件類型
                m_LOS02 = "" & rsRd.Fields("LOS02")
                '案源單號
                m_LOS15 = "" & rsRd.Fields("LOS15")
                '介紹人, 介紹人(第一位)
                m_LOS04 = "" & rsRd.Fields("LOS04")
                If m_LOS04 <> "" And InStr(m_LOS04, ",") > 0 Then
                    m_LOS04_1 = Mid(m_LOS04, 1, InStr(m_LOS04, ",") - 1)
                Else
                    m_LOS04_1 = m_LOS04
                End If
                If m_LOS04_1 <> "" Then
                    m_LOS04_1st15 = GetST15(m_LOS04_1, , , m_LOS04_1st06)
                End If
                
                '(原)介紹客戶:
                m_LOS05 = "" & rsRd.Fields("LOS05")
                '介紹日
                m_LOS12 = "" & rsRd.Fields("LOS12")
            End If
        Else
        End If
    'Added by Morgan 2022/4/14
    ElseIf frm010001.intModifyKind = 1 Then
      If txtSystem = "LA" Then
         'Modified by Lydia 2022/09/14
         'strR = "select cp162 from caseprogress where cp09='" & txtRecieveCode & "' and cp162 is not null"
         'Modified by Lydia 2022/09/21 +los04
         strR = "select los02,cp162,los04 from caseprogress, LawOfficeSource where cp09='" & txtRecieveCode & "' and cp162 is not null and cp162=los15(+) "
         intR = 1
         Set rsRd = ClsLawReadRstMsg(intR, strR)
         If intR = 1 Then
            m_LOS02 = "" & rsRd.Fields("los02")
            m_LOS15 = "" & rsRd.Fields("cp162")
            'Added by Lydia 2022/09/21 介紹人, 介紹人(第一位)
            m_LOS04 = "" & rsRd.Fields("LOS04")
            If m_LOS04 <> "" And InStr(m_LOS04, ",") > 0 Then
                m_LOS04_1 = Mid(m_LOS04, 1, InStr(m_LOS04, ",") - 1)
            Else
                m_LOS04_1 = m_LOS04
            End If
            If m_LOS04_1 <> "" Then
                m_LOS04_1st15 = GetST15(m_LOS04_1, , , m_LOS04_1st06)
            End If
            'end 2022/09/21
         End If
      End If
    'end 2022/4/14
    End If
    Set rsRd = Nothing
End Sub
