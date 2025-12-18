VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010006_1 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   6570
   ClientLeft      =   5550
   ClientTop       =   1545
   ClientWidth     =   9015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9015
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7620
      TabIndex        =   25
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5664
      TabIndex        =   23
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6492
      TabIndex        =   24
      Top             =   70
      Width           =   1100
   End
   Begin VB.Frame fraWindow1 
      BorderStyle     =   0  '沒有框線
      Height          =   5805
      Left            =   60
      TabIndex        =   26
      Top             =   600
      Width           =   8895
      Begin VB.CheckBox Check2 
         Caption         =   "有★★的應收帳款簽核控管"
         Height          =   285
         Left            =   4440
         TabIndex        =   22
         Top             =   5100
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "現金或支票"
         Height          =   285
         Left            =   6810
         TabIndex        =   20
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Frame fraPromoter 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Height          =   372
         Left            =   90
         TabIndex        =   51
         Top             =   5100
         Width           =   4185
         Begin MSForms.TextBox txtAdviser 
            Height          =   300
            Index           =   11
            Left            =   990
            TabIndex        =   21
            Top             =   30
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
            Left            =   2100
            TabIndex        =   53
            Top             =   30
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
            Left            =   30
            TabIndex        =   52
            Top             =   53
            Width           =   975
         End
      End
      Begin VB.TextBox txtRecieveCode 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1092
         TabIndex        =   27
         Top             =   120
         Width           =   1452
      End
      Begin VB.Frame fraWindow2 
         Height          =   2535
         Left            =   30
         TabIndex        =   38
         Top             =   780
         Width           =   8805
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   3720
            MaxLength       =   2
            TabIndex        =   45
            Top             =   240
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   3240
            MaxLength       =   1
            TabIndex        =   44
            Top             =   240
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Enabled         =   0   'False
            Height          =   300
            Index           =   0
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   43
            Top             =   240
            Width           =   1212
         End
         Begin VB.TextBox txtSystem 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   42
            Top             =   240
            Width           =   732
         End
         Begin MSForms.TextBox txtAdviser 
            Height          =   300
            Index           =   17
            Left            =   1080
            TabIndex        =   8
            Top             =   1800
            Width           =   1092
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
            Left            =   1080
            TabIndex        =   7
            Top             =   1500
            Width           =   1092
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
            Left            =   1080
            TabIndex        =   6
            Top             =   1200
            Width           =   1092
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
            Left            =   1080
            TabIndex        =   5
            Top             =   900
            Width           =   1092
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
            Left            =   1560
            TabIndex        =   9
            Top             =   2145
            Width           =   6612
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
            Left            =   1080
            TabIndex        =   3
            Top             =   600
            Width           =   1092
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
            Left            =   6750
            TabIndex        =   4
            Top             =   600
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
            Left            =   2220
            TabIndex        =   64
            Top             =   1815
            Width           =   3555
            VariousPropertyBits=   27
            Size            =   "6271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label15 
            Caption         =   "當事人5："
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1800
            Width           =   1005
         End
         Begin MSForms.Label lblPetition 
            Height          =   300
            Index           =   3
            Left            =   2220
            TabIndex        =   62
            Top             =   1515
            Width           =   3555
            VariousPropertyBits=   27
            Size            =   "6271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label12 
            Caption         =   "當事人4："
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1515
            Width           =   1005
         End
         Begin MSForms.Label lblPetition 
            Height          =   300
            Index           =   2
            Left            =   2220
            TabIndex        =   60
            Top             =   1215
            Width           =   3555
            VariousPropertyBits=   27
            Size            =   "6271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label10 
            Caption         =   "當事人3："
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   1215
            Width           =   1005
         End
         Begin MSForms.Label lblPetition 
            Height          =   300
            Index           =   1
            Left            =   2220
            TabIndex        =   58
            Top             =   915
            Width           =   3555
            VariousPropertyBits=   27
            Size            =   "6271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label8 
            Caption         =   "當事人2："
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   915
            Width           =   1005
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "接洽人："
            Height          =   180
            Left            =   6000
            TabIndex        =   56
            Top             =   630
            Width           =   720
         End
         Begin VB.Label Label4 
            Caption         =   "案件名稱 (160)："
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "本所案號："
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "當事人1："
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   615
            Width           =   1005
         End
         Begin MSForms.Label lblPetition 
            Height          =   300
            Index           =   0
            Left            =   2220
            TabIndex        =   39
            Top             =   615
            Width           =   3555
            VariousPropertyBits=   27
            Size            =   "6271;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   19
         Left            =   1080
         TabIndex        =   18
         Top             =   4770
         Width           =   1095
         VariousPropertyBits=   671105051
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         Caption         =   "點數："
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   4770
         Width           =   975
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   18
         Left            =   3420
         TabIndex        =   17
         Top             =   4440
         Width           =   1095
         VariousPropertyBits=   671105051
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         Caption         =   "規費："
         Height          =   255
         Left            =   2790
         TabIndex        =   65
         Top             =   4440
         Width           =   975
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   13
         Left            =   5550
         TabIndex        =   19
         Top             =   4800
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
         Left            =   5940
         TabIndex        =   13
         Top             =   3720
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
         Index           =   6
         Left            =   2520
         TabIndex        =   11
         Top             =   3360
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
         Left            =   1080
         TabIndex        =   10
         Top             =   3360
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
         Index           =   10
         Left            =   4230
         TabIndex        =   15
         Top             =   4110
         Width           =   495
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   2
         Left            =   5160
         TabIndex        =   2
         Top             =   480
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "656;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtAdviser 
         Height          =   300
         Index           =   1
         Left            =   1092
         TabIndex        =   1
         Top             =   480
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
         Index           =   8
         Left            =   1080
         TabIndex        =   14
         Top             =   4080
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
         Left            =   1080
         TabIndex        =   16
         Top             =   4440
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
         Left            =   1080
         TabIndex        =   12
         Top             =   3720
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
         Left            =   5160
         TabIndex        =   0
         Top             =   120
         Width           =   1092
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1926;529"
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
         TabIndex        =   55
         Top             =   4860
         Width           =   1080
      End
      Begin VB.Label Label30 
         Caption         =   "分所案號："
         Height          =   255
         Left            =   5010
         TabIndex        =   54
         Top             =   3720
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   2280
         X2              =   2400
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label6 
         Caption         =   "顧問期間："
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "是否開電腦收據：           （N：不開)"
         Height          =   255
         Left            =   2790
         TabIndex        =   49
         Top             =   4110
         Width           =   3015
      End
      Begin VB.Label lblDepartment 
         Height          =   255
         Left            =   4050
         TabIndex        =   48
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "業務區："
         Height          =   255
         Left            =   3270
         TabIndex        =   47
         Top             =   3720
         Width           =   735
      End
      Begin MSForms.Label lblSales 
         Height          =   300
         Left            =   2190
         TabIndex        =   37
         Top             =   3750
         Width           =   1065
         VariousPropertyBits=   27
         Size            =   "1879;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCaseProperty 
         Height          =   252
         Left            =   1812
         TabIndex        =   36
         Top             =   552
         Width           =   2172
      End
      Begin VB.Label lblCaseSource 
         Height          =   252
         Left            =   5640
         TabIndex        =   35
         Top             =   528
         Width           =   2772
      End
      Begin VB.Label Label5 
         Caption         =   "案件來源："
         Height          =   255
         Left            =   4110
         TabIndex        =   34
         Top             =   510
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "案件性質："
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "智權人員："
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "費用："
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "郵遞區號："
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "收文日："
         Height          =   255
         Left            =   4110
         TabIndex        =   29
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "收文號："
         Height          =   255
         Left            =   150
         TabIndex        =   28
         Top             =   150
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm010006_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/04/29 Form2.0已修改(txtAdviser(index)、lblPetition(index)、lblSales、lblPromoter、cboContact
'Create by Lydia 2021/04/29 ACS智財顧問收文
Option Explicit

'bolLeave判斷離開時，是否要彈出詢問視窗
'LastData上一次存檔時，所輸入之收文日
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim bolLeave As Boolean, LastDate As String, intLeaveKind As Integer
Dim strNation As String
'是否已觸發 Form Active 事件
Dim bolActive As Boolean
Dim IsSaveData As Boolean
Dim strAppNo1 As String '申請人1編號
Dim dblAmt As Double, dblPFee As Double, dblTFee As Double, m_CP150 As String
Dim dblChkAmt As Double
Dim dblCu183 As Double '個人之應收帳款上限
Dim dblAmtR As Double, dblPFeeR As Double, dblTFeeR As Double '關係企業之應收帳款金額
Dim m_SalesST15 As String '畫面上智權人員的收文部門
Dim m_Tuser As String '創新業務部預設收文人員
Dim m_SalesST06 As String '智權人員的所別

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
      modBase(5) = txtAdviser(3)  '案件名稱(中)
      modBase(16) = txtAdviser(12)  '分所案號
      '當事人1~5
      modBase(11) = ChangeCustomerL(txtAdviser(4))
      modBase(43) = ChangeCustomerL(txtAdviser(14))
      modBase(44) = ChangeCustomerL(txtAdviser(15))
      modBase(45) = ChangeCustomerL(txtAdviser(16))
      modBase(46) = ChangeCustomerL(txtAdviser(17))
      
      '申請人聯絡人編號
      If cboContact.Locked = False Then
         If cboContact.ListIndex >= 0 Then
            modBase(42) = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
            If Val(modBase(42)) > 0 Then
            '若個案接洽人與客戶檔的預設接洽人相同時不必設定
               PUB_GetContact modBase(11), strTmpA, True
               If modBase(42) = strTmpA Then
                  modBase(42) = ""
               End If
            '排除空白=00
            ElseIf modBase(42) = "00" And Trim(cboContact.Text) = "" Then
               modBase(42) = ""
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
      modCP(17) = txtAdviser(18)     '規費
      modCP(18) = txtAdviser(19)     '點數
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
   End If
   
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      '分所智權人員則多一天
      If PUB_GetST06(txtAdviser(7)) <> "1" Then
         txtAdviser(13) = PUB_GetWorkDayAfterSysDate(CDbl(txtAdviser(0)) + 19110000, 6)
      Else
         txtAdviser(13) = PUB_GetWorkDayAfterSysDate(CDbl(txtAdviser(0)) + 19110000, 5)
      End If
      txtAdviser(13).Locked = True
   Else
      txtAdviser(13).Locked = False
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim varSaveCursor, strAuto1 As String, strAuto2 As String, i As Integer
Dim mBillNo As String, mMemo As String
Dim bolSaveOK As Boolean  'Added by Lydia 2022/09/14

If Index = 0 Then
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
   End If
   
   m_SalesST15 = GetST15(txtAdviser(7))
     
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
      '重新檢查欄位有效性
      If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
      
    '加入檢查特殊客戶
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

    If IsSpecCu Then
          If MsgBox("請確認此客戶接洽單主管是否核示??", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
              Screen.MousePointer = vbDefault
              Exit Sub
          End If
    End If
      
      'Add By Sindy 2010/12/31 費用檢查提到存檔前檢查
      '郭 請作單 X14843050 不管
      'Modify By Sindy 2011/1/18 增加當事人2,3,4,5檢查
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
      
      'Added by Lydia 2021/11/19 輸入規費、點數檢查
      '費用
      If Val(txtAdviser(9)) <> 0 Then
          If Val(txtAdviser(19)) = 0 Then
             ShowMsg MsgText(1034) '點數必須輸入
             txtAdviser(19).SetFocus
             Call txtAdviser_GotFocus(19)
             Screen.MousePointer = vbDefault
             Exit Sub
          ElseIf Format((Val(txtAdviser(9)) - Val(txtAdviser(18))) / 1000, "0.0") <> Format(Val(txtAdviser(19)), "0.0") Then
             ShowMsg MsgText(1036) '點數不符
             txtAdviser(19).SetFocus
             Call txtAdviser_GotFocus(19)
             Screen.MousePointer = vbDefault
             Exit Sub
          End If
      End If
      '規費
      If Val(txtAdviser(18)) <> 0 Then
          If Val(txtAdviser(9)) = 0 Then
              ShowMsg MsgText(1037)  '費用必須輸入
              txtAdviser(9).SetFocus
              Call txtAdviser_GotFocus(9)
              Screen.MousePointer = vbDefault
              Exit Sub
          End If
      End If
      '點數
      If Val(txtAdviser(19)) <> 0 Then
          If Val(txtAdviser(9)) = 0 Then
              ShowMsg MsgText(1037)  '費用必須輸入
              txtAdviser(9).SetFocus
              Call txtAdviser_GotFocus(9)
              Screen.MousePointer = vbDefault
              Exit Sub
          End If
      End If
      'end 2021/11/19

      '檢查當事人的輸入順序
      If (Trim(txtAdviser(14)) <> "" And Trim(txtAdviser(4)) = "") Or _
         (Trim(txtAdviser(15)) <> "" And Trim(txtAdviser(14)) = "") Or _
         (Trim(txtAdviser(16)) <> "" And Trim(txtAdviser(15)) = "") Or _
         (Trim(txtAdviser(17)) <> "" And Trim(txtAdviser(16)) = "") Then
         ShowMsg "請依序輸入當事人!"
         If Trim(txtAdviser(14)) <> "" And Trim(txtAdviser(4)) = "" Then txtAdviser(14).SetFocus: Call txtAdviser_GotFocus(14)
         If Trim(txtAdviser(15)) <> "" And Trim(txtAdviser(14)) = "" Then txtAdviser(15).SetFocus: Call txtAdviser_GotFocus(15)
         If Trim(txtAdviser(16)) <> "" And Trim(txtAdviser(15)) = "" Then txtAdviser(16).SetFocus: Call txtAdviser_GotFocus(16)
         If Trim(txtAdviser(17)) <> "" And Trim(txtAdviser(16)) = "" Then txtAdviser(17).SetFocus: Call txtAdviser_GotFocus(17)
         Screen.MousePointer = vbDefault
         Exit Sub
      End If

Dim strLC42 As String, strContact As String

      If cboContact.Locked = False Then
         strContact = ""
         If cboContact.ListCount > 2 Then
            strLC42 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
            PUB_GetContact strAppNo1, strContact, True
            If strLC42 = strContact Or strLC42 = "00" Then
               If MsgBox("請確定接洽人欄是否有為★, 是否要選擇其他接洽人!!", vbYesNo, "警告！") = vbYes Then
                   Screen.MousePointer = varSaveCursor
                   cboContact.SetFocus
                   Exit Sub
               End If
            End If
         End If
      End If

      '應收帳款管制
      If Left(m_SalesST15, 1) <> "F" And txtAdviser(4).Text <> "" And Val(txtAdviser(9).Text) > 0 Then
          'Modified by Lydia 2022/06/13 傳入收文之本所案號,案件性質(可用,串接)
          'Call PUB_GetBillDataAll("3", txtAdviser(4), dblAmt, dblPFee, dblTFee, , , TransDate(txtAdviser(0), 2), mBillNo, mMemo)
          Call PUB_GetBillDataAll("3", txtAdviser(4), txtSystem & IIf(txtCode(0) <> "", txtCode(0) & Left(txtCode(1) & "0", 1) & Left(txtCode(2) & "00", 2), ""), txtAdviser(1), dblAmt, dblPFee, dblTFee, , , TransDate(txtAdviser(0), 2), mBillNo, mMemo)
      End If
      
      ' 非T*案件(TF要含)若已送件之應收款超過15萬以上,智權人員非國外部且有費用者須做下列控管
      If (Left(Trim(txtSystem), 1) <> "T" Or Trim(txtSystem) = "TF") And _
         Left(m_SalesST15, 1) <> "F" And _
         Val(txtAdviser(9)) > 0 And _
         Check2.Value = 0 And Trim(txtAdviser(4)) <> "" Then
         dblCu183 = PUB_GetCustRecAmtLmt(txtAdviser(4), dblChkAmt)
         '判斷是否有集團上限
         If dblChkAmt = 0 Then
             dblAmtR = 0: dblPFeeR = 0: dblTFeeR = 0
         Else   '有集團上限才抓關係企業的應收帳款金額
             GetBillData txtAdviser(4), dblAmtR, dblPFeeR, dblTFeeR
         End If
         
         '已送件之應收款超過30萬以上(不含T*案件應收款),提醒
         ' 應收帳款上限分開管制為個人"應收帳款上限"和"集團應收帳款上限"
         If dblAmt >= dblCu183 Or (dblAmtR >= dblChkAmt And dblChkAmt > 0) Then
            If MsgBox("請注意接洽單上是否有註明應收帳款超額，需主管簽核才可收文！是否可收文？" & vbCrLf & _
                      "（接洽單上若有★★的應收帳款簽核控管，是否已勾選畫面上的註記欄位了？）", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               Screen.MousePointer = varSaveCursor
               Exit Sub
            End If
         End If
      End If

      '應收帳款管控
      If Left(m_SalesST15, 1) <> "F" And txtAdviser(4).Text <> "" And Val(txtAdviser(9).Text) > 0 Then
         If mMemo <> "" Then
             If MsgBox("請注意接洽單上是否有註明" & vbCrLf & mMemo & "，請交主管簽核。" & vbCrLf & "並且有主管簽核，是否可收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                Screen.MousePointer = varSaveCursor
                Exit Sub
             End If
         End If
      End If
      'end 2018/08/22
      
      '增加檢查重複聘任期間，彈訊息與智權人員確認後，方可收文。 ex. LA-003219於109/1/9已有顧問聘任期間(輸錯109/2/1~111/1/31)，又於110/1/20重複顧問聘任期間
      If txtSystem = "ACS" And Len(txtCode(0)) = 6 And txtAdviser(1) = "112" Then
        strSql = "select cp53,cp54 from caseprogress Where Cp09 In (" & _
                     "Select Substr(Max(Cp05||Cp09),9,9) Mno From Caseprogress Where Cp01='" & txtSystem & "' And Cp02='" & txtCode(0) & "' And Cp03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' And Cp04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' And Cp10='" & txtAdviser(1) & "' And Cp158=0 And Cp159=0) "
        CheckOC3
        AdoRecordSet3.CursorLocation = adUseClient
        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If AdoRecordSet3.RecordCount <> 0 Then
            If Val("" & AdoRecordSet3.Fields("cp54")) >= Val(DBDATE(txtAdviser(5))) Then
                strExc(1) = "已有顧問期間：" & ChangeWStringToTDateString("" & AdoRecordSet3.Fields("cp53")) & "-" & ChangeWStringToTDateString("" & AdoRecordSet3.Fields("cp54")) & vbCrLf & "請與智權人員聯繫，確認是否繼續收文？"
                If MsgBox(strExc(1), vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                   Screen.MousePointer = varSaveCursor
                   Exit Sub
                End If
            End If
        End If
      End If

      'Modified by Lydia 2022/09/14 判斷啟用日
      'If SaveDatabase(strAuto1, strAuto2) Then
      bolSaveOK = False
      If strSrvDate(1) < 收文存檔模組化啟用日 Then
         bolSaveOK = SaveDatabase(strAuto1, strAuto2)
      Else
         Call SetDBArray(False, txtRecieveCode, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
         bolSaveOK = PUB_SaveFrm010006_1(Me.Name, frm010001.intSaveMode, frm010001.intModifyKind, frm010001.intCaseKind, frm010001.intChoose, modBase, modCP, txtAdviser(8), IsSaveData, mType, mCaseNo)
         
         If frm010001.intModifyKind = 0 And bolSaveOK = True Then
             txtCode(0) = modBase(2)
             strAuto1 = modCP(9)
             strAuto2 = modBase(2)
         End If
      End If
      If bolSaveOK = True Then
      'end 2022/09/14
         PUB_SendMailCache 'Add by Sindy 2022/9/29
         frm010001.ClearForm strAuto1, strAuto2
         bolLeave = True
         intLeaveKind = 1
         If frm010001.intModifyKind = 0 Then LastDate = txtAdviser(0).Text
         Unload Me
      End If
   End If
   Screen.MousePointer = vbDefault
Else
   If Index = 2 Then
      intLeaveKind = 0
   Else
      intLeaveKind = 1
   End If
   Unload Me
End If
End Sub

Private Function SaveDatabase(ByRef strRecieveAuto As String, ByRef strCaseAuto As String) As Boolean
Dim adoquery As New ADODB.Recordset
Dim strLC42 As String, strContact As String

   If cboContact.Locked = False Then
      If cboContact.ListIndex >= 0 Then
         strLC42 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
         If Val(strLC42) > 0 Then
            '若個案接洽人與客戶檔的預設接洽人相同時不必設定
            PUB_GetContact strAppNo1, strContact, True
            If strLC42 = strContact Then
               strLC42 = ""
            End If
         'Added by Lydia 2022/09/16 排除空白=00
         ElseIf strLC42 = "00" And Trim(cboContact.Text) = "" Then
             strLC42 = ""
         'end 2022/09/16
         End If
      End If
   Else
      strLC42 = "LC42"
   End If
      
   m_SalesST15 = GetST15(txtAdviser(7).Text)
   If frm010001.intModifyKind = 0 Then
      If strLC42 = "LC42" Then strLC42 = ""
      'Modified by Lydia 2021/11/19 +cp17,cp18=> txtAdviser(18), txtAdviser(19)
      SaveDatabase = InsertCaseDatabase(frm010001.intSaveMode, frm010001.intCaseKind, txtSystem, txtCode(0), _
               IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtAdviser(4), txtAdviser(3), txtAdviser(12), txtAdviser(0), txtAdviser(1), _
               txtAdviser(2), txtAdviser(5), txtAdviser(6), txtAdviser(7), txtAdviser(9), txtAdviser(18), txtAdviser(19), txtAdviser(10), txtAdviser(11), strRecieveAuto, strCaseAuto, strLC42, _
               txtAdviser(14), txtAdviser(15), txtAdviser(16), txtAdviser(17), txtAdviser(8))
   Else
      'Modified by Lydia 2021/11/19 +cp17,cp18=> txtAdviser(18), txtAdviser(19)
      SaveDatabase = UpdateCaseDatabase(frm010001.intSaveMode, frm010001.intCaseKind, txtSystem, txtCode(0), _
               IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), txtAdviser(4), txtAdviser(3), txtAdviser(12), txtRecieveCode, txtAdviser(0), txtAdviser(1), _
               txtAdviser(2), txtAdviser(5), txtAdviser(6), txtAdviser(7), txtAdviser(9), txtAdviser(18), txtAdviser(19), txtAdviser(10), txtAdviser(11), strLC42, _
               txtAdviser(14), txtAdviser(15), txtAdviser(16), txtAdviser(17), txtAdviser(8))
   End If
   
   If SaveDatabase = False Then Exit Function '存檔失敗,後續不檢查
   
   '測試解決mail 發不到的時候會存兩筆的錯誤
   On Error GoTo 0    '歸零

   If frm010001.intModifyKind = 0 Then
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
            oContext = "本所案號： " + txtSystem + "-" + txtCode(0) + "-" + txtCode(1) + "-" + txtCode(2) + vbCrLf + "案件名稱： " + txtAdviser(3) + vbCrLf + "收文日： " + ChangeTStringToTDateString(txtAdviser(0)) + vbCrLf + "案件性質： " + lblCaseProperty.Caption + vbCrLf
            oMailCount = ""
            If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(4).Text) <> "" Then
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1)), 1) = "F" Then
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales1 & ";"
                  oContext = oContext & vbCrLf + "當事人1： " + GetCustomerName(ChangeCustomerL(txtAdviser(4).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales1)
               End If
             '秀玲說，其中一個符合就不發了
             Else
                   If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(4).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            '檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtAdviser(4)) <> "" And Trim(txtAdviser(7)) <> "" Then
                If PUB_ChkOldCustomer(True, txtAdviser(4), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If

            If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales2) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(14).Text) <> "" Then
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales2)), 1) = "F" Then
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales2 & ";"
                  oContext = oContext & vbCrLf + "當事人2： " + GetCustomerName(ChangeCustomerL(txtAdviser(14).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales2)
               End If
             '秀玲說，其中一個符合就不發了
             Else
                   If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(14).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            If m_SalesST06 <> "" And Trim(txtAdviser(14)) <> "" And Trim(txtAdviser(7)) <> "" Then
                If PUB_ChkOldCustomer(True, txtAdviser(14), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If

            If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales3) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(15).Text) <> "" Then
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales3)), 1) = "F" Then
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales3 & ";"
                  oContext = oContext & vbCrLf + "當事人3： " + GetCustomerName(ChangeCustomerL(txtAdviser(15).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales3)
               End If
             '秀玲說，其中一個符合就不發了
             Else
                   If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(15).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            '檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtAdviser(15)) <> "" And Trim(txtAdviser(7)) <> "" Then
                If PUB_ChkOldCustomer(True, txtAdviser(15), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If

            If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales4) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(16).Text) <> "" Then
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales4)), 1) = "F" Then
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales4 & ";"
                  oContext = oContext & vbCrLf + "當事人4： " + GetCustomerName(ChangeCustomerL(txtAdviser(16).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales4)
               End If
             Else
                   If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(16).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            '檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtAdviser(16)) <> "" And Trim(txtAdviser(7)) <> "" Then
                If PUB_ChkOldCustomer(True, txtAdviser(16), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If

            If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales5) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(17).Text) <> "" Then
               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales5)), 1) = "F" Then
                  '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail
               Else
                  oMailCount = oMailCount & oStrCuSales5 & ";"
                  oContext = oContext & vbCrLf + "當事人5： " + GetCustomerName(ChangeCustomerL(txtAdviser(17).Text)) + vbCrLf + "原智權人員： " + GetPrjSalesNM(oStrCuSales5)
               End If
             Else
                   If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(17).Text) <> "" Then
                       IsMail = False
                   End If
            End If
            '檢查是否為待活化客戶,並且更新DB
            If m_SalesST06 <> "" And Trim(txtAdviser(17)) <> "" And Trim(txtAdviser(7)) <> "" Then
                If PUB_ChkOldCustomer(True, txtAdviser(17), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
                   IsMail = False
               End If
            End If
            
            '若申請人全空白，不發
            If IsMail = False Or (Trim(txtAdviser(4)) = "" And Trim(txtAdviser(14)) = "" And Trim(txtAdviser(15)) = "" And Trim(txtAdviser(16)) = "" And Trim(txtAdviser(17)) = "") Then
                 oMailCount = ""
            End If
            
            'TXTSYSTEM只判斷1碼,因為FG
            If UCase(Mid(txtSystem, 1, 1)) <> "F" And oMailCount <> "" Then
               '申請人為 X65299 或 X03072 的所有關係企業都不檢查業務區
               If Left(Trim(txtAdviser(4)), 6) <> "X65299" And Left(Trim(txtAdviser(4)), 6) <> "X03072" And _
                  Left(Trim(txtAdviser(14)), 6) <> "X65299" And Left(Trim(txtAdviser(14)), 6) <> "X03072" And _
                  Left(Trim(txtAdviser(15)), 6) <> "X65299" And Left(Trim(txtAdviser(15)), 6) <> "X03072" And _
                  Left(Trim(txtAdviser(16)), 6) <> "X65299" And Left(Trim(txtAdviser(16)), 6) <> "X03072" And _
                  Left(Trim(txtAdviser(17)), 6) <> "X65299" And Left(Trim(txtAdviser(17)), 6) <> "X03072" Then
                  MsgBox "收文智權人員與客戶智權人員不同業務區，準備發 mail ！", , "注意！"
                  '加發秀玲
                  oMailCount = oMailCount & Trim(txtAdviser(7).Text) & ";83002"
                  oContext = oContext & vbCrLf + "收文智權人員： " + lblSales.Caption + vbCrLf + vbCrLf + "智權人員(區)不同！"
                  PUB_SendMail strUserNum, oMailCount, "", "案件收文通知--此案收文非原智權人員(區)！", oContext
               End If
            End If
        End If

End Function

Private Sub Form_Activate()

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
   ReadCaseDatabaseR
End If

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
   IsSaveData = False

   If frm010001.m_blnNewCase = True Then
      cboContact.Locked = False
   Else
      cboContact.Locked = True
   End If
   
   '應收帳款管控:取消預定收款日,改成付款週期
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

PUB_SendMailCache
Where01ToGo intLeaveKind
intLeaveKind = 0

'Set frm010006_1 = Nothing 'Remove by Lydia 2021/12/16 Form2.0會有問題，改在呼叫時清除記憶體變數
stChkForm = Me.Name 'Add by Amy 2021/12/21
End Sub

'Remove by Lydia 2021/04/29 Form 2.0的Label沒有Change模組
'Private Sub lblPetition_Change(Index As Integer)
'If Me.txtSystem.Text = "ACS" Then
'   If frm010001.intModifyKind = 0 Then '新增狀態
'      Me.txtAdviser(3).Text = Me.lblPetition(0).Caption
'   End If
'End If
'End Sub
'end 2021/04/29

Private Sub txtAdviser_Change(Index As Integer)
Select Case Index
             Case 2
                        lblCaseSource.Caption = ""
             Case 4 '當事人1
                        lblPetition(0).Caption = ""
                        txtAdviser(8).Text = ""
             Case 7
                        lblSales.Caption = ""
                        lblDepartment = ""
                        m_SalesST15 = ""
             Case 14, 15, 16, 17 '當事人2,3,4,5
                        lblPetition(Index - 13).Caption = ""
End Select
End Sub

Private Sub txtAdviser_Validate(Index As Integer, Cancel As Boolean)

If Index = 7 Then
   If txtAdviser(Index).Text <> "" And txtAdviser(Index) < "63001" Then
      MsgBox "智權人員不可小於 63001！", , "注意！"
      Cancel = True
      Exit Sub
   End If

   Dim strTemp As String, strTemp1 As String
   If Not ClsPDGetStaff(txtAdviser(Index).Text, strTemp, strTemp1) Then
       Cancel = True
       Exit Sub
   End If
   m_SalesST15 = GetST15(txtAdviser(Index).Text, strTemp1)
   lblSales.Caption = strTemp
   lblDepartment = strTemp1
   
   '創新業務部人員收文控管
   If PUB_ChkIsT10T20("2", txtAdviser(Index).Text, m_Tuser, strTemp) = True Then
        txtAdviser(Index) = m_Tuser
        lblSales.Caption = strTemp
        txtAdviser(Index).SetFocus
        Call txtAdviser_GotFocus(Index)
        Cancel = True
        Exit Sub
   End If
   
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
        
        If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(4).Text) <> "" Then
        '秀玲說，其中一個符合就不發了
        Else
              If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(4).Text) <> "" Then
                  IsMail = False
              End If
        End If
        '檢查是否為待活化客戶
        If m_SalesST06 <> "" And Trim(txtAdviser(4)) <> "" And Trim(txtAdviser(7)) <> "" Then
            If PUB_ChkOldCustomer(False, txtAdviser(4), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
               IsMail = False
            End If
        End If
        
        If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales2) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(14).Text) <> "" Then
        '秀玲說，其中一個符合就不發了
        Else
              If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(14).Text) <> "" Then
                  IsMail = False
              End If
        End If
        '檢查是否為待活化客戶
        If m_SalesST06 <> "" And Trim(txtAdviser(14)) <> "" And Trim(txtAdviser(7)) <> "" Then
            If PUB_ChkOldCustomer(False, txtAdviser(14), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
               IsMail = False
            End If
        End If

        If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales3) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(15).Text) <> "" Then
        '秀玲說，其中一個符合就不發了
        Else
              If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(15).Text) <> "" Then
                  IsMail = False
              End If
        End If
        '檢查是否為待活化客戶
        If m_SalesST06 <> "" And Trim(txtAdviser(15)) <> "" And Trim(txtAdviser(7)) <> "" Then
            If PUB_ChkOldCustomer(False, txtAdviser(15), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
               IsMail = False
            End If
        End If

        If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales4) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(16).Text) <> "" Then
        '秀玲說，其中一個符合就不發了
        Else
              If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(16).Text) <> "" Then
                  IsMail = False
              End If
        End If
        '檢查是否為待活化客戶
        If m_SalesST06 <> "" And Trim(txtAdviser(16)) <> "" And Trim(txtAdviser(7)) <> "" Then
            If PUB_ChkOldCustomer(False, txtAdviser(16), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
               IsMail = False
            End If
        End If

        If m_SalesST15 <> GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales5) And Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(17).Text) <> "" Then
        '秀玲說，其中一個符合就不發了
        Else
              If Trim(txtAdviser(7).Text) <> "" And Trim(txtAdviser(17).Text) <> "" Then
                  IsMail = False
              End If
        End If
        '檢查是否為待活化客戶
        If m_SalesST06 <> "" And Trim(txtAdviser(17)) <> "" And Trim(txtAdviser(7)) <> "" Then
            If PUB_ChkOldCustomer(False, txtAdviser(17), Trim(txtAdviser(7)), m_SalesST15, m_SalesST06) = True Then
               IsMail = False
            End If
        End If
   
        If UCase(Mid(txtSystem, 1, 1)) <> "F" And IsMail = True And (txtAdviser(4) <> "" Or txtAdviser(14) <> "" Or txtAdviser(15) <> "" Or txtAdviser(16) <> "" Or txtAdviser(17) <> "") Then
             '若收文智權人員之ST15為F字頭並且客戶智權人員之ST15也為F字頭則不發Mail，不顯示訊息
             oMailCount = ""
             If txtAdviser(4) <> "" Then
                If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(4).Text), oStrCuSales1)), 1) = "F" Then
                Else
                   oMailCount = "Y"
                End If
             End If
             If txtAdviser(14) <> "" Then
                If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(14).Text), oStrCuSales1)), 1) = "F" Then
                Else
                   oMailCount = "Y"
                End If
             End If
             If txtAdviser(15) <> "" Then
                If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(15).Text), oStrCuSales1)), 1) = "F" Then
                Else
                   oMailCount = "Y"
                End If
             End If
             If txtAdviser(16) <> "" Then
                If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(16).Text), oStrCuSales1)), 1) = "F" Then
                Else
                   oMailCount = "Y"
                End If
             End If
             If txtAdviser(17) <> "" Then
                If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(txtAdviser(17).Text), oStrCuSales1)), 1) = "F" Then
                Else
                   oMailCount = "Y"
                End If
             End If
             
             If Trim(oMailCount) <> "" Then
                '申請人為 X65299 或 X03072 的所有關係企業都不檢查業務區
                If Left(Trim(txtAdviser(4)), 6) <> "X65299" And Left(Trim(txtAdviser(4)), 6) <> "X03072" And _
                   Left(Trim(txtAdviser(14)), 6) <> "X65299" And Left(Trim(txtAdviser(14)), 6) <> "X03072" And _
                   Left(Trim(txtAdviser(15)), 6) <> "X65299" And Left(Trim(txtAdviser(15)), 6) <> "X03072" And _
                   Left(Trim(txtAdviser(16)), 6) <> "X65299" And Left(Trim(txtAdviser(16)), 6) <> "X03072" And _
                   Left(Trim(txtAdviser(17)), 6) <> "X65299" And Left(Trim(txtAdviser(17)), 6) <> "X03072" Then
                   MsgBox "收文智權人員與客戶智權人員不同業務區！", , "注意！"
                End If
             End If
        End If
End If

If CheckKeyIn(Index) = -1 Then
   Cancel = True
   txtAdviser(Index).SetFocus 'Added by Lydia 2021/06/09
   txtAdviser_GotFocus (Index)
End If
End Sub

Private Function CheckKeyIn(ByRef intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, strCusTemp As String
Static strLastCus As String

CheckKeyIn = -1
Select Case intIndex
             Case 0, 5 '收文日CP05, 顧問期間(起)CP53
                        If CheckIsTaiwanDate(txtAdviser(intIndex).Text) Then
                            CheckKeyIn = 1
                        End If
             Case 2  '來源CP11
                        If ClsPDGetCaseSource(txtAdviser(intIndex).Text, strTemp) Then
                           lblCaseSource.Caption = strTemp
                           CheckKeyIn = 1
                        End If
             Case 3  '案件名稱LC05
                        If txtAdviser(intIndex) = "" Then
                           ShowMsg MsgText(1041)
                        ElseIf CheckLengthIsOK(txtAdviser(intIndex), 40) Then
                           CheckKeyIn = 1
                        End If
             Case 4 '當事人1
                        If txtAdviser(intIndex) = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        If intIndex = 4 Then
                           If txtAdviser(intIndex) = txtAdviser(14) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(15) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(16) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(17) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        
                        strCusTemp = txtAdviser(intIndex)
                        '檢查該申請人或代理人狀態，若為不再使用則停在原地
                        'Modified by Lydia 2023/03/06 傳入本所案號 , , , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))
                        If GetCustomerAndState(strCusTemp, strTemp, strTemp1, , , txtSystem, , , , txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) Then
                           txtAdviser(intIndex) = strCusTemp
                           lblPetition(0).Caption = strTemp
                           If strLastCus <> strCusTemp Or txtAdviser(8) = "" Then
                              txtAdviser(8).Text = strTemp1
                              strLastCus = strCusTemp
                           End If
                           CheckKeyIn = 1
                           If ChangeCustomerL(strCusTemp) <> strAppNo1 Then
                              strAppNo1 = ChangeCustomerL(strCusTemp)
                              strExc(10) = cboContact.Tag
                              'Added by Lydia 2022/11/25 區分有無輸入接洽人; ex.P-130652接洽人不是客戶預設接洽人
                              If cboContact.Text <> "" Then
                                 strExc(9) = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
                                 PUB_AddContact strAppNo1, cboContact, strExc(9), True, True, strExc(10)
                              Else
                              'end 2022/22/25
                                 PUB_AddContact strAppNo1, cboContact, , True, True, strExc(10)
                              End If 'Added by Lydia 2022/11/25
                              cboContact.Tag = strExc(10)
                           End If
                        End If
                        If CheckKeyIn <> -1 Then
                           If ClsPDGetCustomerNation(strCusTemp, strNation) Then
                           End If
                        End If
                        If CheckKeyIn = 1 Then
                            If frm010001.m_blnNewCase = True And frm010001.intModifyKind = 0 Then
                                '若輸入9碼且最後一碼不為"0"
                                If Len(Me.txtAdviser(intIndex).Text) = 9 And Right(Me.txtAdviser(intIndex).Text, 1) <> "0" Then
                                    MsgBox "此客戶已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
                                    CheckKeyIn = -1
                                End If
                            End If
                        End If
             Case 6  '顧問期間(止)CP54
                        If CheckIsTaiwanDate(txtAdviser(intIndex).Text) Then
                           If Val(txtAdviser(intIndex - 1)) < Val(txtAdviser(intIndex)) Then
                              CheckKeyIn = 1
                           Else
                              ShowMsg MsgText(1042)
                           End If
                        End If
             Case 7  '智權人員CP13
                        If ClsPDGetStaff(txtAdviser(intIndex).Text, strTemp, strTemp1) Then
                           CheckKeyIn = 1
                        End If
                        lblSales.Caption = strTemp

                        m_SalesST15 = GetST15(txtAdviser(intIndex).Text, strTemp1)
                        lblDepartment = strTemp1
             Case 10  '是否開電腦收據CP32
                        If txtAdviser(intIndex) = "" Or txtAdviser(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 11  '承辦人員CP14
                        If txtAdviser(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else

                        If ClsPDGetStaff(txtAdviser(intIndex), strTemp) Then
                           lblPromoter = strTemp
                           CheckKeyIn = 1
                        End If
                        End If
             Case 12  '分所案號LC16
                        If CheckLengthIsOK(txtAdviser(intIndex), 50) Then
                            CheckKeyIn = 1
                        End If

             Case 13  '預定收款日
                        If txtAdviser(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtAdviser(intIndex).Text) Then
                                If DBDATE(txtAdviser(intIndex).Text) >= DBDATE(txtAdviser(0).Text) Then
                                   CheckKeyIn = 1
                                Else
                                    MsgBox "預定收款日必須>= 收文日", vbOKOnly + vbCritical, "輸入錯誤！"
                                End If
                           End If
                        End If
             Case 14, 15, 16, 17 '當事人2,3,4,5 (LC43,LC44,LC45,LC46)
                        If txtAdviser(intIndex) = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        If intIndex = 14 Then
                           If txtAdviser(intIndex) = txtAdviser(4) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(15) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(16) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(17) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        If intIndex = 15 Then
                           If txtAdviser(intIndex) = txtAdviser(4) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(14) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(16) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(17) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        If intIndex = 16 Then
                           If txtAdviser(intIndex) = txtAdviser(4) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(14) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(15) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(17) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        If intIndex = 17 Then
                           If txtAdviser(intIndex) = txtAdviser(4) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(14) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(15) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                           If txtAdviser(intIndex) = txtAdviser(16) Then
                              MsgBox "當事人不可重覆!", vbOKOnly + vbCritical, "警告!!"
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        strCusTemp = txtAdviser(intIndex)
                        '檢查該申請人或代理人狀態，若為不再使用則停在原地
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

Private Sub txtAdviser_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
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
'                        OpenIme
'             Case Else
'                        CloseIme
'End Select
End Sub

Private Sub txtAdviser_LostFocus(Index As Integer)
'關閉輸入法
'edit by nickc 2007/06/06 切換輸入法改用API
'txtAdviser(Index).IMEMode = 2
'CloseIme 'Removed by Morgan 2016/10/20 會造成 Win7 的切換錯誤
End Sub

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

If (Trim(txtAdviser(4)) = "" Or Trim(lblPetition(0).Caption) = "") Then
    MsgBox "當事人1不可空白！", vbCritical, "檢核資料"
    txtAdviser(4).SetFocus
    Call txtAdviser_GotFocus(4)
    Exit Function
End If

If Trim(txtAdviser(5)) = "" Then
    MsgBox "顧問期間起不可空白！", vbCritical, "檢核資料"
    txtAdviser(5).SetFocus
    Call txtAdviser_GotFocus(5)
    Exit Function
End If

If Trim(txtAdviser(6)) = "" Then
    MsgBox "顧問期間止不可空白！", vbCritical, "檢核資料"
    txtAdviser(6).SetFocus
    Call txtAdviser_GotFocus(6)
    Exit Function
End If

TxtValidate = True
End Function

'新增資料至資料庫
'Modified by Lydia 2021/11/19 +pCP17, pCP18
Private Function InsertCaseDatabase(ByRef intSaveMode As Integer, ByRef intCaseKind As Integer, ByRef pCP01 As String, _
             ByRef pCP02 As String, ByRef pCP03 As String, ByRef pCP04 As String, ByRef pLC11 As String, ByRef pLC05 As String, ByRef pLC16 As String, ByRef pCP05 As String, ByRef pCP10 As String, _
             ByRef pCP11 As String, ByRef pCP53 As String, ByRef pCP54 As String, ByRef pCP13 As String, ByRef pCP16 As String, ByRef pCP17 As String, ByRef pCP18 As String, ByRef pCP32 As String, ByRef pCP14 As String, ByRef pCP09 As String, ByRef pLC02 As String, ByRef pLC42 As String, _
             ByRef pLC43 As String, ByRef pLC44 As String, ByRef pLC45 As String, ByRef pLC46 As String, ByRef pCU30 As String) As Boolean

Dim strSql As String, strAutoNumber As String, bolError As Boolean
Dim cp31 As String, cp12 As String
Dim adoquery As New ADODB.Recordset
Dim strCusReceipt As String  '收據公司別

If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand
   '傳入0為重複之本所案號(新增舊案)，1為正確之本所案號(新增新案)
    pCP05 = ChangeTStringToWString(pCP05)
    pCP53 = ChangeTStringToWString(pCP53)
    pCP54 = ChangeTStringToWString(pCP54)
    pLC11 = ChangeCustomerL(pLC11) '當事人1
    pLC43 = ChangeCustomerL(pLC43) '當事人2
    pLC44 = ChangeCustomerL(pLC44) '當事人3
    pLC45 = ChangeCustomerL(pLC45) '當事人4
    pLC46 = ChangeCustomerL(pLC46) '當事人5

   cnnConnection.BeginTrans
   If intSaveMode = 1 Then
      If pCP02 = "" Then
         If ClsPDGetAutoNumber(pCP01, strAutoNumber, True, False) Then
            pCP02 = strAutoNumber
         Else
            bolError = True
         End If
      End If
      If bolError = False Then
         pLC02 = pCP02
         '收據公司別
         If intCaseKind <> 顧問 Then
            strCusReceipt = GetReceiptCmp(Mid(pLC11, 1, 8), Mid(pLC11, 9, 1), pCP01, "000")
         End If
         Select Case intCaseKind
                Case 法務
                    strSql = "insert into lawcase (lc01,lc02,lc03,lc04,lc05,lc06,lc07,lc11,lc15,lc16,lc42,lc43,lc44,lc45,lc46,lc48) " + _
                        "values (" + CNULL(pCP01) + "," + CNULL(pCP02) + "," + CNULL(pCP03) + "," + CNULL(pCP04) + "," + CNULL(ChgSQL(pLC05)) + "," + _
                        "null, null," + CNULL(pLC11) + ",'000' ," + CNULL(ChgSQL(pLC16)) + "," + CNULL(pLC42) + "," + CNULL(pLC43) + "," + CNULL(pLC44) + "," + CNULL(pLC45) + "," + CNULL(pLC46) + "," + CNULL(strCusReceipt) + ")"
                    cnnConnection.Execute strSql
         End Select
         cp31 = "Y"
      Else
         bolError = True
      End If
   End If
   If bolError = False Then

      If ClsPDGetAutoNumber(Left(pCP09, 1), strAutoNumber, True, True) Then
         pCP09 = pCP09 + strAutoNumber
         
         '有★★的應收帳款簽核控管
         m_CP150 = ""
         If Check2.Value = 1 Then m_CP150 = "Y"
         
         cp12 = PUB_GetStaffST15(pCP13, 1)
        'Modified by Lydia 2021/11/19 輸入規費、點數
         'strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp16, " + _
             "cp17,cp18,cp31,cp32,cp53,cp54,CP150) values (" + CNULL(pCP01) + "," + CNULL(pCP02) + "," + CNULL(pCP03) + "," + CNULL(pCP04) + "," + CNULL(pCP05) + "," + _
              CNULL(pCP09) + "," + CNULL(pCP10) + "," + CNULL(pCP11) + "," + CNULL(cp12) + "," + CNULL(pCP13) + "," + CNULL(pCP14) + "," + CNULL(pCP16) + "," + _
             "0, " + CNULL(IIf(Val(pCP16) / 1000 = 0, "", Val(pCP16) / 1000)) + "," + CNULL(cp31) + "," + CNULL(pCP32) + "," + CNULL(pCP53) + "," + CNULL(pCP54) + "," + CNULL(m_CP150) + ")"
         strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp16, " + _
             "cp17,cp18,cp31,cp32,cp53,cp54,CP150) values (" + CNULL(pCP01) + "," + CNULL(pCP02) + "," + CNULL(pCP03) + "," + CNULL(pCP04) + "," + CNULL(pCP05) + "," + _
              CNULL(pCP09) + "," + CNULL(pCP10) + "," + CNULL(pCP11) + "," + CNULL(cp12) + "," + CNULL(pCP13) + "," + CNULL(pCP14) + "," + CNULL(pCP16) + "," + _
              CNULL(pCP17) + "," + CNULL(pCP18) + "," + CNULL(cp31) + "," + CNULL(pCP32) + "," + CNULL(pCP53) + "," + CNULL(pCP54) + "," + CNULL(m_CP150) + ")"
         cnnConnection.Execute strSql, intI
         
         '若為接洽記錄單(櫃台收文), 費用可改時才做，否則已收款資料會被還原
         If frm010001.intChoose = 0 And txtAdviser(9).Enabled = True Then
             '未收金額 = 費用
             strSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(pCP09)
             cnnConnection.Execute strSql
         End If
         'Added by Lydia 2022/11/29 非內部收文並且有費用，先統一設定CP20=Null ;
         If frm010001.intChoose = 0 And Val(pCP16) > 0 Then
             strSql = "update caseprogress set cp20=null where cp09=" + CNULL(pCP09)
             cnnConnection.Execute strSql
         End If
         'end 2022/11/29
         '若為內部收文作業時, 案件進度檔的是否向客戶收款設定為"N"
         If frm010001.intChoose = 1 Then
            strSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(pCP09)
            cnnConnection.Execute strSql
         End If
   
         strSql = "update customer set cu30=" + CNULL(pCU30) + " where cu01=" + CNULL(Mid(pLC11, 1, 8)) + " and cu02=" + CNULL(Mid(pLC11, 9, 1))
         cnnConnection.Execute strSql
        
        '若在下一程序檔只抓到一筆資料時, 才要抓下一程序檔的總收文號更新案件進度檔的相關總收文號
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select np01 from nextprogress where np02 = '" & pCP01 & "' and np03 = '" & pCP02 & "' and np04 = '" & pCP03 & "' and np05 = '" & pCP04 & "' and np06 is null and np07 = '" & pCP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
            If IsNull(adoquery.Fields(0).Value) = False Then
               cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & pCP09 & "'"
            End If
         End If
         adoquery.Close
      Else
         bolError = True
      End If
   End If

   
   If bolError Then
      cnnConnection.RollbackTrans
      ShowMsg MsgText(9004)
      IsSaveData = False
   Else
      cnnConnection.CommitTrans
      InsertCaseDatabase = True
      txtCode(0) = pCP02
   End If
   txtCode(0) = pCP02
   Exit Function
ErrHand:
   cnnConnection.RollbackTrans
   ShowMsg MsgText(9004)
   IsSaveData = False
End Function

'修改資料庫
'Modified by Lydia 2021/11/19 +pCP17, pCP18
Private Function UpdateCaseDatabase(ByRef intSaveMode As Integer, ByRef intCaseKind As Integer, ByRef pCP01 As String, _
             ByRef pCP02 As String, ByRef pCP03 As String, ByRef pCP04 As String, ByRef pLC11 As String, ByRef pLC05 As String, ByRef pLC16 As String, ByRef pCP09 As String, ByRef pCP05 As String, ByRef pCP10 As String, _
             ByRef pCP11 As String, ByRef pCP53 As String, ByRef pCP54 As String, ByRef pCP13 As String, ByRef pCP16 As String, ByRef pCP17 As String, ByRef pCP18 As String, ByRef pCP32 As String, ByRef pCP14 As String, ByRef pLC42 As String, _
             ByRef pLC43 As String, ByRef pLC44 As String, ByRef pLC45 As String, ByRef pLC46 As String, ByRef pCU30 As String) As Boolean

Dim strSql As String
Dim adoquery As New ADODB.Recordset
Dim strCusReceipt As String  '收據公司別

'add by nickc 2007/12/12
If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand
 pCP05 = ChangeTStringToWString(pCP05)
 pCP53 = ChangeTStringToWString(pCP53)
 pCP54 = ChangeTStringToWString(pCP54)
 pLC11 = ChangeCustomerL(pLC11) '當事人1
 pLC43 = ChangeCustomerL(pLC43) '當事人2
 pLC44 = ChangeCustomerL(pLC44) '當事人3
 pLC45 = ChangeCustomerL(pLC45) '當事人4
 pLC46 = ChangeCustomerL(pLC46) '當事人5
 If intCaseKind <> 顧問 Then
    strCusReceipt = GetReceiptCmp(Mid(pLC11, 1, 8), Mid(pLC11, 9, 1), pCP01, "000")
 End If
 
cnnConnection.BeginTrans

    Select Case intCaseKind
          Case 法務
                 strSql = "update lawcase set lc05=" + CNULL(ChgSQL(pLC05)) + ", lc11=" + CNULL(pLC11) + ", lc16=" + CNULL(ChgSQL(pLC16)) + _
                       ", lc43=" + CNULL(pLC43) + ", lc44=" + CNULL(pLC44) + ", lc45=" + CNULL(pLC45) + ", lc46=" + CNULL(pLC46) + ", lc48=" + CNULL(strCusReceipt)
                 strSql = strSql + " where lc01=" + CNULL(pCP01) + " and lc02=" + CNULL(pCP02) + " and lc03=" + CNULL(pCP03) + " and lc04=" + CNULL(pCP04)
                 cnnConnection.Execute strSql
    End Select
    
    '有★★的應收帳款簽核控管
    m_CP150 = ""
    If Check2.Value = 1 Then m_CP150 = "Y"
    'Modified by Lydia 2021/11/19 輸入規費、點數
    'strSql = "update caseprogress set cp05=" + CNULL(pCP05) + ",cp10=" + CNULL(pCP10) + ",cp11=" + CNULL(pCP11) + ",cp53=" + CNULL(pCP53) + ",cp54=" + CNULL(pCP54) + ",cp13=" + CNULL(pCP13) + _
       ",cp14=" + CNULL(pCP14) + ",cp16=" + CNULL(pCP16) + ",cp32=" + CNULL(pCP32) + ",cp18=" & CNULL(IIf(Val(pCP16) / 1000 = 0, "", Val(pCP16) / 1000)) & ",cp150=" & CNULL(m_CP150) & " where cp09=" + CNULL(pCP09)
    strSql = "update caseprogress set cp05=" + CNULL(pCP05) + ",cp10=" + CNULL(pCP10) + ",cp11=" + CNULL(pCP11) + ",cp53=" + CNULL(pCP53) + ",cp54=" + CNULL(pCP54) + ",cp13=" + CNULL(pCP13) + _
       ",cp14=" + CNULL(pCP14) + ",cp16=" + CNULL(pCP16) + " ,cp17=" + CNULL(pCP17) + ",cp18=" + CNULL(pCP18) + " ,cp32=" + CNULL(pCP32) + " ,cp150=" & CNULL(m_CP150) & " where cp09=" + CNULL(pCP09)
    cnnConnection.Execute strSql
    strSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(pCP13) + ") where cp09=" + CNULL(pCP09)
    cnnConnection.Execute strSql

    '若為接洽記錄單(櫃台收文), 費用可改時才做，否則已收款資料會被還原
    If frm010001.intChoose = 0 And txtAdviser(9).Enabled = True Then
        '未收金額 = 費用
        strSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(pCP09)
        cnnConnection.Execute strSql
    End If
    'Added by Lydia 2022/11/29 非內部收文並且有費用，先統一設定CP20=Null ;
    If frm010001.intChoose = 0 And Val(pCP16) > 0 Then
        strSql = "update caseprogress set cp20=null where cp09=" + CNULL(pCP09)
        cnnConnection.Execute strSql
    End If
    'end 2022/11/29
    '若為內部收文作業時, 案件進度檔的是否向客戶收款設定為"N"
    If frm010001.intChoose = 1 Then
       strSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(pCP09)
       cnnConnection.Execute strSql
    End If

    strSql = "update customer set cu30=" + CNULL(pCU30) + " where cu01=" + CNULL(Mid(pLC11, 1, 8)) + " and cu02=" + CNULL(Mid(pLC11, 9, 1))
    cnnConnection.Execute strSql
    
    adoquery.CursorLocation = adUseClient
    adoquery.Open "select np01 from nextprogress where np02 = '" & pCP01 & "' and np03 = '" & pCP02 & "' and np04 = '" & pCP03 & "' and np05 = '" & pCP04 & "' and np06 is null and np07 = '" & pCP10 & "'", cnnConnection, adOpenStatic, adLockReadOnly
    '若在下一程序檔只抓到一筆資料時, 才要抓下一程序檔的總收文號更新案件進度檔的相關總收文號
    If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
       If IsNull(adoquery.Fields(0).Value) = False Then
          cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & pCP09 & "'"
       End If
    End If
    adoquery.Close

cnnConnection.CommitTrans
UpdateCaseDatabase = True
Exit Function
ErrHand:
cnnConnection.RollbackTrans
ShowMsg MsgText(9004)

IsSaveData = False
End Function

'讀取資料庫
'Modified by Lydia 2021/11/19 +pCP17,pCP18
Private Function ReadCaseDatabase(ByRef intModifyKind As Integer, ByRef intCaseKind As Integer, ByRef pCP01 As String, _
             ByRef pCP02 As String, ByRef pCP03 As String, ByRef pCP04 As String, ByRef pLC11 As String, ByRef pLC05 As String, ByRef pLC16 As String, ByRef pCP09 As String, ByRef pCP05 As String, ByRef pCP10 As String, _
             ByRef pCP11 As String, ByRef pCP53 As String, ByRef pCP54 As String, ByRef pCP13 As String, ByRef pCP16 As String, ByRef pCP17 As String, ByRef pCP18 As String, ByRef pCP32 As String, ByRef pCP14 As String, ByRef pLC42 As String, _
             ByRef pLC43 As String, ByRef pLC44 As String, ByRef pLC45 As String, ByRef pLC46 As String, ByRef pCU30 As String, ByRef pCP150 As String) As Boolean

Dim strSql As String, rsRecordset As New ADODB.Recordset, strTemp As String
Dim stCP60 As String '收據號碼
   
On Error GoTo ErrHand

If intModifyKind <> 0 Then
   'Modified by Lydia 2021/11/19 +cp17,cp18
   strSql = "select cp05,cp10,cp11,cp13,cp14,cp16,cp17,cp18,cp32,cp53,cp54,cp60,cp150 from caseprogress where cp09='" + pCP09 + "'"
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      pCP05 = IIf(IsNull(rsRecordset.Fields("cp05")), "", rsRecordset.Fields("cp05"))
      If pCP05 <> "" Then pCP05 = ChangeWStringToTString(pCP05)
      pCP10 = IIf(IsNull(rsRecordset.Fields("cp10")), "", rsRecordset.Fields("cp10"))
      pCP11 = IIf(IsNull(rsRecordset.Fields("cp11")), "", rsRecordset.Fields("cp11"))
      pCP13 = IIf(IsNull(rsRecordset.Fields("cp13")), "", rsRecordset.Fields("cp13"))
      pCP14 = IIf(IsNull(rsRecordset.Fields("cp14")), "", rsRecordset.Fields("cp14"))
      pCP16 = IIf(IsNull(rsRecordset.Fields("cp16")), "", rsRecordset.Fields("cp16"))
      'Added by Lydia 2021/11/19
      pCP17 = IIf(IsNull(rsRecordset.Fields("cp17")), "", rsRecordset.Fields("cp17"))
      pCP18 = IIf(IsNull(rsRecordset.Fields("cp18")), "", rsRecordset.Fields("cp18"))
      'end 2021/11/19
      pCP53 = IIf(IsNull(rsRecordset.Fields("cp53")), "", rsRecordset.Fields("cp53"))
      If pCP53 <> "" Then pCP53 = ChangeWStringToTString(pCP53)
      pCP54 = IIf(IsNull(rsRecordset.Fields("cp54")), "", rsRecordset.Fields("cp54"))
      If pCP54 <> "" Then pCP54 = ChangeWStringToTString(pCP54)
      pCP32 = IIf(IsNull(rsRecordset.Fields("cp32")), "", rsRecordset.Fields("cp32"))
      stCP60 = IIf(IsNull(rsRecordset.Fields("cp60")), "", rsRecordset.Fields("cp60"))
      pCP150 = IIf(IsNull(rsRecordset.Fields("cp150")), "", rsRecordset.Fields("cp150"))
      If stCP60 <> "" Then
         '鎖定：LC16,預定收款日,LC45,CP13
         txtAdviser(12).Enabled = False: txtAdviser(13).Enabled = False: txtAdviser(16).Enabled = False
         txtAdviser(10).Enabled = False
      End If
   Else
      ShowMsg MsgText(1502)
      rsRecordset.Close
      Exit Function
   End If
   rsRecordset.Close
Else
'      If GetNextProgressDate(pCP01, pCP02, pCP03, pCP04, pCP10, cp06, cp07, CP64, cp13) = False Then
'         Exit Function
'      End If
End If

strSql = ""
    Select Case intCaseKind
             Case 法務
                    strSql = "SELECT LC05, LC11, LC16, LC43, LC44, LC42 ,LC45 ,LC46  FROM LAWCASE WHERE LC01=" + CNULL(pCP01) + " AND LC02=" + CNULL(pCP02) + " AND LC03=" + CNULL(pCP03) + " AND LC04=" + CNULL(pCP04)
    End Select

If strSql = "" Then
  ShowMsg "找不到此本所案號在基本檔之資料"
  Exit Function
End If

rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSql, cnnConnection
If rsRecordset.RecordCount > 0 Then
   '案件名稱_中文
   pLC05 = IIf(IsNull(rsRecordset.Fields("lc05")), "", rsRecordset.Fields("lc05"))
   '分所案號
   pLC16 = IIf(IsNull(rsRecordset.Fields("lc16")), "", rsRecordset.Fields("lc16"))
   '當事人聯絡人編號
   pLC42 = IIf(IsNull(rsRecordset.Fields("lc42")), "", rsRecordset.Fields("lc42"))
   strExc(10) = cboContact.Tag
   PUB_AddContact pLC11, cboContact, pLC42, True, True, strExc(10)
   cboContact.Tag = strExc(10)
   '當事人1~5
   pLC11 = IIf(IsNull(rsRecordset.Fields("lc11")), "", rsRecordset.Fields("lc11"))
   pLC43 = IIf(IsNull(rsRecordset.Fields("lc43")), "", rsRecordset.Fields("lc43"))
   pLC44 = IIf(IsNull(rsRecordset.Fields("lc44")), "", rsRecordset.Fields("lc44"))
   pLC45 = IIf(IsNull(rsRecordset.Fields("lc45")), "", rsRecordset.Fields("lc45"))
   pLC46 = IIf(IsNull(rsRecordset.Fields("lc46")), "", rsRecordset.Fields("lc46"))

   
   rsRecordset.Close
   strSql = "select cu30 from customer where cu01=" + CNULL(Mid(pLC11, 1, 8)) + " AND cu02=" + CNULL(Mid(pLC11, 9, 1))
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      pCU30 = IIf(IsNull(rsRecordset.Fields("cu30")), "", rsRecordset.Fields("cu30"))
      pLC11 = ChangeCustomerS(pLC11)
      pLC43 = ChangeCustomerS(pLC43)
      pLC44 = ChangeCustomerS(pLC44)
      pLC45 = ChangeCustomerS(pLC45)
      pLC46 = ChangeCustomerS(pLC46)
      ReadCaseDatabase = True
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

Set rsRecordset = Nothing
Exit Function
ErrHand:
   ShowMsg "資料讀取失敗,請洽系統管理者!"
End Function

Private Sub ReadCaseDatabaseR()
Dim rCP01 As String, rCP02 As String, rCP03 As String, rCP04 As String, rCP05 As String, rCP10 As String, rCP11 As String
Dim rCP53 As String, rCP54 As String, rCP13 As String, rCP16 As String, rCP32 As String, rCP14 As String
Dim rLC05 As String, rLC11 As String, rLC16 As String, rLC42 As String
Dim rLC43 As String, rLC44 As String, rLC45 As String, rLC46 As String, rCU30 As String
Dim rCP150 As String
Dim rCP17 As String, rCP18 As String 'Added by Lydia 2021/11/19 +規費CP17、點數CP18
Dim rt As Boolean, strTemp As String

'Modified by Lydia 2021/11/19 + rCP17, rCP18
rt = ReadCaseDatabase(frm010001.intModifyKind, frm010001.intCaseKind, txtSystem, txtCode(0), _
       IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), rLC11, rLC05, rLC16, txtRecieveCode, rCP05, rCP10, _
       rCP11, rCP53, rCP54, rCP13, rCP16, rCP17, rCP18, rCP32, rCP14, rLC42, rLC43, rLC44, rLC45, rLC46, rCU30, rCP150)

If rt Then
   If frm010001.intModifyKind <> 0 Then
      txtAdviser(0) = rCP05
      txtAdviser(2) = rCP11
      txtAdviser(5) = rCP53
      txtAdviser(6) = rCP54
      txtAdviser(7) = rCP13
      txtAdviser(8) = rCU30
      txtAdviser(9) = rCP16
      'Added by Lydia 2021/11/19
      txtAdviser(18) = rCP17
      txtAdviser(19) = rCP18
      'end 2021/11/19
      txtAdviser(10) = rCP32
      txtAdviser(11) = rCP14
      CheckKeyIn 7
      CheckKeyIn 2
      CheckKeyIn 11
      txtAdviser(1) = rCP10
      If ClsPDGetCaseProperty(txtSystem, rCP10, strTemp) Then
         lblCaseProperty.Caption = strTemp
      End If
      If rCP150 = "Y" Then
         Check2.Value = 1
      End If
   End If
   txtAdviser(3) = rLC05
   txtAdviser(4) = rLC11
   CheckKeyIn 4
   txtAdviser(14) = rLC43
   txtAdviser(15) = rLC44
   txtAdviser(16) = rLC45
   txtAdviser(17) = rLC46
   CheckKeyIn 14
   CheckKeyIn 15
   CheckKeyIn 16
   CheckKeyIn 17
Else
   If frm010001.intModifyKind <> 0 Then
      MsgBox "讀取資料時發生錯誤!!", vbCritical
      bolLeave = True
      Unload Me
   Else
      txtAdviser(0) = rCP05
      txtAdviser(2) = rCP11
      txtAdviser(5) = rCP53
      txtAdviser(6) = rCP54
      'txtAdviser(7) = rcp13  '2011/5/11 cancel by sonia 偶而改智權人員收文會忘記打所以不自動帶
      txtAdviser(8) = rCU30
      txtAdviser(9) = rCP16
      'Added by Lydia 2021/11/19
      txtAdviser(18) = rCP17
      txtAdviser(19) = rCP18
      'end 2021/11/19
      txtAdviser(10) = rCP32
      txtAdviser(11) = rCP14
      CheckKeyIn 7
      CheckKeyIn 2
      CheckKeyIn 11
      txtAdviser(1) = rCP10
      If ClsPDGetCaseProperty(txtSystem, txtAdviser(1), strTemp) Then
         lblCaseProperty.Caption = strTemp
      End If
   End If
End If
'分所案號: 舊案若有分所案號則帶出並鎖住，若無分所案號則開放可輸入並更新回基本檔。
txtAdviser(12).Locked = False
txtAdviser(12) = rLC16
If Trim(rLC16) <> "" Then
    txtAdviser(12).Locked = True
End If

If frm010001.intChoose = 1 Then
   txtAdviser(2) = "90"
   CheckKeyIn (2)
End If

End Sub

