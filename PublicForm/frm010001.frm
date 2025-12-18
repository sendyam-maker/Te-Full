VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010001 
   BorderStyle     =   1  '單線固定
   ClientHeight    =   4530
   ClientLeft      =   5700
   ClientTop       =   5770
   ClientWidth     =   8050
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8050
   Begin VB.Frame FraRecvList 
      Caption         =   "多案收文: 首筆案號="
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1965
      Left            =   4560
      TabIndex        =   57
      Top             =   990
      Visible         =   0   'False
      Width           =   3465
      Begin VB.CommandButton Command1 
         Caption         =   "刪除"
         Height          =   400
         Index           =   1
         Left            =   330
         TabIndex        =   59
         Top             =   930
         Width           =   720
      End
      Begin VB.CommandButton Command1 
         Caption         =   "新增"
         Height          =   400
         Index           =   0
         Left            =   330
         TabIndex        =   58
         Top             =   420
         Width           =   720
      End
      Begin VB.ListBox List1 
         Height          =   1120
         ItemData        =   "frm010001.frx":0000
         Left            =   1110
         List            =   "frm010001.frx":0002
         MultiSelect     =   2  '進階多重選取
         TabIndex        =   60
         Top             =   210
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "件數:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   120
         TabIndex        =   63
         Top             =   1530
         Width           =   495
      End
      Begin VB.Label LblCaseNo 
         Caption         =   "LblCaseNo"
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
         Height          =   315
         Left            =   1890
         TabIndex        =   62
         Top             =   0
         Width           =   1635
      End
      Begin VB.Label LblCnt 
         Caption         =   "LblCnt"
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
         Height          =   315
         Left            =   660
         TabIndex        =   61
         Top             =   1530
         Width           =   585
      End
   End
   Begin VB.Frame FraRecv 
      BorderStyle     =   0  '沒有框線
      Caption         =   "FrmRecv"
      Height          =   495
      Left            =   2910
      TabIndex        =   56
      Top             =   30
      Visible         =   0   'False
      Width           =   1965
      Begin VB.CommandButton cmdMRecv 
         Caption         =   "多案收文(&R)"
         Height          =   400
         Left            =   30
         Style           =   1  '圖片外觀
         TabIndex        =   17
         Top             =   30
         Width           =   1155
      End
   End
   Begin VB.Frame fraLastCaseCode 
      BorderStyle     =   0  '沒有框線
      Height          =   372
      Left            =   3600
      TabIndex        =   36
      Top             =   1020
      Visible         =   0   'False
      Width           =   3132
      Begin VB.Label lblCaseCode 
         Height          =   252
         Left            =   1080
         TabIndex        =   37
         Top             =   0
         Width           =   1932
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   225
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5940
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5100
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   45
      Width           =   800
   End
   Begin VB.Frame fraRecieve 
      BorderStyle     =   0  '沒有框線
      Height          =   495
      Left            =   1710
      TabIndex        =   26
      Top             =   900
      Width           =   1812
      Begin VB.TextBox txtRecieveCode 
         Height          =   300
         Index           =   1
         Left            =   720
         MaxLength       =   6
         TabIndex        =   0
         Top             =   120
         Width           =   1092
      End
      Begin VB.TextBox txtRecieveCode 
         Height          =   300
         Index           =   0
         Left            =   384
         MaxLength       =   2
         TabIndex        =   24
         Top             =   120
         Width           =   372
      End
      Begin VB.Label lblReciveCode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   252
      End
   End
   Begin VB.Frame fraCode 
      BorderStyle     =   0  '沒有框線
      Height          =   3015
      Left            =   300
      TabIndex        =   27
      Top             =   1440
      Width           =   6940
      Begin VB.Frame fraFMP 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   315
         Left            =   2340
         TabIndex        =   52
         Top             =   450
         Width           =   6015
         Begin VB.TextBox txtNA01 
            Height          =   300
            Left            =   3720
            MaxLength       =   3
            TabIndex        =   47
            Top             =   0
            Width           =   585
         End
         Begin VB.TextBox txtFMP 
            Height          =   300
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   46
            Top             =   0
            Width           =   372
         End
         Begin VB.Label lblNation 
            Height          =   255
            Left            =   4350
            TabIndex        =   55
            Top             =   23
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "申請國家："
            Height          =   255
            Left            =   2760
            TabIndex        =   54
            Top             =   23
            Width           =   945
         End
         Begin VB.Label Label6 
            Caption         =   "是否為FMP案：             (Y:是)"
            Height          =   255
            Left            =   0
            TabIndex        =   53
            Top             =   23
            Width           =   2595
         End
      End
      Begin VB.Frame fraNA239 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   315
         Left            =   5100
         TabIndex        =   50
         Top             =   1440
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox txtCaseNa239 
            Height          =   270
            Left            =   1230
            MaxLength       =   12
            TabIndex        =   16
            Top             =   22
            Width           =   1400
         End
         Begin VB.Label Label8 
            Caption         =   "歐盟案案號："
            Height          =   195
            Left            =   0
            TabIndex        =   51
            Top             =   60
            Width           =   1095
         End
      End
      Begin VB.Frame fraLOS 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Height          =   315
         Left            =   3930
         TabIndex        =   48
         Top             =   1440
         Width           =   2895
         Begin VB.TextBox txtLOS15 
            Height          =   270
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   7
            Top             =   37
            Width           =   1300
         End
         Begin VB.Label Label7 
            Caption         =   "案源單號："
            Height          =   225
            Left            =   0
            TabIndex        =   49
            Top             =   60
            Width           =   915
         End
      End
      Begin VB.Frame fraElse 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1800
         TabIndex        =   28
         Top             =   -90
         Width           =   2892
         Begin VB.TextBox txtCode 
            Height          =   300
            Index           =   2
            Left            =   1830
            MaxLength       =   2
            TabIndex        =   4
            Top             =   150
            Width           =   492
         End
         Begin VB.TextBox txtCode 
            Height          =   300
            Index           =   1
            Left            =   1380
            MaxLength       =   1
            TabIndex        =   3
            Top             =   150
            Width           =   372
         End
         Begin VB.TextBox txtCode 
            Height          =   300
            Index           =   0
            Left            =   0
            MaxLength       =   6
            TabIndex        =   2
            Top             =   150
            Width           =   1332
         End
      End
      Begin VB.Frame fraTF 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Left            =   1920
         TabIndex        =   29
         Top             =   -60
         Visible         =   0   'False
         Width           =   2772
         Begin VB.TextBox txtTFCode 
            Height          =   300
            Index           =   3
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   23
            Top             =   120
            Width           =   492
         End
         Begin VB.TextBox txtTFCode 
            Height          =   300
            Index           =   2
            Left            =   1680
            MaxLength       =   1
            TabIndex        =   22
            Top             =   120
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   300
            Index           =   1
            Left            =   1200
            MaxLength       =   1
            TabIndex        =   21
            Top             =   120
            Width           =   372
         End
         Begin VB.TextBox txtTFCode 
            Height          =   300
            Index           =   0
            Left            =   0
            MaxLength       =   5
            TabIndex        =   20
            Top             =   120
            Width           =   1092
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   435
         Left            =   0
         TabIndex        =   41
         Top             =   1260
         Visible         =   0   'False
         Width           =   2805
         Begin VB.TextBox textCP24 
            Height          =   264
            Left            =   1080
            MaxLength       =   1
            TabIndex        =   8
            Top             =   120
            Width           =   405
         End
         Begin VB.Label Label59 
            Caption         =   "案件准駁："
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   180
            Width           =   945
         End
         Begin VB.Label Label60 
            Caption         =   "(1:准 , 2:駁)"
            Height          =   255
            Left            =   1530
            TabIndex        =   42
            Top             =   150
            Width           =   1035
         End
      End
      Begin VB.TextBox textYear 
         Height          =   270
         Left            =   4290
         MaxLength       =   5
         TabIndex        =   14
         Top             =   936
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   5370
         MaxLength       =   5
         TabIndex        =   15
         Top             =   936
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox textCP05 
         Height          =   300
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   5
         Top             =   480
         Width           =   1164
      End
      Begin VB.Frame fraPatition 
         BorderStyle     =   0  '沒有框線
         Height          =   1560
         Left            =   -15
         TabIndex        =   33
         Top             =   1410
         Visible         =   0   'False
         Width           =   6492
         Begin VB.TextBox txtPetitionx 
            Height          =   300
            Index           =   5
            Left            =   1710
            MaxLength       =   9
            TabIndex        =   13
            Top             =   1200
            Width           =   1332
         End
         Begin VB.TextBox txtPetitionx 
            Height          =   300
            Index           =   4
            Left            =   1710
            MaxLength       =   9
            TabIndex        =   12
            Top             =   900
            Width           =   1332
         End
         Begin VB.TextBox txtPetitionx 
            Height          =   300
            Index           =   3
            Left            =   1710
            MaxLength       =   9
            TabIndex        =   11
            Top             =   600
            Width           =   1332
         End
         Begin VB.TextBox txtPetitionx 
            Height          =   300
            Index           =   2
            Left            =   1710
            MaxLength       =   9
            TabIndex        =   10
            Top             =   300
            Width           =   1332
         End
         Begin VB.TextBox txtPetition 
            Height          =   300
            Left            =   1710
            MaxLength       =   9
            TabIndex        =   9
            Top             =   0
            Width           =   1332
         End
         Begin MSForms.Label lblPetitionNamex 
            Height          =   255
            Index           =   5
            Left            =   3150
            TabIndex        =   64
            Top             =   1200
            Width           =   3255
            VariousPropertyBits=   27
            Size            =   "5741;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetitionNamex 
            Height          =   255
            Index           =   4
            Left            =   3150
            TabIndex        =   68
            Top             =   906
            Width           =   3255
            VariousPropertyBits=   27
            Size            =   "5741;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetitionNamex 
            Height          =   255
            Index           =   3
            Left            =   3150
            TabIndex        =   67
            Top             =   614
            Width           =   3255
            VariousPropertyBits=   27
            Size            =   "5741;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetitionNamex 
            Height          =   255
            Index           =   2
            Left            =   3150
            TabIndex        =   66
            Top             =   322
            Width           =   3255
            VariousPropertyBits=   27
            Size            =   "5741;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.Label lblPetitionName 
            Height          =   255
            Left            =   3150
            TabIndex        =   65
            Top             =   30
            Width           =   3255
            VariousPropertyBits=   27
            Size            =   "5741;450"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label5 
            Caption         =   "移轉、讓與申請人："
            Height          =   330
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   1710
         End
      End
      Begin VB.TextBox txtCaseProperty 
         Height          =   300
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   6
         Top             =   936
         Width           =   732
      End
      Begin VB.TextBox txtSystem 
         Height          =   300
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   1
         Top             =   60
         Width           =   675
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "繳費年度：第           年至第           年"
         Height          =   180
         Index           =   1
         Left            =   3180
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   2856
      End
      Begin VB.Label Label4 
         Caption         =   "收文日："
         Height          =   252
         Left            =   0
         TabIndex        =   39
         Top             =   504
         Width           =   972
      End
      Begin VB.Label lblCasePropertyName 
         Height          =   375
         Left            =   1890
         TabIndex        =   32
         Top             =   960
         Width           =   1740
      End
      Begin VB.Label Label3 
         Caption         =   "案件性質："
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "本所案號："
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   90
         Width           =   975
      End
   End
   Begin VB.Label lblTCT 
      Caption         =   "中說或其他收文號："
      Height          =   300
      Left            =   240
      TabIndex        =   44
      Top             =   495
      Width           =   1635
   End
   Begin VB.Label lblRecieveKind 
      Caption         =   "上一筆之收文號："
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1050
      Width           =   1455
   End
   Begin VB.Label lblTCTNO 
      Caption         =   "lblTCTNO"
      Height          =   540
      Left            =   1920
      TabIndex        =   45
      Top             =   480
      Width           =   4875
   End
End
Attribute VB_Name = "frm010001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/5/7 改成Form2.0 (無)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

'intChoose   0:收文   1:內部收文
Public intChoose As Integer
'intSaveMode，To A:-1為錯誤之本所案號，0為重複之本所案號(新增舊案)，1為正確之本所案號(新增新案)
'intCaseKind，1為專利，2為商標，3為法務，4為顧問，5為專利(服)，6為商標(服)，7為法務(服)，8為顧問(服)
'strCaseName，取得案件名稱
Public intSaveMode As Integer, intCaseKind As Integer, strCaseName As String
'strReceiveKind ，A為接洽記錄單，B為政府機關來函
Dim strReceiveKind As String
'intReceiveKind=0為接洽紀錄單;=1為政府來函
'intModifyKind=0為新增;=1為修改;=2為查詢
Public intReceiveKind As Integer, intModifyKind As Integer
Dim adoquery As New ADODB.Recordset
'Add By Cheng 2003/09/08
Public m_blnNewCase As Boolean '判斷是否為新案(無流水號或基本檔無資料)
'Dim m_TM10 As String   '94.1.12 add by sonia 'Remove by Lydia 2020/06/08 統一用m_Nation
'Add by Morgan 2006/6/27
Dim m_bolStopOnTxtPetition As Boolean '是否停駐於受讓人1
'Add By Sindy 2009/07/06
Dim m_bolStopOntextYear As Boolean '是否停駐於起始年度
'Add By Sindy 2012/2/23
Dim strSK02 As String
Dim strSK03 As String
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, m_CP09 As String
Dim m_CP12 As String, m_CP13 As String, m_CP13_2 As String
Dim oSubject As String, oContext As String
Public strMailNote As String, strTo As String
Dim strApp As String, strNation As String
'2012/2/23 End
Dim intCP135 As Integer, intCP136 As Integer 'Add by Amy 2013/08/27 頁數及項數
Public Tmpfrm090130 As Form 'Added by Lydia 2015/11/12 新增查名單對應
Dim m_Nation As String 'Add by Amy 2016/08/16 申請國家
Dim m_UpdPA163 As String 'Added by Lydia 2016/09/16 內部收文(假收文)更新初審階段提分割
Public mPrevForm As Form 'Added by Lydia 2018/02/01 母表單(FCP客戶提供文件處理要進入內部收文）
'Add By Sindy 2018/2/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String
Dim m_Done As Boolean
'2018/2/22 END
'Added by Lydia 2018/08/30 新增-外專後續案收文A類
Public mRole As String '從外專後續案收文進來
Dim mPreCaseNo As String  '前一筆收文的本所案號
'Added by Morgan 2020/4/9 批次收文
Dim m_bBatch As Boolean
Dim m_CP11 As String
'end 2020/4/9
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_strLOSkind As String '案件性質=>案源案件類型: A, B1P, B1T, B2P, B2T, CP,CT
Dim m_bolStopOntxtLOS15 As Boolean '是否停駐於案源單號
'end 2020/05/20
Dim m_bolStopOntxtCaseNa239 As Boolean 'Added by Lydia 2020/12/04 CFT脫歐案是否停駐於歐盟案案號
Public m_GetB202CP09 As String 'Added by Lydia 2021/02/22 FCP客戶提供文件處理之收文號
Dim bolChkChange As Boolean 'Added by Lydia 2022/06/30
Dim m_CP10 As String 'Added y Lydia 2021/04/29
Dim m_strfirCP01 As String, m_strfirCP02 As String, m_strfirCP03 As String, m_strfirCP04 As String 'Add By Sindy 2022/7/7
Dim m_strCP14 As String 'Added by Lydia 2023/01/10 外商臺灣案收文:預設承辦人
Dim pa() As String 'Add By Sindy 2023/3/8
Dim bolChild013 As Boolean 'Added by Lydia 2024/02/20 增加FMP案之子案(新案)
Dim m_strCP27 As String 'Add By Sindy 2025/7/29


'Add By Sindy 2020/5/27
Public Sub SetParent(ByRef fm As Form)
   Set mPrevForm = fm
End Sub

'Modify By Sindy 2022/7/6
Private Sub cmdExit_Click()
   Unload Me
End Sub

'Add By Sindy 2022/7/7
Private Sub cmdMRecv_Click()
   If m_strIR01 = "" Or m_strfirCP01 = "" Then Exit Sub
   
   If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> m_strfirCP01 & m_strfirCP02 & m_strfirCP03 & m_strfirCP04 Then
      MsgBox "多案收文第一筆案號必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
      Exit Sub
   End If
   
   If List1.ListCount <= 1 Then
      MsgBox "請輸入多筆案號才能進行多案收文！"
      Exit Sub
   'Add By Sindy 2023/5/15
   Else
      List1.Selected(0) = True
      '2023/5/15 END
   End If
   
   'Add By Sindy 2023/5/15
   If Trim(Me.txtCaseProperty.Text) = "" And intModifyKind = 0 Then
       MsgBox "案件性質不可空白！"
       txtCaseProperty.SetFocus
       Exit Sub
   End If
   '2023/5/15 END
   
   If MsgBox("確定要進行多案收文嗎？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbNo Then Exit Sub
   
   Call OnSaveMRecv
End Sub

'Add By Sindy 2022/7/7 多案收文
'0.人員按下多案收文按鍵
'1.回訊息,繼續收文
Private Sub OnSaveMRecv(Optional intRunSta As Integer = 0)
Dim strQ As String, intQ As Integer
Dim RsQ As ADODB.Recordset
Dim ii As Integer
Dim arrData As Variant
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   
   If intRunSta = 1 Then GoTo RunRecv '前頭回訊息後,繼續收文
   
   '寫入多案收文記錄檔
   '先刪未收文
   strSql = "delete from multiCaseRecv" & _
            " where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
            " and mcr11 is null"
   cnnConnection.Execute strSql
   For ii = 0 To List1.ListCount - 1
      arrData = Split(List1.List(ii), " ")
      strCP01 = SystemNumber(CStr(arrData(0)), 1)
      strCP02 = SystemNumber(CStr(arrData(0)), 2)
      strCP03 = SystemNumber(CStr(arrData(0)), 3)
      strCP04 = SystemNumber(CStr(arrData(0)), 4)
      '檢查案號是否已存在
      strQ = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
             " and mcr02='" & strCP01 & "' and mcr03='" & strCP02 & "' and mcr04='" & strCP03 & "' and mcr05='" & strCP04 & "'"
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, strQ)
      If intQ = 0 Then
         '再寫入
         strSql = "INSERT INTO multiCaseRecv " & _
                  "(mcr01,mcr02,mcr03,mcr04,mcr05,mcr06,mcr07,mcr08,mcr09,mcr10) " & _
                  "VALUES ('" & m_strIR01 & m_strIR03 & "','" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & _
                  "','" & txtCaseProperty & "','" & m_strfirCP01 & "','" & m_strfirCP02 & "'" & _
                  ",'" & m_strfirCP03 & "','" & m_strfirCP04 & "')"
         cnnConnection.Execute strSql
      End If
   Next ii
   
RunRecv:
   
   '呼叫收文程式
   strQ = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
          " and mcr11 is null" & _
          " order by decode(mcr02||mcr03||mcr04||mcr05,mcr07||mcr08||mcr09||mcr10,1,2) asc,mcr02,mcr03,mcr04,mcr05 asc"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      RsQ.MoveFirst
      txtSystem = "" & RsQ.Fields("mcr02")
      txtCode(0) = "" & RsQ.Fields("mcr03")
      txtCode(1) = "" & RsQ.Fields("mcr04")
      txtCode(2) = "" & RsQ.Fields("mcr05")
      txtCaseProperty = "" & RsQ.Fields("mcr06")
      cmdok(0).Value = True
   End If
   
   Set RsQ = Nothing
End Sub

'Modify By Sindy 2022/7/6 將檢查的程式碼抽出來,變成一個函數
Private Function CheckDataIsOk() As Boolean
Dim bolErr As Boolean
''Add By Sindy 2009/07/06
Dim strYear As String '抓下次繳費年度
Dim m_Nexttimes As String '抓下次繳費次數
Dim intQ As Integer 'Add By Sindy 2025/3/31

   CheckDataIsOk = False
   
   'Added by Morgan 2020/2/13
   '若有輸入前6碼的本所號時後面沒輸的要補0否則有些檢查可能會抓不到資料
   If txtCode(0) <> "" Then
      If txtCode(1) = "" Then txtCode(1) = "0"
      If txtCode(2) = "" Then txtCode(2) = "00"
   End If
   'end 2020/2/13
   
   'Added by Lydia 2021/09/03 從外商臺灣案收文進入，沒有帶入收文類別A; ex.FCT-046258的9/3註冊費收文號存成B0031663, 人工修改為AB0036995
   'Modified by Lydia 2021/10/14 發現有外商、外專自行收文的收文號少掉年份的兩碼; ex.FCT-047364的10/13收文號存成A042506,人工修改為AB0042506(共19筆)
   'If lblReciveCode.Caption = "" Then
   If Len(Trim(lblReciveCode.Caption & txtRecieveCode(0))) <> 3 Then
      'Modified by Lydia 2021/10/15 + 客戶提供文件處理 m_GetB202CP09
      'If Left(mRole, 1) = "F" And InStr(Me.Caption, "新增") > 0 Then
      If (Left(mRole, 1) = "F" Or Me.m_GetB202CP09 = "B") And InStr(Me.Caption, "新增") > 0 Then
         strReceiveKind = 接洽記錄單
         lblReciveCode.Caption = strReceiveKind
         'Added by Lydia 2021/10/14
         txtRecieveCode(0).Text = CompAutoNumberYear(GetTaiwanThisYear)
         If Len(Trim(lblReciveCode.Caption & txtRecieveCode(0))) <> 3 Then
            MsgBox "收文號有問題，請關閉收文畫面後，再重新進入！"
            Exit Function
         End If
         'end 2021/10/14
      Else
         MsgBox "收文號有問題，請關閉收文畫面後，再重新進入！"
         Exit Function
      End If
   End If
   'end 2021/09/03

   Call ChkAndCloseForm 'Add by Amy 2021/12/21 Unload 沒 set Nothing 會殘留前次變數值,故由此清->改成共用,怕有沒清到的

    'add by nick 2004/08/23 判斷案件性質不能空白
     'edit by nick 2004/08/26
     'If Trim(Me.txtCaseProperty.Text) = "" Then
     
   ' 91.03.25 modify by louis
   '91.11.10 MODIFY BY SONIA
   'txtCaseProperty_Validate False
   bolErr = False
   txtCaseProperty_Validate bolErr
   If bolErr = True Then
      Exit Function
   End If
   
   'Add By Sindy 2018/2/22
   If m_strIR01 <> "" Then
      If FraRecv.Visible = True Then
         If m_strfirCP01 <> "" And m_strfirCP02 <> "" Then
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> m_strfirCP01 & m_strfirCP02 & m_strfirCP03 & m_strfirCP04 Then
               MsgBox "多案收文第一筆案號必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
               Exit Function
            End If
         End If
      Else
         If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> txtSystem & txtCode(0) & txtCode(1) & txtCode(2) Then
            MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
            Exit Function
         End If
      End If
   End If
   '2018/2/22 END
   
   If Trim(Me.txtCaseProperty.Text) = "" And intModifyKind = 0 Then
       MsgBox "案件性質不可空白！"
       txtCaseProperty.SetFocus
       Exit Function
   End If
   
   'add by sonia 2021/12/29
   If InStr(frm010001.Caption, "內部收文") > 0 Then
      If Left(Pub_strUserST05, 1) = "3" Then
         'Added by Lydia 2022/02/24 FCP程序可以進行內部收文FMP非寰華案之在途期間442。
         If intModifyKind = 0 And txtSystem = "P" And PUB_ChkIsFMP(Trim(txtSystem), Trim(txtCode(0)), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
         Else
         'end 2022/02/24
            If PUB_FMPtoCheck(0, 1, Pub_strUserST05, txtSystem, txtCode(0), txtCode(1), txtCode(2)) = False Then
               txtCode(0).SetFocus
               Exit Function
            End If
         End If 'Added by Lydia 2022/02/24
      'Moeified by Lydia 2023/03/27
      'ElseIf ChkSysName(txtSystem) = False Then
      Else
         If txtSystem = "" Then
         Else
           If ChkSysName(txtSystem) = False Then
      'end 2023/03/27
              txtSystem.SetFocus
              Exit Function
         'Added by Lydia 2023/03/27
           End If
         End If
         'end 2023/03/27
      End If
   End If
   'end 2021/12/29
   
   'Added by Lydia 2022/02/24 外專後續案收文:先限制不可收法務案源
   If Left(mRole, 2) = "F2" And txtSystem = "FCP" And FraLOS.Visible = False Then
      strExc(2) = PUB_GetLOSkind(txtSystem, txtCaseProperty, "000")
      If Left(strExc(2), 1) = "B" Then
         MsgBox "法務案源請交給櫃台收文！", vbCritical, "檢核案源單號"
         Exit Function
      End If
   End If
   'end 2022/02/24
   
   'Added by Lydia 2020/07/02 非新增,帶出本所案號
   If intModifyKind <> 0 And txtSystem = "" And txtRecieveCode(1) <> "" Then
      'Modified by Lydia 2021/04/29 +CP10
      'Modify by Sindy 2023/2/22 +CP140
      strSql = "select cp01,cp02,cp03,cp04,CP10,cp158,cp159,CP140 from caseprogress where cp09='" & Trim(lblReciveCode.Caption) & Trim(txtRecieveCode(0)) & Trim(txtRecieveCode(1)) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If intModifyKind = 1 Then '修改
            'Add by Sindy 2023/2/22
            If Len("" & RsTemp.Fields("CP140")) > 0 Then
               MsgBox "此收文號為電子收文，不可在此作業做修改，" & vbCrLf & vbCrLf & "有需要請通知專業部！", vbCritical, "檢核資料"
               Exit Function
            End If
            '2023/2/22 END
            If Val("" & RsTemp.Fields("cp158")) > 0 Then
               MsgBox "收文號已發文！", vbCritical, "檢核資料"
               Exit Function
            End If
         Else
            If Val("" & RsTemp.Fields("cp159")) > 0 Then
               MsgBox "收文號已取消收文！", vbCritical, "檢核資料"
               Exit Function
            End If
         End If
         txtSystem = "" & RsTemp.Fields("cp01")
         txtCode(0) = "" & RsTemp.Fields("cp02")
         txtCode(1) = "" & RsTemp.Fields("cp03")
         txtCode(2) = "" & RsTemp.Fields("cp04")
         m_CP10 = "" & RsTemp.Fields("CP10") 'Added by Lydia 2021/04/29
      Else
         MsgBox "查無此收文號！", vbCritical, "檢核資料"
         Exit Function
      End If
   End If
   'end 2020/07/02
   
   'Added by Lydia 2018/02/01 檢查是否有D類客戶提供文件
   'Modified by Lydia 2021/02/22 改判斷
   'If intModifyKind = 0 And txtSystem = "FCP" And Trim(Me.txtCaseProperty.Text) = "202" And Trim(txtCode(0).Text) <> "" And lblReciveCode.Caption = "B" And TypeName(mPrevForm) <> "frm060121_1" Then
   If intModifyKind = 0 And txtSystem = "FCP" And Trim(Me.txtCaseProperty.Text) = "202" And Trim(txtCode(0).Text) <> "" And lblReciveCode.Caption = "B" And m_GetB202CP09 = "" Then
      strSql = "select count(*) from custsupportdoc where nvl(csd11,0) = 0 and csd01='" & txtSystem & "' and csd02='" & txtCode(0) & "' and csd03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and csd04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Val("" & RsTemp(0)) > 0 Then
            If MsgBox("本案尚有客戶提供文件未做補文件，是否繼續做內部收文？", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
         End If
      End If
   End If
   'end 2018/02/01
   
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   If FMP2open = True Then
      If intModifyKind = 0 Then '新增-用本所案號
         'Added by Lydia 2020/07/16 FCP程序可以進行內部收文FMP非寰華案之在途期間442。
         If txtSystem = "P" And txtCaseProperty = "442" And PUB_ChkIsFMP(Trim(txtSystem), Trim(txtCode(0)), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
         Else
         'end 2020/07/16
            If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Trim(txtSystem), Trim(txtCode(0)), _
                 IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = False Then
               txtCode(0).SetFocus
               Exit Function
            End If
         End If 'Added by Lydia 2020/07/16
      Else  '修改-用內部收文號
         If txtSystem = "P" Then   'Added by Lydia 2020/07/02 因為原本控制外專人員在Patpro有限制，反而在Patpro1反而不受限；改成不限制exe，只針對外專人員+P案。
            'Added by Lydia 2020/07/16 FCP程序可以進行內部收文FMP非寰華案之在途期間442。
            strSql = "select cp01,cp02,cp03,cp04 from caseprogress where cp31='Y' and substr(cp12,1,1) = 'F' and (cp01,cp02,cp03,cp04) " & _
                        "in (select cp01,cp02,cp03,cp04 from caseprogress where cp09='" & Trim(lblReciveCode.Caption) & Trim(txtRecieveCode(0)) & Trim(txtRecieveCode(1)) & "' and cp10='442' ) "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
            Else
            'end 2020/07/16
               strExc(0) = " select f0.CP01, f0.CP02, f0.CP03, f0.CP04, f0.CP05, f0.CP09, f0.CP10, f0.CP12 " & _
               " from caseprogress f0 where  f0.CP09='" & Trim(lblReciveCode.Caption) & Trim(txtRecieveCode(0)) & Trim(txtRecieveCode(1)) & "' " & FMP2openSQL
               If PUB_FMPtoCheck(0, 1, Pub_strUserST05, "CHANGE_SQL", strExc(0)) = False Then
                  txtRecieveCode(1).SetFocus
                  Exit Function
               End If
            End If 'Added by Lydia 2020/07/16
         End If 'Added by Lydia 2020/07/02
      End If
   End If
   
   'Add By Sindy 2012/4/18
   If intModifyKind = 0 And (txtSystem = "S" Or txtSystem = "TS") And Trim(txtCode(0)) <> "" Then
      strSql = "select * from servicepractice " & _
               "where sp01='" & txtSystem & "' and sp02='" & txtCode(0) & "' and sp03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and sp04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' " & _
               "and instr(sp18,'轉入商標')>0 "
      CheckOC3
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If Not AdoRecordSet3.EOF And Not AdoRecordSet3.BOF Then
         MsgBox "此案號已轉入商標案，不可再收文！"
         txtCode(0).SetFocus
         Exit Function
      End If
   End If
   '2012/4/18 End
   
   m_strCP14 = strUserNum 'Added by Lydia 2023/01/10
   '911111 nick 檢查前兩欄不可空白，新增時，內部收文
   '***** start
   If intModifyKind = 0 Then '新增
      'Added by Lydia 2018/08/30 外專後續案收文
      'Modified by Lydia 2020/12/16 區分外專F2x
      'If Left(mRole, 1) = "F" Then
      If Left(mRole, 2) = "F2" Then
         If txtCode(0) = "" Then
            MsgBox "本所案號不可空白！"
            txtCode(0).SetFocus
            Exit Function
         'Modified by Lydia 2022/06/21 +P案 (FMP案)
         ElseIf txtSystem <> "FCP" And txtSystem <> "FG" And txtSystem <> "P" Then
            MsgBox "不可收" & txtSystem & "案！"
            txtSystem.SetFocus
            Exit Function
         End If
         'Added by Lydia 2022/06/21 外專後續案收文，請開放P的寰華案也可以操作
         If txtSystem = "P" Then
            bolErr = False
            'Modified by Lydia 2022/07/27 開放P(非寰華案)=FMP案的收文權限
            'bolErr = PUB_FMPtoCheck(1, 2, Pub_strUserST05, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
            bolErr = PUB_ChkIsFMP(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
            If bolErr = False Then
               'Modified by Lydia 2022/07/27 開放P(非寰華案)=FMP案的收文權限
               'MsgBox "只可收文寰華案！"
               MsgBox "只可收文寰華案／FMP案！"
               txtCode(0).SetFocus
               Exit Function
            End If
         End If
         'end 2022/06/21
         '檢查是否為管制人
         If Pub_StrUserSt03 <> "M51" Then
            strExc(1) = PUB_GetFCPSalesNo(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
            If strExc(1) <> "" And strExc(1) <> strUserNum Then
               If MsgBox(txtSystem & "-" & txtCode(0) & "的管制人為" & GetStaffName(strExc(1), True) & "，是否繼續收文？", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                  txtCode(0).SetFocus
                  Exit Function
               End If
            End If
         End If
      
      'Added by Lydia 2020/12/16 外商臺灣案收文
      'modify by sonia 2022/1/25 外商程序不能內部收文台灣T案
      'ElseIf Left(mRole, 2) = "F1" Then
      '     If txtSystem <> "FCT" And txtSystem <> "S" Then
      ElseIf Left(mRole, 2) = "F1" Or Left(Pub_StrUserSt03, 2) = "F1" Then
         If txtSystem <> "FCT" And txtSystem <> "CFT" And txtSystem <> "S" And m_Nation <> "020" Then
      'end 2022/1/25
            MsgBox "不可收" & txtSystem & "案！"
            txtSystem.SetFocus
            Exit Function
         End If
         'modify by sonia 2022/1/28外商程序內收文不彈此訊息
         'If txtCode(0) <> ""  Then
         If txtCode(0) <> "" And Pub_StrUserSt03 <> "F12" Then
            If ClsPDCheckCaseCodeIsExist(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
               strExc(1) = PUB_GetAKindSalesNo(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)))
               If strExc(1) <> "" And strExc(1) <> strUserNum Then
                  If MsgBox(txtSystem & "-" & txtCode(0) & "的智權人員為" & GetStaffName(strExc(1), True) & "，是否繼續收文？", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                     txtCode(0).SetFocus
                     Exit Function
                  End If
                  'Added by Lydia 2023/01/10 系統增加提醒「此次收文是否為代收文」：選擇「是」系統自動設定為原智權人員，選擇「否」系統自動設定輸入者為智權人員。
                  If MsgBox("此次收文是否為代收文", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                      m_strCP14 = strExc(1)
                  End If
                  'end 2023/01/10
               End If
            Else
               txtCode(0).SetFocus
               Exit Function
            End If
         End If
      'end 2020/12/16
      'Added by Lydia 2020/05/20 判斷是否為法律所案源收文
      ElseIf FraLOS.Visible = True Then
         If Trim(txtLOS15) = "" Then
            'Modified by Lydia 2020/07/20 預設"否"要輸入案源單號 vbDefaultButton2 (原本預設vbDefaultButton1)
            If MsgBox("請確認接洽單左上角是否沒有案源單號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""要輸入案源單號", vbInformation + vbYesNo + vbDefaultButton2, "檢核案源單號") = vbNo Then
               txtLOS15.SetFocus
               txtLOS15_GotFocus
               Exit Function
            End If
         Else
            If GetStateLOS(txtSystem, txtCaseProperty, txtCode(0), txtLOS15, m_strLOSkind) = False Then
               txtLOS15.SetFocus
               txtLOS15_GotFocus
               Exit Function
            End If
         End If
      'end 2020/05/20
      'Added by Lydia 2020/11/19 CFP和CFT英國脫歐案管制：歐盟案案號
      ElseIf fraNA239.Visible = True Then
         If Trim(txtCaseNa239) = "" Then
            If MsgBox("請確認接洽單左上角是否沒有歐盟案案號？" & vbCrLf & "選擇""是""會繼續作業，選擇""否""要輸入歐盟案案號", vbInformation + vbYesNo + vbDefaultButton2, "檢核歐盟案案號") = vbNo Then
               txtCaseNa239.SetFocus
               txtCaseNa239_GotFocus
               Exit Function
            End If
         Else
            bolErr = False
            txtCaseNa239_Validate bolErr
            If bolErr = True Then
               Exit Function
            End If
         End If
      'end 2020/11/19
      End If
      'end 2018/08/30
      
      'Add by Morgan 2010/4/2 集體設計必須收母案之1,2...
      If txtSystem = "CFP" And txtCaseProperty = "105" Then
         '2011/9/9 modify by sonia CFP-024311-A 無法輸入
         'If txtCode(0) = "" Or Val(txtCode(1)) = 0 Then
         If txtCode(0) = "" Or (txtCode(1) = "" Or txtCode(1) = "0") Then
            MsgBox "集體設計必須收母案之1,2...！"
            Exit Function
         End If
         If ChkPCode(txtSystem, txtCode(0), txtCode(1), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
            MsgBox "此集體設計子案已存在，請收其他子案案號！"
            txtCode(1).SetFocus
            Exit Function
         End If
      End If
      'end 2010/4/2
      
      '2012/6/18 ADD BY SONIA 聯合申請案必須收母案之1,2...
      If txtSystem = "P" And txtCaseProperty = "105" Then
         If txtCode(0) = "" Or (txtCode(1) = "" Or txtCode(1) = "0") Then
            MsgBox "聯合申請案必須收母案之1,2...！"
            Exit Function
         End If
         If ChkPCode(txtSystem, txtCode(0), txtCode(1), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
            MsgBox "此聯合申請案子案已存在，請收其他子案案號！"
            txtCode(1).SetFocus
            Exit Function
         End If
      End If
      '2012/6/18 END
      
      'Added by Morgan 2012/10/8 衍生設計案必須收母案之1,2...
      If txtSystem = "P" And txtCaseProperty = "125" Then
         If txtCode(0) = "" Or (txtCode(1) = "" Or txtCode(1) = "0") Then
            MsgBox "衍生設計案必須收母案之1,2...！"
            Exit Function
         End If
         If ChkPCode(txtSystem, txtCode(0), txtCode(1), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
            MsgBox "此衍生設計子案已存在，請收其他子案案號！"
            txtCode(1).SetFocus
            Exit Function
         End If
      End If
      'end 2012/10/8
      
      'Added by Morgan 2020/2/13
      '收文442在途期間提醒
      If txtSystem = "P" And Me.txtCode(0).Text <> "" And txtCaseProperty = "442" Then
         '接洽單
         If intChoose = 0 Then
            '非FMP案時提醒
            'Modified by Morgan 2021/2/2
            'If Check_IsFMP(txtSystem, txtCode(0), txtCode(1), IIf(txtCode(2) = "", "00", txtCode(2))) = False Then
            If PUB_ChkIsFMP(txtSystem, txtCode(0), txtCode(1), IIf(txtCode(2) = "", "00", txtCode(2))) = False Then
               If MsgBox("請確認是否有區主管核可？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Function
               End If
            End If
         '內部收文
         Else
            '外專程序操作時提醒(109/2/12郭說內專不用)--淑華
            If Left(Pub_StrUserSt03, 1) = "F" Then
               If MsgBox("請確認工程師是否已呈報協理核准？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Function
               End If
            End If
         End If
      End If
      'end 2020/2/13
      
      'Added by Lydia 2021/11/10 寰華和FMP大陸案衍生的香港、澳門案不走命名之流程，及其相關管控。
      'Modifiedby Lydia 2024/02/20 有關" 香港013專利開放收文集體設計申請105"，請比照" 香港013專利收文設計申請103"
      'If txtSystem = "P" And Trim(txtCode(0)) = "" And fraFMP.Visible = True Then
      If txtSystem = "P" And (Trim(txtCode(0)) = "" Or bolChild013 = True) And fraFMP.Visible = True Then
         If txtFMP <> "" And txtNA01 = "" Then
            MsgBox "請輸入申請國家！", vbCritical
            txtNA01.SetFocus
            txtNA01_GotFocus
            Exit Function
         ElseIf txtFMP = "" And txtNA01 <> "" Then
            If MsgBox("是否為FMP案？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                txtFMP = "Y"
            Else
                txtNA01 = "": lblNation.Caption = ""
            End If
         ElseIf txtFMP <> "" And txtNA01 <> "" Then
            '檢查
            bolErr = False
            txtNA01_Validate bolErr
            If bolErr = True Then
               Exit Function
            End If
         End If
         'Added by Lydia 2024/02/20 有關" 香港013專利開放收文集體設計申請105"，請比照" 香港013專利收文設計申請103"
         If bolChild013 = True Then
            If txtFMP = "" Then
               If MsgBox("是否為FMP案？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                   txtFMP = "Y"
                   txtNA01.SetFocus
                   txtNA01_GotFocus
                   Exit Function
               Else
                   txtNA01 = "": lblNation.Caption = ""
               End If
            End If
            If txtCaseProperty <> "105" Then
               MsgBox "請輸入案件性質105", vbCritical
               txtCaseProperty.SetFocus
               txtCaseProperty_GotFocus
               Exit Function
            End If
         End If
         'end 2024/02/20
      End If
      'end 2021/11/10
      
      '內部收文
      If intChoose = 1 Then
         If txtSystem = "TF" Then
            If txtTFCode(0) = "" Then
               MsgBox "本所案號不可空白！"
               txtTFCode(0).SetFocus
               Exit Function
            End If
         Else
            If txtCode(0) = "" Then
               MsgBox "本所案號不可空白！"
               txtCode(0).SetFocus
               Exit Function
            End If
             'add by nick 2004/10/15
             'CFP  不可以收專利調查(903) 及調卷(904)  'cancel by sonia 2020/1/8
             'P                       調卷(904)       'cancel by sonia 2020/1/8
             'PS                     專利調查(903) '2015/10/5 CANCEL BY SONIA 郭副理要取消
             'P 的聯合申請(105) 不可收新流水號的新案，要收子號的新案
             '2008/1/14 modify by sonia CFP可收專利調查
             'If txtSystem = "CFP" And (txtCaseProperty = "903" Or txtCaseProperty = "904") Then
         'cancel by sonia 2020/1/8
         '                 If txtSystem = "CFP" And txtCaseProperty = "904" Then
         '                      MsgBox "CFP 不可以收" & IIf(txtCaseProperty = "903", "專利調查", "調卷") & "！"
         '                      txtCaseProperty.SetFocus
         '                      Exit Function
         '                 End If
         '                 If txtSystem = "P" And txtCaseProperty = "904" Then
         '                      MsgBox "P 不可以收調卷！"
         '                      txtCaseProperty.SetFocus
         '                      Exit Function
         '                 End If
         'end 2020/1/8
             '2015/10/5 CANCEL BY SONIA 郭副理要取消
         '                 If txtSystem = "PS" And txtCaseProperty = "903" Then
         '                      MsgBox "PS 不可以收專利調查！"
         '                      txtCaseProperty.SetFocus
         '                      Exit Function
         '                 End If
            If txtSystem = "P" And txtCaseProperty = "105" Then
               If txtCode(1) = "" Then
                  MsgBox "本所案號不可空白！"
                  txtCode(1).SetFocus
                  Exit Function
               End If
               If ChkPCode(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
                  MsgBox "此聯合申請子案已存在，請收其他子案案號！"
                  txtCode(1).SetFocus
                  Exit Function
               End If
            End If
            '2005/12/5 ADD BY SONIA
             
            'Added by Morgan 2012/12/27
            If txtSystem = "P" And txtCaseProperty = "125" Then
               If txtCode(1) = "" Then
                  MsgBox "本所案號不可空白！"
                  txtCode(1).SetFocus
                  Exit Function
               End If
               If ChkPCode(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
                  MsgBox "此衍生設計子案已存在，請收其他子案案號！"
                  txtCode(1).SetFocus
                  Exit Function
               End If
            End If
            'end 2012/12/27
            
            If txtSystem = "CFT" And txtCaseProperty = "312" Then
               MsgBox "CFT 不可以收 緩衝期限！"
               txtCaseProperty.SetFocus
               Exit Function
            End If
            '2005/12/5 END
         End If
      '收文
      Else
         'add by nick 2004/10/15
         'CFP  不可以收專利調查(903) 及調卷(904)   'cancel by sonia 2020/1/8
         'P                       調卷(904)        'cancel by sonia 2020/1/8
         'PS                     專利調查(903) '2015/10/5 CANCEL BY SONIA 郭副理要取消
         'P 的聯合申請(105) 不可收新流水號的新案，要收子號的新案
         '2008/1/14 modify by sonia CFP可收專利調查
         'If txtSystem = "CFP" And (txtCaseProperty = "903" Or txtCaseProperty = "904") Then
   'cancel by sonia 2020/1/8
   '            If txtSystem = "CFP" And txtCaseProperty = "904" Then
   '                 MsgBox "CFP 不可以收" & IIf(txtCaseProperty = "903", "專利調查", "調卷") & "！"
   '                 txtCaseProperty.SetFocus
   '                 Exit Function
   '            End If
   '            If txtSystem = "P" And txtCaseProperty = "904" Then
   '                 MsgBox "P 不可以收調卷！"
   '                 txtCaseProperty.SetFocus
   '                 Exit Function
   '            End If
   'end 2020/1/8
         '2015/10/5 CANCEL BY SONIA 郭副理要取消
   '            If txtSystem = "PS" And txtCaseProperty = "903" Then
   '                 MsgBox "PS 不可以收專利調查！"
   '                 txtCaseProperty.SetFocus
   '                 Exit Function
   '            End If
         If txtSystem = "P" And txtCaseProperty = "105" Then
             If txtCode(0) = "" Then
                  MsgBox "本所案號不可空白！"
                  txtCode(0).SetFocus
                  Exit Function
             End If
             If txtCode(1) = "" Then
                  MsgBox "本所案號不可空白！"
                  txtCode(1).SetFocus
                  Exit Function
             End If
             If ChkPCode(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
                 MsgBox "此聯合申請子案已存在，請收其他子案案號！"
                 txtCode(1).SetFocus
                 Exit Function
             End If
         End If
         
         'Added by Morgan 2012/12/29
         If txtSystem = "P" And txtCaseProperty = "125" Then
             If txtCode(0) = "" Then
                  MsgBox "本所案號不可空白！"
                  txtCode(0).SetFocus
                  Exit Function
             End If
             If txtCode(1) = "" Then
                  MsgBox "本所案號不可空白！"
                  txtCode(1).SetFocus
                  Exit Function
             End If
             If ChkPCode(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
                 MsgBox "此衍生設計子案已存在，請收其他子案案號！"
                 txtCode(1).SetFocus
                 Exit Function
             End If
         End If
         'end 2012/12/29
         
         'add by nick 2005/01/27 P 台灣案櫃檯收文時，若該案號已有該案件性質且未發文之b 類收文資料，則不可收文，
         '顯示"此案件性質已有內部收文，請智權人員與專業部確認，此程序是否需要收文！"
         If txtSystem = "P" And txtCode(0) <> "" Then
            strSql = "select count(*) from caseprogress,patent where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp10='" & txtCaseProperty & "' and cp27 is null and cp09>'B' and cp09<'C' and cp57 is null and pa09='000' and cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and cp04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' "
            'Added byLydia 2018/09/10 剔除案件性質408面詢不檢查
            strSql = strSql & " and cp10 not in ('408') "
            CheckOC3
            AdoRecordSet3.CursorLocation = adUseClient
            AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If AdoRecordSet3.Fields(0).Value <> 0 Then
                MsgBox "此案件性質已有內部收文，請智權人員與專業部確認，此程序是否需要收文！"
                txtCode(0).SetFocus
                CheckOC3
                Exit Function
            End If
         End If
         '2005/12/5 ADD BY SONIA
         If txtSystem = "CFT" And txtCaseProperty = "312" Then
            MsgBox "CFT 不可以收 緩衝期限！"
            txtCaseProperty.SetFocus
            Exit Function
         End If
         '2005/12/5 END
         'add by nickc 2007/04/17 加入，若是收文 CFT 的領証時，若以收過註冊証來函，也有費用的，提示：此案已有註冊証帳款，不用再收領証，並將此訊息轉知智權人員
         If txtSystem = "CFT" And txtCaseProperty = "701" And txtCode(0) <> "" Then
             CheckOC3
             strSql = "select * from caseprogress where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and cp04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' and cp10='1701' and cp57 is null and cp16>0 "
             AdoRecordSet3.CursorLocation = adUseClient
             AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If AdoRecordSet3.RecordCount <> 0 Then
                 MsgBox "此案已有註冊証帳款，不用再收領証，並將此訊息轉知智權人員"
                 txtSystem = ""
                 txtCode(0) = ""
                 txtCode(1) = ""
                 txtCode(2) = ""
                 txtCaseProperty = ""
                 txtSystem.SetFocus
                 Exit Function
             End If
         End If
         
         'Add by Sindy 2013/8/7 詢問是否要重覆收文
         strSql = "select cp01 from caseprogress where cp05=" & DBDATE(IIf(textCP05 = "", strSrvDate(1), textCP05)) & " and cp10='" & txtCaseProperty & "'" & _
                  " and cp01='" & Trim(txtSystem) & "'" & _
                  " and cp02='" & Trim(txtCode(0)) & "'" & _
                  " and cp03='" & IIf(Trim(txtCode(1)) = "", "0", Trim(txtCode(1))) & "'" & _
                  " and cp04='" & IIf(Trim(txtCode(2)) = "", "00", Trim(txtCode(2))) & "'"
         CheckOC3
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If AdoRecordSet3.RecordCount > 0 Then
            If MsgBox("此案號今日已收相同案件性質，請確認是否要收文？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               txtCaseProperty.SetFocus
               Exit Function
            End If
         End If
         '2013/8/7 END
            
'Removed by Morgan 2012/7/4 不用
'            'Added by Morgan 2012/7/3
'            If (txtSystem = "P" Or txtSystem = "FCP") And txtCaseProperty = "124" And txtCode(0) <> "" Then
'               strExc(0) = "select 1 from patent where pa01='" & txtSystem & "' and pa02='" & txtCode(0) & "' and pa03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and pa04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' and pa16 is not null"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  MsgBox "本案已審定不可再收文" & lblCasePropertyName & "！"
'                  Exit Function
'               End If
'            End If
'            'end 2012/7/3
      End If
      
      'Add By Sindy 2009/09/18
      If txtSystem = "S" And txtCode(0) <> "" Then
         strSql = "select * from servicepractice where sp01='" & txtSystem & "' and sp02='" & txtCode(0) & "' and sp03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and sp04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' " & _
                         "and instr(sp18,'轉入商標') >0 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            MsgBox "此案已" & Trim(RsTemp.Fields("sp18")) & "，不可收文！"
            txtCode(0).SetFocus
            CheckOC3
            Exit Function
         End If
      End If
      '2009/09/18 End
      
      'Add By Sindy 2023/1/18
      If txtSystem = "FCT" And txtCaseProperty = "214" Then
         MsgBox "陳述聲明由專業部管控處理，不可收文214陳述聲明！"
         txtCaseProperty.SetFocus
         CheckOC3
         Exit Function
      End If
      '2023/1/18 End
      
      'Add By Sindy 2023/3/8 當FCT、T、CFT收文之案件性質為「延展」，
      '                      請在收文時控管，若有未發文之「延展」，不可重複收文。
      If (txtSystem = "CFT" Or txtSystem = "FCT" Or txtSystem = "T") And txtCaseProperty = "102" Then
          '不可重複收文(有已收未發的延展)
          ReDim Preserve pa(1 To TF_PA) As String
          pa(1) = txtSystem
          pa(2) = txtCode(0)
          pa(3) = txtCode(1)
          pa(4) = txtCode(2)
          If PUB_ChkCPExist(pa, "102", 1) Then
              If GetPrjNation1(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)) = "000" Then
                 Call ClsPDGetCasePropertyL(1, txtSystem, "102", strExc(10))
              Else
                 Call ClsPDGetCasePropertyL(2, txtSystem, "102", strExc(10))
              End If
              MsgBox "本案目前有<" & strExc(10) & ">尚未發文，不可再收文<" & strExc(10) & ">!!!", vbExclamation + vbOKOnly
              txtCaseProperty.SetFocus
              CheckOC3
              Exit Function
          End If
      End If
      '2023/3/8 END
      
      '2007/6/28 ADD BY SONIA P,FCP已收過928重新委任(未取消收文)都不可再收文,不管接洽單或內部收文
      If (txtSystem = "P" Or txtSystem = "FCP") And txtCaseProperty = "928" Then
         strSql = "select * from caseprogress where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and cp04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' and cp10='928' and cp57 is null "
         CheckOC3
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If Not AdoRecordSet3.EOF And Not AdoRecordSet3.BOF Then
             MsgBox "此案已收過928重新委任, 不可再收此案件性質, 請退回給智權人員！"
             txtCode(0).SetFocus
             Exit Function
         End If
         '2007/7/4 add by sonia 檢查是否同意重新委任
         If txtCode(0) <> "" Then
            'Modify By Sindy 2022/11/15
            'If ChkAgree928 = False Then
            If PUB_ChkAgree928(txtSystem, txtCode(0), txtCode(1), txtCode(2)) = False Then
            '2022/11/15 END
               txtCaseProperty.SetFocus
               Exit Function
            End If
         End If
         '2007/7/4 end
      End If
      '2007/6/28 end
      
   ElseIf intModifyKind = 1 Then '修改
      'Add By Sindy 2011/1/11 修改時若該筆資料已收款cp75>0時不可修改
      'Modify by Morgan 2011/7/12 已發文已不可再修改,否則FCP的收文費用會被覆蓋 Ex.FCP-019354
      'Modified by Morgan 2018/3/1
      'strSql = "select cp75,cp27 from caseprogress where cp09='A" & Trim(txtRecieveCode(0)) & Trim(txtRecieveCode(1)) & "' "
      strSql = "select cp75,cp27 from caseprogress where cp09='" & Trim(lblReciveCode) & Trim(txtRecieveCode(0)) & Trim(txtRecieveCode(1)) & "' "
      CheckOC3
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount > 0 Then
         If Val("" & AdoRecordSet3.Fields("cp75").Value) > 0 Then
             MsgBox "此筆收文資料已收款，不可修改！"
             txtRecieveCode(1).SetFocus
             CheckOC3
             Exit Function
         End If
         'Add by Morgan 2011/7/12
         If Val("" & AdoRecordSet3.Fields("cp27").Value) > 0 Then
             MsgBox "此筆收文資料已發文，不可修改！"
             txtRecieveCode(1).SetFocus
             CheckOC3
             Exit Function
         End If
      End If
   End If
   '***** end
   
   'add by nick 2004/11/04
   If intReceiveKind = 0 Then
       If (txtSystem = "T" Or txtSystem = "CFT" Or txtSystem = "FCT" Or txtSystem = "TF") And txtCaseProperty = "308" And Trim(txtCode(0)) <> "" Then
           'edit by nickc 2006/07/24 台灣案的跳過
           'Modify by Amy 2014/10/22 大陸案跳過
           'strSql = "select * from trademark where tm10<>'000' and tm01='" & txtSystem & "' and tm02='" & txtCode(0) & "' and tm03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and tm04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' "
           strSql = "select * from trademark where tm10 not in ('000','020')  and tm01='" & txtSystem & "' and tm02='" & txtCode(0) & "' and tm03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and tm04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' "
           CheckOC3
           AdoRecordSet3.CursorLocation = adUseClient
           AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
           If Not AdoRecordSet3.EOF And Not AdoRecordSet3.BOF Then
               MsgBox "商標分割非台灣案及非大陸案只能收新本所案號，本所案號只能輸入系統別，其餘應該空白！"
               txtCode(0).SetFocus
               Exit Function
           End If
           'end 2014/10/22
       End If
   End If
   
   ' 90.12.19 modify by louis
   If txtPetition.Visible = True Then
      If IsEmptyText(txtPetition) Then
         'Modified by Lydia 2022/06/30
         'MsgBox "請輸入移轉、讓與申請人", vbOKOnly + vbCritical, "檢核資料"
         MsgBox "請輸入" & Mid(Me.Label5.Caption, 1, Len(Me.Label5.Caption) - 1), vbOKOnly + vbCritical, "檢核資料"
         txtPetition.SetFocus
         Exit Function
      '911111 nick 檢查申請人是否有輸入
      '***** start
      Else
          bolErr = False
          txtPetition_Validate bolErr
          If bolErr = True Then
             Exit Function
          End If
      '***** end
      End If
   End If
   
   'Add By Sindy 2009/07/06
   If txtCode(0) <> "" Then 'Modify by Sindy 2010/8/4 舊案才檢查
      'Added by Morgan 2012/12/28
      If txtSystem = "P" And txtCaseProperty = "601" Then
         If txtCode(1) = "" Then txtCode(1).Text = "0"
         If txtCode(2) = "" Then txtCode(2).Text = "00"
         strExc(0) = "select pa11 from patent where pa01='" & txtSystem & "' and pa02='" & txtCode(0) & "' and pa03='" & txtCode(1) & "' and pa04='" & txtCode(2) & "' and pa09='000' and pa11 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If PUB_ChkPriDate(RsTemp("pa11"), strExc(3), False) = True Then
               MsgBox "本案已被 " & strExc(3) & " 主張國內優先權，不可領證!!", vbExclamation
               Exit Function
            End If
         End If
      End If
      'end 2012/12/28
      
      If textYear.Visible = True And Text1(0).Visible = True Then
         ' 檢查繳費年度/次數是否不正確
         If IsEmptyText(textYear) = False And txtCaseProperty <> "601" Then
            If txtCode(1) = "" Then txtCode(1).Text = "0"
            If txtCode(2) = "" Then txtCode(2).Text = "00"
            'Modified by Morgan 2022/6/9 +txtCaseProperty
            m_Nexttimes = PUB_Getnexttimes(txtSystem, txtCode(0), txtCode(1), txtCode(2), strYear, , txtCaseProperty)
            If m_Nexttimes <> "" Then
               If txtCaseProperty = "601" Or txtCaseProperty = "605" Then '繳費年度
                  '2010/12/29 add by sonia P-097385
                  If m_Nexttimes = "1" And txtCaseProperty = "605" And Me.txtSystem.Text = "P" Then
                     MsgBox "此案件無繳費記錄, 請不要輸入繳費年度起迄資料！"
                     textYear.SetFocus
                     Exit Function
                  '2010/12/29 end
                  ElseIf Val(textYear) <> Val(strYear) Then
                     MsgBox "繳費(起)年度有誤，應為" & strYear & "！"
                     textYear.SetFocus
                     Exit Function
                  End If
               Else '繳費次數
                  If Val(textYear) <> Val(m_Nexttimes) Then
                     MsgBox "繳費(起)次數有誤，應為" & m_Nexttimes & "！"
                     textYear.SetFocus
                     Exit Function
                  End If
               End If
            Else
               If txtCaseProperty = "601" Or txtCaseProperty = "605" Then '繳費年度
                  MsgBox "無下次繳費年度！"
               Else '繳費次數
                  MsgBox "無下次繳費次數！"
               End If
               If textYear.Enabled = False Then
                  Text1(0).SetFocus
               Else
                  textYear.SetFocus
               End If
               Exit Function
            End If
         End If
         If IsEmptyText(textYear) = False Or IsEmptyText(Text1(0)) = False Then
            If textYear = "" Or textYear = "0" Then
               If txtCaseProperty = "601" Or txtCaseProperty = "605" Then '繳費年度
                  MsgBox "無繳費(起)年度，請清空起迄年度！"
               Else '繳費次數
                  MsgBox "無繳費(起)次數，請清空起迄次數！"
               End If
               If textYear = "0" Then
                  textYear.SetFocus
               Else
                  Text1(0).SetFocus
               End If
               Exit Function
            End If
            If txtCaseProperty = "601" And Text1(0) = "" Then
               '不跑else段程式
            Else
               If Text1(0) = "" Then Text1(0) = "0"
               If Val(textYear) > Val(Text1(0)) Then
                  If txtCaseProperty = "601" Or txtCaseProperty = "605" Then '繳費年度
                     MsgBox "繳費(迄)年度不可小於(起)年度！"
                  Else '繳費次數
                     MsgBox "繳費(迄)次數不可小於(起)次數！"
                  End If
                  Text1(0).SetFocus
                  Exit Function
               End If
            End If
         End If
      End If
   End If
   '2009/07/06 End
   
   ' 91.09.16 modify by louis
   'Add By Cheng 2001/12/12
   'If Me.txtSystem.Visible Then
   '   bolErr = False
   '   txtSystem_Validate bolErr
   '   If bolErr Then txtSystem.SetFocus: Exit Sub
   'End If
   If Me.txtSystem.Visible = True Then
      If CheckEverythingOK() = False Then
         Exit Function
      End If
      If GetSysTemKind = False Then
         Exit Function
      End If
   End If
   
   lblTCT.Caption = "" 'Added by Lydia 2017/11/14
   lblTCTNO.Caption = "" 'Added by Lydia 2018/05/10
   
   If intModifyKind <> 0 Then
      If CheckEverythingOK = False Then
         Exit Function
      Else
         If txtCaseProperty.Visible = True And lblCasePropertyName = "" Then
            ShowMsg MsgText(1014)
            txtCaseProperty.SetFocus
            txtCaseProperty_GotFocus
            Exit Function
         Else
            If txtPetition.Visible = True Then
               txtPetition_Validate False
               If lblPetitionName = "" Then
                  ShowMsg MsgText(1015)
                  txtPetition.SetFocus
                  txtPetition_GotFocus
                  Exit Function
               End If
               'Add by Morgan 2006/6/23
               For intI = 2 To 5
                  txtPetitionx_Validate intI, False
                  If lblPetitionNamex(intI) = "" Then
                     ShowMsg MsgText(1015)
                     txtPetitionx(intI).SetFocus
                     txtPetitionx_GotFocus intI
                     Exit Function
                  End If
               Next
            End If
         End If
      End If
   'Add By Sindy 2025/3/31 檢查不得代理
   Else
      If txtSystem.Text <> "" And txtCode(0).Text <> "" Then '舊案
         If txtPetition.Visible = True Then
            If GetCustomerAndState(txtPetition.Text, strExc(10), , , , txtSystem.Text, , IIf(txtCode(0).Text <> "", True, False), Me.Name, txtCode(0).Text, IIf(txtCode(1).Text = "", "0", txtCode(1).Text), IIf(txtCode(2).Text = "", "00", txtCode(2).Text), txtCaseProperty.Text) = False Then
               txtPetition.SetFocus
               txtPetition_GotFocus
               Exit Function
            End If
            For intQ = 2 To 5
               If txtPetitionx(intQ) <> "" Then
                  If GetCustomerAndState(txtPetitionx(intQ).Text, strExc(10), , , , txtSystem.Text, , IIf(txtCode(0).Text <> "", True, False), Me.Name, txtCode(0).Text, IIf(txtCode(1).Text = "", "0", txtCode(1).Text), IIf(txtCode(2).Text = "", "00", txtCode(2).Text), txtCaseProperty.Text) = False Then
                     txtPetitionx(intQ).SetFocus
                     txtPetitionx_GotFocus intQ
                     Exit Function
                  End If
               End If
            Next
         End If
      End If
   '2025/3/31 END
   End If
   
   'Add By Cheng 2003/09/08
   'Begin
   If Me.txtSystem.Text = 馬德里案 Then
       m_blnNewCase = CheckNewCase(Me.txtSystem.Text, Me.txtTFCode(0).Text & Me.txtTFCode(1).Text, Me.txtTFCode(2).Text, Me.txtTFCode(3).Text)
   Else
       m_blnNewCase = CheckNewCase(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text)
   End If
   'End
   
   'Modify By Sindy 2022/7/6 接洽紀錄單收文時的檢查(下列程式段抽出來的)
   'Modify By Sindy 2022/9/30 + And intChoose = 0
   If intReceiveKind = 0 And intModifyKind = 0 And intChoose = 0 Then
      '911017 nick
      'If txtCode(0) <> "" Then
      If txtCode(0) <> "" Or txtTFCode(0) <> "" Then
         '911017 nick 將 1 和 2  對調
         If CheckCaseNo Then   '----1
            Exit Function
         End If
         If fraFMP.Visible = False Then   'Added by Lydia 2024/02/20 提前到輸入本所案號檢查
            If CheckExist(1) Then    '------2
               Exit Function
            End If
         End If   'Added by Lydia 2024/02/20
      Else
         intSaveMode = 1
      End If
      'edit by nickc 2007/02/02 不用 dll 了
      'objPublicData.GetSystemKind txtSystem.Text, intCaseKind, strCaseName
      ClsPDGetSystemKind txtSystem.Text, intCaseKind, strCaseName
      
      If intCaseKind = 專利 Then
         'Add By Sindy 2013/4/11 大陸107.復審案下一程序只有一筆復審期限時,不可收文延期404
         If txtSystem = "P" And Me.txtCode(0).Text <> "" And txtCaseProperty = "404" Then
            adoquery.CursorLocation = adUseClient
            strSql = "Select NP07 FROM nextprogress,Patent WHERE NP02=PA01 and NP03=PA02 and NP04=PA03 and NP05=PA04" & _
                     " and NP02='" & Trim(txtSystem) & "' and NP03='" & Trim(txtCode(0)) & "' and NP04='" & IIf(txtCode(1).Text = "", "0", txtCode(1).Text) & "' and NP05='" & IIf(txtCode(2).Text = "", "00", txtCode(2).Text) & "'" & _
                     " and NP06 is null and PA09='020'" & strNpSqlOfNoSalesDuty
            adoquery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount = 1 Then
               If adoquery.Fields("NP07").Value = "107" Then
                  MsgBox "請通知智權同仁依大陸專利法規定復審案不得延期。", vbCritical
                  adoquery.Close
                  Exit Function
               End If
            End If
            adoquery.Close
         End If
         '2013/4/11 End
         
         '專利修法問題
         If (txtSystem = "P" Or txtSystem = "FCP") And Me.txtCode(0).Text <> "" Then
            'Modify by Morgan 2004/7/5
            '加判斷申請國家=台灣
            If GetPA09(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) = "000" Then
               '2005/9/22 CANCEL BY SONIA 指案件准駁確定後 60 天, 因P-63193電腦無法判斷何謂准駁確定, 故先取消
               ''國內改請案不可逾准駁日60天
               If InStr("301,302,303,304,305,306", txtCaseProperty) > 0 Then
               '   If PUB_Check301(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) = False Then
               '      Exit Sub
               '   End If
               '2005/9/22 END
               '國內延緩公告案需控制領證未發文
               ElseIf txtCaseProperty = "412" Then
                  'Modify By Sindy 2022/11/15 PUB_Check412 改共用 PUB_RecvCheck412
                  If PUB_RecvCheck412(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) = False Then
                     Exit Function
                  End If
               '技術報告發文時無公告日不可發文，無公告日不可發文或公告日<93.7.1者不可收文也不可發文
               'Modify by Morgan 2007/8/29 加807第三人申請技術報告
               'ElseIf txtCaseProperty = "421" Then
               ElseIf txtCaseProperty = "421" Or txtCaseProperty = "807" Then
                  If PUB_Check421(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text, 1, txtCaseProperty) = False Then
                     Exit Function
                  End If
               'Add by Morgan 2004/9/13 新型主動修正期限為申請日起兩個月
               ElseIf txtCaseProperty = "203" Then
                  If PUB_Check203(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) = False Then
                     Exit Function
                  End If
               End If
            End If
         End If
      End If
   End If
   
   CheckDataIsOk = True
End Function

'Modified by Lydia 2018/02/01
'Private Sub cmdok_Click(Index As Integer)
Public Sub cmdok_Click(Index As Integer)
Dim bolErr As Boolean
Dim strCode0 As String, strCode1 As String, strCode2 As String, strCode3 As String
Dim intRt As Integer
'Dim nFrm As Form 'Added by Lydia 2021/04/29
   
   'Modify By Sindy 2022/7/6 將檢查的程式碼抽出來,變成一個函數
'******************************
'     檢查資料是否正確
   If CheckDataIsOk = False Then Exit Sub
'******************************
   
   'Add By Sindy 2012/2/23
   '內商轉案至他所
   If ((strSK02 = "2" And strSK03 = "0") Or (strSK02 = "6" And strSK03 = "0")) And _
      txtCaseProperty = "728" And Me.Frame1.Visible = True Then
      If CheckExist(0) = True Then
         MsgBox "找不到此本所案號在基本檔之資料"
         txtSystem.SetFocus
         Exit Sub
      End If
      
      If textCP24.Text = "" Then
         MsgBox "內商轉案至他所，案件准駁不可空白！"
         textCP24.SetFocus
         Exit Sub
      End If
      
      bolErr = False
      textCP24_Validate bolErr
      If bolErr = True Then
         Exit Sub
      End If
      
      Call T728Progress
      Exit Sub
   End If
   '2012/2/23 End
   
   ' 91.09.04 modify by louis
   If textCP05 = "111111" Then
      'Added by Morgan 2013/6/24 檢查基本檔是否存在
      If CheckExist(0) = True Then
         MsgBox "找不到此本所案號在基本檔之資料", vbCritical
         txtSystem.SetFocus
         Exit Sub
      End If
      'end 2013/6/24
      
      'Add By Sindy 2025/7/29
      '針對轉案進來的案件，若內部收文日為111111'案件性質為[（107）再審申請]，則
      '按確定後'彈出訊息：請輸入再審查送件日期,將日期回寫到進度檔的[再審查申請]的發文日
      If txtSystem = "FCP" And txtCaseProperty = "107" Then
input_CP27:
         m_strCP27 = InputBox("請輸入再審查送件日期！" & vbCrLf & vbCrLf & _
                            "※若不確定正確之送件日期，請輸入大概之日期即可（僅供判斷新、舊法）", , m_strCP27)
         If Trim(m_strCP27) = "" Then
            MsgBox "再審查送件日期不可空白！", vbInformation, "檢核資料"
            GoTo input_CP27
         Else
            If CheckIsTaiwanDate(m_strCP27, False) = False Then
               MsgBox "再審查送件日期(" & m_strCP27 & ")，日期格式不正確！", vbExclamation, "檢核資料"
               GoTo input_CP27
            End If
         End If
      End If
      '2025/7/29 END
      
      'Add By Cheng 2003/11/25
      '檢查流水號是否大於自動編號
      If CheckCaseNo = True Then
         If Me.txtSystem.Text = "TF" Then
            Me.txtTFCode(0).SetFocus
            txtTFCode_GotFocus 0
         Else
            Me.txtCode(0).SetFocus
            txtCode_GotFocus 0
         End If
         Exit Sub
      End If
      'End
      
      'Add by Amy 2013/08/27 判斷系統別及案件性質為416者輸入頁數及項數
      If txtSystem = "FCP" And txtCaseProperty = "416" Then
         intCP135 = Val(InputBox("請輸入頁數!!"))
         intCP136 = Val(InputBox("請輸入項數!!"))
         If Not (intCP135 > 0 And intCP136 > 0) Then
            MsgBox "頁數及項數需大於0", vbCritical
            Exit Sub
         End If
      End If
      'end 2013/08/27
      
      'Added by Lydia 2016/09/21 內部收文(假收文)更新初審階段提分割
      m_UpdPA163 = ""
      If txtSystem = "FCP" And txtCaseProperty = "307" Then
         strExc(0) = "select pa163 from patent where pa01='" & txtSystem & "' and pa02='" & txtCode(0) & "' and pa03='" & txtCode(1) & "' and pa04='" & txtCode(2) & "' and pa09='000' and pa11 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(0) = "" & RsTemp(0)
         End If
         If strExc(0) = "" Then
            If MsgBox("請問本案是否為初審階段提分割??", vbYesNo + vbDefaultButton2) = vbYes Then
               m_UpdPA163 = "PA163='Y'"
            Else
               m_UpdPA163 = "PA163='N'"
            End If
         End If
      End If
      'end 2016/09/21
      
      'edit by nickc 2007/01/11
      If OnSaveNewData = True Then
         '2014/9/5 add by sonia
         If txtSystem = "FCP" Then
            Select Case txtCaseProperty
               Case "101"
                  MsgBox "若尚未提實審，請至下一程序檔新增實審期限！" & vbCrLf & vbCrLf & _
                         "若已提實審則請再做實審的內部收文, 收文日為111111！"
               Case "416"
                  MsgBox "若已收到實審通知，請再做 實審通知日輸入！" & vbCrLf & vbCrLf & _
                         "並請注意來函收文日請輸入111111！"
            End Select
         End If
         'end 2014/9/5
      
         'Modify By Sindy 2012/3/1 寫成共用函數
         Call SetColClearVal(False)
      End If
      Exit Sub
   End If
   
   '執行內部收文
   ' 91.09.04 modify by louis
   If intChoose <> 0 Then
      'Add by Morgan 2007/1/15 台灣新型主動修正控制,內部收文也要
      If intModifyKind = 0 And (txtSystem = "P" Or txtSystem = "FCP") And Me.txtCode(0).Text <> "" Then
         If txtCaseProperty = "203" Then
            If PUB_Check203(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text, intChoose) = False Then
               Exit Sub
            End If
         End If
      End If
      'end 2007/1/15
      '911111 nick 若為內部收文，案號不存再不新增
      If intModifyKind = 0 And CheckExist(0) = True Then
         MsgBox "找不到此本所案號在基本檔之資料"
         txtSystem.SetFocus
         Exit Sub
      End If
      'Modify By Cheng 2003/03/28
'      OnNextForm
      If OnNextForm = True Then
         Me.Hide
         Exit Sub
      Else
         Exit Sub
      End If
   End If
   
   '執行收文
   Select Case intReceiveKind
      Case 0 '0為接洽紀錄單
         Select Case intModifyKind
            Case 0 '新增
'               '911017 nick
'               'If txtCode(0) <> "" Then
'               If txtCode(0) <> "" Or txtTFCode(0) <> "" Then
'                  '911017 nick 將 1 和 2  對調
'                  If CheckCaseNo Then   '----1
'                     Exit Sub
'                  End If
'                  If CheckExist(1) Then    '------2
'                     Exit Sub
'                  End If
'               Else
'                  intSaveMode = 1
'               End If
'               If intChoose = 1 Then
'                  If txtCode(0).Text <> "" Then
'                     adoquery.CursorLocation = adUseClient
'                     adoquery.Open "select np06, np07 from nextprogress where np02 = '" & txtSystem & "' and np03 = '" & txtCode(0).Text & "' and np04 = '" & IIf(txtCode(1).Text = "", "0", txtCode(1).Text) & "' and np05 = '" & IIf(txtCode(2).Text = "", "00", txtCode(2).Text) & "'", cnnConnection, adOpenStatic, adLockReadOnly
'                     If adoquery.RecordCount < 2 And adoquery.RecordCount > 0 Then
'                        If adoquery.Fields("np07").Value <> txtCaseProperty.Text Then
'                           If IsNull(adoquery.Fields(0).Value) = True Or adoquery.Fields(0).Value = "" Then
'                              ShowMsg "下一程序的人有未收文之資料，請自行處理"
'                           End If
'                        End If
'                     End If
'                     adoquery.Close
'                   End If
'               End If
'               'edit by nickc 2007/02/02 不用 dll 了
'               'objPublicData.GetSystemKind txtSystem.Text, intCaseKind, strCaseName
'               ClsPDGetSystemKind txtSystem.Text, intCaseKind, strCaseName
               
               Select Case intCaseKind
                  Case 專利
'                     If intModifyKind <> 0 Then
'                        If CheckKeyInOkay = False Then
'                           Exit Sub
'                        End If
'                     'Add by Morgan 2004/5/29
'                     Else  '新增
'                        'Add By Sindy 2013/4/11 大陸107.復審案下一程序只有一筆復審期限時,不可收文延期404
'                        If txtSystem = "P" And Me.txtCode(0).Text <> "" And txtCaseProperty = "404" Then
'                           adoquery.CursorLocation = adUseClient
'                           strSql = "Select NP07 FROM nextprogress,Patent WHERE NP02=PA01 and NP03=PA02 and NP04=PA03 and NP05=PA04" & _
'                                    " and NP02='" & Trim(txtSystem) & "' and NP03='" & Trim(txtCode(0)) & "' and NP04='" & IIf(txtCode(1).Text = "", "0", txtCode(1).Text) & "' and NP05='" & IIf(txtCode(2).Text = "", "00", txtCode(2).Text) & "'" & _
'                                    " and NP06 is null and PA09='020'" & strNpSqlOfNoSalesDuty
'                           adoquery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                           If adoquery.RecordCount = 1 Then
'                              If adoquery.Fields("NP07").Value = "107" Then
'                                 MsgBox "請通知智權同仁依大陸專利法規定復審案不得延期。", vbCritical
'                                 adoquery.Close
'                                 Exit Sub
'                              End If
'                           End If
'                           adoquery.Close
'                        End If
'                        '2013/4/11 End
'
'                        '專利修法問題
'                        If (txtSystem = "P" Or txtSystem = "FCP") And Me.txtCode(0).Text <> "" Then
'                           'Modify by Morgan 2004/7/5
'                           '加判斷申請國家=台灣
'                           If GetPA09(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) = "000" Then
'                              '2005/9/22 CANCEL BY SONIA 指案件准駁確定後 60 天, 因P-63193電腦無法判斷何謂准駁確定, 故先取消
'                              ''國內改請案不可逾准駁日60天
'                              If InStr("301,302,303,304,305,306", txtCaseProperty) > 0 Then
'                              '   If PUB_Check301(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) = False Then
'                              '      Exit Sub
'                              '   End If
'                              '2005/9/22 END
'                              '國內延緩公告案需控制領證未發文
'                              ElseIf txtCaseProperty = "412" Then
'                                 If Check412(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) = False Then
'                                    Exit Sub
'                                 End If
'                              '技術報告發文時無公告日不可發文，無公告日不可發文或公告日<93.7.1者不可收文也不可發文
'                              'Modify by Morgan 2007/8/29 加807第三人申請技術報告
'                              'ElseIf txtCaseProperty = "421" Then
'                              ElseIf txtCaseProperty = "421" Or txtCaseProperty = "807" Then
'                                 If PUB_Check421(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text, 1, txtCaseProperty) = False Then
'                                    Exit Sub
'                                 End If
'                              'Add by Morgan 2004/9/13 新型主動修正期限為申請日起兩個月
'                              ElseIf txtCaseProperty = "203" Then
'                                 If PUB_Check203(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) = False Then
'                                    Exit Sub
'                                 End If
'                              End If
'                           End If
'                        End If
'                     End If

                     'Mark by Amy 2021/12/21 Unload 沒 set Nothing 會殘留前次變數值,故由此清->改成共用,怕有沒清到的
'                     'Add by Amy 2021/12/20 改Form2.0 Unload 沒 set Nothing 會殘留前次變數值,故由此清
'                     If PUB_CheckFormExist("frm010005") = False Then
'                        Set frm010005 = Nothing
'                     End If
                     'Added by Lydia 2018/09/18 外專承辦後續案收文時,不知何故ClsPDGetAutoNumber沒有傳入收文號開頭造成錯誤,重新進入後又沒有錯誤(FCP-59592收123主張優惠期)
                     If strReceiveKind = "" Then
                          strReceiveKind = lblReciveCode.Caption
                     End If
                     'end 2018/09/18
                     frm010005.Caption = frm010001.Caption + "－" + strCaseName
                     frm010005.txtRecieveCode.Text = strReceiveKind + txtRecieveCode(0).Text
                     frm010005.txtPatent(1) = txtCaseProperty
                     frm010005.lblCaseProperty = lblCasePropertyName
                     'Added by Lydia 2018/08/30 外專後續案收文
                     'Modified by Lydia 2020/12/16 區分外專F2x
                     'If Left(mRole, 1) = "F" Then
                     If Left(mRole, 2) = "F2" Then
                         frm010005.txtPatent(2) = "07" '案件來源CP11
                         If ClsPDGetCaseSource("07", strExc(1)) Then
                              frm010005.lblCaseSource.Caption = strExc(1)
                         End If
                         frm010005.txtPatent(15) = strUserNum '智權人員=操作者
                     End If
                     'end 2018/08/30
                     'Add By Sindy 2009/07/06
                     If textYear.Visible = True Then
                        frm010005.textYear.Visible = True
                        frm010005.textYear.Enabled = textYear.Enabled
                        frm010005.Text1(0).Visible = True
                        frm010005.Label11(1).Visible = True
                        frm010005.Label11(1).Caption = Label11(1).Caption
                        If IsEmptyText(textYear) = False Or IsEmptyText(Text1(0)) = False Then
                           frm010005.textYear = textYear.Text
                           frm010005.Text1(0) = Text1(0).Text
                        End If
                     Else
                        frm010005.textYear.Visible = False
                        frm010005.Text1(0).Visible = False
                        frm010005.Label11(1).Visible = False
                     End If
                     '2009/07/06 End
                     If txtPetition.Visible = True Then
                        frm010005.fraPatition.Visible = True
                        frm010005.txtPatent(23) = txtPetition
                        frm010005.lblPetitionName = lblPetitionName
                        'Add by Morgan 2006/6/23
                        For intI = 2 To 5
                           frm010005.txtPetitionx(intI) = txtPetitionx(intI)
                           frm010005.lblPetitionNamex(intI) = lblPetitionNamex(intI)
                        Next
                        'end 2006/6/23
                        'Add By Cheng 2002/01/14
                        Select Case Me.txtCaseProperty.Text
                        Case 合併
                           frm010005.Label29.Caption = "合併申請人1："
                           For intI = 2 To 5
                              frm010005.Label31(intI).Caption = "合併申請人" & intI & "："
                           Next
                        Case 繼承
                           frm010005.Label29.Caption = "繼承申請人1："
                           For intI = 2 To 5
                              frm010005.Label31(intI).Caption = "繼承申請人" & intI & "："
                           Next
                        End Select
                     Else
                        frm010005.fraPatition.Visible = False
                     End If
                     frm010005.txtSystem.Text = txtSystem.Text
                     frm010005.txtCode(0).Text = txtCode(0).Text
                     frm010005.txtCode(1).Text = txtCode(1).Text
                     frm010005.txtCode(2).Text = txtCode(2).Text
                     'modify by sonia 90.10.8
                     If frm010005.txtSystem = "FCP" And frm010001.intChoose = 1 Then
                        frm010005.txtPatent(24) = strUserNum
                     'Add By Cheng 2001/12/12
                     Else
                        frm010005.txtPatent(24) = ""
                     End If
                     '2015/1/9 add by sonia FMP寰華案件後續收文時,提醒收文人員, 接洽單交外專
                     'Remove by Lydia 2018/05/07 影響代入舊案資料(FormActive)
'                     If txtCode(0) <> "" Then
'                        If PUB_FMPtoCheck(1, 2, "", txtSystem, txtCode(0), txtCode(1), txtCode(2)) = True Then
'                           MsgBox "此案號為FMP寰華案件, 收文後請將文件將 外專！", , "注意！"
'                        End If
'                     End If
                     '2015/1/9 end
                     
                     'Added by Morgan 2020/4/9
                     If m_bBatch = True Then
                        frm010005.m_bBatch = True
                        frm010005.m_CP11 = m_CP11
                        frm010005.m_CP13 = m_CP13
                     End If
                     'end 2020/4/9
                     
                     'Add By Sindy 2022/6/29
                     frm010005.m_strIR01 = m_strIR01
                     frm010005.m_strIR02 = m_strIR02
                     frm010005.m_strIR03 = m_strIR03
                     frm010005.m_strIR04 = m_strIR04
                     If Me.FraRecv.Visible = True Then '多案收文
                        frm010005.m_bMRecvBatch = True
                     End If
                     Set frm010005.m_PrevForm = mPrevForm
                     '2022/6/29 END
                     
                     frm010005.Show
                     
                  Case 商標
'                     If intModifyKind <> 0 Then
'                        If CheckKeyInOkay = False Then
'                           Exit Sub
'                        End If
'                     End If
                     
                     'Added by Lydia 2020/12/16
                     If strReceiveKind = "" Then
                          strReceiveKind = lblReciveCode.Caption
                     End If
                     'end by Lydia 2020/12/16
                     
                     'Added by Lydia 2015/11/12 新增查名單對應
                     Set frm010004.Tmpfrm090130 = Tmpfrm090130
                     frm010004.SetParent Me
                     'end 2015/11/12
                     frm010004.Caption = frm010001.Caption + "－" + strCaseName
                     frm010004.txtRecieveCode.Text = strReceiveKind + txtRecieveCode(0).Text
                     frm010004.txtSystem.Text = txtSystem.Text
                     frm010004.txtTrademark(1) = txtCaseProperty
                     frm010004.lblCaseProperty = lblCasePropertyName
                     'Added by Lydia 2020/12/16 外商臺灣案收文
                     If Left(mRole, 2) = "F1" Then
                         frm010004.txtTrademark(2) = "07" '案件來源CP11
                         If ClsPDGetCaseSource("07", strExc(1)) Then
                              frm010004.lblCaseSource.Caption = strExc(1)
                         End If
                         'Modified by Lydia 2023/01/10 系統增加提醒「此次收文是否為代收文」：選擇「是」系統自動設定為原智權人員，選擇「否」系統自動設定輸入者為智權人員。
                         'frm010004.txtTrademark(12) = strUserNum '智權人員=操作者
                         frm010004.txtTrademark(12) = m_strCP14
                         frm010004.lblSales.Caption = strUserName
                     End If
                     'end 2020/12/16
                     If txtPetition.Visible = True Then
                        frm010004.fraPatition.Visible = True
                        frm010004.txtTrademark(20) = txtPetition
                        'edit by nickc 2006/11/22
                        'frm010004.lblPetitionName = lblPetitionName
                        frm010004.lblPetitionName(0) = lblPetitionName
                        'Add by nickc 2006/11/22
                        frm010004.txtTrademark(28) = txtPetitionx(2)
                        frm010004.lblPetitionName(1) = lblPetitionNamex(2)
                        frm010004.txtTrademark(29) = txtPetitionx(3)
                        frm010004.lblPetitionName(2) = lblPetitionNamex(3)
                        frm010004.txtTrademark(30) = txtPetitionx(4)
                        frm010004.lblPetitionName(3) = lblPetitionNamex(4)
                        frm010004.txtTrademark(31) = txtPetitionx(5)
                        frm010004.lblPetitionName(4) = lblPetitionNamex(5)
                        'end 2006/6/23
                     Else
                        frm010004.fraPatition.Visible = False
                     End If
                     'TF為馬德里案，另外判斷
                     If txtSystem.Text <> 馬德里案 Then
                        frm010004.fraElse.Visible = True
                        frm010004.fraTF.Visible = False
                        frm010004.txtCode(0).Text = txtCode(0).Text
                        frm010004.txtCode(1).Text = txtCode(1).Text
                        frm010004.txtCode(2).Text = txtCode(2).Text
                     Else
                        frm010004.fraElse.Visible = False
                        frm010004.fraTF.Visible = True
                        frm010004.txtTFCode(0).Text = txtTFCode(0).Text
                        frm010004.txtTFCode(1).Text = txtTFCode(1).Text
                        frm010004.txtTFCode(2).Text = txtTFCode(2).Text
                        frm010004.txtTFCode(3).Text = txtTFCode(3).Text
                     End If
                     'Add by Morgan 2003/11/26
                     frm010004.fraTM15.Visible = txtTM15Control()
                     '---End
                      
                     'Add By Sindy 2023/6/29
                     frm010004.m_strIR01 = m_strIR01
                     frm010004.m_strIR02 = m_strIR02
                     frm010004.m_strIR03 = m_strIR03
                     frm010004.m_strIR04 = m_strIR04
                     Set frm010004.m_PrevFormIR = mPrevForm
                     '2022/6/29 END
                     
                     frm010004.Show
                  
                  Case Else
'                     If intModifyKind <> 0 Then
'                        If CheckKeyInOkay Then
'                           Exit Sub
'                        End If
'                     End If
                     
                     'Added by Lydia 2021/04/29 先判斷表單是否存在
                     'Remove by Lydia 2021/12/16 已加入有收文的所有vbp
                     'If strSrvDate(1) >= ACS_PFrateStart And txtSystem.Text = "ACS" And txtCaseProperty = "112" Then
                     '   Set nFrm = Forms(0).GetForm("frm010006_1")
                     'End If
                     ''end 2021/04/29
                     'end 2021/12/16
                     
                     If frm010001.intCaseKind = 顧問 And txtCaseProperty = 顧問聘任 Then
                        'Mark by Amy 2021/12/21 改成共用,怕有沒清到的
'                        'Added by Lydia 2021/12/13
'                        '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
'                        If PUB_CheckFormExist("frm010006") = False Then
'                           Set frm010006 = Nothing
'                        End If
'                        'end 2021/12/13
                        frm010006.Caption = frm010001.Caption + "－" + strCaseName
                        frm010006.txtRecieveCode.Text = strReceiveKind + txtRecieveCode(0).Text
                        frm010006.txtAdviser(1) = txtCaseProperty
                        frm010006.lblCaseProperty = lblCasePropertyName
                        frm010006.txtSystem.Text = txtSystem.Text
                        frm010006.txtCode(0).Text = txtCode(0).Text
                        frm010006.txtCode(1).Text = txtCode(1).Text
                        frm010006.txtCode(2).Text = txtCode(2).Text
                        frm010006.Show
                     'Added by Lydia 2021/04/29 ACS智財顧問專業分配比例管制：另開收文畫面
                     'Modified by Lydia 2021/12/16  已加入有收文的所有vbp
                     'ElseIf strSrvDate(1) >= ACS_PFrateStart And Not nFrm Is Nothing And txtSystem.Text = "ACS" And txtCaseProperty = "112" Then
                     ElseIf strSrvDate(1) >= ACS_PFrateStart And txtSystem.Text = "ACS" And txtCaseProperty = "112" Then
                     'Modified by Lydia 2021/12/16 nfrm=>frm010006_1
                        frm010006_1.Caption = frm010001.Caption + "－" + strCaseName
                        frm010006_1.txtRecieveCode.Text = strReceiveKind + txtRecieveCode(0).Text
                        frm010006_1.txtAdviser(1) = txtCaseProperty
                        frm010006_1.lblCaseProperty = lblCasePropertyName
                        frm010006_1.txtSystem.Text = txtSystem.Text
                        frm010006_1.txtCode(0).Text = txtCode(0).Text
                        frm010006_1.txtCode(1).Text = txtCode(1).Text
                        frm010006_1.txtCode(2).Text = txtCode(2).Text
                        frm010006_1.Show
                     'end 2021/12/16
                     'end 2021/04/29
                     Else
                        'Added by Lydia 2016/04/25 新增查名單對應
                        Set frm010007.Tmpfrm090130 = Tmpfrm090130
                        frm010007.SetParent Me
                        'end 2016/04/25
                        frm010007.Caption = frm010001.Caption + "－" + strCaseName
                        frm010007.txtRecieveCode.Text = strReceiveKind + txtRecieveCode(0).Text
                        frm010007.txtOther(1) = txtCaseProperty
                        frm010007.lblCaseProperty = lblCasePropertyName
                        'Added by Lydia 2018/08/30 外專後續案收文
                        'Memo by Lydia 2020/12/16 含外商
                        If Left(mRole, 1) = "F" Then
                            frm010007.txtOther(2) = "07" '案件來源CP11
                            If ClsPDGetCaseSource("07", strExc(1)) Then
                               frm010007.lblCaseSource.Caption = strExc(1)
                            End If
                            frm010007.txtOther(10) = strUserNum '智權人員=操作者
                        End If
                        'end 2018/08/30
                        frm010007.txtSystem.Text = txtSystem.Text
                        frm010007.txtCode(0).Text = txtCode(0).Text
                        frm010007.txtCode(1).Text = txtCode(1).Text
                        frm010007.txtCode(2).Text = txtCode(2).Text
                        frm010007.intCaseKind = intCaseKind
                        
                        'Add By Sindy 2022/8/17
                        frm010007.m_strIR01 = m_strIR01
                        frm010007.m_strIR02 = m_strIR02
                        frm010007.m_strIR03 = m_strIR03
                        frm010007.m_strIR04 = m_strIR04
                        If Me.FraRecv.Visible = True Then '多案收文
                           frm010007.m_bMRecvBatch = True
                        End If
                        Set frm010007.m_PrevForm = mPrevForm
                        '2022/8/17 END
                        
                        frm010007.Show
                     End If
               End Select
              
            Case 1, 2 '修改/查詢
               'edit by nickc 2007/02/02 不用 dll 了
               'intRt = objPublicData.CheckRecieveCode(strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, strCode0, strCode1, strCode2, strCode3)
               intRt = ClsPDCheckRecieveCode(strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, strCode0, strCode1, strCode2, strCode3)
               
               If intRt <> 0 Then
                   'edit by nickc 2007/02/02 不用 dll 了
                   'If objPublicData.GetSystemKind(strCode0, intCaseKind, strCaseName) = False Then
                   If ClsPDGetSystemKind(strCode0, intCaseKind, strCaseName) = False Then
                      Exit Sub
                   End If
                   Select Case intCaseKind
                     Case 專利
                        'Memo by Amy 2021/12/21 Unload 沒 set Nothing 會殘留前次變數值,故由此清->改成共用,怕有沒清到的
                        'Add by Amy 2021/12/20 改Form2.0 Unload 沒 set Nothing 會殘留前次變數值,故由此清
                        frm010005.Caption = frm010001.Caption + "－" + strCaseName
                        frm010005.txtRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
                        frm010005.txtSystem.Text = strCode0
                        frm010005.txtCode(0).Text = strCode1
                        frm010005.txtCode(1).Text = IIf(strCode2 = "0", "", strCode2)
                        frm010005.txtCode(2).Text = IIf(strCode3 = "00", "", strCode3)
                        frm010005.Show
                     Case 商標
                        'Added by Lydia 2015/11/12 新增查名單對應
                        Set frm010004.Tmpfrm090130 = Tmpfrm090130
                        frm010004.SetParent Me
                        'end 2015/11/12
                        frm010004.Caption = frm010001.Caption + "－" + strCaseName
                        frm010004.txtRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
                        frm010004.txtSystem.Text = strCode0
                        'TF為馬德里案，另外判斷
                        If strCode0 <> 馬德里案 Then
                           frm010004.fraElse.Visible = True
                           frm010004.fraTF.Visible = False
                           frm010004.txtCode(0).Text = strCode1
                           frm010004.txtCode(1).Text = IIf(strCode2 = "0", "", strCode2)
                           frm010004.txtCode(2).Text = IIf(strCode3 = "00", "", strCode3)
                        Else
                           frm010004.fraElse.Visible = False
                           frm010004.fraTF.Visible = True
                           frm010004.txtTFCode(0).Text = Left(strCode1, 5)
                           frm010004.txtTFCode(1).Text = IIf(Right(strCode1, 1) = "0", "", Right(strCode1, 1))
                           frm010004.txtTFCode(2).Text = IIf(strCode2 = "0", "", strCode2)
                           frm010004.txtTFCode(3).Text = IIf(strCode3 = "00", "", strCode3)
                        End If
                        frm010004.Show
                     Case Else
                        'Added by Lydia 2021/04/29 先判斷表單是否存在
                        'Remove by Lydia 2021/12/16 已加入有收文的所有vbp
                        'If strSrvDate(1) >= ACS_PFrateStart And txtSystem.Text = "ACS" And m_CP10 = "112" Then
                        '   Set nFrm = Forms(0).GetForm("frm010006_1")
                        'End If
                        ''end 2021/04/29
                        'end 2021/12/16
                        
                        If frm010001.intCaseKind = 顧問 And intRt = 2 Then
                            'Mark by amy 2021/12/21 改成共用,怕有沒清到的
'                            'Added by Lydia 2021/12/13
'                            '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
'                            If PUB_CheckFormExist("frm010006") = False Then
'                               Set frm010006 = Nothing
'                            End If
'                            'end 2021/12/13
                           frm010006.Caption = frm010001.Caption + "－" + strCaseName
                           frm010006.txtRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
                           frm010006.txtSystem.Text = strCode0
                           frm010006.txtCode(0).Text = strCode1
                           frm010006.txtCode(1).Text = IIf(strCode2 = "0", "", strCode2)
                           frm010006.txtCode(2).Text = IIf(strCode3 = "00", "", strCode3)
                           frm010006.Show
                        'Added by Lydia 2021/04/29 ACS智財顧問專業分配比例管制：另開收文畫面
                        'Modified by Lydia 2021/12/16 已加入有收文的所有vbp
                        'ElseIf strSrvDate(1) >= ACS_PFrateStart And Not nFrm Is Nothing And txtSystem.Text = "ACS" And m_CP10 = "112" Then
                        ElseIf strSrvDate(1) >= ACS_PFrateStart And txtSystem.Text = "ACS" And m_CP10 = "112" Then
                        'Modified by Lydia 2021/12/16 nfrm=>frm010006_1
                            frm010006_1.Caption = frm010001.Caption + "－" + strCaseName
                            frm010006_1.txtRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
                            frm010006_1.txtSystem.Text = strCode0
                            frm010006_1.txtCode(0).Text = strCode1
                            frm010006_1.txtCode(1).Text = IIf(strCode2 = "0", "", strCode2)
                            frm010006_1.txtCode(2).Text = IIf(strCode3 = "00", "", strCode3)
                            frm010006_1.Show
                        'end 2021/12/16
                        'end 2021/04/29
                        Else
                          'Added by Lydia 2016/04/25 新增查名單對應
                           Set frm010007.Tmpfrm090130 = Tmpfrm090130
                           frm010007.SetParent Me
                           'end 2016/04/25
                           frm010007.Caption = frm010001.Caption + "－" + strCaseName
                           frm010007.txtRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
                           frm010007.txtSystem.Text = strCode0
                           frm010007.txtCode(0).Text = strCode1
                           frm010007.txtCode(1).Text = IIf(strCode2 = "0", "", strCode2)
                           frm010007.txtCode(2).Text = IIf(strCode3 = "00", "", strCode3)
                           frm010007.Show
                        End If
                  End Select
               Else
                   bolErr = True
               End If
         End Select
         
      Case 1 '1為政府來函
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.CheckRecieveCode(strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, strCode0, strCode1, strCode2, strCode3) <> 0 Then
         If ClsPDCheckRecieveCode(strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, strCode0, strCode1, strCode2, strCode3) <> 0 Then
            frm010002.lblRecieveCode = strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text
            frm010002.Caption = Me.Caption
         Else
            Exit Sub
         End If
   End Select
   
   'If bolErr = False Then
      frm010001.Hide
   'End If
End Sub

'由於使用Validate()，以致無法正確跳躍Focus，因此在txtPetition_GotFocus()及
'cmdOK_GotFocus()加入判斷，以得正確之跳躍
Private Sub cmdOK_GotFocus(Index As Integer)
   Static boltxtPetition As Boolean
   
   If Index = 0 Then
      If fraPatition.Visible Then
         'Modify by Morgan 2006/6/27
         'If boltxtPetition = False Then
         If m_bolStopOnTxtPetition = True Then
            txtPetition.SetFocus
            boltxtPetition = True
         End If
      Else
         boltxtPetition = False
      End If
      'Add By Sindy 2009/07/06
      If textYear.Visible = True Then
         If m_bolStopOntextYear = True Then
            If textYear.Enabled = True Then
               textYear.SetFocus
            Else
               Text1(0).SetFocus
            End If
         End If
      End If
      'Add by Sindy 2012/3/1
      If ((strSK02 = "2" And strSK03 = "0") Or (strSK02 = "6" And strSK03 = "0")) And _
         txtCaseProperty = "728" And Me.Frame1.Visible = True And textCP24 = "" Then
         textCP24.SetFocus
      End If
      '2012/3/1 End
      
      'Added by Lydia 2020/05/20 是否停駐於案源單號
      If FraLOS.Visible = True And m_bolStopOntxtLOS15 = True Then
          txtLOS15.SetFocus
          txtLOS15_GotFocus
          m_bolStopOntxtLOS15 = False
      End If
      'end 2020/05/20
      
      'Added by Lydia 2020/12/04 CFT脫歐案是否停駐於歐盟案案號
      If fraNA239.Visible = True And m_bolStopOntxtCaseNa239 = True Then
          txtCaseNa239.SetFocus
          txtCaseNa239_GotFocus
          m_bolStopOntxtCaseNa239 = False
      End If
      'end 2020/12/04
   End If
End Sub

Public Sub ClearForm(ByVal strAuto1 As String, ByVal strAuto2 As String)
Dim i As Integer, bolNewCaseCode As Boolean

   If intModifyKind = 0 Then
      txtRecieveCode(1).Text = Mid(strAuto1, 4)
      If strAuto2 <> "" Then
         If txtSystem = 馬德里案 Then
            If txtTFCode(0) <> strAuto2 Then bolNewCaseCode = True
         Else
            If txtCode(0) <> strAuto2 Then bolNewCaseCode = True
         End If
      End If
      If bolNewCaseCode Then
         lblCaseCode = txtSystem + "- " + strAuto2
      Else
         lblCaseCode = txtSystem
         If txtSystem = 馬德里案 Then
            For i = 0 To 3
                   If txtTFCode(i) <> "" Then
                      lblCaseCode = lblCaseCode + "- " + txtTFCode(i)
                   Else
                      Exit For
                   End If
            Next
         Else
            For i = 0 To 2
                   If txtCode(i) <> "" Then
                      lblCaseCode = lblCaseCode + "- " + txtCode(i)
                   Else
                      Exit For
                   End If
            Next
         End If
      End If
   End If
   'Added by Lydia 2020/12/16 外商臺灣案收文：
   If Left(mRole, 2) = "F1" Then
      If txtCode(0) = "" Then '新案
          mPreCaseNo = txtSystem & "-" & strAuto2 & "-0-00"
      Else                             '非新案
          mPreCaseNo = txtSystem & "-" & txtCode(0) & "-" & IIf(txtCode(1) = "", "0", txtCode(1)) & "-" & IIf(txtCode(2) = "", "00", txtCode(2))
      End If
   ElseIf Left(mRole, 2) = "F2" Then
   'end 2020/12/16
       mPreCaseNo = txtSystem & "-" & txtCode(0) & "-" & IIf(txtCode(1) = "", "0", txtCode(1)) & "-" & IIf(txtCode(2) = "", "00", txtCode(2)) 'Added by Lydia 2018/08/30 記錄前一筆收文的本所案號
   End If 'Added by Lydia 2020/12/26
   txtSystem = ""
   For i = 0 To 2
      txtCode(i) = ""
   Next
   For i = 0 To 3
      txtTFCode(i) = ""
   Next
   txtCaseProperty = ""
   txtCaseProperty.Tag = "" 'Added by Lydia 2020/05/20
   txtPetition = ""
   For intI = 2 To 5
      txtPetitionx(intI) = Empty
      lblPetitionNamex(intI) = Empty
   Next
   'Add By Sindy 2009/07/06
   textYear.Visible = False
   Text1(0).Visible = False
   Label11(1).Visible = False
   textYear.Text = ""
   Text1(0).Text = ""
   '2009/07/06 End
   
   'Add By Sindy 2012/2/24
   Frame1.Visible = False
   textCP24 = ""
   strMailNote = ""
   strTo = ""
   '2012/2/24 End
   
   txtFMP = "" 'Added by Lydia 2018/05/07 清空-前案是否為FMP案
   txtNA01 = "": lblNation.Caption = ""   'Added by Lydia 2021/11/10 清空FMP案的申請國家
   bolChild013 = False 'Added by Lydia 2024/02/20 'Added by Lydia 2024/02/20 清空-增加FMP案之子案(新案)
   
   'Added by Lydia 2020/05/20
   txtLOS15 = ""
   m_strLOSkind = ""
   FraLOS.Visible = False
   
   'Added by Lydia 2020/11/19
   txtCaseNa239 = ""
   fraNA239.Visible = False
End Sub

'Added by Morgan 2020/4/9 FCP批次收文
Private Sub BatchAdd()
   Dim strQ As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
         
   If strUserNum <> "QPGMR" Then MsgBox "請先切換使用者為 QPGMR !!", vbExclamation: Exit Sub
   
   If MsgBox("是否 SQL 語法有改好？" & vbCrLf & vbCrLf & "注意:是否新案(CP31)需手動更新！", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then Exit Sub
   
   strQ = "select * from patent where pa93=" & strSrvDate(1) & " and pa92='QPGMR' and pa01='FCP' order by pa01,pa02,pa03,pa04"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      If MsgBox("將有 " & RsQ.RecordCount & " 件案號要收文，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then GoTo ExitPort
      With RsQ
      m_bBatch = True
      Do While Not .EOF
         txtSystem = .Fields("pa01")
         txtCode(0) = .Fields("pa02")
         txtCode(1) = .Fields("pa03")
         txtCode(2) = .Fields("pa04")
         txtCaseProperty = "935"
         m_CP11 = "07"
         m_CP13 = "A6010"
         cmdok(0).Value = True
         DoEvents
         
         txtSystem = .Fields("pa01")
         txtCode(0) = .Fields("pa02")
         txtCode(1) = .Fields("pa03")
         txtCode(2) = .Fields("pa04")
         txtCaseProperty = "401"
         m_CP11 = "07"
         m_CP13 = "A6010"
         cmdok(0).Value = True
         DoEvents
         
         .MoveNext
      Loop
      m_bBatch = False
      End With
   End If
   
ExitPort:

   Set RsQ = Nothing
End Sub

'Add By Sindy 2022/7/7
Private Sub Command1_Click(Index As Integer)
Dim strTemp As String
Dim arrData As Variant
Dim ii As Integer
Dim adoRst As ADODB.Recordset
Dim strData1 As String, strData2 As String, strData3 As String, strData4 As String, strData5 As String
   
   If Index = 1 Then
      If List1.ListCount = 0 Then
         m_strfirCP01 = "": m_strfirCP02 = "": m_strfirCP03 = "": m_strfirCP04 = ""
         MsgBox "無資料列可刪除！", vbExclamation
         Exit Sub
      End If
   End If
   
   '新增
   If Index = 0 And txtCode(0).Text <> "" Then
      'Modify By Sindy 2022/11/18
      If txtCode(1) = "" Then txtCode(1) = "0"
      If txtCode(2) = "" Then txtCode(2) = "00"
      '2022/11/18 END
      '檢查欲新增的案號,是否同第一筆案號的基礎資料
      If m_strfirCP01 <> "" And m_strfirCP02 <> "" Then
         For ii = 0 To List1.ListCount - 1
            If InStr(List1.List(ii), txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2)) > 0 Then
               MsgBox "此案號( " & txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2) & " )已重覆，新增失敗！", vbExclamation
               Exit Sub
            End If
         Next ii
         
         '是否同FC代理人
         strSql = "select pa75 from patent where pa01='" & m_strfirCP01 & "' and pa02='" & m_strfirCP02 & "' and pa03='" & m_strfirCP03 & "' and pa04='" & m_strfirCP04 & "'" & _
                  " union select sp26 from servicepractice where sp01='" & m_strfirCP01 & "' and sp02='" & m_strfirCP02 & "' and sp03='" & m_strfirCP03 & "' and sp04='" & m_strfirCP04 & "'"
         intI = 1
         Set adoRst = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strData1 = "" & adoRst.Fields(0)
            If strData1 <> "" Then
               strSql = "select pa75 from patent where pa01='" & txtSystem & "' and pa02='" & txtCode(0) & "' and pa03='" & txtCode(1) & "' and pa04='" & txtCode(2) & "'" & _
                        " union select sp26 from servicepractice where sp01='" & txtSystem & "' and sp02='" & txtCode(0) & "' and sp03='" & txtCode(1) & "' and sp04='" & txtCode(2) & "'"
               intI = 1
               Set adoRst = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  If strData1 <> "" & adoRst.Fields(0) Then
                     If MsgBox("第一筆案號的代理人為(" & strData1 & ")" & vbCrLf & vbCrLf & _
                        "此案號為(" & "" & adoRst.Fields(0) & ")" & vbCrLf & vbCrLf & _
                        "FC代理人不同，確定要繼續嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                        Set adoRst = Nothing
                        Exit Sub
                     End If
                  End If
               End If
            End If
         End If
         
         '是否同申請人
         strSql = "select pa26,pa27,pa28,pa29,pa30 from patent where pa01='" & m_strfirCP01 & "' and pa02='" & m_strfirCP02 & "' and pa03='" & m_strfirCP03 & "' and pa04='" & m_strfirCP04 & "'" & _
                  " union select sp08,sp58,sp59,sp65,sp66 from servicepractice where sp01='" & m_strfirCP01 & "' and sp02='" & m_strfirCP02 & "' and sp03='" & m_strfirCP03 & "' and sp04='" & m_strfirCP04 & "'"
         intI = 1
         Set adoRst = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strData1 = "" & adoRst.Fields(0)
            strData2 = "" & adoRst.Fields(1)
            strData3 = "" & adoRst.Fields(2)
            strData4 = "" & adoRst.Fields(3)
            strData5 = "" & adoRst.Fields(4)
            If strData1 <> "" Then
               strSql = "select pa26||','||pa27||','||pa28||','||pa29||','||pa30 from patent where pa01='" & txtSystem & "' and pa02='" & txtCode(0) & "' and pa03='" & txtCode(1) & "' and pa04='" & txtCode(2) & "'" & _
                  " union select sp08||','||sp58||','||sp59||','||sp65||','||sp66 from servicepractice where sp01='" & txtSystem & "' and sp02='" & txtCode(0) & "' and sp03='" & txtCode(1) & "' and sp04='" & txtCode(2) & "'"
               intI = 1
               Set adoRst = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  If (strData1 <> "" And InStr(adoRst.Fields(0), strData1) = 0) Or _
                     (strData2 <> "" And InStr(adoRst.Fields(0), strData2) = 0) Or _
                     (strData3 <> "" And InStr(adoRst.Fields(0), strData3) = 0) Or _
                     (strData4 <> "" And InStr(adoRst.Fields(0), strData4) = 0) Or _
                     (strData5 <> "" And InStr(adoRst.Fields(0), strData5) = 0) Then
                     If MsgBox("第一筆案號的申請人為(" & strData1 & IIf(strData2 <> "", "," & strData2, "") & IIf(strData3 <> "", "," & strData3, "") & IIf(strData4 <> "", "," & strData4, "") & IIf(strData5 <> "", "," & strData5, "") & ")" & vbCrLf & vbCrLf & _
                        "此案號為(" & Replace(Replace(adoRst.Fields(0), ",,", ""), ",)", ")") & ")" & vbCrLf & vbCrLf & _
                        "申請人不同，確定要繼續嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                        Set adoRst = Nothing
                        Exit Sub
                     End If
                  End If
               Else
                  Set adoRst = Nothing
                  MsgBox txtSystem & "-" & txtCode(0) & _
                     IIf(txtCode(1) & txtCode(2) <> "000", "-" & txtCode(1) & "-" & txtCode(2), "") & _
                     "無資料！", vbExclamation
                  Exit Sub
               End If
            End If
         Else
            Set adoRst = Nothing
            MsgBox m_strfirCP01 & "-" & m_strfirCP02 & _
               IIf(m_strfirCP03 & m_strfirCP04 <> "000", "-" & m_strfirCP03 & "-" & m_strfirCP04, "") & _
               "無資料！", vbExclamation
            Exit Sub
         End If
      End If
      
      '檢查資料是否正確
      If CheckDataIsOk = False Then Exit Sub
      
      If m_strfirCP01 = "" Or m_strfirCP02 = "" Then
         '記錄第一筆案號
         m_strfirCP01 = txtSystem
         m_strfirCP02 = txtCode(0)
         m_strfirCP03 = txtCode(1)
         m_strfirCP04 = txtCode(2)
      End If
      
      '新增至案號區
      List1.AddItem txtSystem & "-" & txtCode(0) & "-" & txtCode(1) & "-" & txtCode(2)
   
   '刪除
   Else
      If List1.ListCount > 0 Then
         ii = 0
         Do While ii < List1.ListCount
            If List1.Selected(ii) = True Then
               strTemp = List1.List(1)
               arrData = Split(strTemp, " ")
               If ii = 0 And List1.ListCount > 1 Then '異動第一筆案號資料
                  m_strfirCP01 = SystemNumber(CStr(arrData(0)), 1)
                  m_strfirCP02 = SystemNumber(CStr(arrData(0)), 2)
                  m_strfirCP03 = SystemNumber(CStr(arrData(0)), 3)
                  m_strfirCP04 = SystemNumber(CStr(arrData(0)), 4)
               End If
               '若已收文,不可以移除案號
               strExc(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                           " and mcr02='" & SystemNumber(CStr(arrData(0)), 1) & "'" & _
                           " and mcr03='" & SystemNumber(CStr(arrData(0)), 2) & "'" & _
                           " and mcr04='" & SystemNumber(CStr(arrData(0)), 3) & "'" & _
                           " and mcr05='" & SystemNumber(CStr(arrData(0)), 4) & "'" & _
                           " and mcr11 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox RsTemp.Fields("mcr02") & "-" & RsTemp.Fields("mcr03") & _
                     IIf(RsTemp.Fields("mcr04") & RsTemp.Fields("mcr05") <> "000", "-" & RsTemp.Fields("mcr04") & "-" & RsTemp.Fields("mcr05"), "") & _
                     " 已收文( " & RsTemp.Fields("mcr11") & " )，不可移除！", vbExclamation
                  Exit Sub
               End If
               List1.RemoveItem ii
               ii = ii - 1
            End If
            ii = ii + 1
         Loop
         
         If List1.ListCount = 0 Then '刪除第一筆案號資料
            m_strfirCP01 = "": m_strfirCP02 = "": m_strfirCP03 = "": m_strfirCP04 = ""
         End If
      End If
   End If
   
   If m_strfirCP01 <> "" Then
      lblCaseNo.Caption = m_strfirCP01 & "-" & m_strfirCP02 & "-" & m_strfirCP03 & "-" & m_strfirCP04
   Else
      lblCaseNo.Caption = ""
   End If
   LblCnt.Caption = List1.ListCount
   Set adoRst = Nothing
End Sub

Private Sub Form_Activate()
Dim intMCR11isNull As Integer
   'edit by nickc 2007/02/06 不用 dll 了
   'If obj001 Is Nothing Then
   '   Set obj001 = CreateObject("prjTaieDll001.cls001")
   '   Set obj001.Connection = cnnConnection
   'End If
   
   'Modify By Sindy 2010/8/17 比對自動編號年度
   'txtRecieveCode(0).Text = GetTaiwanThisYear
   txtRecieveCode(0).Text = CompAutoNumberYear(GetTaiwanThisYear)
   'Add By Cheng 2002/07/17
   strReceiveKind = ""
   Select Case intReceiveKind
      Case 0
         If intChoose = 1 Then
            strReceiveKind = 內部收文
         Else
            strReceiveKind = 接洽記錄單
         End If
         Select Case intModifyKind
            Case 0
               '新增：輸入本所案號
               fraRecieve.Enabled = False
               fraCode.Visible = True
               lblRecieveKind = "上一筆之收文號："
               fraLastCaseCode.Visible = True
               'Modified by Lydia 2016/04/29
               'txtSystem.SetFocus
               'Modified by Lydia 2018/02/01
               'Me.txtSystem.SetFocus
               'Modified by Lydia 2018/03/13 exe檔會出錯
               'If Me.Visible = True Then Me.txtSystem.SetFocus
               If Me.Visible = True And Me.Enabled = True Then
                   Me.txtSystem.SetFocus
               End If
               'end 2018/03/13
               m_bolStopOnTxtPetition = False '2011/4/12 add by sonia FCP-015522合併收文後回此畫面不會停在txtSystem
               'Added by Lydia 2018/08/30 外專後續案收文完成,跳卷宗區
               'Memo by Lydia 2020/12/16 F含外商臺灣案收文
               If Left(frm010001.mRole, 1) = "F" And txtRecieveCode(0) <> "" And txtRecieveCode(1) <> "" And mPreCaseNo <> "" Then
                    Screen.MousePointer = vbHourglass
                    frm100101_L.m_strKey = mPreCaseNo
                    frm100101_L.SetParent Me
                    If frm100101_L.QueryData = True Then
                       mPreCaseNo = ""
                       frm100101_L.Show
                       Me.Hide
                    End If
                    Screen.MousePointer = vbDefault
               End If
               'end 2018/08/30
            Case 1, 2
               '修改，刪除：輸入收文號
               lblRecieveKind = "收文號："
               fraRecieve.Enabled = True
               fraCode.Visible = False
               'Modified by Lydia 2016/04/29
               'txtRecieveCode(1).SetFocus
               If Me.Visible = True And Me.Enabled = True Then 'Added by Lydia 2018/03/13 exe檔會出錯
                    Me.txtRecieveCode(1).SetFocus
               End If
         End Select
      Case 1
         strReceiveKind = 政府機關來函
         fraCode.Visible = False
         '修改，刪除：可輸入收文號
         fraRecieve.Enabled = True
   End Select
   lblReciveCode.Caption = strReceiveKind
   
   'Added by Sindy 2018/2/22
   If m_strIR01 <> "" And m_Done = False Then
      txtSystem.Text = m_strCP01 '第一筆
      txtCode(0).Text = m_strCP02
      txtCode(1).Text = m_strCP03
      txtCode(2).Text = m_strCP04
      'Modify By Sindy 2022/7/7
      If FraRecv.Visible = False Then
         Label2.ForeColor = &H80000010 '深灰
         txtSystem.Locked = True
         txtCode(0).Locked = True
         txtCode(1).Locked = True
         txtCode(2).Locked = True
      Else '多案收文
         Label2.ForeColor = &H80000012 '黑色
'         If fraPatition.Visible = True Then
'            frm010001.FraRecvList.Height = 1965
'            frm010001.List1.Height = 1680
'         Else
'            frm010001.FraRecvList.Height = 3355
'            frm010001.List1.Height = 3040
'         End If
      End If
      '2022/7/7 END
      'textCP05.Text = m_RDate
      'cmdOK(0).Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
      
      '讀取是否有多案收文的資料
      strExc(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                  " order by decode(mcr02||mcr03||mcr04||mcr05,mcr07||mcr08||mcr09||mcr10,1,2) asc,mcr11 asc,mcr02,mcr03,mcr04,mcr05 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_strfirCP01 = "" & RsTemp.Fields("mcr07")
         m_strfirCP02 = "" & RsTemp.Fields("mcr08")
         m_strfirCP03 = "" & RsTemp.Fields("mcr09")
         m_strfirCP04 = "" & RsTemp.Fields("mcr10")
         txtCaseProperty = "" & RsTemp.Fields("mcr06")
         
         If FraRecv.Visible = True Then
            lblCaseNo.Caption = m_strfirCP01 & "-" & m_strfirCP02 & "-" & m_strfirCP03 & "-" & m_strfirCP04
            '新增案號區
            List1.Clear
            RsTemp.MoveFirst
            intMCR11isNull = 0 '未收文
            Do While Not RsTemp.EOF
               If "" & RsTemp.Fields("mcr11") = "" Then
                  intMCR11isNull = intMCR11isNull + 1
               End If
               List1.AddItem RsTemp.Fields("mcr02") & "-" & RsTemp.Fields("mcr03") & "-" & RsTemp.Fields("mcr04") & "-" & RsTemp.Fields("mcr05")
               RsTemp.MoveNext
            Loop
            LblCnt.Caption = List1.ListCount
            If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> m_strfirCP01 & m_strfirCP02 & m_strfirCP03 & m_strfirCP04 Then
               MsgBox "尚有未收文，但第一筆案號(" & lblCaseNo.Caption & ")與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")不符！"
               Exit Sub
            End If
            List1.Selected(0) = True 'Add By Sindy 2023/5/15
'            If MsgBox("尚有 " & intMCR11isNull & " 件案號未收文，是否繼續執行收文？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
'               'Modify By Sindy 2023/5/15
'               'Call OnSaveMRecv(1)
'               Call OnSaveMRecv
'            End If
         Else
            If MsgBox("此信件操作過多案收文，尚有 " & RsTemp.RecordCount & " 件案號，未執行完畢。" & vbCrLf & _
                      "要刪除資料，繼續執行嗎？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
               strSql = "delete from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                        " and mcr11 is null"
               cnnConnection.Execute strSql
            Else
               Unload Me
            End If
         End If
      End If
   End If
   '2018/2/22 END
End Sub

Private Sub Form_Load()
   intSaveMode = 0
   MoveFormToCenter Me
   Me.Width = 6990
   
   ' 91.09.03 modify by louis
   textCP05 = Empty
   If intChoose = 0 Then
      EnableTextBox textCP05, False
      Label4.Visible = False
      textCP05.Visible = False
   Else
      Label4.Visible = True
      textCP05.Visible = True
      EnableTextBox textCP05, True
   End If
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
   
   lblTCT.Caption = "" 'Added by Lydia 2017/11/14
   'Added by Lydia 2018/05/07
   'Modified by Lydia 2018/05/17 北所才顯示
   'If intChoose <> 0 Then
   'Modified by Lydia 2018/08/30 排除外專舊案
   'If intChoose <> 0 Or pub_strUserOffice <> "1" Then
   If intChoose <> 0 Or pub_strUserOffice <> "1" Or mRole <> "" Then
      fraFMP.Visible = False 'Modified by Lydia 2021/11/10 改用Frame
   End If
   fraFMP.Left = 0   'Modified by Lydia 2021/11/10 改用Frame
   lblTCTNO.Caption = ""
   'end 2018/05/07
      
   'Added by Lydia 2020/05/20
   FraLOS.Left = 0
   FraLOS.BackColor = &H8000000F
   FraLOS.Visible = False
   'end 2020/05/20
   m_Nation = "000" 'Added by Lydia 2020/06/08 預設台灣案
   
   'Added by Lydia 2020/11/19
   fraNA239.Left = 0
   fraNA239.BackColor = &H8000000F
   fraNA239.Visible = False
   'Added by Lydia 2021/11/10
   fraFMP.BackColor = &H8000000F
   
   LblCnt.Caption = "": lblCaseNo.Caption = "" 'Add By Sindy 2022/7/5
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Added by Lydia 2018/02/01
   'Modified by Lydia 2021/02/22 改判斷；
   'If TypeName(mPrevForm) = "frm060121_1" Then
   '   Call mPrevForm.GetB202(Me.Tag)
   If m_GetB202CP09 <> "" And Len(m_GetB202CP09) = 9 Then
      Call frm060121_1.GetB202(m_GetB202CP09)
      If Not mPrevForm Is Nothing Then
         Set mPrevForm = Nothing
      End If
      m_GetB202CP09 = ""
   'end 2021/02/22
   End If
   'end 2018/02/01
   
   'Add By Sindy 2018/2/23
   If m_strIR01 <> "" Then
      If Not mPrevForm Is Nothing Then
         Set mPrevForm = Nothing
      End If
   End If
   '2018/2/23 END
   
   'Add By Cheng 2002/07/18
   Set frm010001 = Nothing
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub

Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If textCP05 <> "111111" Then
         Cancel = True
         strMsg = "收文日只可為空白或111111"
         strTit = "收文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2012/2/20
Private Sub textCP24_GotFocus()
   InverseTextBox textCP24
End Sub

'Add By Sindy 2012/2/20
Private Sub textCP24_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP24) = False Then
      If textCP24 <> "1" And textCP24 <> "2" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件准駁只可輸入1或2"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP24_GotFocus
      End If
   End If
End Sub

Private Sub txtCaseProperty_Change()
   lblCasePropertyName = ""
End Sub

'Add by Morgan 2003/11/26
Private Function txtTM15Control() As Boolean
   If (txtCaseProperty = "102") And ((txtTFCode(0).Text = Empty And txtSystem.Text = "TF") Or _
      (txtCode(0).Text = Empty And (txtSystem.Text = "T" Or txtSystem.Text = "FCT"))) Then
      txtTM15Control = True
   Else
      txtTM15Control = False
   End If
End Function

Private Sub txtCaseProperty_Validate(Cancel As Boolean)
Dim strTemp As String
Dim bolIsChina As Boolean
Dim adocase As New ADODB.Recordset  '2010/3/31 ADD BY SONIA
Dim strMsg As String 'Add by Amy 2016/08/16
   
   'm_TM10 = "" 'Remove by Lydia 2020/06/08 統一用m_Nation
   
   'Add By Sindy 2012/3/1
   Me.Frame1.Visible = False
   Me.textCP05.Enabled = True
   '2012/3/1 End
   
   If txtCaseProperty.Visible = False Then
      Exit Sub
   End If
   If txtCaseProperty <> "" Then
      '94.1.12 ADD BY SONIA CFT新加坡的跨類107要收新案號
'edit by nickc 2007/07/27 新加坡跨類已經不用收新案號了
'    If Me.txtSystem.Text = "CFT" And Me.Caption = "接洽紀錄單－新增" And m_TM10 = "014" And Me.txtCaseProperty.Text = "107" And txtCode(0) <> "" Then
'            MsgBox "新加坡的跨類107要收新案號!!!", vbExclamation
'            Cancel = True
'            txtCode_GotFocus (0)
'            Me.txtCode(0).SetFocus
'            Exit Sub
'    End If
    '94.1.12 end
      'add by nick 2004/08/23 專利案件性質是306 或是 307 時要收新案
      '2010/3/10 modify by sonia FCP的306不必收新案號
      'If (Me.txtSystem.Text = "P" Or Me.txtSystem.Text = "FCP" Or Me.txtSystem.Text = "CFP") And Me.Caption = "接洽紀錄單－新增" And (Me.txtCaseProperty.Text = "306" Or Me.txtCaseProperty.Text = "307") And txtCode(0) <> "" Then
      'Modified by Lydia 2021/06/16 + 接洽單自動轉收文－新增
      'If (Me.txtSystem.Text = "P" Or Me.txtSystem.Text = "CFP") And Me.Caption = "接洽紀錄單－新增" And (Me.txtCaseProperty.Text = "306" Or Me.txtCaseProperty.Text = "307") And txtCode(0) <> "" Then
      'Modified by Lydia 2021/09/10 取消專利案件306改請獨立一定要收新案號的限制
      'If (Me.txtSystem.Text = "P" Or Me.txtSystem.Text = "CFP") And InStr("接洽紀錄單－新增,接洽單自動轉收文－新增", Me.Caption) > 0 And (Me.txtCaseProperty.Text = "306" Or Me.txtCaseProperty.Text = "307") And txtCode(0) <> "" Then
      If (Me.txtSystem.Text = "P" Or Me.txtSystem.Text = "CFP") And InStr("接洽紀錄單－新增,接洽單自動轉收文－新增", Me.Caption) > 0 And Me.txtCaseProperty.Text = "307" And txtCode(0) <> "" Then
            'Mark by Lydia 2021/09/10
            'If Me.txtCaseProperty.Text = "306" Then
            '    MsgBox "改請獨立必須收新本所案號!!!", vbExclamation
            'Else
            'end 2021/09/10
                If Me.txtCaseProperty.Text = "307" Then
                    MsgBox "分割必須收新本所案號!!!", vbExclamation
                End If
            'End If  'Mark by Lydia 2021/09/10
            Cancel = True
            txtCode_GotFocus (0)
            Me.txtCode(0).SetFocus
            Exit Sub
      End If
      If Me.txtSystem.Text = "FCP" And Me.Caption = "接洽紀錄單－新增" And Me.txtCaseProperty.Text = "307" And txtCode(0) <> "" Then
            MsgBox "分割必須收新本所案號!!!", vbExclamation
            Cancel = True
            txtCode_GotFocus (0)
            Me.txtCode(0).SetFocus
            Exit Sub
      End If
'cancel by sonia 2020/1/8
'      'add by nick 2004/08/23 P 收新案時，若是 904 要提示收 PS
'      If Me.txtSystem.Text = "P" And Me.Caption = "接洽紀錄單－新增" And Me.txtCaseProperty = "904" Then
'            MsgBox "調卷新案，請改收 PS !!!", vbExclamation
'            Cancel = True
'            txtSystem_GotFocus
'            Me.txtSystem.SetFocus
'            Exit Sub
'      End If
'end 2020/1/8
      '2005/5/24 ADD BY SONIA
      'Modified by Lydia 2021/06/16 + 接洽單自動轉收文－新增
      'If (Me.txtSystem.Text = "P" Or Me.txtSystem.Text = "FCP") And Me.Caption = "接洽紀錄單－新增" And (Me.txtCaseProperty.Text = "801" Or Me.txtCaseProperty.Text = "802") Then
      If (Me.txtSystem.Text = "P" Or Me.txtSystem.Text = "FCP") And InStr("接洽紀錄單－新增,接洽單自動轉收文－新增", Me.Caption) > 0 And (Me.txtCaseProperty.Text = "801" Or Me.txtCaseProperty.Text = "802") Then
            MsgBox "P 或 FCP 案已無 異議 或 異議答辯 程序!!!", vbExclamation
            Cancel = True
            txtCaseProperty_GotFocus
            Me.txtCaseProperty.SetFocus
            Exit Sub
       End If
       '2005/5/24 END
      'Add By Cheng 2002/01/08
      '若系統類別非"L", "CFL", "FCL", "LA"檢查案件性質必須為三碼
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      If Me.txtSystem.Text <> "L" And Me.txtSystem.Text <> "CFL" And _
         Me.txtSystem.Text <> "FCL" And Me.txtSystem.Text <> "LA" And _
         Me.txtSystem.Text <> "LIN" And Me.txtSystem.Text <> "ACS" Then
         If Len(Me.txtCaseProperty.Text) <> 3 Then
            MsgBox "案件性質必須為三碼!!!", vbExclamation
            Cancel = True
            txtCaseProperty_GotFocus
            Me.txtCaseProperty.SetFocus
            Exit Sub
         End If
      End If
      
      If CheckExist(0) = False Then
         If CheckEverythingOK = False Then
            Cancel = True
            txtCaseProperty_GotFocus
            Exit Sub
         End If
      Else
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(txtSystem, txtCaseProperty, strTemp, bolIsChina) Then
         If ClsPDGetCaseProperty(txtSystem, txtCaseProperty, strTemp, bolIsChina) Then
            'add by sonia 2013/8/12
            If (strTemp = "" Or strTemp = "（無）") And txtCode(0) = "" Then
               If ClsPDGetCaseProperty(txtSystem, txtCaseProperty, strTemp, True) Then
               End If
            End If
            'end 2013/8/12
         
            lblCasePropertyName = strTemp
            '92.2.19 ADD BY SONIA
            If txtCaseProperty = "001" And txtSystem <> "TS" And txtSystem <> "S" Then
               MsgBox "查名案件性質之系統類別只可為 TS(內商) 或 S(FCT,CFT)  !!!", vbExclamation
               Cancel = True
               txtCaseProperty_GotFocus
               Me.txtCaseProperty.SetFocus
            End If
            '92.2.19 END
         Else
            lblCasePropertyName = ""
            Cancel = True
            txtCaseProperty_GotFocus
         End If
      End If
      'Add by Amy 2016/08/16 +已收放棄專利權且已發文及已通知專利權公告作廢,不能再收文年費控制制
      If InStr(Me.txtSystem.Text, "P") > 0 And InStr(Me.Caption, "新增") > 0 And (Me.txtCaseProperty.Text = "605" Or Me.txtCaseProperty.Text = "606" Or Me.txtCaseProperty.Text = "607") Then
           'Add by Amy 2016/09/10 臺灣一案兩請新型案是否不可收文年費控制 P-108013
           If GetPA09(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) = "000" Then
               If PUB_TwDualCaseUtyNoAdd605(Me.txtSystem.Text, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), "000") = True Then
                   MsgBox "發明案有收文「擇一申復」且已公告，故不可再收文年費！", vbExclamation + vbOKOnly
                   Cancel = True
                   txtCaseProperty_GotFocus
                   Me.txtCaseProperty.SetFocus
                   Exit Sub
               End If
           End If
           strMsg = ChkP429Or1606(Me.txtSystem.Text, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), m_Nation)
           If strMsg <> MsgText(601) Then
               MsgBox "本案已「" & strMsg & "」無需再繳納年費！", vbExclamation + vbOKOnly
               Cancel = True
               txtCaseProperty_GotFocus
               Me.txtCaseProperty.SetFocus
               Exit Sub
           End If
      End If
      'end 2016/08/16
       
      '2010/3/31 ADD BY SONIA 大陸案檢查421或423
      adocase.CursorLocation = adUseClient
      'Modified by Lydia 2021/06/16 + 接洽單自動轉收文－新增
      'If Me.txtSystem.Text = "P" And Me.Caption = "接洽紀錄單－新增" And txtCode(0) <> "" And (Me.txtCaseProperty = "421" Or Me.txtCaseProperty = "423") Then
      If Me.txtSystem.Text = "P" And InStr("接洽紀錄單－新增,接洽單自動轉收文－新增", Me.Caption) > 0 And txtCode(0) <> "" And (Me.txtCaseProperty = "421" Or Me.txtCaseProperty = "423") Then
         strExc(0) = "select PA09,PA08,PA10 from patent where pa01 = '" & txtSystem & "' and pa02 = " & CNULL(txtCode(0)) & " and pa03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and pa04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
         intI = 1
         Set adocase = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If adocase.Fields(0) = "020" And Not IsNull(adocase.Fields(2)) Then
               If adocase.Fields(1) = "3" And Val(adocase.Fields(2)) < 20091001 Then
                  MsgBox "大陸外觀設計,申請日小於2009/10/1，不可收此案件性質 !!!", vbExclamation
                  Cancel = True
                  txtCaseProperty_GotFocus
                  Me.txtCaseProperty.SetFocus
                  adocase.Close
                  Exit Sub
               ElseIf Me.txtCaseProperty = "421" And (adocase.Fields(1) <> "2" Or Val(adocase.Fields(2)) >= 20091001) Then
                  MsgBox "大陸實用新型或外觀設計,申請日大於2009/10/1，請改收 423申請專利權評價報告 !!!", vbExclamation
                  Cancel = True
                  txtCaseProperty_GotFocus
                  Me.txtCaseProperty.SetFocus
                  adocase.Close
                  Exit Sub
               ElseIf Me.txtCaseProperty = "423" And adocase.Fields(1) = "2" And Val(adocase.Fields(2)) < 20091001 Then
                  MsgBox "大陸實用新型申請日小於2009/10/1, 請改收 421申請檢索報告 !!!", vbExclamation
                  Cancel = True
                  txtCaseProperty_GotFocus
                  Me.txtCaseProperty.SetFocus
                  adocase.Close
                  Exit Sub
               End If
            End If
         End If
         adocase.Close
      End If
      '2010/3/31 END
      
      ' 90.12.19 add by louis
      Cancel = UpdateCtrlState
      
      'Added by Lydia 2020/05/20 是否停駐於案源單號
      'Modified by Lydia 2020/12/16 外商臺灣案收文：可輸入案源
      'If txtCaseProperty.Tag <> txtCaseProperty.Text And mRole = "" Then '排除外專後續案收文
      'Modified by Lydia 2025/05/08 外專後續案收文可輸入案源 Left(mRole, 2) = "F1" => Left(mRole, 1) = "F"
      If txtCaseProperty.Tag <> txtCaseProperty.Text And (mRole = "" Or Left(mRole, 1) = "F") Then
           FraLOS.Visible = False
           m_bolStopOntxtLOS15 = False
           If Trim(txtFMP) = "" Then
               If GetStateLOS(txtSystem, txtCaseProperty, txtCode(0), txtLOS15, m_strLOSkind) = True Then
                   If m_strLOSkind <> "" Then
                       FraLOS.Visible = True
                       m_bolStopOntxtLOS15 = True
                   End If
               End If
           End If
           'Added by Lydia 2020/11/19 CFP和CFT英國脫歐案管制：收文CFP及CFT英國新案「延展費CFP.607」/「延展CFT.102」時，系統顯示歐盟案案號欄位(接洽單左上角)供輸入，若未輸入時提醒並確認。
           fraNA239.Visible = False
           txtCaseNa239 = ""
           m_bolStopOntxtCaseNa239 = False 'Added by Lydia 2020/12/04
           'Modified by Lydia 2020/12/01 + 「委任代理人(CFP.444, CFT.710)」時，系統顯示歐盟案案號欄位供輸入
           'If ((txtSystem = "CFP" And txtCaseProperty = "607") Or (txtSystem = "CFT" And txtCaseProperty = "102")) And txtCode(0) & txtCode(1) & txtCode(2) = "" Then
           If txtCode(0) & txtCode(1) & txtCode(2) = "" And ((txtSystem = "CFP" And (txtCaseProperty = "607" Or txtCaseProperty = "444")) _
                                                                                    Or (txtSystem = "CFT" And (txtCaseProperty = "102" Or txtCaseProperty = "710"))) Then
               fraNA239.Visible = True
               'Added by Lydia 2020/12/04 CFT脫歐案是否停駐於歐盟案案號
               If txtSystem = "CFT" Then
                   m_bolStopOntxtCaseNa239 = True
               End If
               'end 2020/12/04
           End If
           'end 2020/11/19
           'Added by Lydia 2021/03/05 CFT歐盟尚未註冊案轉換英國申請案收文控管：針對2021.9.30前收文之英國新「申請101」案建立關聯案
           If strSrvDate(1) <= "20210930" And txtSystem = "CFT" And txtCode(0) & txtCode(1) & txtCode(2) = "" And txtCaseProperty = "101" Then
               fraNA239.Visible = True
               m_bolStopOntxtCaseNa239 = True
           End If
           'end 2021/03/05
      End If
      txtCaseProperty.Tag = txtCaseProperty.Text
      'end 2020/05/20
      
      'Added by Sindy 2022/7/7
      If m_strIR01 <> "" And FraRecv.Visible = True Then '多案收文時
         If fraPatition.Visible = True Then
            frm010001.FraRecvList.Height = 1965
            frm010001.List1.Height = 1680
         Else
            frm010001.FraRecvList.Height = 3355
            frm010001.List1.Height = 3040
         End If
      End If
      '2022/7/7 END
   End If
End Sub

Private Sub txtCaseProperty_GotFocus()
   txtCaseProperty.SelStart = 0
   txtCaseProperty.SelLength = Len(txtCaseProperty)
End Sub

Private Sub txtCode_Change(Index As Integer)
   Call txtTM15Control
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
   txtCode(Index).SelStart = 0
   txtCode(Index).SelLength = Len(txtCode(Index).Text)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii = 45 Then KeyAscii = 0    '2009/10/14 ADD BY SONIA 不可輸入-
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   'Added by Lydia 2020/06/08 統一用m_Nation
   If txtCode(Index) <> "" Then
       m_Nation = "000" '預設台灣案
   End If
   'end 2020/06/08
   If Len(txtCode(Index)) = txtCode(Index).MaxLength Then
      If CheckExist(0) = False Then
         If CheckEverythingOK = False Then
            If Index = 0 Then
               SendKeys "{TAB}"
               SendKeys "{TAB}"
            Else
               Cancel = True
               txtCode_GotFocus Index
            End If
         End If
      End If
      ' 90.12.19 add by louis
      UpdateCtrlState
   ElseIf Len(txtCode(Index)) <> 0 Then
      ShowMsg MsgText(1017)
      Cancel = True
      txtCode_GotFocus Index
   End If
   
   'Added by Lydia 2024/02/20 提前到輸入本所案號檢查
   If pub_strUserOffice = "1" And mRole = "" And txtSystem = "P" And Index = 2 And Len(txtCode(0)) = 6 And txtCode(1) <> "" And bolChild013 = False Then
     If intReceiveKind = 0 And intModifyKind = 0 And intChoose = 0 Then
        If CheckCaseNo Then
           Cancel = True
           txtCode_GotFocus 0
           Exit Sub
        End If
        If CheckExist(1) = True Then
        End If
     End If
   End If
   'end 2024/02/20
End Sub

'Added by Lydia 2020/06/10
Private Sub txtFMP_Validate(Cancel As Boolean)
   If Trim(txtFMP) = "Y" Then
       FraLOS.Visible = False
       m_strLOSkind = ""
   End If
End Sub

Private Sub txtPetition_Change()
   lblPetitionName.Caption = ""
End Sub

Private Sub txtPetition_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPetition_Validate(Cancel As Boolean)
Dim strTemp As String, intLength As Integer, strPetition As String

   If Len(txtPetition) > 0 Then
      strPetition = txtPetition
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCustomer(strPetition, strTemp) Then
      If ClsPDGetCustomer(strPetition, strTemp) Then
      
         txtPetition = strPetition
         lblPetitionName = strTemp
      Else
         Cancel = True
         txtPetition_GotFocus
      End If
   End If
End Sub

Private Sub txtPetition_GotFocus()
   '由於使用Validate()，以致無法正確跳躍Focus，因此在txtPetition_GotFocus()及
   'cmdOK_GotFocus()加入判斷，以得正確之跳躍
   If txtPetition.Visible = False Then
      'Modify By Sindy 2022/7/7
      'cmdOK(0).SetFocus
      If cmdok(0).Visible = True Then
         cmdok(0).SetFocus
      Else
         cmdMRecv.SetFocus
      End If
      '2022/7/7 END
   Else
      m_bolStopOnTxtPetition = False 'Add by Morgan 2006/6/27
      txtPetition.SelStart = 0
      txtPetition.SelLength = Len(txtPetition)
   End If
End Sub

'Add by Morgan 2006/6/23
'移轉申請人2-5
Private Sub txtPetitionx_Change(Index As Integer)
   lblPetitionNamex(Index).Caption = ""
End Sub

Private Sub txtPetitionx_GotFocus(Index As Integer)
   If txtPetitionx(Index).Visible = True Then
      txtPetitionx(Index).SelStart = 0
      txtPetitionx(Index).SelLength = Len(txtPetitionx(Index))
   End If
End Sub

Private Sub txtPetitionx_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPetitionx_Validate(Index As Integer, Cancel As Boolean)
   Dim strTemp As String, strPetition As String

   If Len(txtPetitionx(Index)) > 0 Then
      strPetition = txtPetitionx(Index)
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCustomer(strPetition, strTemp) Then
      If ClsPDGetCustomer(strPetition, strTemp) Then
      
         txtPetitionx(Index) = strPetition
         lblPetitionNamex(Index) = strTemp
      Else
         Cancel = True
         txtPetitionx(Index).SetFocus
      End If
   End If
End Sub
'end 2006/6/23
Private Sub txtRecieveCode_GotFocus(Index As Integer)
   txtRecieveCode(Index).SelStart = 0
   txtRecieveCode(Index).SelLength = Len(txtRecieveCode(Index).Text)
End Sub
Private Sub txtSystem_Change()
   If txtSystem.Text = 馬德里案 Then
      fraTF.Visible = True
      fraElse.Visible = False
      '911111 nick
      txtTFCode(0).Enabled = True
      txtTFCode(1).Enabled = True
      txtTFCode(2).Enabled = True
      txtTFCode(3).Enabled = True
      txtCode(0).Enabled = False
      txtCode(1).Enabled = False
      txtCode(2).Enabled = False
      txtTFCode(0).Text = Empty
      txtTFCode(1).Text = Empty
      txtTFCode(2).Text = Empty
      txtTFCode(3).Text = Empty
      txtTFCode(0).TabIndex = 2
      txtTFCode(1).TabIndex = 3
      txtTFCode(2).TabIndex = 4
      txtTFCode(3).TabIndex = 5
      textCP05.TabIndex = 6
      '911017 nick
      txtCaseProperty.TabIndex = 7
   Else
      fraTF.Visible = False
      fraElse.Visible = True
      '911111 nick
      txtCode(0).Enabled = True
      txtCode(1).Enabled = True
      txtCode(2).Enabled = True
      txtTFCode(0).Enabled = False
      txtTFCode(1).Enabled = False
      txtTFCode(2).Enabled = False
      txtTFCode(3).Enabled = False
      txtCode(0).Text = Empty
      txtCode(1).Text = Empty
      txtCode(2).Text = Empty
      txtCode(0).TabIndex = 2
      txtCode(1).TabIndex = 3
      txtCode(2).TabIndex = 4
      'Added by Lydia 2018/05/07 顯示是否為FMP案
      If intChoose = 0 Then
            txtFMP.TabIndex = 5
            txtNA01.TabIndex = 6 'Added by Lydia 2021/11/10  後面的index + 1
            textCP05.TabIndex = 7
            txtCaseProperty.TabIndex = 8
      Else
      'end 2018/05/07
            textCP05.TabIndex = 5
            '911017 nick
            txtCaseProperty.TabIndex = 6
      End If
   
   End If
   'Add by Morgan 2003/11/24
   If (txtSystem.Text = "T" Or txtSystem.Text = "TF" Or txtSystem.Text = "FCT") Then
      Call txtTM15Control
   End If
   '---End
   
   'Added by Lydia 2018/05/07 顯示是否為FMP案
   'Modified by Lydia 2018/05/17 北所才顯示
   'If txtSystem = "P" Then
   If intChoose = 0 And txtSystem = "P" And pub_strUserOffice = "1" Then
       fraFMP.Visible = True 'Modified by Lydia 2021/11/10 改成Frame
   Else
       fraFMP.Visible = False 'Modified by Lydia 2021/11/10 改成Frame
   End If
   'end 2018/05/17
    
   'Added by Lydia 2020/07/22 因為先收CFT案所以國籍未變更
   If txtSystem.Tag <> txtSystem.Text Then
        m_Nation = "000" '預設
   End If
   'end 2020/07/22
   
   'Added by Lydia 2020/05/20 判斷是否為法律所案源收文
   If txtSystem.Tag <> "" And txtSystem.Tag <> txtSystem.Text And mRole = "" Then '排除外專後續案收文
       FraLOS.Visible = False
       txtLOS15 = ""
       If Trim(txtFMP) = "" Then
            If GetStateLOS(txtSystem, txtCaseProperty, txtCode(0), txtLOS15, m_strLOSkind) = True Then
                If m_strLOSkind <> "" Then
                    FraLOS.Visible = True
                End If
            End If
       End If
       'Added by Lydia 2020/11/19 CFP和CFT英國脫歐案管制：收文CFP及CFT英國新案「延展費CFP.607」/「延展CFT.102」時，系統顯示歐盟案案號欄位(接洽單左上角)供輸入，若未輸入時提醒並確認。
       fraNA239.Visible = False
       'Modified by Lydia 2020/12/01 + 「委任代理人(CFP.444, CFT.710)」時，系統顯示歐盟案案號欄位供輸入
       'If ((txtSystem = "CFP" And txtCaseProperty = "607") Or (txtSystem = "CFT" And txtCaseProperty = "102")) And txtCode(0) & txtCode(1) & txtCode(2) = "" Then
       If txtCode(0) & txtCode(1) & txtCode(2) = "" And ((txtSystem = "CFP" And (txtCaseProperty = "607" Or txtCaseProperty = "444")) _
                                                                                 Or (txtSystem = "CFT" And (txtCaseProperty = "102" Or txtCaseProperty = "710"))) Then
           fraNA239.Visible = True
           'Added by Lydia 2020/12/04 CFT脫歐案是否停駐於歐盟案案號
           If txtSystem = "CFT" Then
               m_bolStopOntxtCaseNa239 = True
           End If
           'end 2020/12/04
       End If
       'end 2020/11/19
   End If
   txtSystem.Tag = txtSystem.Text
   'end 2020/05/20
End Sub

Private Sub txtSystem_GotFocus()
   txtSystem.SelStart = 0
   txtSystem.SelLength = Len(txtSystem.Text)
   CloseIme
End Sub

Private Sub txtSystem_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSystem_LostFocus()
   If txtSystem.Text = 馬德里案 Then
      fraTF.Visible = True
      fraElse.Visible = False
      '911111 nick
      txtTFCode(0).Enabled = True
      txtTFCode(1).Enabled = True
      txtTFCode(2).Enabled = True
      txtTFCode(3).Enabled = True
      txtCode(0).Enabled = False
      txtCode(1).Enabled = False
      txtCode(2).Enabled = False
      txtTFCode(0).TabIndex = 2
      txtTFCode(1).TabIndex = 3
      txtTFCode(2).TabIndex = 4
      txtTFCode(3).TabIndex = 5
   
      '911017 nick
      txtCaseProperty.TabIndex = 11
   Else
      fraTF.Visible = False
      fraElse.Visible = True
      '911111 nick
      txtCode(0).Enabled = True
      txtCode(1).Enabled = True
      txtCode(2).Enabled = True
      txtTFCode(0).Enabled = False
      txtTFCode(1).Enabled = False
      txtTFCode(2).Enabled = False
      txtTFCode(3).Enabled = False
      txtCode(0).TabIndex = 2
      txtCode(1).TabIndex = 3
      txtCode(2).TabIndex = 4
      '911017 nick
      'Added by Lydia 2021/11/10 判斷FMP案
      If fraFMP.Visible = True Then
           txtFMP.TabIndex = 5
           txtNA01.TabIndex = 6
           textCP05.TabIndex = 7
           txtCaseProperty.TabIndex = 8
      Else
      'end 2021/11/10
          txtCaseProperty.TabIndex = 6
      End If 'Added by Lydia 2021/11/10
   End If
   
   'Add By Sindy 2012/2/23
   If txtSystem <> "" Then
      If GetSysTemKind = False Then
         txtSystem_GotFocus
         txtSystem.SetFocus
         Exit Sub
      End If
   End If
   
   'Added by Lydia 2021/09/03 從外商臺灣案收文進入，沒有帶入收文類別A; ex.FCT-046258的9/3註冊費收文號存成B0031663, 人工修改為AB0036995
   'Modified by Lydia 2021/10/14 發現有外商、外專自行收文的收文號少掉年份的兩碼; ex.FCT-047364的10/13收文號存成A042506,人工修改為AB0042506(共19筆)
   'If lblReciveCode.Caption = "" Then
   If Len(Trim(lblReciveCode.Caption & txtRecieveCode(0))) <> 3 Then
       'Modified by Lydia 2021/10/15 + 客戶提供文件處理 m_GetB202CP09
       'If Left(mRole, 1) = "F" And InStr(Me.Caption, "新增") > 0 Then
       If (Left(mRole, 1) = "F" Or Me.m_GetB202CP09 = "B") And InStr(Me.Caption, "新增") > 0 Then
          strReceiveKind = 接洽記錄單
          lblReciveCode.Caption = strReceiveKind
          'Added by Lydia 2021/10/14
          txtRecieveCode(0).Text = CompAutoNumberYear(GetTaiwanThisYear)
          If Len(Trim(lblReciveCode.Caption & txtRecieveCode(0))) <> 3 Then
              MsgBox "收文號有問題，請關閉收文畫面後，再重新進入！"
              Exit Sub
          End If
          'end 2021/10/14
       Else
           MsgBox "收文號有問題，請關閉收文畫面後，再重新進入！"
           Exit Sub
       End If
   End If
   'end 2021/09/03
End Sub

Private Sub txtSystem_Validate(Cancel As Boolean)
   ' 91.09.16 modify by louis (變更Flag值)
   'If objPublicData.GetSystemKind(txtSystem.Text, intCaseKind, strCaseName) = False Then
   '   Cancel = True
   '   txtSystem_GotFocus
   'Else
   '   CheckEverythingOK
   'End If
    'Modify By Cheng 2003/03/28
    '若系統類別欄有顯示才要檢查系統類別
    If Me.txtSystem.Visible = True Then
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetSystemKind(txtSystem.Text, intCaseKind, strCaseName) = False Then
        If ClsPDGetSystemKind(txtSystem.Text, intCaseKind, strCaseName) = False Then
           Cancel = True
           txtSystem_GotFocus
        End If
    End If
End Sub

Private Sub txtTFCode_GotFocus(Index As Integer)
   txtTFCode(Index).SelStart = 0
   txtTFCode(Index).SelLength = Len(txtTFCode(Index).Text)
End Sub

Private Sub txtTFCode_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then
      If Len(txtTFCode(Index)) = txtTFCode(Index).MaxLength Then
         If CheckEverythingOK = False Then
            Cancel = True
            txtTFCode_GotFocus Index
         End If
      ElseIf Len(txtTFCode(Index)) <> 0 Then
         ShowMsg MsgText(1017)
         Cancel = True
         txtTFCode_GotFocus Index
      End If
   End If
End Sub

Private Function CheckKeyInOkay() As Boolean
   'TF為馬德里案，另外判斷
   If txtSystem = 馬德里案 Then
      'edit by nickc 2007/02/06 不用 dll 了
      'If obj001.CheckTFTextOkay(txtSystem.Text, intSaveMode, txtTFCode(0), txtTFCode(1), txtTFCode(2), txtTFCode(3)) Then
      If Cls001CheckTFTextOkay(txtSystem.Text, intSaveMode, txtTFCode(0), txtTFCode(1), txtTFCode(2), txtTFCode(3)) Then
         CheckKeyInOkay = True
      End If
   Else
      'edit by nickc 2007/02/06 不用 dll 了
      'If obj001.CheckTextOkay(intCaseKind, txtSystem.Text, intSaveMode, txtCode(0), txtCode(1), txtCode(2)) Then
      If Cls001CheckTextOkay(intCaseKind, txtSystem.Text, intSaveMode, txtCode(0), txtCode(1), txtCode(2)) Then
         CheckKeyInOkay = True
      End If
   End If
End Function

Private Function CheckEverythingOK() As Boolean
Dim strTemp As String, bolIsChina As Boolean
'Add By Cheng 2003/08/28
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

If Len(txtCaseProperty) > 0 Then
   If CheckKeyInOkay Then
      If intCaseKind = 顧問 And intSaveMode = 1 And txtCaseProperty <> 顧問聘任 Then
         ShowMsg MsgText(1018)
         Exit Function
      End If
      If intSaveMode = 0 Then
         'If intCaseKind <> 顧問 And intCaseKind <> 法務 Then '2013/8/22 cancel by sonia(LA-003178舊案)
            If txtSystem = 馬德里案 Then
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetCaseNation(intCaseKind, txtSystem, txtTFCode(0) + txtTFCode(1), IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strTemp) = False Then
               If ClsPDGetCaseNation(intCaseKind, txtSystem, txtTFCode(0) + txtTFCode(1), IIf(txtTFCode(2) = "", "0", txtTFCode(2)), IIf(txtTFCode(3) = "", "00", txtTFCode(3)), strTemp) = False Then
                  Exit Function
               End If
            Else
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetCaseNation(intCaseKind, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strTemp) = False Then
               If ClsPDGetCaseNation(intCaseKind, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strTemp) = False Then
                  Exit Function
               End If
            End If
         'End If    '2013/8/22 cancel by sonia
         If strTemp <> 台灣國家代號 Then bolIsChina = True Else bolIsChina = False
      End If
        'Add By Cheng 2003/08/28
        If (Me.txtSystem.Text = "P" Or Me.txtSystem.Text = "CFP" Or Me.txtSystem.Text = "FCP") And Me.txtCaseProperty.Text = "601" Then
            '是否有A類未取消收文的領證(601)
            StrSQLa = "Select Count(*) From CaseProgress Where " & ChgCaseprogress(Me.txtSystem.Text & Me.txtCode(0).Text & Me.txtCode(1).Text & Me.txtCode(2).Text) & " And CP09<'B' And CP10='601' And CP57 Is Null "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               '已有領證
               If rsA.Fields(0).Value > 0 Then
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  '是否有C類被異議(理由)(1801)
                  StrSQLa = "Select Count(*) From CaseProgress Where " & ChgCaseprogress(Me.txtSystem.Text & Me.txtCode(0).Text & Me.txtCode(1).Text & Me.txtCode(2).Text) & " And CP09>'C' And CP10='1801' "
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     If rsA.Fields(0).Value <= 0 Then
                        MsgBox "本案不可再收<領證>!!!", vbExclamation + vbOKOnly
                        Me.txtCaseProperty.SetFocus
                        txtCaseProperty_GotFocus
                        CheckEverythingOK = False
                        Exit Function
                     End If
                  End If
               '未領證
               Else
                  'Add By Sindy 2015/4/15 ex.P-109425
                  If Me.txtSystem.Text = "P" Then
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                     '檢查是否已有公告號或專利號數
                     StrSQLa = "Select pa15,pa22 From Patent Where " & ChgPatent(Me.txtSystem.Text & Me.txtCode(0).Text & Me.txtCode(1).Text & Me.txtCode(2).Text)
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsA.RecordCount > 0 Then
                        If ("" & rsA.Fields("pa15")) <> "" Or ("" & rsA.Fields("pa22")) <> "" Then
                           MsgBox "此本所案號已有公告號或專利號數, 不可收文領證！", vbExclamation + vbOKOnly
                           Me.txtCaseProperty.SetFocus
                           txtCaseProperty_GotFocus
                           CheckEverythingOK = False
                           Exit Function
                        End If
                     End If
                  End If
               End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCaseProperty(txtSystem, txtCaseProperty, strTemp, bolIsChina) Then
      If ClsPDGetCaseProperty(txtSystem, txtCaseProperty, strTemp, bolIsChina) Then
         'add by sonia 2013/8/12
         If (strTemp = "" Or strTemp = "（無）") And txtCode(0) = "" Then
            If ClsPDGetCaseProperty(txtSystem, txtCaseProperty, strTemp, True) Then
            End If
         End If
         'end 2013/8/12
         
         lblCasePropertyName = strTemp
         '92.9.15 MODIFY BY SONIA
         'If (((txtCaseProperty = 申請 Or txtCaseProperty = 異議 Or txtCaseProperty = 評定 Or txtCaseProperty = 廢止) And intCaseKind Mod 4 = 商標) Or _
         '      ((txtCaseProperty = 發明申請 Or txtCaseProperty = 新型申請 Or txtCaseProperty = 設計申請 Or txtCaseProperty = 追加申請 Or txtCaseProperty = 聯合申請 Or txtCaseProperty = 異議_專 Or txtCaseProperty = 舉發) _
         '      And intCaseKind Mod 4 = 專利)) And intSaveMode <> 1 Then
         'edit by nick 開放聯合申請的 cfp 可收舊案 2004/07/05
         'If (((txtCaseProperty = 申請 Or txtCaseProperty = 異議 Or txtCaseProperty = 評定 Or txtCaseProperty = 廢止) And intCaseKind Mod 4 = 商標) Or _
               ((txtCaseProperty = 發明申請 Or txtCaseProperty = 新型申請 Or txtCaseProperty = 設計申請 Or txtCaseProperty = 追加申請 Or txtCaseProperty = 聯合申請 Or txtCaseProperty = 異議_專 Or txtCaseProperty = 舉發 Or txtCaseProperty = PCT申請 Or txtCaseProperty = 記錄請求_標準專利 Or txtCaseProperty = 短期專利申請 Or txtCaseProperty = CIP申請 Or txtCaseProperty = CPA申請 Or txtCaseProperty = 再發行) _
               And intCaseKind Mod 4 = 專利)) And intSaveMode <> 1 Then
         '2007/5/31 加 "618"註冊不當撤銷 by sonia
         'Modify by Amy 2016/09/01 取消 T/CFT/FCT 601異議/603評定/605廢止/618註冊不當撤銷 須收新案
         'If (((txtCaseProperty = 申請 Or txtCaseProperty = 異議 Or txtCaseProperty = 評定 Or txtCaseProperty = 廢止 Or txtCaseProperty = "618") And intCaseKind Mod 4 = 商標) Or _
               ((txtCaseProperty = 發明申請 Or txtCaseProperty = 新型申請 Or txtCaseProperty = 設計申請 Or txtCaseProperty = 追加申請 Or txtCaseProperty = 異議_專 Or txtCaseProperty = 舉發 Or txtCaseProperty = PCT申請 Or txtCaseProperty = 記錄請求_標準專利 Or txtCaseProperty = 短期專利申請 Or txtCaseProperty = CIP申請 Or txtCaseProperty = CPA申請 Or txtCaseProperty = 再發行) _
               And intCaseKind Mod 4 = 專利)) And intSaveMode <> 1 Then
        'Modify by Amy 2022/03/08 2016/09/01未取消控制前,若輸某CFT舊案號,性質603 應彈訊息收新案;1081023與外商陳金蓮確認後FC/CF仍維持原控制 -->當時只改接洽單(請作單號1050921-02) ex:CFT-022838;故CFT/FCT 改回2016/09/01未取消前的控制
'         If ((txtCaseProperty = 申請 And intCaseKind Mod 4 = 商標) Or _
'               ((txtCaseProperty = 發明申請 Or txtCaseProperty = 新型申請 Or txtCaseProperty = 設計申請 Or txtCaseProperty = 追加申請 Or txtCaseProperty = 異議_專 Or txtCaseProperty = 舉發 Or txtCaseProperty = PCT申請 Or txtCaseProperty = 記錄請求_標準專利 Or txtCaseProperty = 短期專利申請 Or txtCaseProperty = CIP申請 Or txtCaseProperty = CPA申請 Or txtCaseProperty = 再發行) _
'               And intCaseKind Mod 4 = 專利)) And intSaveMode <> 1 Then
        If ((txtCaseProperty = 申請 And intCaseKind Mod 4 = 商標) Or _
               ((txtCaseProperty = 異議 Or txtCaseProperty = 評定 Or txtCaseProperty = 廢止 Or txtCaseProperty = "618") And (Me.txtSystem.Text = "CFT" Or Me.txtSystem.Text = "FCT")) Or _
               ((txtCaseProperty = 發明申請 Or txtCaseProperty = 新型申請 Or txtCaseProperty = 設計申請 Or txtCaseProperty = 追加申請 Or txtCaseProperty = 異議_專 Or txtCaseProperty = 舉發 Or txtCaseProperty = PCT申請 Or txtCaseProperty = 記錄請求_標準專利 Or txtCaseProperty = 短期專利申請 Or txtCaseProperty = CIP申請 Or txtCaseProperty = CPA申請 Or txtCaseProperty = 再發行) _
               And intCaseKind Mod 4 = 專利)) And intSaveMode <> 1 Then
         '92.9.15 END
            ' 91.09.03 marked by louis
            'ShowMsg MsgText(1019)
            'txtCaseProperty.SetFocus
            'CheckEverythingOK = False
            ' 91.09.03 modify by louis
            If textCP05 <> "111111" Then
                'Added by Lydia 2021/09/07 修改若舊案只有903專利調查、426新穎性調查時，可以開放收文新申請案的案件性質。
                strExc(1) = "N"
                If (txtSystem = "CFP" Or txtSystem = "P") Then
                    strExc(0) = "select count(*) cnt from caseprogress where cp01='" & txtSystem & "' and cp02='" & txtCode(0) & "' and cp03='" & Left(txtCode(1) & "0", 1) & "' and cp04='" & Left(txtCode(2) & "00", 2) & "' " & _
                                      "and substr(cp09,1,1)='A' and cp10 not in ('903','426') and cp159=0"
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    If intI = 1 Then
                        If Val("" & RsTemp.Fields("cnt")) = 0 Then
                            strExc(1) = "Y"
                        End If
                    End If
                End If
                'end 2021/09/07
               If strExc(1) <> "Y" Then 'Added by Lydia 2021/09/07 判斷舊案只有903專利調查、426新穎性調查
                   ShowMsg MsgText(1019)
                   txtCaseProperty.SetFocus
                   CheckEverythingOK = False
               'Added by Lydia 2021/09/07
               Else
                   CheckEverythingOK = True
               End If
               'end 2021/09/07
            Else
               CheckEverythingOK = True
            End If
         ElseIf txtCaseProperty = 移轉 And intCaseKind Mod 4 = 商標 Then
            fraPatition.Visible = True
            CheckEverythingOK = True
         ElseIf txtCaseProperty = 讓與 And intCaseKind Mod 4 = 專利 Then
            fraPatition.Visible = True
            CheckEverythingOK = True
         'Add By Cheng 2002/01/11
         ElseIf txtCaseProperty = 專利權讓與 And intCaseKind Mod 4 = 專利 Then
            If Me.txtSystem.Text = "P" Then
               fraPatition.Visible = True
               CheckEverythingOK = True
            End If
         'add by nick 2004/07/05
         'Modify by Morgan 2004/8/13 需排除新案號
         'ElseIf txtCaseProperty = 聯合申請 And intCaseKind Mod 4 = 專利 Then
         ElseIf txtCaseProperty = 聯合申請 And intCaseKind Mod 4 = 專利 And intSaveMode <> 1 Then
            '2012/6/18 MODIFY BY SONIA 加入P
            'If Me.txtSystem.Text = "CFP" Then
            If Me.txtSystem.Text = "CFP" Or Me.txtSystem.Text = "P" Then
               fraPatition.Visible = True
               CheckEverythingOK = True
            Else
                  If textCP05 <> "111111" Then
                     ShowMsg MsgText(1019)
                     txtCaseProperty.SetFocus
                     CheckEverythingOK = False
                  Else
                     CheckEverythingOK = True
                  End If
            End If
         'add end
         Else
            'Added by Lydia 2022/06/30 內部收文之假收文控制：專利案件性質401變更或商標301時先彈訊息詢問「是否變更申請人？」若選擇是且收文日也是111111時，也做上述更新基本檔之申請人及申請地址。
            If intChoose = 1 And textCP05.Visible = True And bolChkChange = True And ((InStr("P,CFP,FCP,", txtSystem & ",") > 0 And txtCaseProperty = "401") Or _
                        (InStr("T,TF,CFT,FCT,TB,TC,", txtSystem & ",") > 0 And txtCaseProperty = "301")) Then
                If textCP05 <> "111111" Then
                    fraPatition.Visible = False
                End If
                CheckEverythingOK = True
            Else
            'end 2022/0630
                fraPatition.Visible = False
                CheckEverythingOK = True
            End If 'Added by Lydia 2022/06/30
         End If
      'Add By Sindy 2012/2/24
      Else
         txtCaseProperty = ""
         Me.Frame1.Visible = False
         textCP24 = ""
      '2012/2/24 End
      End If
   End If
Else
   lblCasePropertyName = ""
   fraPatition.Visible = False
   CheckEverythingOK = True
End If
End Function

'檢查本所案號是否存在
Public Function CheckExist(intShow As Integer) As Boolean
Dim adocase As New ADODB.Recordset
  
   adocase.CursorLocation = adUseClient
   Select Case txtSystem
      Case "P", "CFP", "FCP"
         'Modify by Amy 2016/08/16 +pa09
         strExc(0) = "select pa01,pa09 from patent where pa01 = '" & txtSystem & "' and pa02 = " & CNULL(txtCode(0)) & " and pa03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and pa04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
      Case "T", "CFT", "FCT"
         '94.1.12 modify by sonia
         'strExc(0) = "select tm01 from trademark where tm01 = '" & txtSystem & "' and tm02 = '" & txtCode(0) & "' and tm03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and tm04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
         strExc(0) = "select tm01,tm10 from trademark where tm01 = '" & txtSystem & "' and tm02 = '" & txtCode(0) & "' and tm03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and tm04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
         '94.1.12 end
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/24 +ACS系統類
      Case "L", "CFL", "FCL", "LIN", "ACS"
         strExc(0) = "select lc01,lc15 from lawcase where lc01 = '" & txtSystem & "' and lc02 = '" & txtCode(0) & "' and lc03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and lc04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
      Case "LA"
         strExc(0) = "select hc01,'000' from hirecase where hc01 = '" & txtSystem & "' and hc02 = '" & txtCode(0) & "' and hc03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and hc04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
      Case "TF"
         strExc(0) = "select tm01,tm10 from trademark where tm01 = '" & txtSystem & "' and tm02 = '" & txtTFCode(0) & txtTFCode(1) & "' and tm03 = '" & IIf(txtTFCode(2) = "", "0", txtTFCode(2)) & "' and tm04 = '" & IIf(txtTFCode(3) = "", "00", txtTFCode(3)) & "'"
      Case Else
         strExc(0) = "select sp01,sp09 from servicepractice where sp01 = '" & txtSystem & "' and sp02 = '" & txtCode(0) & "' and sp03 = '" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and sp04 = '" & IIf(txtCode(2) = "", "00", txtCode(2)) & "'"
   End Select
   intI = 1
   Set adocase = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 1 Then
      If intShow = 1 Then
         If MsgBox(MsgText(36), vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
            CheckExist = True
         Else
            CheckExist = False
            intSaveMode = 1
            'Added by Lydia 2024/02/20 有關" 香港013專利開放收文集體設計申請105"，請比照" 香港013專利收文設計申請103"
            If txtSystem = "P" And pub_strUserOffice = "1" And mRole = "" Then
               strExc(0) = "select pa01,pa08,pa09 from patent where pa01 = '" & txtSystem & "' and pa02 = " & CNULL(txtCode(0)) & " and pa03='0' and pa04='00' "
               intI = 1
               Set adocase = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If "" & adocase.Fields("pa08") = "3" And "" & adocase.Fields("pa09") = "013" Then
                     bolChild013 = True
                     fraFMP.Visible = True
                     CheckExist = True
                  End If
               End If
            End If
            'end 2024/02/20
         End If
      Else
         CheckExist = True
      End If
   Else
      bolChild013 = False 'Added by Lydia 2024/02/20
      CheckExist = False
      '94.1.12 add by sonia
      'Modify by Amy 2016/08/16 +申請國家
      'Remove by Lydia 2020/06/08 統一用m_Nation
      'If InStr(txtSystem, "P") > 0 Then
        m_Nation = "" & adocase.Fields(1)
      'ElseIf (txtSystem = "CFT" Or txtSystem = "FCT" Or txtSystem = "T") Then
      '   m_TM10 = adocase.Fields(1)
      'End If
      ''94.1.12 end
      'end 2020/06/08
      
   End If
   adocase.Close
End Function

'檢查本所案號是否大於目前流水號
Public Function CheckCaseNo() As Boolean
Dim adocase As New ADODB.Recordset
Dim strSql As String
    strSql = "select au03 from autonumber where au01 = '" & txtSystem & "'"
    intI = 1
    'edit by nickc 2007/02/05 不用 dll 了
    'Set RsTemp = objLawDll.ReadRstMsg(intI, strSQL)
    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
    If intI <> 1 Then
'   adocase.CursorLocation = adUseClient
   'adocase.Open "select au03 from autonumber where au01 = '" & txtSystem & "'", cnnConnection, adOpenStatic, adLockReadOnly
'   adocase.Open strSQL, cnnConnection, adOpenDynamic, adLockReadOnly
'   If adocase.RecordCount = 0 Then
         ShowMsg MsgText(37)
      CheckCaseNo = True
   Else
      'Added by Lydia 2020/03/23 排除LA-999999和TT-999999
      'modify by sonia 2021/2/8 再排除L-999999
      If (Me.txtSystem = "LA" Or Me.txtSystem = "TT" Or Me.txtSystem = "L") And Me.txtCode(0) = "999999" Then
           CheckCaseNo = False
      'add by sonia 2022/9/22 再排除L-888888
      ElseIf Me.txtSystem = "L" And Me.txtCode(0) = "888888" Then
           CheckCaseNo = False
      'end 2022/9/22
      Else
      'end 2020/03/23
            If Val(txtCode(0)) > RsTemp.Fields(0).Value Then
               '91.12.22 MODIFY BY SONIA
               'ShowMsg MsgText(37)
               'CheckCaseNo = True
               If txtSystem = "FCT" And txtCode(1) = "T" Then
                  CheckCaseNo = False
               Else
                  ShowMsg MsgText(37)
                  CheckCaseNo = True
               End If
               '91.12.22 END
            Else
               CheckCaseNo = False
            End If
      End If 'Added by Lydia 2020/03/23
   End If
'   adocase.Close
End Function

' 90.12.19 add by louis
Private Function UpdateCtrlState() As Boolean
Dim bShow As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2012/2/23
Dim pShow As Boolean 'Added by Lydia 2022/06/30

   UpdateCtrlState = False
   bShow = False
   
   'Add By Sindy 2009/07/06
   textYear.Visible = False
   Text1(0).Visible = False
   Label11(1).Visible = False
   '2009/07/06 End
   
   'Modify By Sindy 2012/2/23
   If ((strSK02 = "2" And strSK03 = "0") Or (strSK02 = "6" And strSK03 = "0")) And _
      txtCaseProperty = "728" Then
      m_CP01 = txtSystem
      m_CP02 = IIf(txtSystem <> "TF", txtCode(0), txtTFCode(0) & txtTFCode(1))
      m_CP03 = IIf(txtSystem <> "TF", txtCode(1), txtTFCode(2))
      m_CP03 = m_CP03 & String(1 - Len(m_CP03), "0")
      m_CP04 = IIf(txtSystem <> "TF", txtCode(2), txtTFCode(3))
      m_CP04 = m_CP04 & String(2 - Len(m_CP04), "0")
      
      Select Case m_CP01
         Case "T", "TF", "CFT", "FCT":
            strSql = "SELECT tm29 FROM trademark " & _
                     "WHERE tm01='" & m_CP01 & "' and tm02='" & m_CP02 & "' and tm03='" & m_CP03 & "' and tm04='" & m_CP04 & "' "
         Case Else:
            strSql = "SELECT sp15 FROM servicepractice " & _
                     "WHERE sp01='" & m_CP01 & "' and sp02='" & m_CP02 & "' and sp03='" & m_CP03 & "' and sp04='" & m_CP04 & "' "
      End Select
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If rsTmp.Fields(0) = "Y" Then
            MsgBox "此案件已閉卷!!!", vbExclamation
            UpdateCtrlState = True
            Call txtCode_GotFocus(0)
            txtCode(0).SetFocus
            rsTmp.Close
            Set rsTmp = Nothing
            Exit Function
         End If
      End If
      rsTmp.Close
      
      'Add By Sindy 2012/3/5
      strSql = "SELECT count(*) FROM caseprogress " & _
               "WHERE cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "' and cp10='728' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If rsTmp.Fields(0) > 0 Then
            MsgBox "此案件已轉案至他所!!!", vbExclamation
            UpdateCtrlState = True
            Call txtCode_GotFocus(0)
            txtCode(0).SetFocus
            rsTmp.Close
            Set rsTmp = Nothing
            Exit Function
         End If
      End If
      rsTmp.Close
      '2012/3/5 End
      
      Set rsTmp = Nothing
      Me.Frame1.Visible = True
      Me.textCP05.Text = ""
      Me.textCP05.Enabled = False
      Me.textCP24.SetFocus
   Else
      Me.Frame1.Visible = False
      Me.textCP05.Enabled = True
      textCP24 = ""
   '2012/2/23 End
      Select Case txtSystem
         Case "T", "TF", "CFT", "FCT", "TB", "TC"
            Select Case txtCaseProperty
               Case "501":
                  bShow = True
               Case Else:
            End Select
         Case "P", "CFP", "FCP"
            Select Case txtCaseProperty
               Case "701":
                  bShow = True
                  Me.Label5.Caption = "移轉、讓與申請人："
               'Add By Cheng 2002/01/11
               Case 專利權讓與:
                  If Me.txtSystem.Text = "P" Then
                     bShow = True
                     Me.Label5.Caption = "專利權讓與申請人："
                  End If
               'Add By Cheng 2002/01/14
               Case 合併
                  bShow = True
                  Me.Label5.Caption = "合併申請人："
               Case 繼承
                  If Me.txtSystem.Text = "FCP" Then
                     bShow = True
                     Me.Label5.Caption = "繼承申請人："
                  End If
               'Add By Sindy 2009/07/06
               Case "601", "605", "606", "607": '601.領證 605.年費 606.維持費 607.延展費
                  'If Trim(txtCode(0)) <> "" Then 'Modify By Sindy 2010/8/3 舊案才需顯示繳費起迄年度
                     If txtSystem = "P" Or (txtSystem = "CFP" And txtCaseProperty = "605") Or _
                                                      (txtSystem = "CFP" And txtCaseProperty = "606") Or _
                                                      (txtSystem = "CFP" And txtCaseProperty = "607") Then
                        textYear.Visible = True: textYear.Enabled = True
                        Text1(0).Visible = True
                        Label11(1).Visible = True
                        If txtCaseProperty = "601" Or txtCaseProperty = "605" Then
                           Label11(1).Caption = "繳費年度：第            年至第            年"
                           '2010/11/18 modify by sonia 大陸案不鎖領證起始年度
                           'If txtCaseProperty = "601" Then textYear = "1": textYear.Enabled = False
                           If txtCaseProperty = "601" And GetPA09(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) <> "020" Then
                              textYear = "1": textYear.Enabled = False
                           End If
                           '2010/11/18 end
                        Else
                           Label11(1).Caption = "繳費次數：第            次至第            次"
                        End If
                        'Added by Lydia 2023/04/25 外專後續案收文：FMP(非寰華案)領證和年費收文，若為舊案未閉卷，控管一定要輸入年度；預設下一繳費年度。
                        If Left(mRole, 2) = "F2" And txtSystem = "P" And (txtCaseProperty = "605" Or txtCaseProperty = "601") And txtCode(0) <> "" Then
                           If PUB_ChkIsFMP(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = True Then
                              'Mark by Lydia 2023/07/07 FMP案收文領證及繳年費(601)以及年費(605)，由系統直接帶入須繳納年費之年度，起迄年度相同
                              'If PUB_FMPtoCheck(1, 2, Pub_strUserST05, txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2))) = False Then
                                 If txtCaseProperty = "601" Then
                                     textYear = "1"
                                     'Added by Lydia 2023/05/24 P大陸案需要另外讀取領證及繳年費的起迄預設值; ex.P-124130(茹曣:電話聯絡)
                                     If GetPA09(Me.txtSystem.Text, Me.txtCode(0).Text, Me.txtCode(1).Text, Me.txtCode(2).Text) = "020" Then
                                         strSql = "SELECT CP53,CP54 FROM NEXTPROGRESS,CASEPROGRESS WHERE NP02='" & txtSystem & "' AND NP03='" & txtCode(0) & "' AND NP04='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' AND NP05='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' " & _
                                                      "and (NP06 is null OR NP06='N') " & strNpSqlOfNoSalesDuty & _
                                                      " AND NP01=CP09 AND CP53||CP54 IS NOT NULL "
                                         intI = 1
                                         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                                         If intI = 1 Then
                                            textYear = "" & RsTemp.Fields("CP53")
                                            'Modified by Lydia 2023/10/13 因為P127762的領證自動帶入核准的第7年-第8年，經過Elaine和秀玲討論決定只帶入起始年度（起迄年度相同）; 參考1120710-01公告，當時需求也是起迄年度一致
                                            'Text1(0) = "" & RsTemp.Fields("CP54")
                                            Text1(0) = "" & RsTemp.Fields("CP53")
                                         End If
                                     End If
                                     'end 2023/05/24
                                 Else
                                     '取得下次繳費次數/年度
                                     strExc(0) = PUB_Getnexttimes(txtSystem, txtCode(0), IIf(txtCode(1) = "", "0", txtCode(1)), IIf(txtCode(2) = "", "00", txtCode(2)), strExc(1))
                                     If strExc(0) <> "" Then
                                        textYear = strExc(1)
                                        Text1(0) = strExc(1)  'Added by Lydia 2023/07/07 起迄年度相同
                                     End If
                                 End If
                              'End If 'Mark by Lydia 2023/07/07
                           End If
                        End If
                        'end 2023/04/25
                        m_bolStopOntextYear = True
                     End If
                  'End If
               '2009/07/06 End
               Case Else:
            End Select
      End Select
   End If
   
   'Added by Lydia 2022/06/30 內部收文之假收文控制：專利案件性質401變更或商標301時先彈訊息詢問「是否變更申請人？」若選擇是且收文日也是111111時，也做上述更新基本檔之申請人及申請地址。
   bolChkChange = False
   pShow = False
   If intChoose = 1 And textCP05.Visible = True And bShow = False And textCP05 = "111111" Then  '判斷假收文日，避免重複詢問
      If (InStr("P,CFP,FCP,", txtSystem & ",") > 0 And txtCaseProperty = "401") Or _
           (InStr("T,TF,CFT,FCT,TB,TC,", txtSystem & ",") > 0 And txtCaseProperty = "301") Then
          If fraPatition.Visible = True Then
              bShow = True
              bolChkChange = True
          Else
             pShow = True
             If MsgBox("是否變更申請人？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                 bShow = True
                 Me.Label5.Caption = "變更後申請人："
                 bolChkChange = True
             Else
                 txtPetition = "": lblPetitionName = ""
                 For intI = 2 To 5
                    txtPetitionx(intI) = Empty
                    lblPetitionNamex(intI) = Empty
                 Next
             End If
          End If
      End If
   End If
   'end 2022/06/30
   
   If bShow Then
      fraPatition.Visible = True
      txtPetition.TabStop = True
   Else
      fraPatition.Visible = False
      txtPetition.TabStop = False
      '911112 nick
      txtPetition.Text = ""
   End If
   'Add by Morgan 2006/6/27
   m_bolStopOnTxtPetition = bShow
   'Added by Lydia 2022/06/30
   If pShow = True Then
       m_bolStopOnTxtPetition = pShow
   End If
End Function

' 新增案件進度資料
'edit by nickc 2007/01/11
Private Function OnSaveNewData() As Boolean
    'add by nickc 2007/01/11
    OnSaveNewData = False
    
    Dim strSql As String
    ' 本所案號
    Dim strCP01 As String
    Dim strCP02 As String
    Dim strCP03 As String
    Dim strCP04 As String
    ' 收文日
    Dim strCP05 As String
    ' 總收文號
    Dim strCP09 As String
    ' 案件性質
    Dim strCP10 As String
    ' 案件來源代號 (固定90)
    Dim strCP11 As String
    ' 業務區別
    Dim strCP12 As String
    ' 智權人員代號
    Dim strCP13 As String
    ' 承辦人代號
    Dim strCP14 As String
    ' 91.11.10 ADD BY SONIA
    Dim strCP20 As String
    Dim strCP26 As String
    Dim strCP32 As String
    '91.11.10 END
    ' 發文日
    Dim strCP27 As String
    '
    Dim strCP56 As String
    'Add by Morgan 2006/6/23
    Dim strCP(89 To 92) As String
    
    Dim strSalesNo As String '上個接洽記錄單的智權人員
    Dim StrSQLa As String
    Dim rsA As New ADODB.Recordset
    '2005/7/19 ADD BY SONIA
    ' CF代理人
    Dim strCP44 As String
    ' 彼所案號
    Dim strCP45 As String
    
    'add by nickc 2007/01/11
    Dim strCP55 As String
    Dim strCP93 As String
    Dim strCP94 As String
    Dim strCP95 As String
    Dim strCP96 As String
    Dim StrSqlB As String

    strCP01 = txtSystem
    'Modify by Morgan 2004/5/10
    '加判斷 "TF" 時抓不同欄位
    If txtSystem = "TF" Then
      '2009/1/13 MODIFY BY SONIA
      'strCP02 = txtTFCode(0)
      'strCP03 = Right("0" & txtTFCode(1), 1)
      'strCP04 = Right("0" & txtTFCode(2), 1) & Right("0" & txtTFCode(3), 1)
      strCP02 = txtTFCode(0) & txtTFCode(1)
      strCP03 = txtTFCode(2)
      strCP03 = strCP03 & String(1 - Len(strCP03), "0")
      strCP04 = txtTFCode(3)
      strCP04 = strCP04 & String(2 - Len(strCP04), "0")
      '2009/1/13 END
    Else
      strCP02 = txtCode(0)
      strCP03 = txtCode(1)
      strCP03 = strCP03 & String(1 - Len(strCP03), "0")
      strCP04 = txtCode(2)
      strCP04 = strCP04 & String(2 - Len(strCP04), "0")
    End If

    strCP05 = "19221111"
    strCP09 = AutoNo("B", 6)
    strCP11 = "90"
    '2005/7/19 ADD BY SONIA
    '抓A,B類最大的CF代理人及彼所案號
    strCP44 = ""
    strCP45 = ""
    StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(strCP01 & strCP02 & strCP03 & strCP04) & " And CP09 <'C' AND CP44 IS NOT NULL Order By CP27 Desc "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    '若有資料
    If rsA.RecordCount > 0 Then
        Do While Not rsA.EOF
            '若有CF代理人
            If "" & rsA("CP44").Value <> "" Then
                strCP44 = rsA("CP44").Value
                If "" & rsA("CP45").Value <> "" Then
                   strCP45 = rsA("CP45").Value
                End If
                Exit Do
            End If
            rsA.MoveNext
        Loop
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '2005/7/19 END
    strCP13 = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04)
    strCP12 = GetSalesArea(strCP13)
    strCP14 = strUserNum
    '91.11.10 ADD BY SONIA
    strCP20 = "N"
    strCP26 = "N"
    strCP32 = "N"
    '91.11.10 END
    
    'Add By Sindy 2025/7/29
    '針對轉案進來的案件，若內部收文日為111111'案件性質為[（107）再審申請]，則
    '按確定後'彈出訊息：請輸入再審查送件日期,將日期回寫到進度檔的[再審查申請]的發文日
    If txtSystem = "FCP" And txtCaseProperty = "107" And textCP05 = "111111" And m_strCP27 <> "" Then
      strCP27 = DBDATE(m_strCP27)
    Else
    '2025/7/29 END
      strCP27 = "19221111"
    End If
    strCP56 = Empty
    
    'add by nickc 2007/01/11
    strCP55 = Empty
    strCP93 = Empty
    strCP94 = Empty
    strCP95 = Empty
    strCP96 = Empty
    
    'Add by Morgan 2006/6/23
    For intI = 89 To 92
      strCP(intI) = Empty
    Next
    If txtPetition.Visible = True And IsEmptyText(txtPetition) = False Then
      'add by nickc 2007/01/11 將 申請人放到移轉人
      'Modified by Lydia 2022/07/04 內部收文之假收文控制：增加專利案件性質401變更或商標301
      'If ((strCP01 = "T" Or strCP01 = "CFT" Or strCP01 = "FCT" Or strCP01 = "CFC" Or strCP01 = "TB" Or strCP01 = "TC" Or strCP01 = "TF") And txtCaseProperty = "501") Or ((strCP01 = "P" Or strCP01 = "FCP" Or strCP01 = "CFP") And (txtCaseProperty = "701" Or txtCaseProperty = "708")) Then
      If ((strCP01 = "T" Or strCP01 = "CFT" Or strCP01 = "FCT" Or strCP01 = "CFC" Or strCP01 = "TB" Or strCP01 = "TC" Or strCP01 = "TF") _
           And (txtCaseProperty = "501" Or txtCaseProperty = "301")) Or _
         ((strCP01 = "P" Or strCP01 = "FCP" Or strCP01 = "CFP") _
           And (txtCaseProperty = "701" Or txtCaseProperty = "708" Or txtCaseProperty = "401")) Then
      'end 202/07/04
         StrSqlB = "select tm23,tm78,tm79,tm80,tm81 from trademark where tm01='" & strCP01 & "' and tm02='" & strCP02 & "' and tm03='" & strCP03 & "' and tm04='" & strCP04 & "' "
         StrSqlB = StrSqlB & " union all select pa26,pa27,pa28,pa29,pa30 from patent where pa01='" & strCP01 & "' and pa02='" & strCP02 & "' and pa03='" & strCP03 & "' and pa04='" & strCP04 & "' "
         StrSqlB = StrSqlB & " union all select sp08,sp58,sp59,sp65,sp66 from servicepractice where sp01='" & strCP01 & "' and sp02='" & strCP02 & "' and sp03='" & strCP03 & "' and sp04='" & strCP04 & "' "
         CheckOC3
         With AdoRecordSet3
            .CursorLocation = adUseClient
            .Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount <> 0 Then
                strCP55 = CheckStr(.Fields(0).Value)
                strCP93 = CheckStr(.Fields(1).Value)
                strCP94 = CheckStr(.Fields(2).Value)
                strCP95 = CheckStr(.Fields(3).Value)
                strCP96 = CheckStr(.Fields(4).Value)
            End If
         End With
      End If
      strCP56 = Left(txtPetition & "000", 9)
      'Add by Morgan 2006/6/23
      For intI = 2 To 5
         If IsEmptyText(txtPetitionx(intI)) = False Then
            strCP(intI + 87) = Left(txtPetitionx(intI) & "000", 9)
         End If
      Next
    End If
    ' 組成SQL語法
    '2005/7/19 MODIFY BY SONIA
    'strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP56) " & _
    '                "VALUES ('" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "'," & _
    '                        strCP05 & ",'" & strCP09 & "','" & txtCaseProperty & "','" & strCP11 & "'," & _
    '                        "'" & strCP12 & "','" & strCP13 & "','" & strCP14 & "','" & strCP20 & "','" & strCP26 & "'," & strCP27 & ",'" & strCP32 & "'," & DBNullString(strCP56) & ") "
    'Modify by Morgan 2006/6/23 加cp89,cp90,cp91,cp92
'    strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP56,CP44,CP45,CP89,CP90,CP91,CP92) " & _
                    "VALUES ('" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "'," & _
                            strCP05 & ",'" & strCP09 & "','" & txtCaseProperty & "','" & strCP11 & "'," & _
                            "'" & strCP12 & "','" & strCP13 & "','" & strCP14 & "','" & strCP20 & "','" & strCP26 & "'," & strCP27 & ",'" & strCP32 & "'," & DBNullString(strCP56) & "," & DBNullString(strCP44) & "," & DBNullString(strCP45) & "," & CNULL(StrCp(89)) & "," & CNULL(StrCp(90)) & "," & CNULL(StrCp(91)) & "," & CNULL(StrCp(92)) & ") "
    On Error GoTo MyErrLog
    cnnConnection.BeginTrans
    strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP56,CP44,CP45,CP89,CP90,CP91,CP92,cp55,cp93,cp94,cp95,cp96) " & _
                    "VALUES ('" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "'," & _
                            strCP05 & ",'" & strCP09 & "','" & txtCaseProperty & "','" & strCP11 & "'," & _
                            "'" & strCP12 & "','" & strCP13 & "','" & strCP14 & "','" & strCP20 & "','" & strCP26 & "'," & strCP27 & ",'" & strCP32 & "'," & DBNullString(strCP56) & "," & DBNullString(strCP44) & "," & DBNullString(strCP45) & "," & CNULL(strCP(89)) & "," & CNULL(strCP(90)) & "," & CNULL(strCP(91)) & "," & CNULL(strCP(92)) & "," & CNULL(strCP55) & "," & CNULL(strCP93) & "," & CNULL(strCP94) & "," & CNULL(strCP95) & "," & CNULL(strCP96) & " ) "
    
    '2005/7/19 END
    cnnConnection.Execute strSql, intI
    
    'Added by Morgan 2012/8/14
    '中間接辦案件若未核准,以系統日期加6個月為催審日
    'Modified by Morgan 2013/6/18 不必限制案件性質(程序自行判斷需催審的才會收文),CFP也要 --郭
    'If strCP01 = "P" And InStr("101,102,103", txtCaseProperty) > 0 Then
    'Modified by Lydia 2018/05/09 主張國際優先權106不用催審 (ex.CFP-029915)
    'If strCP01 = "P" Or strCP01 = "CFP" Then
    If (strCP01 = "P" Or strCP01 = "CFP") And txtCaseProperty <> "106" Then
    'end 2013/6/18
         strExc(1) = CompDate(1, 6, strSrvDate(1))
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
               "NP07,NP08,NP09,NP10,NP22) select '" & strCP09 & "'" & _
               ",'" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "'," & 催審 & "," & strExc(2) & _
               "," & strExc(1) & ",'" & strUserNum & "',NP22" & _
               " from PATENT,(select nvl(max(np22),0)+1 NP22 from nextprogress)" & _
               " WHERE PA01='" & strCP01 & "' AND PA02='" & strCP02 & "' AND PA03='" & strCP03 & "' AND PA04='" & strCP04 & "' AND PA16 is null"
         cnnConnection.Execute strSql, intI
    End If
    'end 2012/8/14
    
    'Add by Amy 2013/08/27 實審內部收文要輸入中說頁數及項數
    If strCP01 = "FCP" And txtCaseProperty = "416" Then
          strSql = "Update CaseProgress set cp135=" & intCP135 & ",cp136=" & intCP136 & " Where cp09='" & strCP09 & "'"
          cnnConnection.Execute strSql, intI
    End If
    'end 2013/08/27
    
    'add by nickc 2007/01/11 將 更新基本檔
    'Modified by Lydia 2022/06/30 內部收文之假收文控制：增加專利案件性質401變更或商標301
    'If ((strCP01 = "T" Or strCP01 = "CFT" Or strCP01 = "FCT" Or strCP01 = "CFC" Or strCP01 = "TB" Or strCP01 = "TC" Or strCP01 = "TF") And txtCaseProperty = "501") Or ((strCP01 = "P" Or strCP01 = "FCP" Or strCP01 = "CFP") And (txtCaseProperty = "701" Or txtCaseProperty = "708")) Then
    If strCP56 <> "" And (((strCP01 = "T" Or strCP01 = "CFT" Or strCP01 = "FCT" Or strCP01 = "CFC" Or strCP01 = "TB" Or strCP01 = "TC" Or strCP01 = "TF") And (txtCaseProperty = "501" Or txtCaseProperty = "301")) Or _
           ((strCP01 = "P" Or strCP01 = "FCP" Or strCP01 = "CFP") And (txtCaseProperty = "701" Or txtCaseProperty = "708" Or txtCaseProperty = "401"))) Then
    'end 2022/06/30
        If (strCP01 = "T" Or strCP01 = "CFT" Or strCP01 = "FCT" Or strCP01 = "TF") Then
            StrSQLa = "Update TradeMark Set TM23=null,TM24=null,TM25=null,TM26=null,tm78=null,tm79=null,tm80=null,tm81=null,tm82=null,tm83=null,tm84=null,tm85=null" & _
                            ",tm86=null,tm87=null,tm88=null,tm89=null,tm90=null,tm91=null,tm92=null,tm93=null " & _
                            " Where " & ChgTradeMark(strCP01 & strCP02 & strCP03 & strCP04)
            cnnConnection.Execute StrSQLa
            If Not IsEmpty(strCP56) Then
                StrSQLa = "Update TradeMark Set TM23='" & ChangeCustomerL(strCP56) & "' " & _
                                " ,TM24='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP56), "1")) & "' " & _
                                " ,TM25='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP56), "2")) & "' " & _
                                " ,TM26='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP56), "3")) & "' " & _
                                " Where " & ChgTradeMark(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(89)) Then
                StrSQLa = "Update TradeMark Set TM78='" & ChangeCustomerL(strCP(89)) & "' " & _
                                " ,TM82='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(89)), "1")) & "' " & _
                                " ,TM86='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(89)), "2")) & "' " & _
                                " ,TM90='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(89)), "3")) & "' " & _
                                " Where " & ChgTradeMark(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(90)) Then
                StrSQLa = "Update TradeMark Set TM79='" & ChangeCustomerL(strCP(90)) & "' " & _
                                " ,TM83='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(90)), "1")) & "' " & _
                                " ,TM87='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(90)), "2")) & "' " & _
                                " ,TM91='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(90)), "3")) & "' " & _
                                " Where " & ChgTradeMark(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(91)) Then
                StrSQLa = "Update TradeMark Set TM80='" & ChangeCustomerL(strCP(91)) & "' " & _
                                " ,TM84='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(91)), "1")) & "' " & _
                                " ,TM88='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(91)), "2")) & "' " & _
                                " ,TM92='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(91)), "3")) & "' " & _
                                " Where " & ChgTradeMark(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(92)) Then
                StrSQLa = "Update TradeMark Set TM81='" & ChangeCustomerL(strCP(92)) & "' " & _
                                " ,TM85='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(92)), "1")) & "' " & _
                                " ,TM89='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(92)), "2")) & "' " & _
                                " ,TM93='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(92)), "3")) & "' " & _
                                " Where " & ChgTradeMark(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
        ElseIf (strCP01 = "CFC" Or strCP01 = "TB" Or strCP01 = "TC") Then
            StrSQLa = "Update ServicePractice Set SP08=null,sp58=null,sp59=null,sp65=null,sp66=null " & _
                            " Where " & ChgService(strCP01 & strCP02 & strCP03 & strCP04)
            cnnConnection.Execute StrSQLa
            If Not IsEmpty(strCP56) Then
                StrSQLa = "Update ServicePractice Set SP08='" & ChangeCustomerL(strCP56) & "' " & _
                                " Where " & ChgService(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(89)) Then
                StrSQLa = "Update ServicePractice Set SP58='" & ChangeCustomerL(strCP(89)) & "' " & _
                                " Where " & ChgService(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(90)) Then
                StrSQLa = "Update ServicePractice Set SP59='" & ChangeCustomerL(strCP(90)) & "' " & _
                                " Where " & ChgService(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(91)) Then
                StrSQLa = "Update ServicePractice Set SP65='" & ChangeCustomerL(strCP(91)) & "' " & _
                                " Where " & ChgService(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(92)) Then
                StrSQLa = "Update ServicePractice Set SP66='" & ChangeCustomerL(strCP(92)) & "' " & _
                                " Where " & ChgService(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
        ElseIf (strCP01 = "P" Or strCP01 = "FCP" Or strCP01 = "CFP") Then
            StrSQLa = "Update patent Set pa26=null,pa27=null,pa28=null,pa29=null,pa30=null,pa31=null,pa32=null,pa33=null,pa34=null,pa35=null " & _
                            ",pa36=null,pa37=null,pa38=null,pa39=null,pa40=null,pa41=null,pa42=null,pa43=null,pa44=null,pa45=null " & _
                            " Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
            cnnConnection.Execute StrSQLa
            If Not IsEmpty(strCP56) Then
                StrSQLa = "Update patent Set pa26='" & ChangeCustomerL(strCP56) & "' " & _
                                " ,pa31='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP56), "1")) & "' " & _
                                " ,pa36='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP56), "2")) & "' " & _
                                " ,pa41='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP56), "3")) & "' " & _
                                " Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(89)) Then
                StrSQLa = "Update patent Set pa27='" & ChangeCustomerL(strCP(89)) & "' " & _
                                " ,pa32='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(89)), "1")) & "' " & _
                                " ,pa37='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(89)), "2")) & "' " & _
                                " ,pa42='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(89)), "3")) & "' " & _
                                " Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(90)) Then
                StrSQLa = "Update patent Set pa28='" & ChangeCustomerL(strCP(90)) & "' " & _
                                " ,pa33='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(90)), "1")) & "' " & _
                                " ,pa38='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(90)), "2")) & "' " & _
                                " ,pa43='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(90)), "3")) & "' " & _
                                " Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(91)) Then
                StrSQLa = "Update patent Set pa29='" & ChangeCustomerL(strCP(91)) & "' " & _
                                " ,pa34='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(91)), "1")) & "' " & _
                                " ,pa39='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(91)), "2")) & "' " & _
                                " ,pa44='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(91)), "3")) & "' " & _
                                " Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
            If Not IsEmpty(strCP(92)) Then
                StrSQLa = "Update patent Set pa30='" & ChangeCustomerL(strCP(92)) & "' " & _
                                " ,pa35='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(92)), "1")) & "' " & _
                                " ,pa40='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(92)), "2")) & "' " & _
                                " ,pa45='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(strCP(92)), "3")) & "' " & _
                                " Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
                cnnConnection.Execute StrSQLa
            End If
        End If
    End If
    ' 顯示所新增的案件號碼
    'Modify by Morgan 2004/5/10
'    txtRecieveCode(0) = Mid(strCP09, 1, 1)
'    txtRecieveCode(1) = Mid(strCP09, 2, Len(strCP09) - 1)
'    lblCaseCode = strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04
    
    'Added by Lydia 2016/09/21 內部收文(假收文)更新初審階段提分割
    If m_UpdPA163 <> "" Then
       StrSQLa = "UPDATE PATENT SET " & m_UpdPA163 & " WHERE " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
       cnnConnection.Execute StrSQLa
    End If
    'end 2016/09/21
    
    'Added by Lydia 2017/03/15 內部收文-更新基本檔的救濟程序或爭議程序
    If (strCP01 = "P" Or strCP01 = "FCP") And (Mid(txtCaseProperty, 1, 1) = "5" Or Mid(txtCaseProperty, 1, 1) = "8") Then
       StrSQLa = "UPDATE PATENT SET " & IIf(Mid(txtCaseProperty, 1, 1) = "5", "PA18='Y'", "PA19='Y'") & " WHERE " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
       cnnConnection.Execute StrSQLa
    ElseIf (strCP01 = "T" Or strCP01 = "TF" Or strCP01 = "FCT" Or strCP01 = "CFT") And (Mid(txtCaseProperty, 1, 1) = "4" Or Mid(txtCaseProperty, 1, 1) = "6") Then
       StrSQLa = "UPDATE TRADEMARK SET " & IIf(Mid(txtCaseProperty, 1, 1) = "4", "TM18='Y'", "TM19='Y'") & " WHERE " & ChgTradeMark(strCP01 & strCP02 & strCP03 & strCP04)
       cnnConnection.Execute StrSQLa
    End If
    'END 2017/03/15
    
    txtRecieveCode(0) = Mid(strCP09, 2, 2)
    txtRecieveCode(1) = Mid(strCP09, 4)
    If txtSystem = "TF" Then
      lblCaseCode = strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & Left(strCP04, 1) & "-" & Mid(strCP04, 2)
    Else
      lblCaseCode = strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04
   End If
'add by nickc 2007/01/11   ****start
   cnnConnection.CommitTrans
   OnSaveNewData = True
   Exit Function
MyErrLog:
    cnnConnection.RollbackTrans
    MsgBox "存檔失敗"
    '***** end
End Function

'Modify By Cheng 2003/03/28
'Private Sub OnNextForm()
Private Function OnNextForm() As Boolean
'   Dim strSql As String
   Dim rsB As New ADODB.Recordset

   ' 本所案號
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
       
    'Add By Cheng 2003/03/28
    OnNextForm = True
   '911111 nick
   'strCP01 = txtSystem
   'strCP02 = txtCode(0)
   'strCP03 = txtCode(1)
   'strCP03 = strCP03 & String(1 - Len(strCP03), "0")
   'strCP04 = txtCode(2)
   'strCP04 = strCP04 & String(2 - Len(strCP04), "0")
   strCP01 = txtSystem
   strCP02 = IIf(txtSystem <> "TF", txtCode(0), txtTFCode(0) & txtTFCode(1))
   strCP03 = IIf(txtSystem <> "TF", txtCode(1), txtTFCode(2))
   strCP03 = strCP03 & String(1 - Len(strCP03), "0")
   strCP04 = IIf(txtSystem <> "TF", txtCode(2), txtTFCode(3))
   strCP04 = strCP04 & String(2 - Len(strCP04), "0")

   '911018 nick 當修改時，是輸入收文號，並沒有本所案號
   '所以要先 select
   '***** start
   If strCP01 = "" And strCP02 = "" And txtRecieveCode(1).Text <> "" Then
        Dim nickstrsql As String
        Dim nick911018rs As New ADODB.Recordset
        Set nick911018rs = New ADODB.Recordset
        nickstrsql = "select cp01,cp02,cp03,cp04 from caseprogress where cp09='" & strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text & "' "
        
        nick911018rs.CursorLocation = adUseClient
        nick911018rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
        If nick911018rs.RecordCount <> 0 Then
             strCP01 = CheckStr(nick911018rs.Fields(0).Value)
             strCP02 = CheckStr(nick911018rs.Fields(1).Value)
             strCP03 = CheckStr(nick911018rs.Fields(2).Value)
             strCP04 = CheckStr(nick911018rs.Fields(3).Value)
        'Add By Cheng 2003/03/28
        '若無資料
        Else
            MsgBox "無此收文號資料!!!", vbExclamation + vbOKOnly
            OnNextForm = False
            nick911018rs.Close
            Set nick911018rs = Nothing
            Exit Function
        End If
   End If
   '*** end

   'Modify By Sindy 2012/2/23 改呼叫函數
   'Modified by Lydia 2014/11/11 原本被註解,重新使用取得查詢-系統
   If Len(strSK02) = 0 Or Len(strSK03) = 0 Then
        strExc(0) = "SELECT * FROM SYSTEMKIND " & _
                 "WHERE SK01 = '" & strCP01 & "' "

        rsB.CursorLocation = adUseClient
        rsB.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
        If rsB.RecordCount > 0 Then
           rsB.MoveFirst
           If IsNull(rsB.Fields("SK02")) = False Then
              strSK02 = rsB.Fields("SK02")
           End If
           If IsNull(rsB.Fields("SK03")) = False Then
              strSK03 = rsB.Fields("SK03")
           End If
         'Add By Cheng 2003/03/28
         '若無系統類別
         Else
             MsgBox "此資料無相關系統類別!!!", vbExclamation + vbOKOnly
             OnNextForm = False
              If rsB.State <> adStateClosed Then rsB.Close
              Set rsB = Nothing
             Exit Function
        End If
        If rsB.State <> adStateClosed Then rsB.Close
        Set rsB = Nothing
   End If
   
   Select Case strSK02
      ' 專利
      Case "1":
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_04.SetData strCP01, 0, True
               frm010012_04.SetData strCP02, 1, False
               frm010012_04.SetData strCP03, 2, False
               frm010012_04.SetData strCP04, 3, False
               frm010012_04.SetData txtCaseProperty, 4, False
               frm010012_04.SetData txtPetition, 5, False
               '92.03.27
               frm010012_04.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               'Add By Sindy 2009/10/19
               frm010012_04.SetData txtPetitionx(2), 8, False
               frm010012_04.SetData txtPetitionx(3), 9, False
               frm010012_04.SetData txtPetitionx(4), 10, False
               frm010012_04.SetData txtPetitionx(5), 11, False
               
               'Add By Sindy 2018/2/22
               frm010012_04.m_strIR01 = m_strIR01
               frm010012_04.m_strIR02 = m_strIR02
               frm010012_04.m_strIR03 = m_strIR03
               frm010012_04.m_strIR04 = m_strIR04
               Set frm010012_04.m_PrevForm = mPrevForm
               '2018/2/22 END
               
               frm010012_04.Show
               frm010012_04.QueryData
            ' 外
            Case "1":
               frm010012_05.SetData strCP01, 0, True
               frm010012_05.SetData strCP02, 1, False
               frm010012_05.SetData strCP03, 2, False
               frm010012_05.SetData strCP04, 3, False
               frm010012_05.SetData txtCaseProperty, 4, False
               frm010012_05.SetData txtPetition, 5, False
               '92.03.27
               frm010012_05.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               'Add By Sindy 2009/10/19
               frm010012_05.SetData txtPetitionx(2), 8, False
               frm010012_05.SetData txtPetitionx(3), 9, False
               frm010012_05.SetData txtPetitionx(4), 10, False
               frm010012_05.SetData txtPetitionx(5), 11, False
               
               'Add By Sindy 2018/2/22
               frm010012_05.m_strIR01 = m_strIR01
               frm010012_05.m_strIR02 = m_strIR02
               frm010012_05.m_strIR03 = m_strIR03
               frm010012_05.m_strIR04 = m_strIR04
               Set frm010012_05.m_PrevForm = mPrevForm
               '2018/2/22 END
               
               frm010012_05.Show
               frm010012_05.QueryData
            ' 外
            Case "2":
               frm010012_06.SetData strCP01, 0, True
               frm010012_06.SetData strCP02, 1, False
               frm010012_06.SetData strCP03, 2, False
               frm010012_06.SetData strCP04, 3, False
               frm010012_06.SetData txtCaseProperty, 4, False
               frm010012_06.SetData txtPetition, 5, False
               '92.03.27
               frm010012_06.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               'Add By Sindy 2009/10/19
               frm010012_06.SetData txtPetitionx(2), 8, False
               frm010012_06.SetData txtPetitionx(3), 9, False
               frm010012_06.SetData txtPetitionx(4), 10, False
               frm010012_06.SetData txtPetitionx(5), 11, False
               
               'Add By Sindy 2022/6/28
               frm010012_06.m_strIR01 = m_strIR01
               frm010012_06.m_strIR02 = m_strIR02
               frm010012_06.m_strIR03 = m_strIR03
               frm010012_06.m_strIR04 = m_strIR04
               Set frm010012_06.m_PrevForm = mPrevForm
               '2022/6/28 END
               
               frm010012_06.Show
               frm010012_06.QueryData
         End Select
      ' 商標
      Case "2":
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_01.SetData strCP01, 0, True
               frm010012_01.SetData strCP02, 1, False
               frm010012_01.SetData strCP03, 2, False
               frm010012_01.SetData strCP04, 3, False
               frm010012_01.SetData txtCaseProperty, 4, False
               frm010012_01.SetData txtPetition, 5, False
               '92.03.27
               frm010012_01.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               'add by nickc 2006/12/01
               frm010012_01.SetData txtPetitionx(2), 8, False
               frm010012_01.SetData txtPetitionx(3), 9, False
               frm010012_01.SetData txtPetitionx(4), 10, False
               frm010012_01.SetData txtPetitionx(5), 11, False
               
               'Add By Sindy 2019/5/27
               frm010012_01.m_strIR01 = m_strIR01
               frm010012_01.m_strIR02 = m_strIR02
               frm010012_01.m_strIR03 = m_strIR03
               frm010012_01.m_strIR04 = m_strIR04
               Set frm010012_01.m_PrevForm = mPrevForm
               '2019/5/27 END
               
               frm010012_01.Show
               frm010012_01.QueryData
            ' 外
            Case "1", "2":
               frm010012_02.SetData strCP01, 0, True
               frm010012_02.SetData strCP02, 1, False
               frm010012_02.SetData strCP03, 2, False
               frm010012_02.SetData strCP04, 3, False
               frm010012_02.SetData txtCaseProperty, 4, False
               frm010012_02.SetData txtPetition, 5, False
               '92.03.27
               frm010012_02.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               'add by nickc 2006/12/01
               frm010012_02.SetData txtPetitionx(2), 8, False
               frm010012_02.SetData txtPetitionx(3), 9, False
               frm010012_02.SetData txtPetitionx(4), 10, False
               frm010012_02.SetData txtPetitionx(5), 11, False
               
               'Add By Sindy 2019/5/27
               frm010012_02.m_strIR01 = m_strIR01
               frm010012_02.m_strIR02 = m_strIR02
               frm010012_02.m_strIR03 = m_strIR03
               frm010012_02.m_strIR04 = m_strIR04
               Set frm010012_02.m_PrevForm = mPrevForm
               '2019/5/27 END
               
               frm010012_02.Show
               frm010012_02.QueryData
         End Select
      ' 法務
      Case "3":
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_07.SetData strCP01, 0, True
               frm010012_07.SetData strCP02, 1, False
               frm010012_07.SetData strCP03, 2, False
               frm010012_07.SetData strCP04, 3, False
               frm010012_07.SetData txtCaseProperty, 4, False
               frm010012_07.SetData txtPetition, 5, False
               '92.03.27
               frm010012_07.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_07.Show
               frm010012_07.QueryData
            ' 外
            Case "1", "2":
               frm010012_08.SetData strCP01, 0, True
               frm010012_08.SetData strCP02, 1, False
               frm010012_08.SetData strCP03, 2, False
               frm010012_08.SetData strCP04, 3, False
               frm010012_08.SetData txtCaseProperty, 4, False
               frm010012_08.SetData txtPetition, 5, False
               '92.03.27
               frm010012_08.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_08.Show
               frm010012_08.QueryData
         End Select
      ' 顧問
      Case "4":
         frm010012_07.SetData strCP01, 0, True
         frm010012_07.SetData strCP02, 1, False
         frm010012_07.SetData strCP03, 2, False
         frm010012_07.SetData strCP04, 3, False
         frm010012_07.SetData txtCaseProperty, 4, False
         frm010012_07.SetData txtPetition, 5, False
         '92.03.27
         frm010012_07.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
         
         frm010012_07.Show
         frm010012_07.QueryData
      ' 專利
      Case "5"
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_04.SetData strCP01, 0, True
               frm010012_04.SetData strCP02, 1, False
               frm010012_04.SetData strCP03, 2, False
               frm010012_04.SetData strCP04, 3, False
               frm010012_04.SetData txtCaseProperty, 4, False
               frm010012_04.SetData txtPetition, 5, False
               '92.03.27
               frm010012_04.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               'Add By Sindy 2009/10/19
               frm010012_04.SetData txtPetitionx(2), 8, False
               frm010012_04.SetData txtPetitionx(3), 9, False
               frm010012_04.SetData txtPetitionx(4), 10, False
               frm010012_04.SetData txtPetitionx(5), 11, False
               
               'Add By Sindy 2018/2/22
               frm010012_04.m_strIR01 = m_strIR01
               frm010012_04.m_strIR02 = m_strIR02
               frm010012_04.m_strIR03 = m_strIR03
               frm010012_04.m_strIR04 = m_strIR04
               Set frm010012_04.m_PrevForm = mPrevForm
               '2018/2/22 END
               
               frm010012_04.Show
               frm010012_04.QueryData
            ' 外
            Case "1":
               frm010012_05.SetData strCP01, 0, True
               frm010012_05.SetData strCP02, 1, False
               frm010012_05.SetData strCP03, 2, False
               frm010012_05.SetData strCP04, 3, False
               frm010012_05.SetData txtCaseProperty, 4, False
               frm010012_05.SetData txtPetition, 5, False
               '92.03.27
               frm010012_05.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               'Add By Sindy 2009/10/19
               frm010012_05.SetData txtPetitionx(2), 8, False
               frm010012_05.SetData txtPetitionx(3), 9, False
               frm010012_05.SetData txtPetitionx(4), 10, False
               frm010012_05.SetData txtPetitionx(5), 11, False
               
               'Add By Sindy 2018/2/22
               frm010012_05.m_strIR01 = m_strIR01
               frm010012_05.m_strIR02 = m_strIR02
               frm010012_05.m_strIR03 = m_strIR03
               frm010012_05.m_strIR04 = m_strIR04
               Set frm010012_05.m_PrevForm = mPrevForm
               '2018/2/22 END
               
               frm010012_05.Show
               frm010012_05.QueryData
            ' 外
            Case "2":
               frm010012_06.SetData strCP01, 0, True
               frm010012_06.SetData strCP02, 1, False
               frm010012_06.SetData strCP03, 2, False
               frm010012_06.SetData strCP04, 3, False
               frm010012_06.SetData txtCaseProperty, 4, False
               frm010012_06.SetData txtPetition, 5, False
               '92.03.27
               frm010012_06.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               'Add By Sindy 2009/10/19
               frm010012_06.SetData txtPetitionx(2), 8, False
               frm010012_06.SetData txtPetitionx(3), 9, False
               frm010012_06.SetData txtPetitionx(4), 10, False
               frm010012_06.SetData txtPetitionx(5), 11, False
               
               'Add By Sindy 2022/6/28
               frm010012_06.m_strIR01 = m_strIR01
               frm010012_06.m_strIR02 = m_strIR02
               frm010012_06.m_strIR03 = m_strIR03
               frm010012_06.m_strIR04 = m_strIR04
               Set frm010012_06.m_PrevForm = mPrevForm
               '2022/6/28 END
               
               frm010012_06.Show
               frm010012_06.QueryData
         End Select
      ' 商標
      Case "6":
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_01.SetData strCP01, 0, True
               frm010012_01.SetData strCP02, 1, False
               frm010012_01.SetData strCP03, 2, False
               frm010012_01.SetData strCP04, 3, False
               frm010012_01.SetData txtCaseProperty, 4, False
               frm010012_01.SetData txtPetition, 5, False
               '92.03.27
               frm010012_01.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               'add by nickc 2006/12/01
               frm010012_01.SetData txtPetitionx(2), 8, False
               frm010012_01.SetData txtPetitionx(3), 9, False
               frm010012_01.SetData txtPetitionx(4), 10, False
               frm010012_01.SetData txtPetitionx(5), 11, False
               
               'Add By Sindy 2019/5/27
               frm010012_01.m_strIR01 = m_strIR01
               frm010012_01.m_strIR02 = m_strIR02
               frm010012_01.m_strIR03 = m_strIR03
               frm010012_01.m_strIR04 = m_strIR04
               Set frm010012_01.m_PrevForm = mPrevForm
               '2019/5/27 END
               
               frm010012_01.Show
               frm010012_01.QueryData
            ' 外
            Case "1", "2":
               frm010012_02.SetData strCP01, 0, True
               frm010012_02.SetData strCP02, 1, False
               frm010012_02.SetData strCP03, 2, False
               frm010012_02.SetData strCP04, 3, False
               frm010012_02.SetData txtCaseProperty, 4, False
               frm010012_02.SetData txtPetition, 5, False
               '92.03.27
               frm010012_02.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               'add by nickc 2006/12/01
               frm010012_02.SetData txtPetitionx(2), 8, False
               frm010012_02.SetData txtPetitionx(3), 9, False
               frm010012_02.SetData txtPetitionx(4), 10, False
               frm010012_02.SetData txtPetitionx(5), 11, False
               
               'Add By Sindy 2019/5/27
               frm010012_02.m_strIR01 = m_strIR01
               frm010012_02.m_strIR02 = m_strIR02
               frm010012_02.m_strIR03 = m_strIR03
               frm010012_02.m_strIR04 = m_strIR04
               Set frm010012_02.m_PrevForm = mPrevForm
               '2019/5/27 END
               
               frm010012_02.Show
               frm010012_02.QueryData
         End Select
      ' 法務
      Case "7":
         Select Case strSK03
            ' 內
            Case "0":
               frm010012_07.SetData strCP01, 0, True
               frm010012_07.SetData strCP02, 1, False
               frm010012_07.SetData strCP03, 2, False
               frm010012_07.SetData strCP04, 3, False
               frm010012_07.SetData txtCaseProperty, 4, False
               frm010012_07.SetData txtPetition, 5, False
               '92.03.27
               frm010012_07.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_07.Show
               frm010012_07.QueryData
            ' 外
            Case "1", "2":
               frm010012_08.SetData strCP01, 0, True
               frm010012_08.SetData strCP02, 1, False
               frm010012_08.SetData strCP03, 2, False
               frm010012_08.SetData strCP04, 3, False
               frm010012_08.SetData txtCaseProperty, 4, False
               frm010012_08.SetData txtPetition, 5, False
               '92.03.27
               frm010012_08.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
               
               frm010012_08.Show
               frm010012_08.QueryData
         End Select
      ' 顧問
      Case "8":
         frm010012_07.SetData strCP01, 0, True
         frm010012_07.SetData strCP02, 1, False
         frm010012_07.SetData strCP03, 2, False
         frm010012_07.SetData strCP04, 3, False
         frm010012_07.SetData txtCaseProperty, 4, False
         frm010012_07.SetData txtPetition, 5, False
         '92.03.27
         frm010012_07.SetData strReceiveKind + txtRecieveCode(0).Text + txtRecieveCode(1).Text, 7, False
         
         frm010012_07.Show
         frm010012_07.QueryData
   End Select
End Function

Public Sub SetData(ByVal strData As String, ByVal nType As Integer, ByVal bClear As Boolean)
   If bClear Then
'      txtSystem = Empty
'      txtCode(0) = Empty
'      txtCode(1) = Empty
'      txtCode(2) = Empty
'      textCP05 = Empty
'      txtCaseProperty = Empty
'      lblCasePropertyName = Empty
'      txtPetition = Empty
'      lblPetitionName = Empty
'      'Add by Morgan 2006/6/23
'      For intI = 2 To 5
'         txtPetitionx(intI) = Empty
'         lblPetitionNamex(intI) = Empty
'      Next
'      fraRecieve.Enabled = False
'      fraCode.Visible = True
'      fraLastCaseCode.Visible = True
      'Modify By Sindy 2012/3/1 寫成共用函數
      Call SetColClearVal(bClear)
   End If
   Select Case nType
      Case 0:
         ' 91.10.15 modify by louis (固定顯示後面六碼即可, 其它自動顯示)
         'txtRecieveCode(0) = Mid(strData, 1, 1)
         'txtRecieveCode(1) = Mid(strData, 2, Len(strData) - 1)
         txtRecieveCode(1) = Right(strData, 6)
      Case 1:
         lblCaseCode = strData
      Case 2:
         lblCaseCode = lblCaseCode & "-" & strData
      Case 3:
         lblCaseCode = lblCaseCode & "-" & strData
      Case 4:
         lblCaseCode = lblCaseCode & "-" & strData
   End Select
End Sub

'Add By Cheng 2003/09/08
Private Function CheckNewCase(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

CheckNewCase = False
'若無流水號
If strCP02 = "" Then
    CheckNewCase = True
'若有流水號
Else
    StrSQLa = "Select PA01 From Patent Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
    StrSQLa = StrSQLa & " Union Select TM01 From Trademark Where " & ChgTradeMark(strCP01 & strCP02 & strCP03 & strCP04)
    StrSQLa = StrSQLa & " Union Select LC01 From Lawcase Where " & ChgLawcase(strCP01 & strCP02 & strCP03 & strCP04)
    StrSQLa = StrSQLa & " Union Select HC01 From Hirecase Where " & ChgHirecase(strCP01 & strCP02 & strCP03 & strCP04)
    StrSQLa = StrSQLa & " Union Select SP01 From Servicepractice Where " & ChgService(strCP01 & strCP02 & strCP03 & strCP04)
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    '若基本檔無資料
    If rsA.RecordCount <= 0 Then
        CheckNewCase = True
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End If

End Function

'Private Function Check412(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As Boolean
'
'   Check412 = False
'
'   strSql = "Select PA09, PA14, CP27  From Patent, caseprogress Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04) & " and CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP10(+)='601'"
'
'On Error GoTo ErrHnd
'
'   CheckOC
'   With adoRecordset
'      .CursorLocation = adUseClient
'      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      '若有資料
'      If .RecordCount > 0 Then
'         If ("" & .Fields("PA09")) = "000" Then
'            '已公告
'            If Val("" & .Fields("PA14")) > 0 Then
'               MsgBox "該案已於 " & ChangeTStringToTDateString(.Fields("PA14") - 19110000) & " 公告，不可收" & lblCasePropertyName & "！", vbCritical
'
'            'Modify by Morgan 2004/12/20 取消控制
''            '領證已發文
''            ElseIf Not IsNull(.Fields("CP27")) Then
''               MsgBox "該案領證已於 " & ChangeTStringToTDateString(.Fields("CP27") - 19110000) & " 發文，不可收" & lblCasePropertyName & "！", vbCritical
'
'            Else
'               Check412 = True
'            End If
'         '非台灣案
'         Else
'            Check412 = True
'         End If
'      End If
'
'   End With
'   CheckOC
'
'ErrHnd:
'
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
'
'End Function

'Add by Morgan 2004/7/5
'抓申請國家
Private Function GetPA09(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As String

   strSql = "Select PA09  From Patent Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04)
   
On Error GoTo ErrHnd
   
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      '若有資料
      If .RecordCount > 0 Then
         GetPA09 = "" & .Fields(0)
      End If
   
   End With
   CheckOC
   
ErrHnd:

   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

''2007/7/4 add by sonia 檢查是否同意重新委任
'Private Function ChkAgree928() As Boolean
'Dim StrSQLa As String, StrSqlB As String
'Dim rsA As New ADODB.Recordset, rsB As New ADODB.Recordset
'Dim iCol As Integer
'
'   ChkAgree928 = True
'   StrSQLa = "Select PA26,PA27,PA28,PA29,PA30 From patent where pa01='" & txtSystem & "' and pa02='" & txtCode(0) & "' and pa03='" & IIf(txtCode(1) = "", "0", txtCode(1)) & "' and pa04='" & IIf(txtCode(2) = "", "00", txtCode(2)) & "' "
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      For iCol = 0 To 4
'         If Not IsNull(rsA.Fields(iCol)) Then
'            StrSqlB = "Select * From LinReasignRec Where LR01='" & rsA.Fields(iCol) & "' "
'            rsB.CursorLocation = adUseClient
'            rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsB.RecordCount > 0 Then
'               If rsB("LR09").Value = "N" Then
'                  MsgBox "此案件客戶不同意重新委任, 請退回原智權人員 !!!", vbExclamation + vbOKOnly
'                  ChkAgree928 = False
'                  If rsB.State <> adStateClosed Then rsB.Close
'                  Set rsB = Nothing
'                  Exit Function
'               End If
'            End If
'            If rsB.State <> adStateClosed Then rsB.Close
'            Set rsB = Nothing
'         Else
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            Exit Function
'         End If
'      Next
'   Else
'      MsgBox "無此案號基本資料 !!!", vbExclamation + vbOKOnly
'      ChkAgree928 = False
'   End If
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'End Function
''2007/7/4 end

'Add By Sindy 2009/07/06
Private Sub textYear_GotFocus()
   InverseTextBox textYear
   m_bolStopOntextYear = False
End Sub
Private Sub Text1_GotFocus(Index As Integer)
   InverseTextBox Text1(Index)
   If Index = 0 Then
      m_bolStopOntextYear = False
   End If
End Sub
Private Sub textYear_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
'2009/07/06 End

'Add By Sindy 2012/2/23
Private Function GetSysTemKind() As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   GetSysTemKind = True
   
   strSql = "SELECT * FROM SYSTEMKIND " & _
            "WHERE SK01 = '" & txtSystem & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("SK02")) = False Then
         strSK02 = rsTmp.Fields("SK02")
      End If
      If IsNull(rsTmp.Fields("SK03")) = False Then
         strSK03 = rsTmp.Fields("SK03")
      End If
   '若無系統類別
   Else
      GetSysTemKind = False
      MsgBox "此資料無相關系統類別!!!", vbExclamation + vbOKOnly
      Call txtSystem_GotFocus
      txtSystem.SetFocus
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Modify By Sindy 2012/3/1 寫成共用函數
Private Sub SetColClearVal(bClear As Boolean)
   txtSystem = Empty
   txtCode(0) = Empty
   txtCode(1) = Empty
   txtCode(2) = Empty
   textCP05 = Empty
   txtCaseProperty = Empty
   lblCasePropertyName = Empty
   txtPetition = Empty
   lblPetitionName = Empty
   'Add by Morgan 2006/6/23
   For intI = 2 To 5
      txtPetitionx(intI) = Empty
      lblPetitionNamex(intI) = Empty
   Next
   fraRecieve.Enabled = False
   fraCode.Visible = True
   If Not bClear Then
      txtSystem.SetFocus
   End If
   fraLastCaseCode.Visible = True
   'Add By Sindy 2012/3/1
   Frame1.Visible = False
   textCP24 = ""
   strMailNote = ""
   strTo = ""
   textCP05.Enabled = True
   '2012/3/1 End
End Sub

'Add By Sindy 2012/2/20
Private Sub T728Progress()
Dim frm As Form
   
   If textCP24 = "1" Or textCP24 = "2" Then
      m_CP01 = txtSystem
      m_CP02 = IIf(txtSystem <> "TF", txtCode(0), txtTFCode(0) & txtTFCode(1))
      m_CP03 = IIf(txtSystem <> "TF", txtCode(1), txtTFCode(2))
      m_CP03 = m_CP03 & String(1 - Len(m_CP03), "0")
      m_CP04 = IIf(txtSystem <> "TF", txtCode(2), txtTFCode(3))
      m_CP04 = m_CP04 & String(2 - Len(m_CP04), "0")
      m_CP13 = PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)
      m_CP13_2 = GetStaffName(m_CP13)
      m_CP12 = GetST15(m_CP13)
      
      '輸入核准時,依系統別帶商標基本檔或服務業務基本檔維護畫面供使用者補資料
      If textCP24 = "1" Then
         If MsgBox("確定是否要新增轉案至他所核准的進度資料？(請務必將基本資料補齊)", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
             Exit Sub
         End If
         If OnSaveTx728Data = False Then Exit Sub
         Call SetColClearVal(False)
         
         'Modify By Sindy 2014/5/27
         'Call mdiMain.SysCallSpecForm("frm010001", m_CP01, m_CP02, m_CP03, m_CP04)
         Call Forms(0).SysCallSpecForm("frm010001", m_CP01, m_CP02, m_CP03, m_CP04)
         '2014/5/27 END
         
      ElseIf textCP24 = "2" Then
         '讀取基本檔資料
         Select Case m_CP01
            Case "T", "TF", "CFT", "FCT":
               strSql = "select NVL(TM05,NVL(TM06,TM07)),tm23,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),tm10,na03 " & _
                        "from trademark,nation,customer " & _
                        "where tm01='" & m_CP01 & "' and tm02='" & m_CP02 & "' and tm03='" & m_CP03 & "' and tm04='" & m_CP04 & "' " & _
                        "and tm10=na01(+) " & _
                        "and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
            Case Else:
               strSql = "select NVL(sp05,NVL(sp06,sp07)),sp08,NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),sp09,na03 " & _
                        "From servicepractice, nation, customer " & _
                        "where sp01='" & m_CP01 & "' and sp02='" & m_CP02 & "' and sp03='" & m_CP03 & "' and sp04='" & m_CP04 & "' " & _
                        "and sp09=na01(+) " & _
                        "and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) "
         End Select
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strCaseName = "" & RsTemp.Fields(0)
            strApp = "" & RsTemp.Fields(1) & " " & "" & RsTemp.Fields(2)
            strNation = "" & RsTemp.Fields(3) & " " & "" & RsTemp.Fields(4)
         End If
         
         'E-Mail主旨
         oSubject = m_CP01 & "-" & m_CP02 & IIf(m_CP03 & m_CP04 = "000", "", "-" & m_CP03 & "-" & m_CP04) & "已核駁且已轉他所辦理，案件已做閉卷處理"
         
         '開啟轉案至他所結果輸入
         frm010012_09.textTM05 = strCaseName
         frm010012_09.LabApp = strApp
         frm010012_09.LabNation = strNation
         If strTo <> "" Then
            frm010012_09.textCP13 = strTo
            Dim strTemp As String, strTemp1 As String, strST15 As String
            Call ClsPDGetStaff(strTo, strTemp, strTemp1)
            strST15 = GetST15(strTo, strTemp1)
            frm010012_09.LabSales = strTemp
            frm010012_09.LabDept = strST15 & " " & strTemp1
         Else
            frm010012_09.textCP13 = m_CP13
            frm010012_09.LabSales = m_CP13_2
            frm010012_09.LabDept = m_CP12 & " " & GetDepartmentName(m_CP12)
         End If
         frm010012_09.LabSubject = oSubject
         frm010012_09.TextNote = strMailNote
         frm010012_09.Show
         Me.Hide
      End If
   End If
End Sub

'Add By Sindy 2012/3/1
Public Sub CP24_2_T728progress()
   strMailNote = frm010012_09.TextNote
   strTo = frm010012_09.textCP13
   '按[回前畫面]
   If frm010012_09.bolOK = False Then
      Unload frm010012_09
      Set frm010012_09 = Nothing
      Me.Show
      Exit Sub
   '按[確定]
   Else
      Unload frm010012_09
      Set frm010012_09 = Nothing
      Me.Show
      
      If OnSaveTx728Data = False Then Exit Sub
      '發E-Mail通知智權人員
      oContext = "案件名稱：" & strCaseName & vbCrLf & _
                 "申請人　：" & strApp & vbCrLf & _
                 "申請國家：" & strNation & vbCrLf
      If strMailNote <> "" Then
         oContext = oContext & vbCrLf & vbCrLf & "備　　註：" & strMailNote & vbCrLf
      End If
      PUB_SendMail strUserNum, strTo, "", oSubject, oContext
   End If
   Call SetColClearVal(False)
End Sub

'Add By Sindy 2012/2/23
Private Function OnSaveTx728Data() As Boolean

On Error GoTo ErrorHandler
   
   OnSaveTx728Data = True
   
   Screen.MousePointer = vbHourglass
   cnnConnection.BeginTrans
   
   '下一程序檔有未續辦的305,997,998期限的總收文號時
   strSql = "select * from nextprogress " & _
           "where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' " & _
           "and np06 is null and np07 in ('305','997','998') "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         '逐筆更新總收文號的進度檔
         strSql = "update caseprogress set cp24='" & textCP24 & "',cp25=" & strSrvDate(1) & " where cp09='" & .Fields("np01") & "'"
         cnnConnection.Execute strSql, intI
         '逐筆更新總收文號的下一程序檔
         strSql = "update nextprogress set np06='Y'" & _
                  " where np01='" & .Fields("np01") & "' and np07='" & .Fields("np07") & "' and np22='" & .Fields("np22") & "' "
         cnnConnection.Execute strSql, intI
         .MoveNext
      Loop
      End With
   End If
   
   '新增B類收文
   m_CP09 = AutoNo("B", 6)
   strSql = "INSERT INTO CaseProgress " & _
            "(CP01,CP02,CP03,CP04,CP09,CP10,CP05,CP12,CP11,CP20,CP32,CP13,CP14,CP26,CP24,CP27) " & _
            "VALUES ('" & m_CP01 & "','" & m_CP02 & "','" & m_CP03 & "','" & m_CP04 & "','" & m_CP09 & _
            "','" & txtCaseProperty & "'," & strSrvDate(1) & ",'" & m_CP12 & "'" & _
            ",'90','N','N','" & m_CP13 & "'" & _
            ",'" & strUserNum & "','N','" & textCP24 & "'," & strSrvDate(1) & ")"
   cnnConnection.Execute strSql
   
   If textCP24 = "2" Then '駁
      '更新基本檔閉卷
      Select Case m_CP01
         Case "T", "TF", "CFT", "FCT":
            strSql = "UPDATE TRADEMARK SET TM29='Y',TM30=" & strSrvDate(1) & ",TM31='02',TM16='2' " & _
                     "WHERE TM01 = '" & m_CP01 & "' AND " & _
                           "TM02 = '" & m_CP02 & "' AND " & _
                           "TM03 = '" & m_CP03 & "' AND " & _
                           "TM04 = '" & m_CP04 & "' "
         Case Else:
            strSql = "UPDATE SERVICEPRACTICE SET SP15='Y',SP16=" & strSrvDate(1) & ",SP17='02' " & _
                     "WHERE SP01 = '" & m_CP01 & "' AND " & _
                           "SP02 = '" & m_CP02 & "' AND " & _
                           "SP03 = '" & m_CP03 & "' AND " & _
                           "SP04 = '" & m_CP04 & "' "
      End Select
      cnnConnection.Execute strSql
   End If
   
   txtRecieveCode(0) = Mid(m_CP09, 2, 2)
   txtRecieveCode(1) = Mid(m_CP09, 4)
   If txtSystem = "TF" Then
      lblCaseCode = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & Left(m_CP04, 1) & "-" & Mid(m_CP04, 2)
   Else
      lblCaseCode = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
   End If
   
   cnnConnection.CommitTrans
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   OnSaveTx728Data = False
End Function

'Added by Lydia 2018/05/07
Private Sub txtFMP_GotFocus()
    TextInverse txtFMP
End Sub

Private Sub txtFMP_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'end 2018/05/07

'Added by Lydia 2018/05/15
Private Sub txtCode_LostFocus(Index As Integer)
   'P非新申請案,隱藏是否為FMP案
   If Index = 0 Then
       'Modified by Lydia 2018/05/17 北所才顯示
       'If txtSystem & txtCode(0) = "P" Then
       If intChoose = 0 And txtSystem & txtCode(0) = "P" And pub_strUserOffice = "1" Then
           'Modified by Lydia 2021/11/10 改用Frame
           fraFMP.Visible = True
       Else
           'Modified by Lydia 2021/11/10 改用Frame
           fraFMP.Visible = False
       End If
   End If
End Sub

'Added by Lydia 2018/05/15
'Modified by Lydia 2020/04/23 更名：因為與basQuery.PUB_IsFMP同名稱，並且未用在其他程式
'Public Function PUB_IsFMP(pa01 As String, pa02 As String, pa03 As String, pa04 As String) As Boolean
'Removed by Morgan 2021/2/2
'Private Function Check_IsFMP(pa01 As String, pa02 As String, pa03 As String, pa04 As String) As Boolean
'   Dim stSQL As String, intQ As Integer
'   Dim RsQ As ADODB.Recordset
'
'   stSQL = "select cp12 from caseprogress where cp01='" & pa01 & "' and cp02='" & pa02 & "' and cp03='" & pa03 & "' and cp04='" & pa04 & "' order by cp05 desc"
'   intQ = 1
'   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
'   If intQ = 1 Then
'      If Left(RsQ(0), 1) = "F" Then
'         Check_IsFMP = True
'      End If
'   End If
'   Set RsQ = Nothing
'End Function
 
'Added by Lydia 2020/05/20
Private Sub txtLOS15_GotFocus()
   TextInverse txtLOS15
End Sub

'Added by Lydia 2020/05/20 案源單號
Private Sub txtLOS15_Validate(Cancel As Boolean)
   If txtLOS15.Tag <> txtLOS15.Text Then
       If GetStateLOS(txtSystem, txtCaseProperty, txtCode(0), txtLOS15, m_strLOSkind) = False Then
           If m_strLOSkind <> "" Then
                txtLOS15.SetFocus
                txtLOS15_GotFocus
                Cancel = True
                Exit Sub
           End If
       End If
   End If
   txtLOS15.Tag = txtLOS15.Text
End Sub
'Added by Lydia 2020/05/20
Private Sub txtLOS15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2020/05/20 判斷是否為法律所案源收文
Private Function GetStateLOS(ByVal pSYS01 As String, ByVal pCasePty As String, Optional ByVal pSYS02 As String, Optional ByVal pKeyNo As String, Optional ByRef pKeyKind As String) As Boolean
'pSYS01: 系統別、pSYS02: 案件流水號
'pCasePty: 案件性質
'pKeyNo: 案源單號
'pKeyKind: 案源案件類型-->依系統別+案件性質判斷屬於A、B、C那一類
'pStartDay: 啟用日控制
Dim strA1 As String, strB1 As String
Dim intA As Integer
Dim rsAD As New ADODB.Recordset
Dim rsRd As New ADODB.Recordset
Dim pSrcKind As String

   GetStateLOS = False

   If strSrvDate(1) < 法律所案源收文啟用日 Then Exit Function  '啟用日控制
    
   If lblReciveCode.Caption <> "A" Then Exit Function 'Added by Lydia 2020/09/17 內部收文不歸案源
   
   'Modified by Lydia 2020/06/08 +申請國家m_Nation
   pKeyKind = PUB_GetLOSkind(pSYS01, pCasePty, m_Nation)
   pSrcKind = PUB_GetLOSplus(pSYS01, txtCode(0), txtCode(1), txtCode(2), pCasePty, m_Nation, pKeyKind) 'Added by Lydia 2020/06/08 判斷是否為補收文=>案源類別
   
   'Modifie by Lydia 2020/06/08 +準備程序、言詞辯論、訴願
   If pKeyKind & pSrcKind <> "" Then
        'Added by Lydia 2020/06/08 C類:若有其他C類已收未發程序則視為同一案源補收文，不算新案源以一般接洽單處理即可
        If Left(pKeyKind, 1) = "C" And pSrcKind = "" And txtCode(0) <> "" And Trim(pKeyNo) = "" Then
            pKeyKind = ""
            GoTo EXITSUB
        End If
       If Trim(pKeyNo) <> "" Then
            '若輸入案源單號LOS15，需檢查法律所案源檔且必須為無法務收文號LOS06且無放棄日期LOS07的資料；
            strA1 = "select los01,los02,los15,los06, los07 from LawOfficeSource where los15=" & CNULL(Trim(pKeyNo))
            intA = 1
            Set rsAD = ClsLawReadRstMsg(intA, strA1)
            If intA = 0 Then
                  MsgBox "該案源單號不存在！", vbCritical, "檢核案源單號"
                  GoTo EXITSUB
            Else
                  If "" & rsAD.Fields("los06") & rsAD.Fields("los07") <> "" Then
                        strB1 = ""
                        'B1收文時若輸入的"案源單號"已有法律所總收文號但無案源總收文號表示為A轉B1案源。
                        If "" & rsAD.Fields("los01") <> "" And "" & rsAD.Fields("los06") <> "" Then strB1 = strB1 & "、已收文"
                        If "" & rsAD.Fields("los07") <> "" Then strB1 = strB1 & "、已放棄"
                        '若輸入之案源單號已有法務總收文號且為同案號同日收文者，則為同一接洽單之其他性質，可繼續但存檔不必回案源。
                        If "" & rsAD.Fields("los06") <> "" And intCaseKind = 法務 Then
                            strA1 = "select cp01,cp02,cp03,cp04,cp05 from caseprogress where cp09=" & CNULL(rsAD.Fields("los06"))
                            intA = 1
                            Set rsRd = ClsLawReadRstMsg(intA, strA1)
                            If intA = 1 Then
                                '同案號: CP01+CP02 ,不同級訴訟CP03會自動加1
                                If strSrvDate(1) = rsRd.Fields("cp05") And pSYS01 & pSYS02 = rsRd.Fields("cp01") & rsRd.Fields("cp02") Then
                                     strB1 = Replace(strB1, "、已收文", "") '可收文
                                End If
                            End If
                            'Added by Lydia 2020/06/24 法務補收款78輸入案源單號若為B1類表示為A4轉B1，不必限制同日收文。
                            'Modified by Morgan 2020/9/9
                            'If "" & rsAD.Fields("B1") And txtCaseProperty = "78" Then
                            If "" & rsAD.Fields("los02") = "B1" And txtCaseProperty = "78" Then
                            'end 2020/9/9
                                 strB1 = Replace(strB1, "、已收文", "") '可收文
                            End If
                        End If
                        If strB1 <> "" Then
                            'Added by Lydia 2020/06/15 A類可能有一筆以上的收文
                            If strB1 = "、已收文" And Left("" & rsAD.Fields("los02"), 1) = "A" Then
                               If MsgBox("該案源單號已收文，是否再次收文？", vbExclamation + vbYesNo + vbDefaultButton2, "檢核案源單號") = vbNo Then
                                   GoTo EXITSUB
                               End If
                            Else
                            'end 2020/06/15
                               MsgBox "該案源單號" & Mid(strB1, 2) & "，不可收文！", vbCritical, "檢核案源單號"
                               GoTo EXITSUB
                            End If 'Added by Lydia 2020/06/15
                        End If
                  Else
                        If pSrcKind <> "D" Then 'Added by Lydia 2020/06/08 排除準備程序、言詞辯論、訴願
                            '判斷：B類前2碼,C類前1碼
                            If (Trim(Left(pKeyKind, 1)) = "B" And Trim(Left(pKeyKind, 2)) <> Trim(Left("" & rsAD.Fields("los02"), 2))) Or _
                               Trim(Left(pKeyKind, 1)) = "C" And Trim(Left(pKeyKind, 1)) <> Trim(Left("" & rsAD.Fields("los02"), 1)) Then
                                MsgBox "該案源單號的案件類型與輸入畫面的案件性質不同，請重新確認接洽單！", vbCritical, "檢核案源單號"
                                GoTo EXITSUB
                            '判斷：BC類的案件系統別
                            ElseIf Trim(Left(pKeyKind, 1)) <> "A" And ((Right(pKeyKind, 1) = "P" And InStr(txtSystem, "P") = 0) Or (Right(pKeyKind, 1) = "T" And InStr(txtSystem, "T") = 0)) Then
                                MsgBox "該案源單號的案件類型與輸入畫面的案件性質不同，請重新確認接洽單！", vbCritical, "檢核案源單號"
                                GoTo EXITSUB
                            End If
                        End If 'Added by Lydia 2020/06/08
                  End If
            End If
       End If
       If pKeyKind = "" And pSrcKind <> "" Then pKeyKind = pSrcKind  'Added by Lydia 2020/06/08 準備程序、言詞辯論、訴願
       
       GetStateLOS = True
   End If

EXITSUB:
    Set rsAD = Nothing
    Set rsRd = Nothing
End Function

'Added by Lydia 2020/11/19
Private Sub txtCaseNa239_GotFocus()
   TextInverse txtCaseNa239
End Sub

'Added by Lydia 2020/11/19
Private Sub txtCaseNa239_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2020/11/19 CFP和CFT英國脫歐案管制：歐盟案案號
Private Sub txtCaseNa239_Validate(Cancel As Boolean)
   
   If txtCaseNa239.Tag <> txtCaseNa239.Text Then
      If txtCaseNa239.Text <> "" Then
          If InStr("CFP,CFT,", Left(txtCaseNa239, 3)) = 0 Or Left(txtCaseNa239, 3) <> txtSystem Then
              MsgBox "請輸入" & txtSystem & "案！", vbCritical, "檢核資料"
              Cancel = True
              GoTo EXITSUB
          End If
          strExc(0) = Left(txtCaseNa239.Text & String(8, "0"), 12)
          Call ChgCaseNo(strExc(0), strExc)
          If Len("" & strExc(2)) <> 6 Then
              MsgBox "請輸入正確的歐盟案案號！", vbCritical, "檢核資料"
              Cancel = True
              GoTo EXITSUB
          Else
              'Added by Lydia 2021/03/05 CFT歐盟尚未註冊案轉換英國申請案收文控管
              If txtCaseProperty = "101" Then
                  'Modifed by Lydia 2021/03/11 改條件：符合「歐盟商標申請日為2021.1.1前」及「2021.1.1前歐盟商標無發證日」
                  'strSql = "Select Tm01,Tm02,Tm03,Tm04 From Trademark Where Tm01='CFT' And Tm10='239' And Tm11 < 20210101 and (tm13 is null or tm13 >=20210101) " & _
                              "and tm01='" & strExc(1) & "' and tm02='" & strExc(2) & "' and tm03='" & strExc(3) & "' and tm04='" & strExc(4) & "' and  tm30||tm57 is null "
                  strSql = "Select Tm01,Tm02,Tm03,Tm04 From Trademark Where Tm01='CFT' And Tm10='239' And Tm11 < 20210101 and (tm20 is null or tm20 >=20210101) " & _
                              "and tm01='" & strExc(1) & "' and tm02='" & strExc(2) & "' and tm03='" & strExc(3) & "' and tm04='" & strExc(4) & "' and  tm30||tm57 is null "
              Else
              'end 2021/03/05
                  strSql = "SELECT PA01,PA02,PA03,PA04 FROM PATENT WHERE PA01='" & strExc(1) & "' AND PA02='" & strExc(2) & "' AND PA03='" & strExc(3) & "' AND PA04='" & strExc(4) & "' AND PA09='239' " & _
                              "UNION ALL SELECT TM01,TM02,TM03,TM04 FROM TRADEMARK WHERE TM01='" & strExc(1) & "' AND TM02='" & strExc(2) & "' AND TM03='" & strExc(3) & "' AND TM04='" & strExc(4) & "' AND TM10='239' "
              End If 'Added by Lydia 2021/03/05
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
              If intI = 0 Then
                  MsgBox "請輸入正確的歐盟案案號！", vbCritical, "檢核資料"
                  Cancel = True
                  GoTo EXITSUB
              End If
              txtCaseNa239.Text = strExc(0)
          End If
      End If
   End If
   
   txtCaseNa239.Tag = txtCaseNa239.Text
   
   Exit Sub
   
EXITSUB:
   txtCaseNa239.SetFocus
   txtCaseNa239_GotFocus
End Sub

'Added by Lydia 2021/11/10
Private Sub txtNA01_Validate(Cancel As Boolean)
Dim strTemp As String
   If Trim(txtNA01) <> "" Then
      lblNation.Caption = ""
      If ClsPDGetNation(txtNA01.Text, strTemp) Then
         lblNation.Caption = strTemp
      End If
      If txtNA01 <> "020" And txtNA01 <> "013" And txtNA01 <> "044" Then
           MsgBox "申請國家只可為 香港, 大陸, 澳門！", vbCritical
           txtNA01.SetFocus
           txtNA01_GotFocus
           Cancel = True
      End If
   End If
End Sub
Private Sub txtNA01_GotFocus()
   TextInverse txtNA01
End Sub

'Add by Amy 2021/12/21 確認表單及關閉(改Form2.0後,存檔按Enter會當掉,改在呼叫時清除記憶體變數)
Private Sub ChkAndCloseForm()
    '內部收文表單中TextBox 非陣列較不可能當掉,可先不考慮
    If stChkForm = MsgText(601) Then Exit Sub
    
    If PUB_CheckFormExist(stChkForm) = False Then
        Select Case UCase(stChkForm)
            Case UCase("Frm010004")
                Set frm010004 = Nothing
            Case UCase("Frm010005")
                Set frm010005 = Nothing
            Case UCase("Frm010006")
                Set frm010006 = Nothing
            Case UCase("frm010006_1")
                Set frm010006_1 = Nothing
            Case UCase("Frm010007")
                Set frm010007 = Nothing
        End Select
        stChkForm = ""
    End If
End Sub
'end 2021/12/21


