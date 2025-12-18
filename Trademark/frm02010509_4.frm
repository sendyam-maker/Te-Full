VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010509_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "爭議案部分勝部分敗輸入"
   ClientHeight    =   5748
   ClientLeft      =   96
   ClientTop       =   1008
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8952
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   3990
      TabIndex        =   74
      Top             =   3420
      Width           =   4215
      Begin VB.TextBox Text12 
         Height          =   252
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   11
         Top             =   150
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   840
         MaxLength       =   2
         TabIndex        =   7
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   9
         Top             =   150
         Width           =   375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "文到          天"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "        月"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   8
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                      日"
         Height          =   225
         Index           =   2
         Left            =   2520
         TabIndex        =   10
         Top             =   180
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   1290
      TabIndex        =   73
      Top             =   3420
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "文到當日"
         Height          =   180
         Index           =   0
         Left            =   144
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "文到次日"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   5
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "商品及服務資料輸入(&I)"
      Height          =   375
      Left            =   3975
      TabIndex        =   22
      Top             =   30
      Width           =   1965
   End
   Begin VB.TextBox textTM16 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5730
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   735
      Width           =   2532
   End
   Begin VB.TextBox textCF15 
      Height          =   264
      Left            =   1290
      MaxLength       =   4
      TabIndex        =   2
      Top             =   3135
      Width           =   732
   End
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2100
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   3135
      Width           =   1692
   End
   Begin VB.TextBox textCP06 
      Height          =   264
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   12
      Top             =   3930
      Width           =   2532
   End
   Begin VB.TextBox textCP07 
      Height          =   264
      Left            =   5730
      MaxLength       =   7
      TabIndex        =   13
      Top             =   3930
      Width           =   2532
   End
   Begin VB.TextBox TextCP64_1 
      Height          =   264
      Left            =   5730
      MaxLength       =   40
      TabIndex        =   1
      Top             =   2835
      Width           =   2532
   End
   Begin VB.TextBox textCP26_S 
      Height          =   264
      Left            =   7800
      MaxLength       =   1
      TabIndex        =   18
      Top             =   4515
      Width           =   372
   End
   Begin VB.TextBox textCP40 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   1665
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   735
      Width           =   2532
   End
   Begin VB.TextBox textTM32 
      Height          =   300
      Left            =   1290
      MaxLength       =   1500
      TabIndex        =   19
      Top             =   4785
      Width           =   7575
   End
   Begin VB.TextBox textTM17 
      Height          =   264
      Left            =   4830
      MaxLength       =   1
      TabIndex        =   17
      Top             =   4515
      Width           =   372
   End
   Begin VB.TextBox textCP26 
      Height          =   264
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   16
      Top             =   4515
      Width           =   372
   End
   Begin VB.TextBox textCP14 
      Height          =   264
      Left            =   1290
      MaxLength       =   6
      TabIndex        =   14
      Top             =   4230
      Width           =   732
   End
   Begin VB.TextBox textCP48 
      Height          =   264
      Left            =   5730
      MaxLength       =   7
      TabIndex        =   15
      Top             =   4230
      Width           =   1095
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   435
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1965
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5730
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1665
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5730
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2565
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5730
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2265
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2265
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   435
      Width           =   2532
   End
   Begin VB.TextBox textCP08 
      Height          =   264
      Left            =   1290
      MaxLength       =   40
      TabIndex        =   0
      Top             =   2835
      Width           =   2532
   End
   Begin VB.TextBox textCP35 
      Height          =   264
      Left            =   5730
      MaxLength       =   32
      TabIndex        =   3
      Top             =   3120
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8025
      TabIndex        =   25
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   5955
      TabIndex        =   23
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Left            =   6780
      TabIndex        =   24
      Top             =   30
      Width           =   1200
   End
   Begin MSForms.TextBox textTM58 
      Height          =   300
      Left            =   1290
      TabIndex        =   21
      Top             =   5385
      Width           =   7575
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13361;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_2 
      Height          =   264
      Left            =   2070
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4230
      Width           =   1692
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   300
      Left            =   1290
      TabIndex        =   20
      Top             =   5085
      Width           =   7575
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13361;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_Src 
      Height          =   264
      Left            =   1290
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2565
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   300
      Left            =   1290
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1335
      Width           =   7575
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13361;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5730
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1965
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1260
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1005
      Width           =   7575
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13361;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label32 
      Caption         =   "來函期限:"
      Height          =   255
      Left            =   210
      TabIndex        =   75
      Top             =   3570
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "目前准駁 :"
      Height          =   255
      Left            =   4770
      TabIndex        =   72
      Top             =   735
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "下一程序 :"
      Height          =   255
      Left            =   210
      TabIndex        =   70
      Top             =   3135
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "本所期限 :"
      Height          =   255
      Left            =   210
      TabIndex        =   69
      Top             =   3930
      Width           =   855
   End
   Begin VB.Label Label25 
      Caption         =   "法定期限 :"
      Height          =   255
      Left            =   4770
      TabIndex        =   68
      Top             =   3930
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "來文字號 :"
      Height          =   255
      Left            =   4770
      TabIndex        =   66
      Top             =   2850
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "案件備註 :"
      Height          =   255
      Left            =   210
      TabIndex        =   63
      Top             =   5385
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "商品群組 :"
      Height          =   255
      Left            =   210
      TabIndex        =   62
      Top             =   4785
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   210
      TabIndex        =   61
      Top             =   735
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "(Y / N)"
      Height          =   255
      Left            =   8220
      TabIndex        =   60
      Top             =   4515
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "是否計算勝訴率 :"
      Height          =   255
      Left            =   6270
      TabIndex        =   59
      Top             =   4515
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "(Y / N)"
      Height          =   255
      Left            =   5250
      TabIndex        =   58
      Top             =   4515
      Width           =   615
   End
   Begin VB.Label Label19 
      Caption         =   "專用權是否存在 :"
      Height          =   255
      Left            =   3300
      TabIndex        =   57
      Top             =   4515
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不算)"
      Height          =   255
      Left            =   1920
      TabIndex        =   56
      Top             =   4515
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "是否算案件數 :"
      Height          =   255
      Left            =   210
      TabIndex        =   55
      Top             =   4515
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   210
      TabIndex        =   54
      Top             =   4230
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "對照名稱 :"
      Height          =   255
      Left            =   210
      TabIndex        =   52
      Top             =   1665
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   255
      Left            =   210
      TabIndex        =   51
      Top             =   5085
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   210
      TabIndex        =   50
      Top             =   2565
      Width           =   855
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   255
      Left            =   4770
      TabIndex        =   49
      Top             =   4230
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   47
      Top             =   435
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   210
      TabIndex        =   46
      Top             =   1035
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   210
      TabIndex        =   45
      Top             =   1335
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   44
      Top             =   1935
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   4770
      TabIndex        =   43
      Top             =   1665
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   255
      Index           =   8
      Left            =   4770
      TabIndex        =   42
      Top             =   2565
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   255
      Index           =   9
      Left            =   4770
      TabIndex        =   41
      Top             =   2265
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   210
      TabIndex        =   40
      Top             =   2265
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4770
      TabIndex        =   39
      Top             =   1965
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "審定號 :"
      Height          =   255
      Left            =   4770
      TabIndex        =   38
      Top             =   435
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   255
      Left            =   210
      TabIndex        =   37
      Top             =   2835
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "審查委員 :"
      Height          =   255
      Left            =   4770
      TabIndex        =   36
      Top             =   3120
      Width           =   855
   End
End
Attribute VB_Name = "frm02010509_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/03 Form2.0已修改 cmbTM05/textTM23/textCP13/textCP14_Src/textCP14_2/textCP64/textTM58/grdList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
'2009/4/16 CREATE BY SINDY
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 申請國家
Dim m_TM10 As String
Dim m_TM09 As String 'Add By Sindy 2011/6/29
' 來函收文日
Dim m_CP05 As String
' 機關文號
Dim m_CP08 As String
' 所選取的收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
' 對照號數
Dim m_CP36 As String
' 對照案件名稱(中)
Dim m_CP37 As String
' 對照案件名稱(英)
Dim m_CP38 As String
' 對照案件名稱(日)
Dim m_CP39 As String
' 對照名稱(中)
Dim m_cp40 As String
' 對照名稱(英)
Dim m_CP41 As String
' 對照名稱(日)
Dim m_CP42 As String
' 預估結果
Dim m_CP23 As String

Dim m_intNumBegin As Integer
Dim m_intNumEnd As Integer
Dim m_strNumBegin As String
Dim m_strNumEnd As String
'Add By Sindy 2011/6/29 檢查是否已經有商品及服務
Public ChkTG As Boolean
Dim BolPrintCaseCheck As Boolean 'Add By Sindy 2012/4/16
'Added by Morgan 2017/4/27 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/4/27
'Add By Sindy 2019/5/27
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/27 END
Dim strLD18 As String 'Add By Sindy 2020/1/7 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2020/1/7 FC代理人
Dim m_TM23 As String 'Add By Sindy 2020/1/7 申請人
Dim m_TM28 As String 'add by sonia 2021/5/3 卷宗性質

'Add By Sindy 2019/5/27
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010509_3.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm02010509_3
   Unload frm02010509_2
   Unload frm02010509_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   BolPrintCaseCheck = CaseCheck(m_TM01, m_TM02, m_TM03, m_TM04, m_TM10)
   
   If CheckDataValid() = True Then
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      'Pub_EndModCashMsg m_TM10   '2009/11/11 CANCEL BY SONIA取消結餘詢問
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      'Add By Sindy 2012/4/16 列印帳款未結清案件資料
      If BolPrintCaseCheck = True Then
          Call GetPrintCaseCheck(m_CP09)
      End If
      '2012/4/16 End
      Call PUB_ChkTemporaryReceipts(m_TM01, m_TM02, m_TM03, m_TM04) 'Add By Sindy 2014/5/28 檢查是否有暫收款
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010509_3
      Unload frm02010509_2
      'Add By Sindy 2019/5/27
      If Me.m_strIR01 <> "" Then
        Unload frm02010509_1
        If Not m_PrevForm Is Nothing Then
           Call m_PrevForm.GoNext
        End If
        Unload Me
      '2019/5/27 END
      'Modified by Morgan 2017/4/27 電子公文
      'frm02010509_1.Show
      ElseIf m_DocNo <> "" Then
         Unload Me
         Unload frm02010509_1
         frm02010412.GoNext
      Else
         frm02010509_1.Show
         Unload Me
      End If
      'end 2017/4/27
   End If
End Sub

'Add By Sindy 2011/6/29
Private Sub Command2_Click()
   frm03010303_04.Hide
   Set frm03010303_04.UpForm = Me
   frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   frm03010303_04.AllClass = m_TM09
   frm03010303_04.cmdok(2).Visible = True
   
   If m_TM09 <> "" Then  '有商品類別才可進入 T-113511團體標章
      Me.Hide
      frm03010303_04.QueryData
      frm03010303_04.Show vbModal '強制回應表單
   Else
      MsgBox ("無商品類別，不可使用此按鈕 !")
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM16.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_Src.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP40.BackColor = &H8000000F
   textCF15_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/27
   m_strIR01 = frm02010509_1.m_strIR01
   m_strIR02 = frm02010509_1.m_strIR02
   m_strIR03 = frm02010509_1.m_strIR03
   m_strIR04 = frm02010509_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/27 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
   End Select
End Sub

' 取得商標基本檔
Private Sub QueryTradeMark()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      'Add By Sindy 2011/6/29 商品類別
      m_TM09 = ""
      If IsNull(rsTmp.Fields("TM09")) = False Then
         m_TM09 = rsTmp.Fields("TM09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
         m_TM23 = rsTmp.Fields("TM23")
      End If
      
      'Add By Sindy 2020/1/7
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("TM44")) = False Then
         m_TM44 = rsTmp.Fields("TM44")
      End If
      '2020/1/7 END
      
      'add by sonia 2021/5/3 卷宗性質
      m_TM28 = Empty
      If IsNull(rsTmp.Fields("TM28")) = False Then
         m_TM28 = rsTmp.Fields("TM28")
      End If
      'end 2021/5/3
      
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 目前准駁
      If IsNull(rsTmp.Fields("TM16")) = False Then
         Select Case rsTmp.Fields("TM16")
            Case 1: textTM16 = "准"
            Case 2: textTM16 = "駁"
         End Select
      End If
      ' 專用權是否存在
      If IsNull(rsTmp.Fields("TM17")) = False Then
         textTM17 = rsTmp.Fields("TM17")
      End If
      If IsNull(rsTmp.Fields("TM28")) = False Then
         If rsTmp.Fields("TM28") <> "1" Then
            textTM17 = "N"
         End If
      End If
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then
         textTM45 = rsTmp.Fields("TM45")
      End If
      ' 商品組群
      If IsNull(rsTmp.Fields("TM32")) = False Then
         textTM32 = rsTmp.Fields("TM32")
      End If
      ' 案件備註
      If IsNull(rsTmp.Fields("TM58")) = False Then
         textTM58 = rsTmp.Fields("TM58")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bCP40 As Boolean
   m_TM10 = Empty
   m_CP13 = Empty
   m_CP12 = Empty
   
   m_CP36 = Empty
   m_CP37 = Empty
   m_CP38 = Empty
   m_CP39 = Empty
   m_cp40 = Empty
   m_CP41 = Empty
   m_CP42 = Empty
   m_CP23 = Empty
   
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   ' 來函收文日
   textCP05S = m_CP05
   ' 收文號
   textCP09 = m_CP09
   
   ' 取得商標基本檔的相關項目
   QueryTradeMark
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 取得案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' AND " & _
                  "CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then
         m_CP08 = rsTmp.Fields("CP08")
      End If
      ' 案件性質
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      '業務區   nick 91.08.22
      If IsNull(rsTmp.Fields("cp12")) = False Then
        m_CP12 = rsTmp.Fields("cp12")
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = rsTmp.Fields("CP14")
         textCP14_Src = GetStaffName(rsTmp.Fields("CP14"))
         textCP14_2 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 對照名稱 (無中文取英文, 無英文取日文)
      bCP40 = False
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               textCP40 = rsTmp.Fields("CP40")
               bCP40 = True
            End If
         End If
      End If
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               textCP40 = rsTmp.Fields("CP41")
               bCP40 = True
            End If
         End If
      End If
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               textCP40 = rsTmp.Fields("CP42")
               bCP40 = True
            End If
         End If
      End If
      ' 預估結果
      If IsNull(rsTmp.Fields("CP23")) = False Then
         m_CP23 = rsTmp.Fields("CP23")
      End If
      ' 程式存檔用資料
      ' 對造號數
      If IsNull(rsTmp.Fields("CP36")) = False Then
         m_CP36 = rsTmp.Fields("CP36")
      End If
      ' 對造案件名稱(中)
      If IsNull(rsTmp.Fields("CP37")) = False Then
         m_CP37 = rsTmp.Fields("CP37")
      End If
      ' 對造案件名稱(英)
      If IsNull(rsTmp.Fields("CP38")) = False Then
         m_CP38 = rsTmp.Fields("CP38")
      End If
      ' 對造案件名稱(日)
      If IsNull(rsTmp.Fields("CP39")) = False Then
         m_CP39 = rsTmp.Fields("CP39")
      End If
      ' 對造名稱(中)
      If IsNull(rsTmp.Fields("CP40")) = False Then
         m_cp40 = rsTmp.Fields("CP40")
      End If
      ' 對造名稱(英)
      If IsNull(rsTmp.Fields("CP41")) = False Then
         m_CP41 = rsTmp.Fields("CP41")
      End If
      ' 對造名稱(日)
      If IsNull(rsTmp.Fields("CP42")) = False Then
         m_CP42 = rsTmp.Fields("CP42")
      End If
   End If
   rsTmp.Close
   
   Call ChgType 'Add By Sindy 2012/4/17 讀取來函期限
   
   ' 預設承辦期限
   ' 承辦期限的日期應為來函收文日加上工作天數
   ' 工作天數由系統別+國家代碼+案件性質(勝訴)搜尋案件收費表的工作天數
   ' 若有值才預設
   
   textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1006", DBDATE(m_CP05), DBDATE(textCP06), textCP09))
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
   
   Set rsTmp = Nothing

   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   
   If m_TM10 < "010" Then    'modify by sonia 2016/11/30 台灣案才預設機關文號,從下面移上來 T-197970
      If TextCP64_1 = "" Then
         TextCP64_1 = "（" & strTmp & "）智商字第號"
      End If
      m_strNumBegin = "商"
      m_strNumEnd = "字"
   
   'If m_TM10 < "010" Then   'cancel by sonia 2016/11/30改至上面
      If textCP08 = "" Then
         Select Case m_CP10
            Case "601", "602"
               textCP08 = "中台異字第G號"
            Case "603", "604"
               textCP08 = "中台評字第H號"
            Case "605", "606"
               '2015/2/4 modify by sonia
               'textCP08 = "中台處字第L號"
               textCP08 = "中台廢字第L號"
            Case "401" '訴願
               textCP08 = "經訴字第號"
            Case "403"
               textCP08 = strTmp & "年度訴字第號"
         End Select
      End If
   End If
   
   'Added by Morgan 2017/4/27 電子公文
   If m_DocNo <> "" Then
      If m_DocWord <> "" Then
         TextCP64_1 = m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號"
      Else
         TextCP64_1 = Replace(TextCP64_1, "第號", "第" & PUB_GetEDocNo(m_DocNo) & "號")
      End If
      textCP64_1_LostFocus
      '期限
      If m_DeadLine <> "" Then
         Option1(1).Value = True
         If Len(m_DeadLine) >= 7 Then
            Option4(2).Value = True
            Text12 = m_DeadLine
            Text12_Validate False
         ElseIf Right(m_DeadLine, 1) = "日" Then
            Option4(0).Value = True
            Text10 = Val(m_DeadLine)
            Text10_Validate False
         ElseIf Right(m_DeadLine, 1) = "月" Then
            Option4(1).Value = True
            Text11 = Val(m_DeadLine)
            Text11_Validate False
         End If
      End If
   End If
   'end 2017/4/27
   
   textCP08_GotFocus
   
End Sub

Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim strSql As String
   Dim strSubTMSQL As String
   Dim bUpdate As Boolean
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP27 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strNP22 As String
   Dim strCP06 As String
   Dim strCP07 As String
   'Add by Amy 2017/11/13
   Dim m_CP06 As String, m_CP07 As String, st_CP09 As String, m_CP14 As String, strMsg As String
   Dim bolUpdCP As Boolean '是否更新進度檔

On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   strSubTMSQL = "WHERE TM01 = '" & m_TM01 & "' AND " & _
                       "TM02 = '" & m_TM02 & "' AND " & _
                       "TM03 = '" & m_TM03 & "' AND " & _
                       "TM04 = '" & m_TM04 & "' "
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 判斷是否更新實際結果 (無實際結果才更新)
   bUpdate = True
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CP24")) = False Then
         If IsEmptyText(rsTmp.Fields("CP24")) = False Then
            bUpdate = False
         End If
      End If
   End If
   rsTmp.Close
   ' 更新原案件進度檔的收文資料其實際結果為准, 准駁日為來函收文日, 審查委員
   'Modify By Sindy 2024/11/1 2024/8/2進度維護已開放可以輸入為3=部分勝敗
   '                          此處也要更新為3
   If bUpdate = True Then
      strSql = "UPDATE CaseProgress SET CP24 = '3', " & _
                                       "CP25 = " & DBDATE(m_CP05) & ", " & _
                                       "CP35 = '" & textCP35 & "' " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 依是否計算勝訴率來更新原案件進度檔資料的是否算案件數欄位
   Select Case textCP26_S
      Case "Y":
         strSql = "UPDATE CaseProgress SET CP26 = NULL " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         cnnConnection.Execute strSql
      Case "N":
         strSql = "UPDATE CaseProgress SET CP26 = 'N' " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
         cnnConnection.Execute strSql
   End Select
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新商標基本檔的專用權是否存在, 商品組群, 案件備註欄
   strSql = "UPDATE TradeMark SET TM17 = '" & textTM17 & "', " & _
                                 "TM32 = '" & textTM32 & "', " & _
                                 "TM58 = '" & ChgSQL(textTM58) & "' "
   strSql = strSql & strSubTMSQL
   cnnConnection.Execute strSql
   
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   
   strCP06 = Empty
   strCP07 = Empty
   If IsEmptyText(textCP06) = False Then: strCP06 = DBDATE(textCP06)
   If IsEmptyText(textCP07) = False Then: strCP07 = DBDATE(textCP07)
   ' 案件性質為部分勝部分敗
   strCP10 = "1006"
   

   Dim strCP64 As String
   
   strCP64 = Trim(textCP64)
   If strCP64 <> "" And Trim(TextCP64_1) <> "" Then
      strCP64 = strCP64 & ",來文字號：" & Trim(TextCP64_1)
   ElseIf Trim(TextCP64_1) <> "" Then
      strCP64 = "來文字號：" & Trim(TextCP64_1)
   End If
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP48,CP64) " & _
                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                          "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP14 & "'," & _
                          "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
                          "'" & textCP35 & "','" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
                          "'" & m_cp40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "'," & DBDATE(textCP48) & "," & _
                          "'" & ChgSQL(strCP64) & "')"
   cnnConnection.Execute strSql
   
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
      
   ' 本所期限
   If IsEmptyText(strCP06) = False Then
      strSql = "UPDATE CaseProgress SET CP06 = " & strCP06 & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   ' 法定期限
   If IsEmptyText(strCP07) = False Then
      strSql = "UPDATE CaseProgress SET CP07 = " & strCP07 & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2012/4/26 儲存官方發文日及官方期限月數
    If Trim(Text11) <> "" Then
      strSql = "UPDATE CaseProgress SET CP133=" & DBDATE(m_CP05) & ",CP134=" & Text11 & " " & _
               "WHERE CP09='" & strCP09 & "' "
      cnnConnection.Execute strSql
    End If
    
   'Added by Lydia 2025/09/12 TF基礎案號設定：基礎案狀態通知Email
   Dim strTFcase As String
   If m_TM01 = "T" Then
      strTFcase = PUB_GetTFbaseInfo(m_TM01, m_TM02, m_TM03, m_TM04, textTM15, m_TM10, "2", textTM12, strCP09)
   End If
   'end 2025/09/12
   
   'Add By Sindy 2019/12/20 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      If Val(textCP06) > 0 Then '有期限者,為掛號
         PUB_AddLetterProgress strLD18, 1, False, "", True, m_TM23, strCP10, m_TM44
      Else
         PUB_AddLetterProgress strLD18, 1, False, "", False, m_TM23, strCP10, m_TM44
      End If
   End If
   '2019/12/20 END
   
   If m_TM01 = "FCT" And Trim(textCP48) = "" Then
        If Trim(textCP07) = "" Then
            strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(4, DBDATE(m_CP05), 0)) & " " & _
                     "WHERE CP09 = '" & strCP09 & "' "
            cnnConnection.Execute strSql
        Else
            If DateDiff("d", ChangeWStringToWDateString(DBDATE(m_CP05)), ChangeWStringToWDateString(DBDATE(textCP07))) <= 30 Then    '無法與上句合併，因為沒有日期時，datediff  會發生  型態不符 的錯誤
                strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(4, DBDATE(m_CP05), 0)) & " " & _
                         "WHERE CP09 = '" & strCP09 & "' "
                cnnConnection.Execute strSql
            Else
                strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(6, DBDATE(m_CP05), 0)) & " " & _
                         "WHERE CP09 = '" & strCP09 & "' "
                cnnConnection.Execute strSql
            End If
        End If
    End If
    
   If m_TM01 = "FCT" And (m_CP10 = "601" Or m_CP10 = "603" Or m_CP10 = "605") Then
      strSql = "UPDATE CaseProgress SET CP20 = NULL " & _
               "WHERE CP09 = '" & strCP09 & "'"
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '將下一程序為催審的資料, 更新其是否續辦欄位"Y"
   strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP07 = '" & "305" & "' "
   cnnConnection.Execute strSql
      
   ' 若有輸入下一程序時, 新增資料到下一程序檔
   strNP22 = GetNextProgressNo()
   If IsEmptyText(textCF15) = False Then
    'Modify by Amy 2017/11/13 +if 判斷進度檔已有相同未發文未取消收文之案件性質,則判斷是否更新本限及法限
    If ChkSameCaseProgress(m_TM01, m_TM02, m_TM03, m_TM04, textCF15, m_CP06, m_CP07, st_CP09, m_CP14) = True Then
      If m_CP06 = MsgText(601) Or m_CP07 = MsgText(601) Then
        If MsgBox("下一程序已收文但無期限，是否要代入新期限？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            bolUpdCP = True
        End If
      ElseIf Val(textCP06) + 19110000 <> Val(m_CP06) Or Val(textCP07) + 19110000 <> Val(m_CP07) Then
        strMsg = "下一程序已收文且期限不同" & vbCrLf & _
                 "已收文本所期限：" & IIf(m_CP06 <> "", Val(m_CP06) - 19110000, "") & " 來函本所期限：" & textCP06 & vbCrLf & _
                 "已收文法定期限：" & IIf(m_CP07 <> "", Val(m_CP07) - 19110000, "") & " 來函法定期限：" & textCP07 & vbCrLf
        
        If MsgBox(strMsg & "是否要更新為來函期限？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            bolUpdCP = True
        End If
      End If
    End If
      
    '更新進度檔,並發Mail通知承辦人
    If bolUpdCP = True Then
        strSql = "Update CaseProgress Set CP06=" & Val(textCP06) + 19110000 & ",CP07=" & Val(textCP07) + 19110000 & " Where CP09='" & st_CP09 & "'"
        cnnConnection.Execute strSql
        
        If m_CP14 = MsgText(601) Then m_CP14 = GetDeptMan("P20") '無承辦人發給P20部門之A0908
        strMsg = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "收到" & "" & GetCaseTypeName(m_TM01, textCF15, IIf(m_TM10 = "000", 0, 1)) & "前已收文,請辦理後續！"
        PUB_SendMail strUserNum, m_CP14, "", strMsg, "本所期限：" & textCP06 & "　　法定期限：" & textCP07
     
    '進度檔未有相同未發文未取消收文之案件性質或上述不更新期限,才新增下一程序
    Else
        strNP14 = Empty
        strNP14 = GetRelatedPerson(m_CP09)
        '智權人員存最近收文A類接洽記錄單的智權人員
        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                            strCP06 & "," & strCP07 & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "'," & strNP22 & ")"
        cnnConnection.Execute strSql
    End If
    'end 2017/11/13
      
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      Select Case textCF15
         Case "102", "105", "702", "708", "305", "998", "997"
         Case Else:
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
            '加入案件回覆單,FCT案不印回覆單
            If m_TM01 <> "FCT" Then
                'Modify by Amy 2017/11/16 未更新進度檔才印回覆單
                If bolUpdCP = False Then
                    Call g_PrtForm001.PrintReturnSheet(strCP09, textCF15, DBDATE(strCP07), False, , , , m_TM01 & m_TM02 & m_TM03 & m_TM04)
                End If
            End If
      End Select
   End If
   
   'Add By Sindy 2009/09/24
   '因為有些來函由內商輸入，內商有自行控管之承辦期限及發文日。改為內商輸入所有C類來函，
   '若業務區為F字頭者，除爭議受理外，自動產生B類收文，案件性質為外商發文722，不上發文日，不向客戶請款
   Dim strCP48 As String, strCP09B As String
   If Left(GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)), 1) = "F" And _
      ((m_TM01 = "T" And m_TM10 = "020") Or (m_TM01 = "FCT" And m_TM10 = "000")) Then
      strCP09B = AutoNo("B", 6)
      '承辦期限為系統日加4個工作天
      strCP48 = DBDATE(Pub_GetHandleDay(m_TM01, m_TM10, "722", strSrvDate(1), , m_CP09))
      '2011/4/28 modify by sonia 智權人員原抓點選收文號之智權人員,改抓該案最後收文在職智權人員
      strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp48,cp20,cp26,cp32,cp43) " & _
                     "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
                     "," & CNULL(m_TM04) & "," & CNULL(strSrvDate(1)) & "," & CNULL(strCP09B) & ",722," & _
                     CNULL(GetSalesArea(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "," & CNULL(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "," & CNULL(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "," & CNULL(strCP48) & ",'N','N','N'," & CNULL(strCP09) & ")"
      cnnConnection.Execute strSql
   End If
   '2009/09/24 End
   
   'Added by Lydia 2017/03/20 T大陸案增加被異議(理由) 、異議答辯、復審、起訴、上訴(1602,602,401,403,408)於勝訴(1003)或部分勝部分敗(1006)時，掛催註冊證期限為8個月。
   'modify by sonia 2021/5/3 加卷宗性質條件T-223726
   If m_TM01 = "T" And m_TM10 = "020" And m_TM28 = "1" And (m_CP10 = "1602" Or m_CP10 = "602" Or m_CP10 = "401" Or m_CP10 = "403" Or m_CP10 = "408") Then
        'modify by sonia 2024/9/12 8個月改6個月，本所期限改為工作日
        strExc(8) = CompDate(1, 6, strSrvDate(1))
        ' 抓承辦人:只抓申請,若申請為B類收文改抓CP31='Y'的A類收文承辦人
        strSql = "SELECT '1' ord ,CP09,CP14,ST04 FROM CaseProgress,STAFF WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' " & _
                  "AND CP10='101' AND SUBSTR(CP09,1,1)='A' AND CP14=ST01(+) " & _
                  "UNION SELECT '2' ord ,CP09,CP14,ST04 FROM CaseProgress,STAFF WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' " & _
                  "AND CP31='Y' AND SUBSTR(CP09,1,1)='A' AND CP14=ST01(+) ORDER BY 1 "
        intI = 1
        strExc(1) = "P2001" '原承辦人若離職改抓P2001
        Set rsTmp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
           If "" & rsTmp.Fields("ST04") <> "2" Then strExc(1) = "" & rsTmp.Fields("CP14")
        End If

        strNP22 = GetNextProgressNo()
        'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
        'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                 "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "', 1701 ," & _
                         strExc(8) & "," & strExc(8) & ",'" & strExc(1) & "'," & strNP22 & ")"
        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                 "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "', 1701 ," & _
                        PUB_GetWorkDay1(strExc(8), True) & "," & strExc(8) & ",'" & strExc(1) & "'," & strNP22 & ")"
        cnnConnection.Execute strSql
   End If
   'end 2017/03/20
   
   'Added by Morgan 2017/4/27 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
   End If
   'end 2017/4/27
   
   'Add by Sindy 2019/5/27
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010509_1"
   End If
   '2019/5/27 END
   
   Set rsTmp = Nothing
   'Pub_UpdateEndModCash m_TM01, m_TM02, m_TM03, m_TM04   '2009/11/11 CANCEL BY SONIA取消結餘詢問

cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2022/01/03檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        GoTo EXITSUB
   End If

   ' 申請國家為台灣時, 機關文號不可為空白
   If m_TM10 < "010" Then
      If IsEmptyText(textCP08) = True Then
         strTit = "檢核資料"
         strMsg = "申請國家為台灣時, 機關文號不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP08.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 申請國家為台灣時且案件性質非(行政訴訟,行政訴訟上訴,行政上訴答辯), 下一程序不可為空白
   If m_TM10 < "010" And m_CP10 <> "403" And m_CP10 <> "408" And m_CP10 <> "410" Then
      If IsEmptyText(textCF15) = True Then
         strTit = "檢核資料"
         strMsg = "申請國家為台灣時, 下一程序不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2012/4/17
   '檢查來函期限--日期
   If m_TM10 = 台灣國家代號 Then
      If Me.Option4(2).Value = True Then
         If Me.Text12.Text = "" Then
            MsgBox "請輸入來函期限!!!", vbExclamation + vbOKOnly
            Me.Text12.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   ' 有輸入下一程序時, 本所期限與法定期限不可為空白
   If IsEmptyText(textCF15) = False Then
      If IsEmptyText(textCP06) = True Or IsEmptyText(textCP07) = True Then
         strTit = "檢核資料"
         strMsg = "有下一程序時, 本所期限與法定期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
      If Me.textCP06.Text <> "" Then
         If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
            MsgBox "本所期限不可小於系統日期!!!", vbExclamation
            Me.textCP06.SetFocus
            textCP06_GotFocus
            GoTo EXITSUB
         End If
      End If
      ' 本所期限必須小於法定期限
      If Val(textCP06) > Val(textCP07) Then
         strTit = "檢核資料"
         strMsg = "本所期限的日期不可超過法定期限的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   'Add By Sindy 2014/8/8
   Else
      If m_TM10 = 大陸國家代號 And m_CP10 <> "601" Then 'add by sonia 2016/11/30 大陸異議601案不掛下一程序T-197970
         strTit = "檢核資料"
         strMsg = "下一程序不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15.SetFocus
         GoTo EXITSUB
      End If    'add by sonia 2016/11/30
   '2014/8/8 END
   End If
   
   ' 承辦期限不可空白
   If IsEmptyText(textCP48) = True Then
      strTit = "檢核資料"
      strMsg = "承辦期限不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP48.SetFocus
      GoTo EXITSUB
   End If
   ' 是否計算勝訴率不可空白
   If Not IsEmptyText(m_CP23) Then
      If IsEmptyText(textCP26_S) = True Then
         strTit = "檢核資料"
         strMsg = "是否計算勝訴率不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP26_S.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 專用權是否存在不可為空白
   If IsEmptyText(textTM17) = True Then
      strTit = "檢核資料"
      strMsg = "專用權是否存在不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM17.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2025/09/12
   
   'Add By Sindy 2019/5/27
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   Set frm02010509_4 = Nothing
End Sub

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCF15) = False Then
      ' 只取得國內的案件性質名稱
      If m_TM10 < "010" Then
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
      Else
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 1)
      End If
      If IsEmptyText(textCF15_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15_GotFocus
      End If
   End If
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/07
      End If
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Cancel = True
         textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

'Add By Sindy 2010/11/26
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TextCP64_1_GotFocus()
   Dim intPos As Integer
   With Me.TextCP64_1
      If Len("" & .Text) > 0 Then
         intPos = InStr("" & .Text, "字")
         If intPos - 1 >= 0 Then
            .SelStart = intPos - 1
            .SelLength = 0
         End If
      End If
   End With
End Sub

Private Sub textCP64_1_LostFocus()
On Error GoTo ErrorHandler
   If Len(Me.TextCP64_1.Text) > 0 Then
      m_intNumBegin = InStr(Me.TextCP64_1.Text, m_strNumBegin)
      m_intNumEnd = InStr(Me.TextCP64_1.Text, m_strNumEnd)
   Else
      m_intNumBegin = 0
      m_intNumEnd = 0
   End If
   If m_intNumBegin < m_intNumEnd Then
      Me.textCP35.Text = Mid(Me.TextCP64_1.Text, m_intNumBegin + 1, (m_intNumEnd - m_intNumBegin - 1))
   End If
   
   Exit Sub
   
ErrorHandler:
      m_intNumBegin = 0
      m_intNumEnd = 0
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14)
      If IsEmptyText(textCP14_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP14_GotFocus
      End If
   End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP26_S_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否計算勝訴率
Private Sub textCP26_S_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP26_S) = False Then
      Select Case textCP26_S
         Case "Y", "N"
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入Y或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_S_GotFocus
      End Select
   End If
End Sub

' 是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
End Sub

' 承辦期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
      End If
   End If
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2020) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註欄位內容太長"
      textCP64_GotFocus
   End If
End Sub

Private Sub textTM17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 案件備註
Private Sub textTM58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM58, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件備註欄位內容太長"
      textTM58_GotFocus
   End If
End Sub

' 專用權是否存在
Private Sub textTM17_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM17) = False Then
      Select Case textTM17
         Case "Y", "N"
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入Y或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM17_GotFocus
      End Select
   End If
End Sub

' 商品組群
Private Sub textTM32_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim lst() As String
   Dim lstCount As Integer
   Dim nIndex As Integer
   Dim bFind As Boolean
   Dim nPos As Integer
   
   lstCount = 0
   
   Cancel = False
   If IsEmptyText(textTM32) = False Then
      ' 檢查欄位是否太長
      If CheckLengthIsOK(textTM32, textTM32.MaxLength) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "商品組群欄位內容太長"
         textTM32_GotFocus
      End If
      
      'Modify By Sindy 2024/4/18 商品組群欄人員貼上資料後將全形或半形的「；」分號，轉為半形的逗號存入TM32。
      textTM32 = Replace(Replace(textTM32, ";", ","), "；", ",")
      '2024/4/18 END
      For nIndex = 1 To GetSubStringCount(textTM32)
         strTemp = GetSubString(textTM32, nIndex)
         
         If IsEmptyText(strTemp) = True Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "商品組群的資料不可有空白的內容"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM32_GotFocus
            GoTo EXITSUB
         End If
         
         bFind = False
         For nPos = 0 To lstCount - 1
            If lst(nPos) = strTemp Then
               bFind = True
               Exit For
            End If
         Next nPos
         
         If bFind = True Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "商品組群的資料不可重覆"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM32_GotFocus
            GoTo EXITSUB
         Else
            ReDim Preserve lst(lstCount + 1)
            lst(lstCount) = strTemp
            lstCount = lstCount + 1
         End If
      Next nIndex
   End If
   
EXITSUB:
   If lstCount > 0 Then: Erase lst
   lstCount = 0
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textTM17_GotFocus()
   InverseTextBox textTM17
End Sub

Private Sub textTM32_GotFocus()
   InverseTextBox textTM32
End Sub

Private Sub textTM58_GotFocus()
   InverseTextBox textTM58
End Sub

Private Sub textCP08_GotFocus()
Dim intPos As Integer
   With Me.textCP08
      '將游標停在某個字的後面
      If Len("" & .Text) > 0 Then
         intPos = InStr("" & .Text, "G")
         If intPos = 0 Then intPos = InStr("" & .Text, "H")
         If intPos = 0 Then intPos = InStr("" & .Text, "L")
         If intPos = 0 Then intPos = InStr("" & .Text, "第")
         If intPos >= 1 Then
            .SelStart = intPos
         End If
      End If
   End With
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP26_S_GotFocus()
   InverseTextBox textCP26_S
End Sub

Private Sub textCP35_GotFocus()
   InverseTextBox textCP35
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse

   TxtValidate = False
   If Me.textCP14.Enabled = True Then
      Cancel = False
      textCP14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP26.Enabled = True Then
      Cancel = False
      textCP26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP26_S.Enabled = True Then
      Cancel = False
      textCP26_S_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP48.Enabled = True Then
      Cancel = False
      textCP48_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP06.Enabled = True Then
      Cancel = False
      textCP06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP07.Enabled = True Then
      Cancel = False
      textCP07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM17.Enabled = True Then
      Cancel = False
      textTM17_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM32.Enabled = True Then
      Cancel = False
      textTM32_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM58.Enabled = True Then
      Cancel = False
      textTM58_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   ' 申請國家為台灣時需檢查來函記錄檔
   If m_TM10 < "010" Then
      strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR16")
      If IsEmptyText(strDate) = False Then
         If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
            strTit = "資料檢核"
            strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               Cancel = True
               textCP06_GotFocus
               Exit Function
            End If
         End If
      '2011/6/15 ADD BY SONIA
      Else
        If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then
        Else
           If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
              strTit = "資料檢核"
              strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
              nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
              If nResponse = vbCancel Then
                 Cancel = True
                 textCP06_GotFocus
                 Exit Function
              End If
           End If
        End If
        '2011/6/15 END
      End If
   End If
   ' 申請國家為台灣時需檢查來函記錄檔
   If m_TM10 < "010" Then
      strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR17")
      If IsEmptyText(strDate) = False Then
         If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
            strTit = "資料檢核"
            strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               Cancel = True
               textCP07_GotFocus
               Exit Function
            End If
         End If
      Else
         If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then  '2011/6/15 ADD BY SONIA
            'modify by sonia 2018/2/8 電子公文都不檢查來函記錄檔
            'If m_DocNo = "" Or textCP07 <> "" Then 'Added by Morgan 2017/4/27 電子公文
            If m_DocNo = "" And textCP07 <> "" Then 'Added by Morgan 2017/4/27 電子公文
               strTit = "資料檢核"
               strMsg = "來函記錄中無該筆記錄"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  Cancel = True
                  textCP07_GotFocus
                  Exit Function
               End If
            End If
         '2011/6/15 ADD BY SONIA
         Else
            If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
               strTit = "資料檢核"
               strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  Cancel = True
                  textCP07_GotFocus
                  Exit Function
               End If
            End If
         End If
         '2011/6/15 END
      End If
   End If
   
   TxtValidate = True
End Function

'Add By Sindy 2012/4/17
Private Sub Option1_Click(Index As Integer)
   If Me.Option4(0).Value Then
      Text10_Validate False
   ElseIf Me.Option4(1).Value Then
      Text11_Validate False
   ElseIf Me.Option4(2).Value Then
      Text12_Validate False
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_LostFocus()
   '非台灣"天"跳離時到"本所期限"欄位
   If m_TM10 <> 台灣國家代號 Then
      If textCP06.Enabled = True Then textCP06.SetFocus
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
   CloseIme
End Sub

Private Sub Text11_LostFocus()
   '非台灣"月"跳離時到"本所期限"欄位
   'If m_TM10 <> 台灣國家代號 Then
   '   If textCP06.Enabled = True Then textCP06.SetFocus
   'End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_LostFocus()
   '非台灣"日"跳離時到"本所期限"欄位
   If m_TM10 <> 台灣國家代號 Then
      If textCP06.Enabled = True Then textCP06.SetFocus
   End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
   Else
      If ChkDate(Text12) Then
         If m_TM10 = 台灣國家代號 Then
            If Val(Text12) < Val(strSrvDate(2)) Then
               MsgBox "來函期限不可小於系統日 !", vbCritical
               Cancel = True
            Else
               textCP07 = Text12
               'Modify By Sindy 2014/10/6 台灣案之本所期限設定
               If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                  textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
               Else
               '2014/10/6 END
                  textCP06 = TransDate(CompDate(2, -2, TransDate(textCP07, 2)), 1)
               End If
               textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
            End If
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub GetTime()
   Dim i As Integer
   Dim strFromDate As String '期限起算日
   
   'Add By Sindy 2012/8/30
   If Option4(0).Value = False And Option4(1).Value = False Then Exit Sub
   '2012/8/30 End
   
   'strFromDate = DBDATE(textCP05)
   strFromDate = DBDATE(frm02010509_1.textCP05)
   
   If m_TM10 = 台灣國家代號 Then
      '文到天數
      If Option4(0).Value = True Then
         textCP07 = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
         If Option1(0).Value = True Then textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
         If Val(Text10) >= 60 Then
            i = -4
         Else
            i = -2
         End If
      '文到月數
      ElseIf Option4(1).Value = True Then
         textCP07 = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
         If Option1(0).Value = True Then textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
      If textCP07 <> "" Then
         'Modify By Sindy 2014/10/6 台灣案之本所期限設定
         If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
            textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
         Else
         '2014/10/6 END
            textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
         End If
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      End If
   End If
End Sub

'讀取來函期限
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
Dim strFromDate As String '期限起算日
   
   'strFromDate = DBDATE(textCP05)
   strFromDate = DBDATE(frm02010509_1.textCP05)
   
   ChgType = False
   If m_TM10 = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If
         
   If ClsPDGetCaseProperty(m_TM01, m_CP10, strTempName, bolTmp) Then
      textCP06 = ""
      textCP07 = ""
      
      If m_TM10 = 台灣國家代號 Then
         strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & m_CP10 & "'"
         If strExc(0) <> "" Then
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            With RsTemp
               If intI = 1 Then
                  If Not IsNull(.Fields(1)) Then
                     '文到天數
                     Option4(0).Value = True
                     Text10 = .Fields(1)
                     textCP07 = TransDate(CompDate(2, Text10, TransDate(strFromDate, 2)), 1)
                  ElseIf Not IsNull(.Fields(2)) Then
                     '文到月數
                     Option4(1).Value = True
                     Text11 = .Fields(2)
                     textCP07 = TransDate(CompDate(1, .Fields(2), TransDate(strFromDate, 2)), 1)
                  Else
                     '文到天數
                     Option4(0).Value = True
                     Text10 = ""
                     Text11 = ""
                  End If
                  If textCP07 <> "" And Not IsNull(.Fields(0)) Then
                     '文到當日
                     If .Fields(0) = "1" Then
                        Option1(0).Value = True
                        textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
                     '文到次日
                     Else
                        Option1(1).Value = True
                     End If
                  End If
                  '文到天數
                  If Text10 <> "" Then
                     If Val(Text10) >= 60 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  '文到月數
                  ElseIf Not IsNull(.Fields(2)) Then
                     If Val(.Fields(2)) >= 2 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  End If
                  If textCP07 <> "" Then
                     'Modify By Sindy 2014/10/6 台灣案之本所期限設定
                     If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                        textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
                     Else
                     '2014/10/6 END
                        textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
                     End If
                     textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                  End If
               End If
            End With
         End If
      End If
      ChgType = True
   End If
End Function
