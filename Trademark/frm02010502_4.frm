VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010502_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "爭議案敗訴輸入"
   ClientHeight    =   5736
   ClientLeft      =   240
   ClientTop       =   996
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9156
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   1170
      TabIndex        =   67
      Top             =   3840
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "文到次日"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "文到當日"
         Height          =   180
         Index           =   0
         Left            =   144
         TabIndex        =   5
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   3900
      TabIndex        =   66
      Top             =   3840
      Width           =   4215
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   10
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   840
         MaxLength       =   2
         TabIndex        =   8
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text12 
         Height          =   252
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   12
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                      日"
         Height          =   225
         Index           =   2
         Left            =   2520
         TabIndex        =   11
         Top             =   180
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "        月"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "文到          天"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.TextBox TextCP64_1 
      Height          =   264
      Left            =   1170
      MaxLength       =   40
      TabIndex        =   1
      Top             =   3270
      Width           =   2532
   End
   Begin VB.TextBox textTM16S 
      Height          =   264
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   17
      Top             =   4980
      Width           =   372
   End
   Begin VB.TextBox textCP07 
      Height          =   264
      Left            =   5610
      MaxLength       =   7
      TabIndex        =   14
      Top             =   4380
      Width           =   2532
   End
   Begin VB.TextBox textCP06 
      Height          =   264
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   13
      Top             =   4380
      Width           =   2532
   End
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2010
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   3540
      Width           =   1692
   End
   Begin VB.TextBox textCF15 
      Height          =   264
      Left            =   1170
      MaxLength       =   4
      TabIndex        =   3
      Top             =   3540
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6996
      TabIndex        =   21
      Top             =   45
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6168
      TabIndex        =   20
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8220
      TabIndex        =   22
      Top             =   45
      Width           =   800
   End
   Begin VB.TextBox textCP35 
      Height          =   264
      Left            =   5610
      MaxLength       =   32
      TabIndex        =   2
      Top             =   3255
      Width           =   2532
   End
   Begin VB.TextBox textCP08 
      Height          =   264
      Left            =   1170
      MaxLength       =   40
      TabIndex        =   0
      Top             =   3000
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5610
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   570
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2700
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2100
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   570
      Width           =   2532
   End
   Begin VB.TextBox textCP48 
      Height          =   264
      Left            =   5610
      TabIndex        =   16
      Top             =   4680
      Width           =   2532
   End
   Begin VB.TextBox textCP14 
      Height          =   264
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   15
      Top             =   4680
      Width           =   732
   End
   Begin VB.TextBox textCP26 
      Height          =   264
      Left            =   6000
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3540
      Width           =   372
   End
   Begin VB.TextBox textTM17 
      Height          =   264
      Left            =   6210
      MaxLength       =   1
      TabIndex        =   18
      Top             =   4980
      Width           =   372
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   870
      Width           =   2532
   End
   Begin VB.TextBox textCP40 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2532
   End
   Begin MSForms.TextBox textCP64 
      Height          =   300
      Left            =   1170
      TabIndex        =   19
      Top             =   5280
      Width           =   7752
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13674;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1170
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1170
      Width           =   7755
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13679;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5610
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2100
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
      Height          =   264
      Left            =   1170
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1500
      Width           =   7752
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13674;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_Src 
      Height          =   264
      Left            =   1170
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2700
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
   Begin MSForms.TextBox textCP14_2 
      Height          =   264
      Left            =   2010
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1692
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      MaxLength       =   20
      Size            =   "13674;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label32 
      Caption         =   "來函期限:"
      Height          =   255
      Left            =   90
      TabIndex        =   68
      Top             =   3990
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "來文字號 :"
      Height          =   255
      Left            =   90
      TabIndex        =   65
      Top             =   3285
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "(1:准 , 2:駁)"
      Height          =   255
      Left            =   1920
      TabIndex        =   63
      Top             =   4980
      Width           =   1515
   End
   Begin VB.Label Label17 
      Caption         =   "案件目前准駁 :"
      Height          =   255
      Left            =   90
      TabIndex        =   64
      Top             =   4980
      Width           =   2295
   End
   Begin VB.Label Label13 
      Caption         =   "進度備註 :"
      Height          =   255
      Left            =   90
      TabIndex        =   62
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "法定期限 :"
      Height          =   255
      Left            =   4650
      TabIndex        =   61
      Top             =   4380
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "本所期限 :"
      Height          =   255
      Left            =   90
      TabIndex        =   60
      Top             =   4380
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "下一程序 :"
      Height          =   255
      Left            =   90
      TabIndex        =   59
      Top             =   3540
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "審查委員 :"
      Height          =   255
      Left            =   4650
      TabIndex        =   57
      Top             =   3255
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   255
      Left            =   90
      TabIndex        =   56
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "審定號 :"
      Height          =   255
      Left            =   4650
      TabIndex        =   55
      Top             =   570
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4650
      TabIndex        =   54
      Top             =   2100
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   90
      TabIndex        =   53
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   255
      Index           =   9
      Left            =   4650
      TabIndex        =   52
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   255
      Index           =   8
      Left            =   4650
      TabIndex        =   51
      Top             =   2700
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   4650
      TabIndex        =   50
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   49
      Top             =   2100
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   90
      TabIndex        =   48
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   90
      TabIndex        =   47
      Top             =   1170
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   46
      Top             =   570
      Width           =   855
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   255
      Left            =   4650
      TabIndex        =   45
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   90
      TabIndex        =   44
      Top             =   2700
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "對照名稱 :"
      Height          =   255
      Left            =   90
      TabIndex        =   43
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   90
      TabIndex        =   42
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "是否算案件數 :"
      Height          =   255
      Left            =   4650
      TabIndex        =   41
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不算)"
      Height          =   255
      Left            =   6540
      TabIndex        =   40
      Top             =   3540
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   "專用權是否存在 :"
      Height          =   255
      Left            =   4650
      TabIndex        =   39
      Top             =   4980
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "(Y / N)"
      Height          =   255
      Left            =   6690
      TabIndex        =   38
      Top             =   4980
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   90
      TabIndex        =   37
      Top             =   870
      Width           =   975
   End
End
Attribute VB_Name = "frm02010502_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/03 Form2.0已修改 cmbTM05/textTM23/textCP13/textCP14_Src/textCP14_2/textCP64
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/6 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 申請國家
Dim m_TM10 As String
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
'Add By Cheng 2002/01/15
Dim m_intNumBegin As Integer
Dim m_intNumEnd As Integer
Dim m_strNumBegin As String
Dim m_strNumEnd As String
Dim BolPrintCaseCheck As Boolean 'Add By Sindy 2012/4/16
'Added by Morgan 2017/4/25 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/4/25
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END
Dim strLD18 As String 'Add By Sindy 2020/1/7 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2020/1/7 FC代理人
Dim m_TM23 As String 'Add By Sindy 2020/1/7 申請人


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010502_3.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010502_3
   Unload frm02010502_2
   Unload frm02010502_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   'Add by Morgan 2003/11/21
   BolPrintCaseCheck = CaseCheck(m_TM01, m_TM02, m_TM03, m_TM04, m_TM10)
   '---end
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      'add by nickc 2005/04/22
      '2011/11/8 modify by sonia TF子案不可結餘故加傳本所案號
      'Pub_EndModCashMsg m_TM10
      Pub_EndModCashMsg m_TM10, m_TM01, m_TM02, m_TM03, m_TM04
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
    'Modify By Cheng 2002/11/07
'      'OnSaveData
    If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
    
      'Add By Sindy 2012/4/16 列印帳款未結清案件資料
      If BolPrintCaseCheck = True Then
          Call GetPrintCaseCheck(m_CP09)
      End If
      '2012/4/16 End
      Call PUB_ChkTemporaryReceipts(m_TM01, m_TM02, m_TM03, m_TM04) 'Add By Sindy 2014/5/28 檢查是否有暫收款
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010502_3
      Unload frm02010502_2
      'Add By Sindy 2019/5/22
      If Me.m_strIR01 <> "" Then
        Unload frm02010502_1
        If Not m_PrevForm Is Nothing Then
           Call m_PrevForm.GoNext
        End If
        Unload Me
        '2019/5/22 END
      'Modified by Morgan 2017/4/25 電子公文
      'frm02010502_1.Show
      ElseIf m_DocNo <> "" Then
         Unload Me
         Unload frm02010502_1
         frm02010412.GoNext
      Else
         frm02010502_1.Show
         Unload Me
      End If
      'end 2017/4/25
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_Src.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP40.BackColor = &H8000000F
   textCP48.BackColor = &H8000000F
   
   textCF15_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010502_1.m_strIR01
   m_strIR02 = frm02010502_1.m_strIR02
   m_strIR03 = frm02010502_1.m_strIR03
   m_strIR04 = frm02010502_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
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
   '2011/7/15 ADD BY SONIA 加入TD,TD-000150
   Select Case m_TM01
      Case "TD":
         ' 設定SQL語法
         strSql = "SELECT SP01 AS TM01,SP02 AS TM02,SP03 AS TM03,SP04 AS TM04,SP05 AS TM05,SP06 AS TM06,SP07 AS TM07,SP09 AS TM10 " & _
            ",'' AS TM12,'' AS TM15,'' AS TM16,'' AS TM28,SP08 AS TM23,SP27 AS TM45,'' AS TM17,SP26 AS TM44 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
      Case Else
   '2011/7/15 END
      strSql = "SELECT * FROM TradeMark " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "'"
   End Select  '2011/7/15 ADD BY SONIA
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"))
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
      
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      'Add By Cheng 2002/07/22
      '顯示目前准駁
      Me.textTM16S.Text = "" & rsTmp.Fields("TM16").Value
      
      ' 專用權是否存在
      If IsNull(rsTmp.Fields("TM17")) = False Then
         textTM17 = rsTmp.Fields("TM17")
      End If
      '2006/6/2 ADD BY SONIA 卷宗性質非申請時,專用權是否存在設定為N
      If IsNull(rsTmp.Fields("TM28")) = False Then
         If rsTmp.Fields("TM28") <> "1" Then
            textTM17 = "N"
         End If
      End If
      '2006/6/2 END
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then
         textTM45 = rsTmp.Fields("TM45")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bCP40 As Boolean
   Dim strDay As String
   Dim strDate As String
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
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
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 機關文號
      m_CP08 = Empty
      If IsNull(rsTmp.Fields("CP08")) = False Then
         m_CP08 = rsTmp.Fields("CP08")
      End If
      ' 案件性質
      'Add By Cheng 2002/07/17
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
   
   ' 承辦期限的日期應為來函收文日加上工作天數
   ' 工作天數由系統別+國家代碼+案件性質(勝訴)搜尋案件收費表的工作天數
   textCP48 = Empty
''''edit by nickc 2007/10/12 改抓有時效的
''''   strDay = GetWorkDays(m_TM01, m_TM10, "1003")
''''   If IsEmptyText(strDay) = False Then
''''      strDate = DBDATE(m_CP05)
''''      ' 90.07.03 modify by louis (承辦期限以實際的工作天數來計算)
''''      'strTemp = DBDATE(Format(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) + Val(strDay))))
''''      strTemp = DBDATE(CompWorkDay(Val(strDay), DBDATE(strDate), 0))
''''      textCP48 = TAIWANDATE(strTemp)
''''   End If
    'edit by nickc 2008/01/10 修正，應該要抓敗訴
    'textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1003", DBDATE(m_CP05), DBDATE(textCP06), textCP09))
    textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1004", DBDATE(m_CP05), DBDATE(textCP06), textCP09))

   ' 無法取得承辦期限的日期
   If IsEmptyText(textCP48) = True Then
      strTit = "資料檢核"
      '2010/12/16 modify by sonia T-168057
      'strMsg = "無法取得承辦期限, 請聯絡電腦中心！"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      ' 回到前一畫面
      'Unload Me
      'frm02010502_3.Show
      If m_CP05 = 111111 Then
         strMsg = "無法取得承辦期限, 請自行輸入！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.Locked = False
         textCP48.BorderStyle = 1
         textCP48.Enabled = True
      End If
   ElseIf textCP48 = m_CP05 Then
      strMsg = "無法取得承辦期限, 請聯絡電腦中心設定工作天！"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      ' 回到前一畫面
      Unload Me
      frm02010502_3.Show
      Exit Sub
   Else
      textCP48.Locked = True
      textCP48.BorderStyle = 0
      textCP48.Enabled = False
      '2010/12/16 end
   End If
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
   
   Set rsTmp = Nothing

   ' 90.11.19 modify by sonia
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
         
   '2015/7/15 modify by sonia 台灣案才預設機關文號
   'If m_TM01 <> "TD" Then '2011/7/15 ADD BY SONIA TD-000150
   If m_TM01 <> "TD" And m_TM10 = "000" Then
      'Add by Morgan 2003/11/26
      If TextCP64_1 = "" Then
         TextCP64_1 = "（" & strTmp & "）智商字第號"
      End If
      m_strNumBegin = "商"
      m_strNumEnd = "字"
      '---End
      
      If m_TM10 < "010" Then
         If textCP08 = "" Then
            Select Case m_CP10
               Case "601", "602"
                  textCP08 = "中台異字第G號"
                   'Modify By Cheng 2002/11/22
                   '不預帶審查委員
   '               'Add By Cheng 2002/01/15
   '               m_strNumBegin = "異"
   '               m_strNumEnd = "字"
               Case "603", "604"
                  textCP08 = "中台評字第H號"
                   'Modify By Cheng 2002/11/22
                   '不預帶審查委員
   '               'Add By Cheng 2002/01/15
   '               m_strNumBegin = "評"
   '               m_strNumEnd = "字"
               Case "605", "606"
                  '2015/2/4 modify by sonia
                  'textCP08 = "中台處字第L號"
                  textCP08 = "中台廢字第L號"
                   'Modify By Cheng 2002/11/22
                   '不預帶審查委員
   '               'Add By Cheng 2002/01/15
   '               m_strNumBegin = "處"
   '               m_strNumEnd = "字"
               Case "401" '訴願
                   'Modify By Cheng 2002/11/05
   '               textCP08 = "經（" & strTmp & "）訴字第號"
                  textCP08 = "經訴字第號"
                   'Modify By Cheng 2002/11/22
                   '不預帶審查委員
   '               'Add By Cheng 2002/01/15
   '               m_strNumBegin = "訴"
   '               m_strNumEnd = "字"
               Case "403"
                  textCP08 = strTmp & "年度訴字第號"
                   'Modify By Cheng 2002/11/22
                   '不預帶審查委員
   '               'Add By Cheng 2002/01/15
   '               m_strNumBegin = "訴"
   '               m_strNumEnd = "字"
            End Select
         End If
      End If
      
      'Added by Morgan 2017/4/25 電子公文
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
      'end 2017/4/25
      
      'Add By Cheng 2002/07/23
      If (m_CP10 >= "401" And m_CP10 <= "405") _
         Or m_CP10 = "602" Or m_CP10 = "604" Or m_CP10 = "606" _
         Or m_CP10 = "610" Or Left(m_CP09, 1) = "C" Then
         Me.textTM17.Text = "N"
      End If
   End If   '2011/7/15 ADD BY SONIA
       
    'Add By Cheng 2002/11/22
    textCP08_GotFocus
    
End Sub

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim nIndex As Integer
Dim strSql As String
Dim strSubTMSQL As String
Dim strCP06 As String
Dim strCP07 As String
Dim strCP09 As String
Dim strCP10 As String
Dim strCP12 As String
Dim strCP27 As String
Dim strNP07 As String
Dim strNP14 As String
Dim strNP22 As String
Dim strNP08 As String   '2010/3/29 ADD BY SONIA
'Add by Amy 2017/11/13
Dim m_CP06 As String, m_CP07 As String, st_CP09 As String, m_CP14 As String, strMsg As String
Dim bolUpdCP As Boolean '是否更新進度檔

'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans

   strSubTMSQL = "WHERE TM01 = '" & m_TM01 & "' AND " & _
                       "TM02 = '" & m_TM02 & "' AND " & _
                       "TM03 = '" & m_TM03 & "' AND " & _
                       "TM04 = '" & m_TM04 & "' "
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新原案件進度檔的收文資料其實際結果為駁, 准駁日為來函收文日, 審查委員
   strSql = "UPDATE CaseProgress SET CP24 = '2', " & _
                                    "CP25 = " & DBDATE(m_CP05) & ", " & _
                                    "CP35 = '" & textCP35 & "' " & _
            "WHERE CP09 = '" & m_CP09 & "' AND " & _
                  "(CP24 IS NULL OR CP24 = '' OR CP24 = ' ')"
   cnnConnection.Execute strSql
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Do Nothing By Cheng 2002/07/23
   '目前准駁欄只顯示不可更改, 故不用更新
   
   ' 更新商標基本檔的目前准駁欄, 審定來函日(准駁通知日)為來函收文日
   'If textTM16S = "Y" Then
   '   strSQL = "UPDATE TradeMark SET TM16 = '2', " & _
   '                                 "TM13 =" & DBDATE(m_CP05) & " "
   '   strSQL = strSQL & strSubTMSQL
   '   cnnConnection.Execute strSQL
   'End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新商標基本檔的專用權是否存在
   strSql = "UPDATE TradeMark SET TM17 = '" & textTM17 & "' "
   strSql = strSql & strSubTMSQL
   cnnConnection.Execute strSql

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   
   ' 案件性質為敗訴
   strCP10 = "1004"
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 發文日為系統日
   'strCP27 = DBDATE(Date)
   strCP06 = Empty
   strCP07 = Empty
   If IsEmptyText(textCP06) = False Then: strCP06 = DBDATE(textCP06)
   If IsEmptyText(textCP07) = False Then: strCP07 = DBDATE(textCP07)
   
   ' 先新增一筆案件進度記錄再更新其本所期限及法定期限
   ' 91.03.25 modify by louis (單引號)
    '承辦人為原程序承辦人, 不上發文日
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    
   'Add by Morgan 2003/11/26
   Dim strCP64 As String
   
   strCP64 = Trim(textCP64)
   If strCP64 <> "" And Trim(TextCP64_1) <> "" Then
      strCP64 = strCP64 & ",來文字號：" & Trim(TextCP64_1)
   ElseIf Trim(TextCP64_1) <> "" Then
      strCP64 = "來文字號：" & Trim(TextCP64_1)
   End If
   
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP48,CP64) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
'                          "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
'                          "'" & textCP35 & "','" & m_CP36 & "','" & m_CP37 & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                          "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "'," & DBDATE(textCP48) & ",'" & ChgSQL(textCP64) & "')"
'
    'Modify By Cheng 2004/02/03
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP48,CP64) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP14 & "'," & _
'                          "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
'                          "'" & textCP35 & "','" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                          "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "'," & DBDATE(textCP48) & ",'" & ChgSQL(strCP64) & "')"
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP48,CP64) " & _
                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                          "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP14 & "'," & _
                          "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
                          "'" & textCP35 & "','" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
                          "'" & m_cp40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "'," & DBDATE(textCP48) & ",'" & ChgSQL(strCP64) & "')"
    'End

'---End
   cnnConnection.Execute strSql
    
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
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
   
   'Add By Sindy 2020/1/7 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(m_TM01, 1) = "T" Then
      strLD18 = strCP09
      If Val(textCP06) > 0 Then
         PUB_AddLetterProgress strLD18, 1, False, "", True, m_TM23, strCP10, m_TM44
      Else
         PUB_AddLetterProgress strLD18, 1, False, "", False, m_TM23, strCP10, m_TM44
      End If
   End If
   '2020/1/7 END
   
   'add by nickc 2008/01/09 FCT 加判斷，有期限用期限判斷(第三或第五個工作天)，無期限以第三個工作日(當日不算)，寫入承辦期限
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
   
   
   '92.11.12 ADD BY SONIA
   '2012/2/10 modify by sonia 改國外部收文案件的勝訴都要請款
   'If m_TM01 = "FCT" And (m_CP10 = "601" Or m_CP10 = "603" Or m_CP10 = "605") Then
   If Left(Trim(m_CP12), 1) = "F" Then
      strSql = "UPDATE CaseProgress SET CP20 = NULL " & _
               "WHERE CP09 = '" & strCP09 & "'"
      cnnConnection.Execute strSql
   End If
   '92.11.12 END
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '將下一程序為催審的資料, 更新其是否續辦欄位"Y"
   strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP07 = '" & "305" & "' "
   cnnConnection.Execute strSql
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
        'modify by sonia 2019/5/22 +更新CP43,否則撰寫信函無法處理T-204490
        'strSql = "Update CaseProgress Set CP06=" & Val(textCP06) + 19110000 & ",CP07=" & Val(textCP07) + 19110000 & " Where CP09='" & st_CP09 & "'"
        strSql = "Update CaseProgress Set CP06=" & Val(textCP06) + 19110000 & ",CP07=" & Val(textCP07) + 19110000 & ",CP43='" & strCP09 & "' Where CP09='" & st_CP09 & "'"
        cnnConnection.Execute strSql
        
        If m_CP14 = MsgText(601) Then m_CP14 = GetDeptMan("P20") '無承辦人發給P20部門之A0908
        strMsg = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "收到" & "" & GetCaseTypeName(m_TM01, textCF15, IIf(m_TM10 = "000", 0, 1)) & "前已收文,請辦理後續！"
        PUB_SendMail strUserNum, m_CP14, "", strMsg, "本所期限：" & textCP06 & "　　法定期限：" & textCP07
      
    '進度檔未有相同未發文未取消收文之案件性質或上述不更新期限,才新增下一程序
    Else
        strNP14 = Empty
        strNP14 = GetRelatedPerson(m_CP09)
        'Modify By Cheng 2002/09/25
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          strCP06 & "," & strCP07 & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "'," & strNP22 & ")"
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                            strCP06 & "," & strCP07 & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "'," & strNP22 & ")"
        cnnConnection.Execute strSql
    End If
    'end 2017/11/13
      
    ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
    '      '92.6.8 SONIA 加 言詞辯論, 準備程序
    Select Case textCF15
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         Case "102", "105", "702", "708", "305", "998", "997"
         Case Else:
            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
            'add by nickc 2008/04/23  加入案件回覆單
            '2008/5/20 MODIFY BY SONIA FCT案不印回覆單
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
   
   '2010/3/29 add by sonia 大陸部分核駁1205來函提復審401敗訴時,重新掛原申請程序收文號的催審期限為來函收文日+6個月T-151877
   If m_TM10 = "020" And m_CP10 = "401" Then
      '先判斷是否為部分核駁1205來函才提的復審 m_CP09
      strSql = "SELECT C2.CP43 FROM CaseProgress C1,CASEPROGRESS C2 " & _
               "WHERE C1.CP01 = '" & m_TM01 & "' AND C1.CP02 = '" & m_TM02 & "' AND " & _
                     "C1.CP03 = '" & m_TM03 & "' AND C1.CP04 = '" & m_TM04 & "' AND " & _
                     "C1.CP09 = '" & m_CP09 & "' AND C1.CP43=C2.CP09(+) AND C2.CP10='1205'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         strNP08 = CompDate(1, 6, ChangeTStringToWString(frm02010502_1.textCP05))
         '智權人員掛此來函之承辦人
         'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                   "VALUES ('" & rsTmp.Fields(0) & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','305'," & _
                             strNP08 & "," & strNP08 & ",'" & textCP14 & "'," & GetNextProgressNo() & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                   "VALUES ('" & rsTmp.Fields(0) & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','305'," & _
                             PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & textCP14 & "'," & GetNextProgressNo() & ")"
         cnnConnection.Execute strSql
      End If
      rsTmp.Close
   End If
   '2010/3/29 end
   
   'add by nickc 2005/04/22
   Pub_UpdateEndModCash m_TM01, m_TM02, m_TM03, m_TM04
   Set rsTmp = Nothing
   
   'Added by Morgan 2017/4/25 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
   End If
   'end 2017/4/25
   
   'Add by Sindy 2019/5/22
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010502_1"
   End If
   '2019/5/22 END
   
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2025/09/12
   
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010502_4 = Nothing
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
      'Add By Cheng 2002/03/11
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Cancel = True
         textCP06_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2002/11/19
'      ' 申請國家為台灣時需檢查來函記錄檔
'      If m_TM10 < "010" Then
'         strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR16")
'         If IsEmptyText(strDate) = False Then
'            If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
'               strTit = "資料檢核"
'               strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
'               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'               If nResponse = vbCancel Then
'                  Cancel = True
'                  textCP06_GotFocus
'                  GoTo EXITSUB
'               End If
'            End If
'         ' 91.05.16 modify by louis
'         Else
'            strTit = "資料檢核"
'            strMsg = "來函記錄無此記錄"
'            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'            If nResponse = vbCancel Then
'               Cancel = True
'               textCP06_GotFocus
'               GoTo EXITSUB
'            End If
'         End If
'      End If
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
        'Modify By Cheng 2002/11/19
        '按下確定時才檢查
'      ' 申請國家為台灣時需檢查來函記錄檔
'      If m_TM10 < "010" Then
'         strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR17")
'         If IsEmptyText(strDate) = False Then
'            If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
'               strTit = "資料檢核"
'               strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
'               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'               If nResponse = vbCancel Then
'                  Cancel = True
'                  textCP07_GotFocus
'                  GoTo EXITSUB
'               End If
'            End If
'         Else
'            strTit = "資料檢核"
'            strMsg = "來函記錄中無該筆記錄"
'            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'            If nResponse = vbCancel Then
'               Cancel = True
'               textCP07_GotFocus
'               GoTo EXITSUB
'            End If
'         End If
'      End If
   End If
EXITSUB:
End Sub

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
      'Add By Cheng 2002/03/11
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
   End If
   
   ' 承辦期限不可空白
   If IsEmptyText(textCP48) = True Then
      strTit = "檢核資料"
      strMsg = "承辦期限不可為空白！"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   ' 是否更新基本檔目前准駁不可空白
   'If IsEmptyText(textTM16S) = True Then
   '   strTit = "檢核資料"
   '   strMsg = "是否更新基本檔目前准駁不可為空白"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   GoTo ExitSub
   'End If
   ' 專用權是否存在不可為空白
   '2011/7/15 MODIFY BY SONIA TD-000150
   'If IsEmptyText(textTM17) = True Then
   If IsEmptyText(textTM17) = True And m_TM01 <> "TD" Then
      strTit = "檢核資料"
      strMsg = "專用權是否存在不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   CheckDataValid = True
EXITSUB:
End Function

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
'Modify By Cheng 2002/11/22
'不預帶審查委員
On Error GoTo ErrorHandler
'
''Add By Cheng 2002/01/15
'If Len(Me.textCP08.Text) > 0 Then
'   m_intNumBegin = InStr(Me.textCP08.Text, m_strNumBegin)
'   m_intNumEnd = InStr(Me.textCP08.Text, m_strNumEnd)
'Else
'   m_intNumBegin = 0
'   m_intNumEnd = 0
'End If
'If m_intNumBegin < m_intNumEnd Then
'   Me.textCP35.Text = Mid(Me.textCP08.Text, m_intNumBegin + 1, (m_intNumEnd - m_intNumBegin - 1))
'End If
'
'Exit Sub
'
'ErrorHandler:
'   m_intNumBegin = 0
'   m_intNumEnd = 0

'Add by Morgan 2003/11/26
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
'---End
End Sub

Private Sub textCP08_LostFocus()
'Modify By Cheng 2002/11/22
'不預帶審查委員
'On Error GoTo ErrorHandler
'
''Add By Cheng 2002/01/15
'If Len(Me.textCP08.Text) > 0 Then
'   m_intNumBegin = InStr(Me.textCP08.Text, m_strNumBegin)
'   m_intNumEnd = InStr(Me.textCP08.Text, m_strNumEnd)
'Else
'   m_intNumBegin = 0
'   m_intNumEnd = 0
'End If
'If m_intNumBegin < m_intNumEnd Then
'   Me.textCP35.Text = Mid(Me.textCP08.Text, m_intNumBegin + 1, (m_intNumEnd - m_intNumBegin - 1))
'End If
'
'Exit Sub
'
'ErrorHandler:
'   m_intNumBegin = 0
'   m_intNumEnd = 0
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


'Private Sub textTM16S_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub

' 是否更新基本檔目前准駁
'Private Sub textTM16S_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'
'   Cancel = False
'   If IsEmptyText(textTM16S) = False Then
'      Select Case textTM16S
'         Case "Y", "N"
'         Case Else:
'            Cancel = True
'            strTit = "資料檢核"
'            strMsg = "只可輸入Y或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM16S_GotFocus
'      End Select
'   End If
'End Sub

Private Sub textTM17_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 89 And KeyAscii <> 78 Then
        KeyAscii = 0
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

'Private Sub textTM16S_GotFocus()
'   InverseTextBox textTM16S
'End Sub

Private Sub textTM17_GotFocus()
   InverseTextBox textTM17
End Sub

Private Sub textCP08_GotFocus()
   'Modify By Cheng 2002/04/22
   '將游標停在某個字的後面
'   InverseTextBox textCP08
Dim intPos As Integer
With Me.textCP08
   If Len("" & .Text) > 0 Then
        'Modify By Cheng 2002/11/22
'      intPos = InStr("" & .Text, "字")
      intPos = InStr("" & .Text, "G")
      If intPos = 0 Then intPos = InStr("" & .Text, "H")
      If intPos = 0 Then intPos = InStr("" & .Text, "L")
      If intPos = 0 Then intPos = InStr("" & .Text, "第")
      
      If intPos >= 1 Then
         .SelStart = intPos
         .SelLength = 0
      End If
   End If
End With
End Sub

Private Sub textCP35_GotFocus()
   InverseTextBox textCP35
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

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'Add By Cheng 2002/11/19
Dim strTit As String
Dim strMsg As String
Dim nResponse

TxtValidate = False
If Me.textCF15.Enabled = True Then
   Cancel = False
   textCF15_Validate Cancel
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

If Me.textTM17.Enabled = True Then
   Cancel = False
   textTM17_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
    'Add By Cheng 2002/11/19
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
       ' 91.05.16 modify by louis
       '2008/11/27 CANCEL BY SONIA
       'Else
       '   strTit = "資料檢核"
       '   strMsg = "來函記錄無此記錄"
       '   nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
       '   If nResponse = vbCancel Then
       '      Cancel = True
       '      textCP06_GotFocus
       '     Exit Function
       '   End If
       '2008/11/27 END
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
            'If m_DocNo = "" Or textCP07 <> "" Then 'Added by Morgan 2017/4/25 電子公文
            If m_DocNo = "" And textCP07 <> "" Then 'Added by Morgan 2017/4/25 電子公文
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
   strFromDate = DBDATE(frm02010502_1.textCP05)
   
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
   strFromDate = DBDATE(frm02010502_1.textCP05)
   
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
