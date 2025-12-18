VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030101_19 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(申請英文證明)"
   ClientHeight    =   6960
   ClientLeft      =   2172
   ClientTop       =   1584
   ClientWidth     =   9024
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9024
   Begin VB.ComboBox textCP44 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4200
      TabIndex        =   7
      Top             =   3456
      Width           =   1620
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   1596
      Left            =   168
      TabIndex        =   50
      Top             =   2736
      Width           =   4020
      Begin VB.TextBox textCP30 
         Height          =   300
         Left            =   1176
         MaxLength       =   20
         TabIndex        =   2
         Top             =   48
         Width           =   2532
      End
      Begin VB.TextBox textCP30_1 
         Height          =   300
         Left            =   1176
         MaxLength       =   30
         TabIndex        =   3
         Top             =   384
         Width           =   2532
      End
      Begin VB.Label Label2 
         Caption         =   "註冊號數 :"
         Height          =   180
         Left            =   48
         TabIndex        =   53
         Top             =   108
         Width           =   912
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   48
         TabIndex        =   52
         Top             =   792
         Width           =   900
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   312
         Left            =   1176
         TabIndex        =   6
         Top             =   720
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;550"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         Caption         =   "申請號數:"
         Height          =   180
         Left            =   48
         TabIndex        =   51
         Top             =   444
         Width           =   768
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "非本所案件"
      Height          =   1236
      Left            =   144
      TabIndex        =   46
      Top             =   5616
      Width           =   8652
      Begin MSForms.TextBox textOther_T 
         Height          =   288
         Left            =   1992
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   888
         Width           =   5724
         VariousPropertyBits=   671105055
         Size            =   "10096;508"
         Value           =   "1.商標, 2.商標(92年修正前服務標章), 3.團體商標, 4.團體標章, 5.證明標章"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         Caption         =   "商標名稱 :"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   288
         Width           =   912
      End
      Begin MSForms.TextBox txtOther 
         Height          =   528
         Index           =   0
         Left            =   1056
         TabIndex        =   10
         Top             =   288
         Width           =   7476
         VariousPropertyBits=   -1467989989
         MaxLength       =   300
         ScrollBars      =   2
         Size            =   "13187;931"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtOther 
         Height          =   300
         Index           =   1
         Left            =   1560
         TabIndex        =   11
         Top             =   864
         Width           =   372
         VariousPropertyBits=   671105051
         MaxLength       =   1
         ScrollBars      =   2
         Size            =   "656;529"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         Caption         =   "商標或標章種類："
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   912
         Width           =   1584
      End
   End
   Begin VB.TextBox textCP84_1 
      Height          =   300
      Left            =   5430
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2734
      Width           =   540
   End
   Begin VB.TextBox txtCP113 
      Height          =   300
      Left            =   5430
      MaxLength       =   4
      TabIndex        =   5
      Top             =   3074
      Width           =   540
   End
   Begin VB.TextBox textCP84 
      Alignment       =   1  '靠右對齊
      Height          =   300
      Left            =   5430
      TabIndex        =   1
      Top             =   2400
      Width           =   1425
   End
   Begin VB.TextBox textCP22 
      Height          =   264
      Left            =   2052
      MaxLength       =   1
      TabIndex        =   12
      Top             =   108
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textCP27 
      Height          =   300
      Left            =   1350
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2400
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Left            =   7056
      TabIndex        =   14
      Top             =   0
      Width           =   1152
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6096
      TabIndex        =   13
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8208
      TabIndex        =   15
      Top             =   0
      Width           =   912
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5424
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   372
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5424
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   672
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5424
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   972
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1344
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   672
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1344
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   372
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1344
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1272
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5424
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1272
      Width           =   2532
   End
   Begin VB.Label Label4 
      Caption         =   "代理人 :"
      Height          =   180
      Left            =   3360
      TabIndex        =   43
      Top             =   3480
      Width           =   756
   End
   Begin VB.Label Label28 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   220
      TabIndex        =   45
      Top             =   4380
      Width           =   972
   End
   Begin VB.Label Label29 
      Caption         =   "案件備註 :"
      Height          =   252
      Left            =   220
      TabIndex        =   44
      Top             =   4992
      Width           =   972
   End
   Begin MSForms.TextBox textTM58 
      Height          =   600
      Left            =   1350
      TabIndex        =   9
      Top             =   4968
      Width           =   7428
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13102;1058"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   600
      Left            =   1350
      TabIndex        =   8
      Top             =   4344
      Width           =   7428
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13102;1058"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP44_2 
      Height          =   756
      Left            =   5880
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3468
      Width           =   3096
      VariousPropertyBits=   -1467989985
      Size            =   "5461;1333"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "工作時數:"
      Height          =   180
      Left            =   4500
      TabIndex        =   41
      Top             =   3134
      Width           =   768
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1344
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1572
      Width           =   2535
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5424
      TabIndex        =   39
      Top             =   1572
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05_1 
      Height          =   495
      Left            =   1350
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1890
      Width           =   7635
      VariousPropertyBits=   -1475330017
      MaxLength       =   20
      Size            =   "13467;873"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      Caption         =   "份    數:"
      Height          =   180
      Left            =   4500
      TabIndex        =   37
      Top             =   2794
      Width           =   600
   End
   Begin VB.Label Label39 
      Caption         =   "發文規費："
      Height          =   180
      Left            =   4500
      TabIndex        =   36
      Top             =   2460
      Width           =   900
   End
   Begin VB.Label Label31 
      Caption         =   "(N:不出名)"
      Height          =   252
      Left            =   2400
      TabIndex        =   35
      Top             =   144
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label30 
      Caption         =   "是否出名 :"
      Height          =   252
      Left            =   1176
      TabIndex        =   34
      Top             =   96
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label25 
      Caption         =   "發文日 :"
      Height          =   180
      Left            =   220
      TabIndex        =   33
      Top             =   2460
      Width           =   855
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱 :"
      Height          =   180
      Left            =   220
      TabIndex        =   32
      Top             =   1920
      Width           =   810
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   220
      TabIndex        =   31
      Top             =   1572
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "業務區別 :"
      Height          =   252
      Index           =   2
      Left            =   4464
      TabIndex        =   30
      Top             =   372
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   252
      Index           =   3
      Left            =   4464
      TabIndex        =   28
      Top             =   672
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   4464
      TabIndex        =   26
      Top             =   972
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   220
      TabIndex        =   24
      Top             =   672
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   220
      TabIndex        =   23
      Top             =   372
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   220
      TabIndex        =   22
      Top             =   1272
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   9
      Left            =   4464
      TabIndex        =   21
      Top             =   1272
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4464
      TabIndex        =   20
      Top             =   1572
      Width           =   972
   End
End
Attribute VB_Name = "frm030101_19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/12 改成Form2.0 ; textCP13、textCP14、textTM05_1、textCP44_2、textCP64、textTM58
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
'create by nickc copy from frm030101_03
Option Explicit

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CP09 As String
' 申請國家
Dim m_TM10 As String
' 案件性質代號
Dim m_CP10 As String
' 承辦人代號
Dim m_CP14 As String
' 原專用期限起日
Dim m_TM21 As String
' 原專用期限止日
Dim m_TM22 As String
' 申請人
Dim m_TM23 As String
' 申請國家的延展年度
Dim m_NA14 As String
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存商標基本檔或服務業務基本檔檔案欄位的串列
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
' 儲存案件進度檔檔案欄位的串列
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer
' 宣告代理人內容結構
Private Type AGENTITEM
   aiCode As String
   aiName As String
End Type
Dim m_AgentList() As AGENTITEM
Dim m_AgentCount As Integer
Dim m_Case(1 To 4) As String '本所案號
Dim m_CP110 As String
Dim m_CP84 As String       '發文規費
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP13 As String 'Add By Sindy 2014/9/11
'Added by Lydia 2023/08/08
Dim m_ProcType As String '1-產生申請書, 2-發文作業
Dim m_CP17 As String '收文規費
Dim m_InCase() As String '輸入申請案號/註冊號數的本所案號

Private Sub cmdCancel_Click()
   'Added by Lydia 2023/08/08
   If m_ProcType = "1" Then '產生申請書
      frm030206_1.Show
   Else
   'end 2023/08/08
      frm030101_01.Show
   End If 'Added by 2023/08/8
   Unload Me
End Sub

Private Sub cmdExit_Click()
   Me.Enabled = False
   'Added by Lydia 2023/08/08
   If m_ProcType = "1" Then '產生申請書
      Unload frm030206_1
   Else
   'end 2023/08/08
      Unload frm030101_01
   End If 'Added by Lydia 2023/08/08
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim ET03 As String, ET03_1 As String 'Added by Lydia 2022/08/04

   If CheckDataValid = True Then
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault

      'Added by Lydia 2023/08/08
      If m_ProcType = "1" Then '產生申請書
         Call GetApplBook_CFT
         frm030206_1.Show
      Else
      'end 2023/08/08
         'Add By Sindy 2012/4/5 CFT,FCT所有案件性質發文時,檢查代表圖是否存在
         'Modify by Amy 2018/07/31 ChkIsExistImg不使用
         'Call ChkIsExistImg(m_TM01, m_TM02, m_TM03, m_TM04)
         If ChkImgByteFile(m_TM01, m_TM02, m_TM03, m_TM04) = False Then MsgBox "本案尚未放代表圖至系統！"
         
         'Add By Sindy 2024/8/19
         If frm030101_01.bolIsEMPFlow = True Then
            frm090202_4.QueryData
         End If
         '2024/8/19 End
         '若有未發文資料顯示警告
         If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = False Then
            'Add By Sindy 2024/8/19
            If frm030101_01.bolIsEMPFlow = True Then
               Unload frm030101_01
               frm090202_4.Show
               Unload Me
               Exit Sub
            End If
            '2024/8/19 End
         End If
         frm030101_01.Show
         frm030101_01.Clear1
      End If 'Added by Lydia 2023/08/08
      Unload Me
   End If
End Sub


Private Sub Form_Activate()
'Modify By Sindy 2012/10/1 下列程式無意義Mark
'    If (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
'        pub_ModifyCaseNum = ""
'        QueryData
'    End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   'Added by Lydia 2023/08/08
   If m_ProcType = "1" Then  '產生申請書
      Me.Height = 6800
   Else
      Me.Height = 5300
   End If
   Frame2.BackColor = &H8000000F
   'end 2023/08/08
   
   MoveFormToCenter Me
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Added by Lydia 2021/08/12 畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 855
   
   'Added by Lydia 2023/08/08
   Frame1.Enabled = False
   ReDim m_InCase(TF_TM)
   Frame1.Top = 5080
   Label4.Left = Label25.Left: textCP44.Left = textCP27.Left: textCP44_2.Left = 3000 '代理人
   textCP84.TabStop = False
   If m_ProcType = "1" Then '產生申請書
      Me.Caption = "CFT申請英文證明申請書"
      Frame1.Visible = True
      Frame2.Visible = True
      Label25.Visible = False: textCP27.Visible = False  '發文日
      Label7.Visible = False: txtCP113.Visible = False  '工作時數
      Label29.Visible = False: textTM58.Visible = False  '案件備註
      Label4.Visible = False: textCP44.Visible = False: textCP44_2.Visible = False '代理人
      Label9.Visible = True: textCP84_1.Visible = True   '份數
   Else  '發文
      Me.Caption = "發文(申請英文證明)"
      Frame1.Visible = False
      Frame2.Visible = False
      Label25.Visible = True: textCP27.Visible = True  '發文日
      Label7.Visible = True: txtCP113.Visible = True  '工作時數
      Label7.Top = 2794: txtCP113.Top = 2734
      Label29.Visible = True: textTM58.Visible = True  '案件備註
      Label4.Visible = True: textCP44.Visible = True: textCP44_2.Visible = True '代理人
      Label4.Top = 3234: textCP44.Top = 3170: textCP44_2.Top = 3170
      Label9.Visible = False: textCP84_1.Visible = False  '份數
      Label28.Top = 3540: textCP64.Top = 3500
      Label29.Top = 4152: textTM58.Top = 4124
   End If
   'end 2023/08/08
End Sub

'Modified by Lydia 2023/08/08 + pType
Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False, Optional ByVal pType As String = "2")
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If

   Select Case nType
      ' 收文號
      Case 0: m_CP09 = strData

   End Select
   
   m_ProcType = pType 'Added by Lydia 2023/08/08
End Sub

Private Sub ClearAgentList()
   If m_AgentCount > 0 Then
      Erase m_AgentList
   End If
   m_AgentCount = 0
End Sub

Private Sub AddAgent(ByVal strAgentCode As String, ByVal strAgentName As String)
   Dim nIndex As Integer
   Dim bFind As Boolean
   bFind = False
   For nIndex = 0 To m_AgentCount - 1
      If m_AgentList(nIndex).aiCode = strAgentCode Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_AgentList(m_AgentCount + 1)
      m_AgentList(m_AgentCount).aiCode = strAgentCode
      m_AgentList(m_AgentCount).aiName = strAgentName
      m_AgentCount = m_AgentCount + 1
   End If
End Sub


' 清除商標基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiOldData = strFieldData
         m_TMSPList(nPos).fiNewData = strFieldData
         m_TMSPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_TMSPList(m_TMSPCount + 1)
      m_TMSPList(m_TMSPCount).fiName = strFieldName
      m_TMSPList(m_TMSPCount).fiOldData = strFieldData
      m_TMSPList(m_TMSPCount).fiNewData = strFieldData
      m_TMSPList(m_TMSPCount).fiType = nFieldType
      m_TMSPCount = m_TMSPCount + 1
   End If
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 清除案件進度檔檔案欄位串列
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiOldData = strFieldData
         m_CPList(nPos).fiNewData = strFieldData
         m_CPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_CPList(m_CPCount + 1)
      m_CPList(m_CPCount).fiName = strFieldName
      m_CPList(m_CPCount).fiOldData = strFieldData
      m_CPList(m_CPCount).fiNewData = strFieldData
      m_CPList(m_CPCount).fiType = nFieldType
      m_CPCount = m_CPCount + 1
   End If
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 取得商標基本檔的欄位內容
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_TM44 = CheckStr(rsTmp.Fields("TM44"))
      m_TM119 = CheckStr(rsTmp.Fields("TM119"))
      m_TM120 = CheckStr(rsTmp.Fields("TM120"))
      ' 註冊號數
'      If IsNull(rsTmp.Fields("TM15")) = False Then
'         textCP30 = rsTmp.Fields("TM15")
'      End If
'      SetTMSPFieldOldData "TM15", textTM15, 0
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         textTM20 = rsTmp.Fields("TM20")
      End If
      ' 案件名稱
      textTM05_1 = Empty
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05_1 = rsTmp.Fields("TM05")
      End If
      ' 申請國家
      m_NA14 = Empty
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         m_NA14 = GetNationExtentYear(m_TM10)
      End If
      'Added by Lydia 2024/01/18
      m_TM23 = "" & rsTmp.Fields("TM23")
      m_TM21 = "" & rsTmp.Fields("TM21")
      m_TM22 = "" & rsTmp.Fields("TM22")
      'end 2024/01/18
      ' 案件備註
      textTM58 = Empty
      If IsNull(rsTmp.Fields("TM58")) = False Then
         textTM58 = rsTmp.Fields("TM58")
      End If
      SetTMSPFieldOldData "TM58", textTM58, 0
   
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSubTmp As New ADODB.Recordset
   Dim strCP27 As String
   Dim strCP44 As String
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst

      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      ' 案件性質
      m_CP10 = Empty: m_CP14 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區別
      If IsNull(rsTmp.Fields("CP12")) = False Then
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      m_CP13 = "" 'Add By Sindy 2014/9/11
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
         m_CP13 = rsTmp.Fields("CP13") 'Add By Sindy 2014/9/11
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         m_CP14 = rsTmp.Fields("CP14")
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      textCP22 = Empty
      If IsNull(rsTmp.Fields("CP22")) = False Then
         textCP22 = rsTmp.Fields("CP22")
      End If
      SetCPFieldOldData "CP22", textCP22, 0
      ' 發文日(預設為系統日)
      If m_ProcType = "2" Then 'Added by Lydia 2023/08/08 發文作業
        textCP27 = strSrvDate(1)
      End If 'Added by Lydia 2023/08/08
      
      strCP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = rsTmp.Fields("CP27")
      End If
      SetCPFieldOldData "CP27", strCP27, 1
      SetCPFieldOldData "CP30", CheckStr(rsTmp.Fields("CP30")), 0
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
      Else
         textCP44 = "Y00000000" '2009/2/3 ADD BY SONIA 預設台一
      End If
      SetCPFieldOldData "CP44", textCP44, 0
      ' 彼所案號
      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         strCP45 = rsTmp.Fields("CP45")
      End If
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
      m_CP17 = "" & rsTmp.Fields("CP17") 'Added by Lydia 2023/08/02 收文規費
      'add by nick 2004/08/12 發文規費
      If IsNull(rsTmp.Fields("CP17")) = False Then
         m_CP84 = CheckStr(rsTmp.Fields("CP17"))
         textCP84.Text = m_CP84
      End If
      
      'Added by Morgan 2012/9/6 電子送件發文規費預設為承辦人已輸入的金額
      If "" & rsTmp.Fields("cp118") = "Y" Then
         textCP84 = Val("" & rsTmp.Fields("cp84"))
      End If
      'end 2012/9/6
      
      'Added by Lydia 2023/08/08 份數控制發文規
      If Val(textCP84) > 0 Then
         textCP84_1 = CInt(Val(textCP84 / 500))
      Else
         textCP84 = 500
         textCP84_1 = 1
      End If
      textCP84.Locked = True
      'end 2023/08/08
      
      SetCPFieldOldData "CP84", CheckStr(rsTmp.Fields("CP84")), 0
      'add by nickc 2006/01/27
      m_CP110 = CheckStr(rsTmp.Fields("cp110"))
      SetCPFieldOldData "CP110", m_CP110, 0
      SetCPFieldOldData "CP123", "", 0     '2009/4/21 ADD BY SONIA 算發文室件數
      ' 代理人
      ClearAgentList
      'Add By Sindy 2013/5/23 若是原先有，也要加入
      If textCP44.Text <> "" Then
'         If InStr(textCP44, "-") > 0 Then
'            If ClsPDGetContact(textCP44, strCP44) Then
'               AddAgent textCP44, strCP44
'            End If
'         Else
            strCP44 = GetFAgentName(textCP44)
            AddAgent textCP44, strCP44
'         End If
      End If
      '2013/5/23 End
      '2010/9/7 Modify by Sindy 文件簽證711及申請英文證明304不要列入
      strSubSQL = "SELECT CP44, MAX(CP27) AS CP27 FROM CASEPROGRESS " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null " & _
                        "AND CP10 NOT IN ('711','304') " & _
                  "GROUP BY CP44 " & _
                  "ORDER BY CP27 DESC "
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSubTmp.RecordCount > 0 Then
         rsSubTmp.MoveFirst
         ' 依序將代理人加入到系統串列中
         Do While rsSubTmp.EOF = False
            If IsNull(rsSubTmp.Fields("CP44")) = False Then
               strCP44 = GetFAgentName(rsSubTmp.Fields("CP44"))
               AddAgent rsSubTmp.Fields("CP44"), GetFAgentName(rsSubTmp.Fields("CP44"))
            End If
            rsSubTmp.MoveNext
         Loop
      End If
      rsSubTmp.Close
      ' 從系統串列中取得所有代理人並放入Combo Box中
      For nIndex = 0 To m_AgentCount - 1
         textCP44.AddItem m_AgentList(nIndex).aiCode
      Next nIndex
      ' 設定顯示為第一筆
      If textCP44.ListCount > 0 Then
         textCP44.ListIndex = 0
         textCP44_Validate False
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 先清除商標基本檔或服務業務基本檔欄位串列
   ClearTMSPFieldList
   ' 先清除案件進度檔欄位串列
   ClearCPFieldList
   
   ' 先取得本所案號
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 本所案號
      If IsNull(rsTmp.Fields("CP01")) = False Then: m_TM01 = rsTmp.Fields("CP01")
      If IsNull(rsTmp.Fields("CP02")) = False Then: m_TM02 = rsTmp.Fields("CP02")
      If IsNull(rsTmp.Fields("CP03")) = False Then: m_TM03 = rsTmp.Fields("CP03")
      If IsNull(rsTmp.Fields("CP04")) = False Then: m_TM04 = rsTmp.Fields("CP04")
   End If
   rsTmp.Close
    'End
   ' 本所案號
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)

   '抓出名代理人
   m_Case(1) = m_TM01
   m_Case(2) = m_TM02
   m_Case(3) = m_TM03
   m_Case(4) = m_TM04
   'Remove by Lydia 2021/08/12 改成Form 2.0,並且移到後方
   'PUB_SetOurAgent lstNameAgent, m_Case(), m_CP110
   
   ' 收文號
   textCP09 = m_CP09
   
   ' 取得國家代碼
   m_TM10 = GetNationNo(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   
   ' 取得基本檔
   QueryTradeMark

   PUB_SetOurAgent lstNameAgent, m_Case(), m_CP110, m_CP10, True 'Added by Lydia 2021/08/12
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030101_19 = Nothing
End Sub

' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' 發文日日期不正確
      If CheckIsDate(textCP27, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 發文日日期不可超過系統日
      'edit by nickc 2007/12/17 改系統日+1天
      'If Val(DBDATE(textCP27)) > Val(strSrvDate(1)) Then
      If Val(DBDATE(textCP27)) > Val(DBDATE(PUB_GetWorkDay(2))) Then
         Cancel = True
         strTit = "資料檢核"
         'edit by nickc 2007/12/17 改系統日+1天
         'strMsg = "發文日不可超過系統日"
         strMsg = "發文日不可超過系統日加一天"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textCP44_Click()
   textCP44_2 = m_AgentList(textCP44.ListIndex).aiName
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP44_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP44.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 代理人
Private Sub textCP44_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTempName As String   '2010/11/24 add by sonia
   
   Cancel = False
   If m_TM10 <> 台灣國家代號 Then
      If Len(Me.textCP44.Text) <= 0 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
   End If
   
   If textCP44.ListIndex >= 0 Then
      textCP44 = m_AgentList(textCP44.ListIndex).aiCode
   End If
   textCP44_2 = Empty
   If IsEmptyText(textCP44) = False Then
      If PUB_GetAgentNameAndState(m_TM01, textCP44.Text, strTempName) Then
         textCP44_2 = strTempName
      Else
         textCP44_2.Text = ""
         If strTempName <> "" Then
            Cancel = True
            Exit Sub
         End If
      End If
      If IsEmptyText(textCP44_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44_GotFocus
      Else
         ' 依所選擇的代理人找出案件進度檔中其收文日最大的一筆其彼所案號更新到畫面上的彼所案號欄位
         strSql = "SELECT CP45 FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP44 = '" & textCP44 & "' AND " & _
                        "CP05 IN (SELECT MAX(CP05) FROM CASEPROGRESS " & _
                                 "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                       "CP02 = '" & m_TM02 & "' AND " & _
                                       "CP03 = '" & m_TM03 & "' AND " & _
                                       "CP04 = '" & m_TM04 & "' AND " & _
                                       "CP44 = '" & textCP44 & "')"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CP45")) = False Then
               textTM45 = rsTmp.Fields("CP45")
            End If
         End If
         rsTmp.Close
      End If
   End If
   Set rsTmp = Nothing
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
Dim strCP64 As String
   
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   SetCPFieldNewData "CP30", textCP30
   ' 代理人
   If IsEmptyText(textCP44) = False Then
      SetCPFieldNewData "CP44", textCP44 & String(9 - Len(textCP44), "0")
      'add by nickc 2008/02/22
      m_CP44New = textCP44 & String(9 - Len(textCP44), "0")
   Else
      SetCPFieldNewData "CP44", textCP44
      'add by nickc 2008/02/22
      m_CP44New = textCP44
   End If
   SetCPFieldNewData "CP22", textCP22
   SetCPFieldNewData "CP84", textCP84
   SetCPFieldNewData "CP110", m_CP110
   'cancel by sonia 2023/12/22  電子送件時不算發文室件數
   'SetCPFieldNewData "CP123", "Y"    '2009/4/21 ADD BY SONIA 算發文室件數
   ' 進度備註
   strCP64 = Me.textCP64.Text
   SetCPFieldNewData "CP64", strCP64
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
   ' 案件備註
   SetTMSPFieldNewData "TM58", textTM58
End Sub

' 更新商標基本檔的相關欄位
Private Sub OnUpdateTradeMark()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
End Sub

Public Function OnSaveData() As Boolean
Dim strTmp As String
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim nIndex As Integer
Dim bFirst As Boolean
Dim bDifference As Boolean
Dim strNP07 As String
Dim strNP08 As String
Dim strNP22 As String
Dim strCF10 As String
Dim strNP10 As String 'Add By Sindy 2014/9/11

OnSaveData = True
'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   'Modified by Lydia 2016/03/11 +案號
   'Call GetNP69("", m_TM10, m_CP13, strNP10) 'Add By Sindy 2014/9/11
   'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
   Call GetNA69("", m_TM10, m_CP13, strNP10, m_TM01, m_TM02, m_TM03, m_TM04)
   
   'Add By Sindy 2010/02/10
   '取得主管機關
   strSql = "SELECT * FROM CaseFee WHERE CF01='CFT' AND CF02='000' AND CF03='" & m_CP10 & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCF10 = "" & RsTemp("CF10")
   End If
   
   ' 更新案件進度檔
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            If UCase(m_CPList(nIndex).fiName) <> "CP30" Then 'Added by Lydia 2023/08/08 改在後面更新
               strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
            End If 'Added by Lydia 2023/08/08
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_CPList(nIndex).fiName & " = " & m_CPList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   ' 設定SQL語法更新的條件
   'Modified by Lydia 2023/08/08  發文作業:更新
   'strSql = strSql & ",CP130='" & strCF10 & "' " 'Add By Sindy 2010/02/10
   'modify by sonia 2024/2/5 申請台灣案的英文證明同時更新承辦人為發文操作人員,僅發文才更新，申請書時不更新
   'If m_ProcType = "2" Then strSql = strSql & ",CP130='" & strCF10 & "'"
   If m_ProcType = "2" Then strSql = strSql & ",CP130='" & strCF10 & "',CP14='" & strUserNum & "'"
   'end 2024/2/5
   
   strSql = strSql & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "

   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
   
   If m_ProcType = "1" Then 'Added by Lydia 2023/08/08 產生申請書
      '2013/9/10 ADD BY SONIA 加註CP64
      'Modified by Lydia 2023/08/08
      'strSql = "UPDATE CASEPROGRESS SET CP64=CP64||'" & ";申請第" & textCP30 & "號之英文證明" & "' WHERE CP09 = '" & m_CP09 & "' "
      strTmp = ""  '申請英文證明號數為註冊第xxxxx號或申請第xxxxx號
      If Trim(textCP30) <> "" Then
         strTmp = strTmp & ";申請英文證明號數為註冊第" & textCP30 & "號"
      ElseIf Trim(textCP30_1) <> "" Then
         strTmp = strTmp & ";申請英文證明號數為申請第" & textCP30_1 & "號"
      End If
      '申請英文證明=>預設為電子送件CP118
     'modify by sonia 2023/12/22  電子送件時不算發文室件數cp123=null
      strSql = "UPDATE CASEPROGRESS SET CP30='" & IIf(Trim(textCP30) <> "", Trim(textCP30), Trim(textCP30_1)) & "', " & _
               "CP64=CP64||'" & ChgSQL(strTmp) & "', CP118='Y',CP123=null WHERE CP09 = '" & m_CP09 & "' "
      'end 2023/08/08
      cnnConnection.Execute strSql
      '2013/9/10 END
   
   'Added by Lydia 2023/08/08
   Else      '發文
      '更新商標基本檔
      OnUpdateTradeMark
      
      'Add By Sindy 2012/9/10
      ' 若有審查天數, 新增一筆催審期限的記錄到下一程序檔
      strSql = "SELECT * FROM CaseFee " & _
               "WHERE CF01 = '" & m_TM01 & "' AND " & _
                     "CF02 = '" & m_TM10 & "' AND " & _
                     "CF03 = '" & m_CP10 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CF05")) = False Then
            strNP07 = "305"
            strNP08 = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
            strNP22 = GetNextProgressNo()
            'Modify By Sindy 2014/9/11 m_CP14=>strNP10
            'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
            'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                               strNP08 & "," & strNP08 & ",'" & strNP10 & "'," & strNP22 & ")"
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                               PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & strNP10 & "'," & strNP22 & ")"
            cnnConnection.Execute strSql
         End If
      End If
      rsTmp.Close
      '2012/9/10 End
      
      Set rsTmp = Nothing
      
      'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
      Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   End If 'Added by Lydia 2023/08/08
   
   cnnConnection.CommitTrans
    
     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44New, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22
    
    Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
    OnSaveData = False
End Function

Private Sub textCP30_GotFocus()
   InverseTextBox textCP30
End Sub

Private Sub textCP30_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP30, 20) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "註冊號內容太長"
      textCP30_GotFocus
   End If
   'Added by Lydia 2023/08/08
   If textCP30.Text <> textCP30.Tag Then
      If ChkTm12Tm15 = "" Then
         Frame1.Enabled = True
      Else
         Frame1.Enabled = False
         txtOther(0) = "": txtOther(1) = ""
      End If
   End If
   textCP30.Tag = textCP30.Text
   'end 2023/08/08
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
      strMsg = "案件備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM58_GotFocus
   End If
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   CheckDataValid = False
   'add by nickc 2008/05/01
   If IsDebt(m_TM10, textCP09) Then
        strTit = "警告！禁止發文！"
        strMsg = "未收款且無 預定收款日 請轉告智權同仁！！"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        GoTo EXITSUB
   End If
   
   If m_ProcType = "2" Then 'Added by Lydia 2023/08/08 發文作業
      ' 發文日
      If IsEmptyText(textCP27) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27.SetFocus
         GoTo EXITSUB
      End If
      ' 代理人
      If IsEmptyText(textCP44) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入代理人"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44.SetFocus
         GoTo EXITSUB
      End If
   End If 'Added by Lydia 2023/08/08
   ' 註冊號
   'Modified by Lydia 2023/08/02
   'If IsEmptyText(textCP30) = True Then
   '   strTit = "檢核資料"
   '   strMsg = "請輸入國內註冊號"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   textCP30.SetFocus
   '   GoTo EXITSUB
   'End If
   If m_ProcType = "1" Then  '產生申請書
      If Trim(textCP30 & textCP30_1) = "" Then
         MsgBox "請輸入申請案號或註冊號數！", vbCritical + vbOKOnly, "檢核資料"
         textCP30_1.SetFocus
         textCP30_1_GotFocus
         GoTo EXITSUB
      Else
         strExc(1) = ChkTm12Tm15
         If strExc(1) = "" And (Trim(txtOther(0)) = "" Or Trim(txtOther(1)) = "") Then
            Frame1.Enabled = True
            If Trim(txtOther(0)) = "" Then
               MsgBox "請輸非本所案件的商標名稱、商標或標章種類!！", vbCritical + vbOKOnly, "檢核資料"
               txtOther(0).SetFocus
               txtOther_GotFocus 0
               GoTo EXITSUB
            End If
            If Trim(txtOther(1)) = "" Then
               MsgBox "請輸入非本所案件的商標種類！", vbCritical + vbOKOnly, "檢核資料"
               txtOther(1).SetFocus
               txtOther_GotFocus 1
               GoTo EXITSUB
            Else
            End If
         ElseIf strExc(1) <> "" Then
            Frame1.Enabled = False
         End If
      End If
      If Val(textCP84_1) <= 0 Then
         MsgBox "請輸入份數！", vbCritical + vbOKOnly, "檢核資料"
         textCP84_1.SetFocus
         textCP84_1_GotFocus
         GoTo EXITSUB
      'Modified by Lydia 2024/01/15 + 判斷金額不一致  Val(m_CP17) <> Val(Trim(textCP84.Text))
      ElseIf Val(m_CP17) <> Val(Trim(textCP84.Text)) Then
         If MsgBox("收文規費[" & Val(m_CP17) & "] 與實際發文規費[" & Val(Trim(textCP84.Text)) & "]不同，是否繼續存檔？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            textCP84_1.SetFocus
            textCP84_1_GotFocus
            GoTo EXITSUB
         End If
      End If
   End If
   'end 2023/08/08
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textTM58_GotFocus()
   InverseTextBox textTM58
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   
   '發文規費，申請國家台灣才檢查
   If Me.textCP84.Enabled = True Then
      Cancel = False
      textCP84_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2009/4/21 CANCEL BY SONIA 費用併入其他案件性質,此處不必檢查
   'If textCP84.Enabled = True And m_TM10 = "000" Then
   '    If Val(textCP84.Text) <> Val(m_CP84) Then
   '        If MsgBox("發文規費[" & Trim(Val(m_CP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同", vbOKCancel) = vbCancel Then
   '            textCP84_GotFocus
   '            Exit Function
   '        End If
   '    End If
   'End If
   '2009/4/21 END
   
   'Modified by Lydia 2023/08/08 產生申請書
   'If lstNameAgent.Enabled = True Then
   If lstNameAgent.Enabled = True And m_ProcType = "1" Then
       Cancel = False
       lstNameAgent_Validate Cancel
       If Cancel = True Then
           Exit Function
       End If
   End If
   
   If m_ProcType = "2" Then 'Added by Lydia 2023/08/08 發文作業
      If Me.textCP27.Enabled = True Then
         Cancel = False
         textCP27_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
      
      If Me.textCP44.Enabled = True Then
         Cancel = False
         textCP44_Validate Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   End If 'Added by Lydia 2023/08/08
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
    'Added by Lydia 2021/06/04 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
    If Pub_ChkACS112isNull(m_TM01, m_TM02, m_TM03, m_TM04, txtCP113) = True Then
        txtCP113.SetFocus
        txtCP113_GotFocus
        Exit Function
    End If
    'end 2021/06/04
       
   'Added by Lydia 2023/08/08
   If Frame1.Enabled = True And Frame1.Visible = True Then
      Cancel = False
      txtOther_Validate 1, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2023/08/08
   
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
   
   TxtValidate = True
End Function

'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/12/5 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Lydia 2021/08/12 改模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      textCP22 = ""
   Else
      textCP22 = "N"
   End If
End Sub
Private Sub textCP84_GotFocus()
   InverseTextBox textCP84
End Sub

Private Sub textCP84_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
If IsEmptyText(textCP84) = False Then
    If IsNumeric(textCP84) = False Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請輸入數字"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP84_GotFocus
    Else
        textCP84.Text = Trim(Val(textCP84.Text))
    End If
End If
End Sub

'Added by Lydia 2021/06/04
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/06/04
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

'Added by Lydia 2023/08/08
Private Sub textCP84_1_GotFocus()
   InverseTextBox textCP84_1
End Sub

Private Sub textCP84_1_Validate(Cancel As Boolean)

   If IsEmptyText(textCP84_1) = False Then
      If IsNumeric(textCP84_1) = False Then
         Cancel = True
         MsgBox "請輸入數字", vbCritical + vbOKOnly, "資料檢核"
         textCP84_1_GotFocus
         textCP84_1.SetFocus
      Else
         textCP84_1.Text = Trim(Val(textCP84_1.Text))
         textCP84.Text = Trim(Val(textCP84_1) * 500)
      End If
   End If
End Sub

Private Sub textCP30_1_GotFocus()
   InverseTextBox textCP30_1
End Sub

Private Sub textCP30_1_Validate(Cancel As Boolean)
   If textCP30_1.Text <> textCP30_1.Tag Then
      If ChkTm12Tm15 = "" Then
         Frame1.Enabled = True
      Else
         Frame1.Enabled = False
         txtOther(0) = "": txtOther(1) = ""
      End If
   End If
   textCP30_1.Tag = textCP30_1.Text
End Sub

Private Sub txtOther_GotFocus(Index As Integer)
   InverseTextBox txtOther(Index)
End Sub

Private Sub txtOther_Validate(Index As Integer, Cancel As Boolean)
   
   If Index = 2 Then
      If Trim(txtOther(Index)) <> "" Then
         If Val(txtOther(Index)) < 0 Or Val(txtOther(Index)) > 5 Then
            MsgBox "請輸入商標種類1~5", vbCritical + vbOKOnly, "資料檢核"
            Cancel = True
            txtOther_GotFocus Index
            txtOther(Index).SetFocus
         End If
      End If
   End If
End Sub

'檢查輸入的申請案號/註冊號數是否為本所案件
Private Function ChkTm12Tm15() As String
Dim intQ As Integer, strQ1 As String
Dim rsQuery As New ADODB.Recordset
   
   ChkTm12Tm15 = ""
   m_InCase(1) = "": m_InCase(2) = "": m_InCase(3) = "": m_InCase(4) = ""
   If Trim(textCP30 & textCP30_1) <> "" Then
      If Trim(textCP30) <> "" Then
         strQ1 = strQ1 & " OR tm15='" & Trim(textCP30) & "'"
      End If
      If Trim(textCP30_1) <> "" Then
         strQ1 = strQ1 & " OR tm12='" & Trim(textCP30_1) & "'"
      End If
      strQ1 = "select tm01,tm02,tm03,tm04,tm12,tm15 from trademark where " & Mid(strQ1, 4)
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         If "" & rsQuery.Fields("tm15") <> "" Then
            ChkTm12Tm15 = "2" '註冊號數
            If Trim(textCP30) = "" Then
               textCP30 = "" & rsQuery.Fields("tm15")
            End If
         ElseIf "" & rsQuery.Fields("tm12") <> "" Then
            ChkTm12Tm15 = "1" '申請案號
         End If
         m_InCase(1) = "" & rsQuery.Fields("tm01")
         m_InCase(2) = "" & rsQuery.Fields("tm02")
         m_InCase(3) = "" & rsQuery.Fields("tm03")
         m_InCase(4) = "" & rsQuery.Fields("tm04")
      End If
      Set rsQuery = Nothing
   End If
End Function

'Added by Lydia 2023/08/08 電子送件申請書
Private Sub GetApplBook_CFT()
Dim intWhere As Integer
Dim strFolder As String, strFileName As String
Dim strContent As String
Dim ET03 As String, ET03_1 As String
Dim strErrMsg As String
Dim m_CaseDocNo As String  'CFT本所案號

   strErrMsg = "讀取基本檔"
   If Frame1.Enabled = True Then
      '非本所案件=>改代入CFT案
      m_InCase(1) = m_TM01: m_InCase(2) = m_TM02: m_InCase(3) = m_TM03: m_InCase(4) = m_TM04
      intWhere = 國外_CF
   Else
      If ClsPDGetSystemKind(m_InCase(1), , , intWhere) = False Then
         GoTo EXITSUB
      End If
   End If
   If ClsPDReadTrademarkDatabase(m_InCase(), intWhere) = False Then
      GoTo EXITSUB
   End If
   
   strErrMsg = "建立資料夾"
   '申請書的檔名和附件統一用本所案號
   m_CaseDocNo = PUB_FCPCaseNo2FileName(m_TM01, m_TM02, m_TM03, m_TM04)
   '桌面上建立案號資料夾
   strFolder = PUB_Getdesktop
   strFolder = strFolder & "\" & m_CaseDocNo
   If Dir(strFolder, vbDirectory) = "" Then
       MkDir strFolder
   End If
   
   ET03 = "21"  '參考FCT案frm03020602_1
   ET03_1 = "11"
   strFileName = "英文證明書申請書"
   
   '申請書:要基本資料表,先不存檔
   If StartLetter2("90", ET03, m_CP09, "2") Then
      NowPrint m_CP09, "90", ET03, False, strUserNum, , , True, strContent
      strFileName = strFolder & "\" & m_CaseDocNo & "." & strFileName
   Else
      strErrMsg = "讀取申請書"
      GoTo EXITSUB
   End If

   '基本資料表
   If StartLetter2("90", ET03_1, m_CP09, "1") Then
      NowPrint m_CP09, "90", ET03_1, False, strUserNum, , strContent, True, strContent
      strContent = Replace(strContent, vbCrLf & Chr(12), vbCrLf & "|#(分節)#|")    '換頁符號Chr(12)替換為分節符號 "|#(分節)#|"
      Call PUB_MakeDoc(strContent, strFileName, , , , , True)  '分節處理頁碼
   Else
      strErrMsg = "讀取基本資料表"
      GoTo EXITSUB
   End If
   
   MsgBox "電子送件申請書已產生在" & strFolder
   Exit Sub
   
EXITSUB:
   MsgBox "產生申請書失敗: " & vbCrLf & strErrMsg, vbCritical
End Sub

'Added by Lydia 2023/08/08 各式申請書-電子送件申請書
Private Function StartLetter2(ByVal iET01 As String, ByVal iET03 As String, ByVal iCp09 As String, ByVal iKind As String) As Boolean
'iKind = 1.基本資料表, 2.申請書
Dim strTxt(1 To 40) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim intA As Integer
Dim m_DocNo As String 'CFT本所案號

   EndLetter iET01, iCp09, iET03, strUserNum
   m_DocNo = m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 <> "000", "-" & m_TM03 & "-" & m_TM04, "")
   ii = 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','本所案號','" & m_DocNo & "')"
   
   '申請人資料=>因為申請T案的證明,所以抓T案的申請人
   'Modified by by Lydia 2023/09/23 不抓案件申請人，改抓申請人資料
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, m_InCase(), True, , , m_TM01)
   Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, m_CP10, m_InCase(), False, , , m_TM01)
   
   '出名代理人: 共用模組取得資料
   strExc(0) = PUB_GetAgentCP110(iCp09, m_CP110, "CFT", "4")
   If strExc(0) <> "" Then
       tmpArr1 = Split(strExc(0), "|")
       For jj = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(jj)) <> "" Then
               tmpArr2 = Empty
               tmpArr2 = Split(tmpArr1(jj), ",")
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-證書字號','" & tmpArr2(0) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-ID','" & tmpArr2(1) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-中文姓名','" & PUB_ConvertNameFormat("" & tmpArr2(2)) & "')"
           End If
       Next jj
   End If

   '基本資料表
   If iKind = "1" Then
      'Modified by Lydia 2023/12/22 傳入本所案號m_TM01~m_TM04
      Call GetNA69(strUserNum, m_TM10, m_CP13, strTmp, m_TM01, m_TM02, m_TM03, m_TM04)
      If strTmp <> "" Then
         strExc(0) = "select st07 from staff where st01='" & strTmp & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            ii = ii + 1
            'CFT承辦分機
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','CFT承辦分機','" & "" & RsTemp.Fields("st07") & "')"
         End If
      End If
   End If
   
   '申請書
   If iKind = "2" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','CFT本所案號','" & m_DocNo & "')"
      '依輸入判斷帶入資料
      If textCP30 <> "" Then
         strExc(1) = "註冊號"
         strExc(2) = Trim(textCP30)
      Else
         strExc(1) = "申請案號"
         strExc(2) = Trim(textCP30_1)
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商標號數標題-輸入','" & strExc(1) & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商標號數-輸入','" & strExc(2) & "')"
      '非本所案
      If Frame1.Enabled = True And txtOther(0) <> "" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商標名稱-輸入','" & ChgSQL(Trim(txtOther(0))) & "')"
         strExc(3) = ""
         Select Case txtOther(1)
            Case "1": strExc(3) = "商標"
            Case "2": strExc(3) = "商標(92年修正前服務標章)"
            Case "3": strExc(3) = "團體商標"
            Case "4": strExc(3) = "團體標章"
            Case "5": strExc(3) = "證明標章"
         End Select
         If strExc(3) <> "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商標種類-輸入','" & ChgSQL(strExc(3)) & "')"
         End If
      '本所案件
      Else
         m_MySt(1) = m_InCase(1)
         m_MySt(2) = m_InCase(2)
         m_MySt(3) = m_InCase(3)
         m_MySt(4) = m_InCase(4)
         m_SysKind = CheckSys(m_InCase(1))
         SetLetterSt
         If m_InCase(5) <> "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商標名稱-輸入','" & ChgSQL(m_InCase(5)) & "')"
         End If
         strExc(3) = ExceptFieldData("商標種類國內名稱電子送件")
         If strExc(3) <> "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商標種類-輸入','" & ChgSQL(strExc(3)) & "')"
         End If
      End If
      
      '繳費金額
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','繳費金額','" & Val(textCP84.Text) & "')"
      ii = ii + 1
      'Added by Lydia 2024/01/18 帶入第1申請人中文名稱(CFT案和T案相同申請人),不用X商
      'Modified by Lydia 2024/03/05 所內案用T案申請人,所外案用CFT案申請人;參考PUB_GetApplFCT_EData 修法:106/12/01開始中文名稱要加外商國名
      'strTmp = GetCustomerName(m_TM23)
      If m_InCase(23) <> "" Then
         strExc(0) = ChangeCustomerL(m_InCase(23))
      Else
         strExc(0) = ChangeCustomerL(m_TM23)
      End If
      strTmp = PUB_GetApplT_CNAME(strExc(0))
      'end 2024/03/05
      If strTmp <> "" Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','收據抬頭','" & ChgSQL(strTmp) & "')"
      Else
      'end 2024/01/18
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','收據抬頭','♀')"
      End If

      '份數
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','份數-輸入','" & Val(textCP84_1.Text) & "')"
            
      '指定使用商品(服務)中文名稱
      If Frame1.Enabled = True Then  '非本所案件
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商品服務類別名稱中文','♀')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商品服務類別名稱英文','♀')"
      Else
         strExc(0) = BeforePrintGetDBData("TMGoods:" & m_InCase(1) & "-" & m_InCase(2) & "-" & m_InCase(3) & "-" & m_InCase(4) & "-中文", True)
         If strExc(0) <> "" Then
             '單一類別的案件,開頭不顯示類別代號 (嘉雯&阿蓮的溝通結果)
             If InStr(m_InCase(9), ",") = 0 Then
                  strExc(0) = Mid(strExc(0), InStr(strExc(0), "：") + 1)
             End If
             If Trim(strExc(0)) <> "" Then
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商品服務類別名稱中文','" & ChgSQL(strExc(0)) & "')"
             End If
         End If
         strExc(1) = BeforePrintGetDBData("TMGoods:" & m_InCase(1) & "-" & m_InCase(2) & "-" & m_InCase(3) & "-" & m_InCase(4) & "-英文", True)
         If strExc(1) <> "" Then
            '單一類別的案件,開頭不顯示類別代號 (嘉雯&阿蓮的溝通結果)
            If InStr(m_InCase(9), ",") = 0 Then
                 strExc(1) = Mid(strExc(1), InStr(strExc(1), "：") + 1)
            End If
            If Trim(strExc(1)) <> "" Then
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商品服務類別名稱英文','" & ChgSQL(strExc(1)) & "')"
            End If
         End If
      End If '---- If Frame1.Enabled = True Then
      
      '附送書件
       ii = ii + 1
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-基本資料表', '" & m_DocNo & ".contact.pdf')"
       ii = ii + 1
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-委任書', '" & m_DocNo & ".poa.pdf')"
       ii = ii + 1
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-廠商基本資料', '" & m_DocNo & ".att.pdf')"
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

