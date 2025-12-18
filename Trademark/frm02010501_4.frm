VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010501_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "爭議案勝訴輸入"
   ClientHeight    =   5508
   ClientLeft      =   96
   ClientTop       =   1008
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5508
   ScaleWidth      =   9144
   Begin VB.CommandButton Command2 
      Caption         =   "商品及服務資料輸入(&I)"
      Height          =   400
      Left            =   4200
      TabIndex        =   12
      Top             =   70
      Width           =   1965
   End
   Begin VB.TextBox textTM14 
      Height          =   264
      Left            =   5700
      MaxLength       =   8
      TabIndex        =   1
      Top             =   3060
      Width           =   1092
   End
   Begin VB.TextBox TextCP64_1 
      Height          =   264
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   2
      Top             =   3330
      Width           =   2532
   End
   Begin VB.TextBox textCP26_S 
      Height          =   264
      Left            =   6300
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3900
      Width           =   372
   End
   Begin VB.TextBox textCP40 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   1860
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   960
      Width           =   2532
   End
   Begin VB.TextBox textTM16 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   960
      Width           =   2532
   End
   Begin VB.TextBox textTM32 
      Height          =   264
      Left            =   1260
      MaxLength       =   699
      TabIndex        =   9
      Top             =   4500
      Width           =   7752
   End
   Begin VB.TextBox textTM17 
      Height          =   264
      Left            =   1740
      MaxLength       =   1
      TabIndex        =   8
      Top             =   4200
      Width           =   372
   End
   Begin VB.TextBox textCP26 
      Height          =   264
      Left            =   1740
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3900
      Width           =   372
   End
   Begin VB.TextBox textCP14 
      Height          =   264
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   4
      Top             =   3600
      Width           =   732
   End
   Begin VB.TextBox textCP48 
      Height          =   264
      Left            =   5700
      MaxLength       =   7
      TabIndex        =   5
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   660
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1860
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2460
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2460
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   660
      Width           =   2532
   End
   Begin VB.TextBox textCP08 
      Height          =   264
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   0
      Top             =   3060
      Width           =   2532
   End
   Begin VB.TextBox textCP35 
      Height          =   264
      Left            =   5700
      MaxLength       =   32
      TabIndex        =   3
      Top             =   3315
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8244
      TabIndex        =   15
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6192
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7020
      TabIndex        =   14
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.TextBox textTM58 
      Height          =   300
      Left            =   1260
      TabIndex        =   11
      Top             =   5100
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
   Begin MSForms.TextBox textCP14_2 
      Height          =   264
      Left            =   2100
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1692
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      MaxLength       =   20
      Size            =   "2984;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   300
      Left            =   1260
      TabIndex        =   10
      Top             =   4800
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
   Begin MSForms.TextBox textCP14_Src 
      Height          =   264
      Left            =   1260
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2760
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
      Left            =   1260
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1560
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
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5700
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2160
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
      Left            =   1230
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1230
      Width           =   7752
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13674;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      Caption         =   "註冊公告日 :"
      Height          =   255
      Left            =   4560
      TabIndex        =   59
      Top             =   3060
      Width           =   1035
   End
   Begin VB.Label Label17 
      Caption         =   "來文字號 :"
      Height          =   255
      Left            =   180
      TabIndex        =   58
      Top             =   3345
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "案件備註 :"
      Height          =   255
      Left            =   180
      TabIndex        =   54
      Top             =   5100
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "商品群組 :"
      Height          =   255
      Left            =   180
      TabIndex        =   53
      Top             =   4500
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "目前准駁 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   52
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label11 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   180
      TabIndex        =   51
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label10 
      Caption         =   "(Y / N)"
      Height          =   255
      Left            =   6780
      TabIndex        =   50
      Top             =   3900
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "是否計算勝訴率 :"
      Height          =   255
      Left            =   4740
      TabIndex        =   49
      Top             =   3900
      Width           =   1455
   End
   Begin VB.Label Label20 
      Caption         =   "(Y / N)"
      Height          =   255
      Left            =   2220
      TabIndex        =   48
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label19 
      Caption         =   "專用權是否存在 :"
      Height          =   255
      Left            =   180
      TabIndex        =   47
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不算)"
      Height          =   255
      Left            =   2220
      TabIndex        =   46
      Top             =   3900
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "是否算案件數 :"
      Height          =   255
      Left            =   180
      TabIndex        =   45
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   180
      TabIndex        =   44
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "對照名稱 :"
      Height          =   252
      Left            =   180
      TabIndex        =   42
      Top             =   1860
      Width           =   972
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   255
      Left            =   180
      TabIndex        =   41
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   180
      TabIndex        =   40
      Top             =   2760
      Width           =   852
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   255
      Left            =   4740
      TabIndex        =   39
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   37
      Top             =   660
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   180
      TabIndex        =   36
      Top             =   1230
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   180
      TabIndex        =   35
      Top             =   1560
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   180
      TabIndex        =   34
      Top             =   2160
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   4740
      TabIndex        =   33
      Top             =   1860
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   252
      Index           =   8
      Left            =   4740
      TabIndex        =   32
      Top             =   2760
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   9
      Left            =   4740
      TabIndex        =   31
      Top             =   2460
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   10
      Left            =   180
      TabIndex        =   30
      Top             =   2460
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4740
      TabIndex        =   29
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "審定號 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   28
      Top             =   660
      Width           =   732
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   255
      Left            =   180
      TabIndex        =   27
      Top             =   3060
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "審查委員 :"
      Height          =   255
      Left            =   4740
      TabIndex        =   26
      Top             =   3315
      Width           =   855
   End
End
Attribute VB_Name = "frm02010501_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/03 Form2.0已修改 cmbTM05/textTM23/textCP13/textCP14_Src/textCP64/textTM58
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
' 商品類別  2011/4/7 ADD BY SONIA
Dim m_TM09 As String
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
' 預估結果
Dim m_CP23 As String
'Add By Cheng 2002/01/15
Dim m_intNumBegin As Integer
Dim m_intNumEnd As Integer
Dim m_strNumBegin As String
Dim m_strNumEnd As String
'2011/4/7 ADD BY SONIA 檢查是否已經有商品及服務
Public ChkTG As Boolean
Dim BolPrintCaseCheck As Boolean 'Add By Sindy 2012/4/16
'Added by Morgan 2017/4/24 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/4/24
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
Dim m_TM28 As String 'add by sonia 2021/5/3 卷宗性質

'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010501_3.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm02010501_3
   Unload frm02010501_2
   Unload frm02010501_1
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
      Unload frm02010501_3
      Unload frm02010501_2
      'Add By Sindy 2019/5/22
      If Me.m_strIR01 <> "" Then
         Unload frm02010501_1
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
         '2019/5/22 END
      'Modified by Morgan 2017/4/25 電子公文
      'frm02010501_1.Show
      ElseIf m_DocNo <> "" Then
         Unload Me
         Unload frm02010501_1
         frm02010412.GoNext
      Else
         frm02010501_1.Show
         Unload Me
      End If
      'end 2017/4/25
   End If
End Sub

'2011/4/7 ADD BY SONIA
Private Sub Command2_Click()
   frm03010303_04.Hide
   Set frm03010303_04.UpForm = Me
   frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   frm03010303_04.AllClass = m_TM09
   frm03010303_04.cmdOK(2).Visible = True
   
   If m_TM09 <> "" Then
      Me.Hide
      frm03010303_04.QueryData
      frm03010303_04.Show vbModal
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
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010501_1.m_strIR01
   m_strIR02 = frm02010501_1.m_strIR02
   m_strIR03 = frm02010501_1.m_strIR03
   m_strIR04 = frm02010501_1.m_strIR04
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
   '2011/7/23 ADD BY SONIA 加入TD,TD-000151
   Select Case m_TM01
      Case "TD":
         ' 設定SQL語法
         strSql = "SELECT SP01 AS TM01,SP02 AS TM02,SP03 AS TM03,SP04 AS TM04,SP05 AS TM05,SP06 AS TM06,SP07 AS TM07,SP09 AS TM10" & _
            ",'' AS TM12,'' AS TM15,'' AS TM16,'' AS TM28,SP08 AS TM23,SP27 AS TM45,'' AS TM17,'' AS TM32,'' AS TM09,SP18 AS TM58,SP26 as TM44 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
      Case Else
   '2011/7/23 END
      strSql = "SELECT * FROM TradeMark " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "'"
   End Select  '2011/7/23 ADD BY SONIA
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
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
      ' 商品組群
      If IsNull(rsTmp.Fields("TM32")) = False Then
         textTM32 = rsTmp.Fields("TM32")
      End If
      ' 案件備註
      If IsNull(rsTmp.Fields("TM58")) = False Then
         textTM58 = rsTmp.Fields("TM58")
      End If
      ' 商品類別 2011/4/7 ADD BY SONIA
      m_TM09 = ""
      If IsNull(rsTmp.Fields("TM09")) = False Then
         m_TM09 = rsTmp.Fields("TM09")
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
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 機關文號
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
      
      'Add By Sindy 2009/06/29
      '相關總收文號若其案件性質為1205.部分核駁時，將其進度備註帶至畫面上
      If IsNull(rsTmp.Fields("CP43")) = False And m_CP10 = "401" And m_TM10 = "020" Then
         strExc(0) = "select * from caseprogress where cp09='" & rsTmp.Fields("CP43") & "' and cp10='1205' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If IsNull(RsTemp.Fields("CP64")) = False Then
               textCP64 = textCP64.Text & RsTemp.Fields("CP64")
               nResponse = MsgBox("此為部分核駁提復審之勝訴，已將部分核駁商品預設在進度備註欄", vbOKOnly, "提示")
            End If
         End If
      End If
      '2009/06/29 End
   End If
   rsTmp.Close
   
   ' 預設承辦期限
   ' 承辦期限的日期應為來函收文日加上工作天數
   ' 工作天數由系統別+國家代碼+案件性質(勝訴)搜尋案件收費表的工作天數
   ' 若有值才預設
''''edit by nickc 2007/10/12 改抓有時效的
''''   strDay = GetWorkDays(m_TM01, m_TM10, "1003")
''''   If IsEmptyText(strDay) = False Then
''''      ' 90.07.03 modify by louis (承辦期限以實際的工作天數來計算)
''''      'textCP48 = TAIWANDATE(Format(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay))))
''''      textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''   End If
   textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1003", DBDATE(m_CP05), , textCP09))
   
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
   
   If m_TM01 <> "TD" Then '2012/7/23 ADD BY SONIA TD-000151
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
               'modify by sonia 2017/9/1 +623,624
               Case "605", "606", "623", "624"
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
                   'Modify By Cheng 2002/11/22
   '               textCP08 = "經訴字第"
                  textCP08 = "經訴字第號"
                   'Modify By Cheng 2002/11/22
                   '不預帶審查委員
   '               'Add By Cheng 2002/01/15
   '               m_strNumBegin = "訴"
   '               m_strNumEnd = "字"
               Case "403"
                  textCP08 = strTmp & "年度訴字第號"
                   'Modify By  Cheng 2002/11/22
                   'Modify By Cheng 2002/11/22
                   '不預帶審查委員
   '               m_strNumBegin = "訴"
   '               m_strNumEnd = "字"
            End Select
         End If
      End If
   End If   '2012/7/23 ADD BY SONIA
   
   
   'Added by Morgan 2017/4/24 電子公文
   If m_DocNo <> "" Then
      If m_DocWord <> "" Then
         TextCP64_1 = m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號"
      Else
         TextCP64_1 = Replace(TextCP64_1, "第號", "第" & PUB_GetEDocNo(m_DocNo) & "號")
      End If
      textCP64_1_LostFocus
   End If
   'end 2017/4/24
   
   'Add By Sindy 2009/06/09
   '馬德里商標領土延伸至大陸被核駁復審案，可在此輸入註冊日，以管制延展期限
   'modify by sonia 2021/5/3 加卷宗性質條T-223726
   If m_TM10 = "020" And m_CP10 = "401" And Left(Trim(textTM15), 1) = "G" And m_TM28 = "1" Then
      textTM14.Visible = True
      Label18.Visible = True
   Else
      textTM14.Visible = False
      Label18.Visible = False
   End If
   '2009/06/09 End
   
    'Add By Cheng 2002/11/12
    textCP08_GotFocus
    
End Sub

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
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
   Dim strNP09 As String
   Dim strNP08 As String
   Dim strTM22 As String
   
'Add By Cheng 2002/11/07
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
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
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
   If bUpdate = True Then
      strSql = "UPDATE CaseProgress SET CP24 = '1', " & _
                                       "CP25 = " & DBDATE(m_CP05) & ", " & _
                                       "CP35 = '" & textCP35 & "' " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 依是否計算勝訴率來更新原案件進度檔資料的是否算案件數欄位
   Select Case textCP26_S
      ' 91.05.16 modify by louis
      'strSQL = "UPDATE CaseProgress SET CP26 = '" & textCP26_S & "' " & _
      '            "WHERE CP09 = '" & m_CP09 & "' "
      'cnnConnection.Execute strSQL
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
   ' 91.03.25 modify by louis (單引號)
   strSql = "UPDATE TradeMark SET TM17 = '" & textTM17 & "', " & _
                                 "TM32 = '" & textTM32 & "', " & _
                                 "TM58 = '" & ChgSQL(textTM58) & "' "
   strSql = strSql & strSubTMSQL
   cnnConnection.Execute strSql
   
   
   'Add By Sindy 2009/06/09
   If textTM14.Visible = True Then
      strTM22 = DBDATE(DateAdd("m", 120, ChangeWStringToWDateString(DBDATE(Me.textTM14.Text))))
      strSql = "UPDATE TradeMark SET TM16 = '1'," & _
                                    "TM17 = 'Y'," & _
                                    "TM14 = " & DBDATE(Me.textTM14.Text) & "," & _
                                    "TM21 = " & DBDATE(Me.textTM14.Text) & "," & _
                                    "TM22 = " & strTM22 & " "
      strSql = strSql & strSubTMSQL
      cnnConnection.Execute strSql
   End If
   '2009/06/09 End
   
   
   ' 更新商標基本檔的目前准駁欄, 審定來函日(准駁通知日)為來函收文日
   'If textTM16S = "Y" Then
   '   strSQL = "UPDATE TradeMark SET TM16 = '1', " & _
   '                                 "TM13 =" & DBDATE(m_CP05) & " "
   '   strSQL = strSQL & strSubTMSQL
   '   cnnConnection.Execute strSQL
   'End If
   
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   
   'Add By Sindy 2009/06/09
   If textTM14.Visible = True Then
      strNP07 = "102"
      ' 法定期限為專用期限截止日
      strNP09 = DBDATE(DateAdd("m", 120, ChangeWStringToWDateString(DBDATE(Me.textTM14.Text))))
      ' 本所期限為法定期限-2天
      'Modify By Sindy 2014/10/6 台灣案之本所期限設定
      If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
         strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
      Else
      '2014/10/6 END
         strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
      End If
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      ' 序號
      strNP22 = GetNextProgressNo()
      '智權人員存最近收文A類接洽記錄單的智權人員
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
                  "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
     cnnConnection.Execute strSql
   End If
   '2009/06/09 End
   
   
   ' 案件性質為勝訴
   strCP10 = "1003"
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 發文日為系統日
   'strCP27 = DBDATE(Date)
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2002/11/27
   '承辦人為原程序承辦人
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    
    'Modify by Morgan 2003/11/26
    
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP48,CP64) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
'                          "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
'                          "'" & textCP35 & "','" & m_CP36 & "','" & m_CP37 & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                          "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "'," & DBDATE(textCP48) & "," & _
'                          "'" & ChgSQL(textCP64) & "')"

   Dim strCP64 As String
   
   strCP64 = Trim(textCP64)
   If strCP64 <> "" And Trim(TextCP64_1) <> "" Then
      strCP64 = strCP64 & ",來文字號：" & Trim(TextCP64_1)
   ElseIf Trim(TextCP64_1) <> "" Then
      strCP64 = "來文字號：" & Trim(TextCP64_1)
   End If
    'Modify By Cheng 2004/02/03
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP48,CP64) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP14 & "'," & _
'                          "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
'                          "'" & textCP35 & "','" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                          "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "'," & DBDATE(textCP48) & "," & _
'                          "'" & ChgSQL(strCP64) & "')"
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP48,CP64) " & _
                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                          "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP14 & "'," & _
                          "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
                          "'" & textCP35 & "','" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
                          "'" & m_cp40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "'," & DBDATE(textCP48) & "," & _
                          "'" & ChgSQL(strCP64) & "')"
    'End
      '---End
   cnnConnection.Execute strSql
   
   'Add By Sindy 2020/1/7 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(m_TM01, 1) = "T" Then
      strLD18 = strCP09
      PUB_AddLetterProgress strLD18, 1, False, "", False, m_TM23, strCP10, m_TM44
   End If
   '2020/1/7 END
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
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
   
   '2010/3/29 add by sonia 大陸部分核駁1205來函提復審401勝訴時,重新掛原申請程序收文號的催審期限為來函收文日+6個月
   If m_TM10 = "020" And m_CP10 = "401" Then
      '先判斷是否為部分核駁1205來函才提的復審 m_CP09
      strSql = "SELECT C2.CP43 FROM CaseProgress C1,CASEPROGRESS C2 " & _
               "WHERE C1.CP01 = '" & m_TM01 & "' AND C1.CP02 = '" & m_TM02 & "' AND " & _
                     "C1.CP03 = '" & m_TM03 & "' AND C1.CP04 = '" & m_TM04 & "' AND " & _
                     "C1.CP09 = '" & m_CP09 & "' AND C1.CP43=C2.CP09(+) AND C2.CP10='1205'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         strNP08 = CompDate(1, 6, ChangeTStringToWString(frm02010501_1.textCP05))
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
   
   'Added by Lydia 2016/09/23 T大陸案1602被異議(理由) 或 602異議答辯之勝訴,新增管制催註冊証(1701)掛期限系統日+8個月
   'Modified by Lydia 2017/03/20 增加為被異議(理由) 、異議答辯、復審、起訴、上訴(1602,602,401,403,408)於勝訴(1003)或部分勝部分敗(1006)時，掛催註冊證期限為8個月。
   'modify by sonia 2021/5/3 加卷宗性質條件T-223726
   If m_TM01 = "T" And m_TM10 = "020" And m_TM28 = "1" And (m_CP10 = "1602" Or m_CP10 = "602" Or m_CP10 = "401" Or m_CP10 = "403" Or m_CP10 = "408") Then
        'modify by sonia 2024/9/12 8個月改6個月，本所期限改為工作日
        strNP08 = CompDate(1, 6, strSrvDate(1))
        strNP09 = strNP08
        strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
        ' 抓1602被異議(理由) 或 602異議答辯的承辦人
        'Modified by Lydia 2016/10/18 只抓申請,若申請為B類收文改抓CP31='Y'的A類收文承辦人
        'strSql = "SELECT CP14,ST04 FROM CaseProgress,STAFF WHERE CP09 = '" & m_CP09 & "' AND CP14=ST01(+) "
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
        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                 "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "', 1701 ," & _
                         strNP08 & "," & strNP09 & ",'" & strExc(1) & "'," & strNP22 & ")"
        cnnConnection.Execute strSql
   End If
   'end 2016/09/23
   
   'Added by Morgan 2017/4/25 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
   End If
   'end 2017/4/25
   
   'Add by Sindy 2019/5/22
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010501_1"
   End If
   '2019/5/22 END
   
   'add by nickc 2005/04/22
   Pub_UpdateEndModCash m_TM01, m_TM02, m_TM03, m_TM04
   
   Set rsTmp = Nothing
'Add By Cheng 2002/11/07
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
   ' 是否更新基本檔目前准駁不可空白
   'If IsEmptyText(textTM16S) = True Then
   '   strTit = "檢核資料"
   '   strMsg = "是否更新基本檔目前准駁不可為空白"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   GoTo ExitSub
   'End If
   
   'Add By Sindy 2009/06/09
   '2012/7/23 MODIFY BY SONIA TD-000151
   'If Me.textTM14.Visible = True Then
   If Me.textTM14.Visible = True And m_TM01 <> "TD" Then
      If Me.textTM14.Text = "" Then
          strTit = "資料檢核"
          strMsg = "請輸入註冊公告日"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          textTM14.SetFocus
          GoTo EXITSUB
      End If
   End If
   '2009/06/09 End
      
   ' 專用權是否存在不可為空白
   '2012/7/23 MODIFY BY SONIA TD-000151
   'If IsEmptyText(textTM17) = True Then
   If IsEmptyText(textTM17) = True And m_TM01 <> "TD" Then
      strTit = "檢核資料"
      strMsg = "專用權是否存在不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM17.SetFocus
      GoTo EXITSUB
   End If
   ' 91.01.22 modify by louis (取消商品組群的檢查)
   ' 商品組群不可為空白
   'If m_TM10 < "010" And (m_CP10 < "601" Or m_CP10 > "606") Then
   '   If IsEmptyText(textTM32) = True Then
   '      strTit = "檢核資料"
   '      strMsg = "商品組群不可為空白"
   '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '      textTM32.SetFocus
   '      GoTo EXITSUB
   '   End If
   'End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010501_4 = Nothing
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
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

'Add By Sindy 2009/06/09
Private Sub textTM14_GotFocus()
    TextInverse Me.textTM14
End Sub

'Add By Sindy 2009/06/09
Private Sub textTM14_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM14) = False Then
      If CheckIsDate(textTM14, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入西元年"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Call textTM14_GotFocus
         Exit Sub
      End If
   End If
End Sub

'Private Sub textTM16S_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub

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
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM58_GotFocus
   End If
End Sub


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
         'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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

'Private Sub textTM16S_GotFocus()
'   InverseTextBox textTM16S
'End Sub

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
'         .SelLength = 0
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

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

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

'Add By Sindy 2009/06/09
If Me.textTM14.Visible = True Then
   If Me.textTM14.Enabled = True Then
      Cancel = False
      textTM14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
End If
'2009/06/09 End

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

TxtValidate = True
End Function

