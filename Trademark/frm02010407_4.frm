VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010407_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "延期受理"
   ClientHeight    =   5748
   ClientLeft      =   156
   ClientTop       =   972
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9144
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3150
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.TextBox textCF15 
      Height          =   264
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   2
      Top             =   3150
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   4020
      TabIndex        =   56
      Top             =   3420
      Width           =   4215
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   840
         MaxLength       =   2
         TabIndex        =   6
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   8
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text12 
         Height          =   252
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   10
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "文到          天"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "        月"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   7
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                      日"
         Height          =   225
         Index           =   2
         Left            =   2520
         TabIndex        =   9
         Top             =   180
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   1200
      TabIndex        =   55
      Top             =   3420
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "文到當日"
         Height          =   180
         Index           =   0
         Left            =   144
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "文到次日"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.TextBox TextCP64_1 
      Height          =   264
      Left            =   5580
      MaxLength       =   40
      TabIndex        =   1
      Top             =   2850
      Width           =   2532
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   4968
      TabIndex        =   15
      Top             =   15
      Width           =   1200
   End
   Begin VB.TextBox textCP40 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   1950
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5580
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   450
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5580
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5580
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2550
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2550
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2532
   End
   Begin VB.TextBox textCP07 
      Height          =   264
      Left            =   5580
      MaxLength       =   7
      TabIndex        =   12
      Top             =   3960
      Width           =   2532
   End
   Begin VB.TextBox textCP06 
      Height          =   264
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   11
      Top             =   3960
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   450
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5580
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1650
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1650
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2250
      Width           =   2532
   End
   Begin VB.TextBox textCP08 
      Height          =   264
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   0
      Top             =   2850
      Width           =   2532
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   13
      Top             =   4260
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8244
      TabIndex        =   17
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6192
      TabIndex        =   14
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7020
      TabIndex        =   16
      Top             =   15
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   1152
      Left            =   1200
      TabIndex        =   58
      Top             =   4560
      Width           =   7812
      _ExtentX        =   13780
      _ExtentY        =   2032
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSForms.TextBox textCP14 
      Height          =   264
      Left            =   5580
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1950
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
      Left            =   1200
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1050
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5580
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2250
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
      Left            =   1170
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   720
      Width           =   7815
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13785;529"
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
      Left            =   120
      TabIndex        =   57
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "收文文號 :"
      Height          =   255
      Left            =   4620
      TabIndex        =   54
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   255
      Index           =   5
      Left            =   4620
      TabIndex        =   52
      Top             =   450
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "審定號數 :"
      Height          =   255
      Index           =   2
      Left            =   4620
      TabIndex        =   49
      Top             =   1350
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   255
      Index           =   7
      Left            =   4620
      TabIndex        =   48
      Top             =   2550
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   46
      Top             =   2550
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "對照名稱 :"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   44
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   43
      Top             =   1350
      Width           =   735
   End
   Begin VB.Label Label25 
      Caption         =   "法定期限 :"
      Height          =   255
      Left            =   4620
      TabIndex        =   41
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   4620
      TabIndex        =   40
      Top             =   1950
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "本所期限 :"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "下一程序 :"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   3150
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   450
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   255
      Index           =   3
      Left            =   4620
      TabIndex        =   34
      Top             =   1650
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   33
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   32
      Top             =   2250
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4620
      TabIndex        =   31
      Top             =   2250
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2850
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   4260
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   255
      Left            =   1680
      TabIndex        =   28
      Top             =   4260
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "本案期限 :"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4560
      Width           =   855
   End
End
Attribute VB_Name = "frm02010407_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/29 Form2.0已修改 grdList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
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
'Add By Cheng 2002/02/01
Dim m_CP43 As String '相關總收文號
'
Dim m_CurrSel As Integer
'Add By Cheng 2002/11/27
Dim m_CP14 As String
'Added by Morgan 2017/4/24 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/4/24
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END
Dim strLD18 As String 'Add By Sindy 2020/1/7 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2020/1/7 FC代理人
Dim m_TM23 As String 'Add By Sindy 2020/1/7 申請人


'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010407_3.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm02010407_3
   Unload frm02010407_2
   Unload frm02010407_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
        'Modify By Cheng 2002/11/07
'      'OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010407_2
      'Add By Sindy 2019/5/10
      If Me.m_strIR01 <> "" Then
        Unload frm02010407_1
        If Not m_PrevForm Is Nothing Then
           Call m_PrevForm.GoNext
        End If
        Unload Me
      '2019/5/10 END
      'Modified by Morgan 2017/4/24 電子公文
      'frm02010407_1.Show
      ElseIf m_DocNo <> "" Then
         Unload Me
         Unload frm02010407_1
         frm02010412.GoNext
      Else
         frm02010407_1.Show
         Unload Me
      End If
      'end 2017/4/24
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   
   textCP05S.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP40.BackColor = &H8000000F
   
   'Modify By Cheng 2002/02/01
'   textCF15_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010407_1.m_strIR01
   m_strIR02 = frm02010407_1.m_strIR02
   m_strIR03 = frm02010407_1.m_strIR03
   m_strIR04 = frm02010407_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/10 END
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

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 10
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "下一程序"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "本所期限"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "法定期限"
   grdList.ColWidth(3) = 1000
   grdList.col = 4
   grdList.Text = "機關文號"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "相關人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "備註"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "收文號"
   grdList.ColWidth(7) = 0
   grdList.col = 8
   grdList.Text = "下一程序代號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "序號"
   grdList.ColWidth(9) = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010407_4 = Nothing
End Sub

Private Sub grdList_Click()
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            grdList.Text = "V"
         End If
      End If
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bCP40 As Boolean
   m_TM10 = Empty
   m_CP13 = Empty
   m_CP08 = Empty
   m_CP10 = Empty
   m_CP12 = Empty
   
   m_CP36 = Empty
   m_CP37 = Empty
   m_CP38 = Empty
   m_CP39 = Empty
   m_cp40 = Empty
   m_CP41 = Empty
   m_CP42 = Empty
   'Add By Cheng 2002/02/01
   m_CP43 = Empty
    'Add By Cheng 2002/11/27
    m_CP14 = Empty
   
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   ' 來函收文日
   textCP05S = m_CP05
   ' 收文號
   textCP09 = m_CP09
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
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
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
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
      
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
   End If
   rsTmp.Close
   
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
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then
         m_CP08 = rsTmp.Fields("CP08")
      End If
      ' 案件性質
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
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
        'Add By Cheng 2002/11/27
        m_CP14 = "" & rsTmp.Fields("CP14")
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
      'Add By Cheng 2002/02/01
      ' 相關總收文號
      If IsNull(rsTmp.Fields("CP43")) = False Then
         m_CP43 = rsTmp.Fields("CP43")
      End If
   End If
   rsTmp.Close
   
   Call ChgType 'Add By Sindy 2012/4/17 讀取來函期限
   
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         ' 是否續辦欄位必須為空白
         If IsNull(rsTmp.Fields("NP06")) = False Then
            If IsEmptyText(rsTmp.Fields("NP06")) = False Then
               GoTo NextRecord
            End If
         End If
         
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         
         ' 收文號
         If IsNull(rsTmp.Fields("NP01")) = False Then
            grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("NP01")
         End If
         ' 下一程序
         If IsNull(rsTmp.Fields("NP07")) = False Then
            grdList.TextMatrix(grdList.row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"))
            grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               grdList.TextMatrix(grdList.row, 2) = ChangeWStringToTString(rsTmp.Fields("NP08"))
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(grdList.row, 3) = ChangeWStringToTString(rsTmp.Fields("NP09"))
            End If
         End If
         ' 機關文號
         If IsNull(rsTmp.Fields("NP13")) = False Then
            grdList.TextMatrix(grdList.row, 4) = rsTmp.Fields("NP13")
         End If
         ' 相關人
         If IsNull(rsTmp.Fields("NP14")) = False Then
            grdList.TextMatrix(grdList.row, 5) = rsTmp.Fields("NP14")
         End If
         ' 備註
         If IsNull(rsTmp.Fields("NP15")) = False Then
            grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("NP15")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("NP22")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/18
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/18
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing

   ' 90.11.19 modify by sonia
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   If m_TM10 < "010" Then
      If textCP08 = "" Then
         textCP08 = "（" & strTmp & "）慧商字第號"
      End If
   End If
   
   'Added by Morgan 2017/4/24 電子公文
   If m_DocWord <> "" Then
      textCP08 = m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號"
   ElseIf m_DocNo <> "" Then
      textCP08 = Replace(textCP08, "第號", "第" & PUB_GetEDocNo(m_DocNo) & "號")
   End If
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
   'end 2017/4/24
   
End Sub

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim nIndex As Integer
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP27 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strNP22 As String
   'Add By Cheng 2002/02/01
   Dim blnCheck As Boolean '是否有勾選本案期限
    Dim strCP64 As String
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   
   ' 案件性質為延期受理
   strCP10 = "1005"
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 發文日為系統日
   'strCP27 = DBDATE(Date)
   strCP27 = DBDATE(SystemDate())
   'Add By Cheng 2004/03/16
    strCP64 = ""
    If Me.TextCP64_1.Text <> "" Then strCP64 = "收文文號：" & Trim(TextCP64_1)
    'End
   ' 判斷是否有輸入本所期限及法定期限
   If IsEmptyText(textCP06) = False Then
        'Modify By Cheng 2002/11/27
        '承辦人為原程序承辦人, 不上發文日
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
'                          "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "'," & _
'                          "'" & m_CP36 & "','" & m_CP37 & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                          "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "')"
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modify By Cheng 2004/02/03
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & m_CP14 & "'," & _
'                          "'" & "N" & "','" & "N" & "','" & "N" & "'," & _
'                          "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                          "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "')"
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43, CP64) " & _
                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
                          "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & m_CP14 & "'," & _
                          "'" & "N" & "','" & "N" & "','" & "N" & "'," & _
                          "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
                          "'" & m_cp40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "','" & ChgSQL(strCP64) & "')"
        'End
      cnnConnection.Execute strSql
      '91.6.25 modify by sonia 'C'來函不必存期限
      ' 本所期限
      'If IsEmptyText(textCP06) = False Then
      '   strSQL = "UPDATE CaseProgress SET CP06 = " & DBDATE(textCP06) & " " & _
      '            "WHERE CP09 = '" & strCP09 & "' "
      '   cnnConnection.Execute strSQL
      'End If
      ' 法定期限
      'If IsEmptyText(textCP07) = False Then
      '   strSQL = "UPDATE CaseProgress SET CP07 = " & DBDATE(textCP07) & " " & _
      '            "WHERE CP09 = '" & strCP09 & "' "
      '   cnnConnection.Execute strSQL
      'End If
      '91.6.25 END
   Else
        'Modify By Cheng 2002/11/27
        '承辦人為原程序承辦人, 不上發文日
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
'                          "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "'," & _
'                          "'" & m_CP36 & "','" & m_CP37 & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                          "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "')"
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modify By Cheng 2004/02/03
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & m_CP14 & "'," & _
'                          "'" & "N" & "','" & "N" & "','" & "N" & "'," & _
'                          "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                          "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "')"
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43, CP64) " & _
                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
                          "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & m_CP14 & "'," & _
                          "'" & "N" & "','" & "N" & "','" & "N" & "'," & _
                          "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
                          "'" & m_cp40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "','" & ChgSQL(strCP64) & "')"
        'End
      cnnConnection.Execute strSql
   End If
   'Add By Sindy 2012/4/26 儲存官方發文日及官方期限月數
   If Trim(Text11) <> "" Then
     strSql = "UPDATE CaseProgress SET CP133=" & DBDATE(m_CP05) & ",CP134=" & Text11 & " " & _
              "WHERE CP09='" & strCP09 & "' "
     cnnConnection.Execute strSql
   End If
   
   'Added by Lydia 2023/06/28 商標之延期受理都請存放畫面上之本所期限及法定期限。
   If (m_TM01 = "T" Or m_TM01 = "FCT" Or m_TM01 = "CFT") And strCP10 = "1005" And (Trim(textCP06) <> "" Or Trim(textCP07) <> "") Then
      strSql = "UPDATE CaseProgress SET CP06=" & TransDate(textCP06, 2) & ",CP07=" & TransDate(textCP07, 2) & " " & _
               "WHERE CP09='" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   'end 2023/06/28
   
   'add by nickc 2008/01/10 FCT 加判斷，有期限用期限判斷(第三或第五個工作天)，無期限以第三個工作日(當日不算)，寫入承辦期限
   If m_TM01 = "FCT" Then
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
   
   'Add By Sindy 2019/12/19 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      PUB_AddLetterProgress strLD18, 1, False, "", False, m_TM23, strCP10, m_TM44
   End If
   '2019/12/19 END
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Modify By Cheng 2002/02/01
   '取消下一程序欄, 因此不新增資料至下一程序檔
'   ' 若有輸入下一程序時, 新增資料到下一程序檔
'   strNP22 = GetNextProgressNo()
'   If IsEmptyText(textCF15) = False Then
'      strNP14 = Empty
'      strNP14 = GetRelatedPerson(m_CP09)
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "'," & strNP22 & ")"
'      cnnConnection.Execute strSQL
'      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      '92.6.8 SONIA 加 言詞辯論, 準備程序
'      Select Case textCF15
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
'         Case Else:
'            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
'      End Select
'   End If
   'Add By Cheng 2002/02/01
   blnCheck = False
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新使用者所選取的本案期限資料
   For nIndex = 1 To grdList.Rows - 1
      ' 判斷該列是否有被選取
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         
         'Add By Cheng 2002/02/01
         blnCheck = True
         
         strNP07 = grdList.TextMatrix(nIndex, 8)
         strNP22 = grdList.TextMatrix(nIndex, 9)
         'Modify By Cheng 2002/02/01
         '若有勾選本案期限, 更新下一程序檔的期限
'         strSQL = "UPDATE NextProgress SET NP06 = 'Y' " & _
'                  "WHERE NP02 = '" & m_TM01 & "' AND " & _
'                        "NP03 = '" & m_TM02 & "' AND " & _
'                        "NP04 = '" & m_TM03 & "' AND " & _
'                        "NP05 = '" & m_TM04 & "' AND " & _
'                        "NP07 = " & strNP07 & " AND " & _
'                        "NP22 = " & strNP22 & " "
         strSql = " UPDATE NextProgress SET NP08 = " & DBDATE(Me.textCP06.Text) & " ,NP09 = " & DBDATE(Me.textCP07.Text) & _
                  " WHERE NP02 = '" & m_TM01 & "' AND " & _
                        "NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND " & _
                        "NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = " & strNP07 & " AND " & _
                        "NP22 = " & strNP22 & " "
         cnnConnection.Execute strSql
      End If
   Next nIndex
   'Add By Cheng 2002/02/01
   '若沒勾選本案期限, 則依本案相關總收文號去更新案件進度檔的期限
   If blnCheck = False Then
      'modify by sonia 2019/12/2 FCT-043669
      'strSql = " Update CaseProgress Set CP06 = " & DBDATE(Me.textCP06.Text) & " ,CP07 = " & DBDATE(Me.textCP07.Text) & _
               " Where CP09 = '" & m_CP43 & "'"
      'modify by sonia 2020/1/16 T-216655
      'Modified by Morgan 2025/10/15 統一改用函數
      'strSql = " Update CaseProgress Set CP06 = " & DBDATE(Me.textCP06.Text) & " ,CP07 = " & DBDATE(Me.textCP07.Text) & _
      '         " Where (CP09 = '" & m_CP43 & "' Or CP43 = '" & m_CP43 & "') AND CP158=0 AND CP159=0"
      ''end 2019/12/2
      'cnnConnection.Execute strSql
      If m_CP43 > "C" Then
         If PUB_GetCPAftExt(m_TM10, m_CP09, strExc(1)) = True Then
            strSql = " Update CaseProgress Set CP06 = " & DBDATE(Me.textCP06.Text) & " ,CP07 = " & DBDATE(Me.textCP07.Text) & " Where CP09 = '" & strExc(1) & "' and cp158=0 and cp159=0"
            cnnConnection.Execute strSql, intI
         End If
      Else
         strSql = " Update CaseProgress Set CP06 = " & DBDATE(Me.textCP06.Text) & " ,CP07 = " & DBDATE(Me.textCP07.Text) & " Where CP09 = '" & m_CP43 & "' and cp158=0 and cp159=0"
         cnnConnection.Execute strSql, intI
      End If
      'end 2025/10/15
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
   
   'Added by Morgan 2017/4/24 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
   End If
   'end 2017/4/24
   'Add by Sindy 2019/5/10
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010407_1", strCP09
   End If
   '2019/5/10 END
   
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
'Modify By Cheng 2002/02/01
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'
'   textCF15_2 = Empty
'   If IsEmptyText(textCF15) = False Then
'      ' 只取得國內的案件性質名稱
'      If m_TM10 < "010" Then
'         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
'      Else
'         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 1)
'      End If
'      If IsEmptyText(textCF15_2) = True Then
'         Cancel = True
'         strTit = "檢核資料"
'         strMsg = "案件性質代號不存在"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCF15_GotFocus
'      End If
'   End If
End Sub

Private Sub TextCP64_1_GotFocus()
    TextInverse Me.TextCP64_1
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
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
         textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2002/11/18
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
'         Else
'            strTit = "資料檢核"
'            strMsg = "來函記錄中無該筆記錄"
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
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
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
        'Modify By Cheng 2002/11/18
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
   
   'Modify By Cheng 2002/02/01
'   ' 有輸入下一程序時, 本所期限與法定期限不可為空白
'   If IsEmptyText(textCF15) = False Then
      If IsEmptyText(textCP06) = True Or IsEmptyText(textCP07) = True Then
         strTit = "檢核資料"
'         strMsg = "有下一程序時, 本所期限與法定期限不可為空白"
         strMsg = "本所期限與法定期限不可為空白"
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
'   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCF15_GotFocus()
'Modify By Cheng 2002/02/01
'   InverseTextBox textCF15
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP08_GotFocus()
   'Modify By Cheng 2002/04/22
   '將游標停在"字"的前面
'   InverseTextBox textCP08
Dim intPos As Integer
With Me.textCP08
   If Len("" & .Text) > 0 Then
      intPos = InStr("" & .Text, "字")
      If intPos - 1 >= 0 Then
         .SelStart = intPos - 1
         .SelLength = 0
      End If
   End If
End With
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'Add By Cheng 2002/11/18
Dim strTit As String
Dim strMsg As String
Dim nResponse

TxtValidate = False
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

If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
    'Add By Cheng 2002/11/18
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
       '2008/11/27 CANCEL BY SONIA
       'Else
       '   strTit = "資料檢核"
       '   strMsg = "來函記錄中無該筆記錄"
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
            'If m_DocNo = "" Or textCP07 <> "" Then 'Added by Morgan 2017/4/24 電子公文
            If m_DocNo = "" And textCP07 <> "" Then 'Added by Morgan 2017/4/24 電子公文
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
   strFromDate = DBDATE(frm02010407_1.textCP05)
   
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
   strFromDate = DBDATE(frm02010407_1.textCP05)
   
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
         strExc(0) = ""
         '若來函性質為延期受理
         If m_CP10 = "303" Then
              '若無相關總收文號
              If m_CP43 = "" Then
                  strExc(0) = "SELECT CF27,CF22,CF25 FROM CASEFEE WHERE CF01='" & m_TM01 & "' And CF02='" & m_TM10 & "' AND CF03='" & m_CP10 & "'"
              '若相關總收文號非C類
              ElseIf Left(m_CP43, 1) <> "C" Then
                  strExc(0) = "SELECT CF27,CF22,CF25 FROM CASEFEE WHERE CF01='" & m_TM01 & "' And CF02='" & m_TM10 & "' AND CF03=(Select CP10 From CaseProgress Where CP09='" & m_CP43 & "') "
              '若相關總收文號為C類
              Else
                  strExc(0) = "SELECT CF27,CF22,CF25 FROM CASEFEE WHERE CF01='" & m_TM01 & "' And CF02='" & m_TM10 & "' AND CF03=(Select NP07 From NextProgress Where NP01='" & m_CP43 & "' And " & ChgNextProgress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & ") "
              End If
         '其他
         Else
             strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & m_CP10 & "'"
         End If
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
