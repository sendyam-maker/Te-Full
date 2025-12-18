VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03010307_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "其它來函輸入"
   ClientHeight    =   6336
   ClientLeft      =   -3876
   ClientTop       =   4116
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6336
   ScaleWidth      =   9156
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   5760
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3480
      Width           =   732
   End
   Begin VB.ComboBox textCP10 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   2517
      Width           =   1332
   End
   Begin VB.TextBox textCP18 
      Height          =   285
      Left            =   8010
      TabIndex        =   11
      Top             =   4140
      Width           =   945
   End
   Begin VB.TextBox textCP16 
      Height          =   285
      Left            =   5790
      TabIndex        =   10
      Top             =   4140
      Width           =   1125
   End
   Begin VB.ComboBox textCF15 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   2849
      Width           =   1332
   End
   Begin VB.TextBox textCP10_2 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2525
      Width           =   1935
   End
   Begin VB.TextBox textCP25 
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   9
      Top             =   4140
      Width           =   2172
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6840
      TabIndex        =   15
      Top             =   60
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5880
      TabIndex        =   14
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   16
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTM29 
      Height          =   285
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3815
      Width           =   732
   End
   Begin VB.TextBox textCP14 
      Height          =   285
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   5
      Top             =   3498
      Width           =   732
   End
   Begin VB.TextBox textCP48 
      Height          =   285
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   12
      Top             =   3181
      Width           =   2532
   End
   Begin VB.TextBox textCP26 
      Height          =   285
      Left            =   6120
      MaxLength       =   1
      TabIndex        =   8
      Top             =   3815
      Width           =   372
   End
   Begin VB.TextBox textCP08 
      Height          =   285
      Left            =   5760
      MaxLength       =   40
      TabIndex        =   1
      Top             =   2525
      Width           =   2532
   End
   Begin VB.TextBox textCP07 
      Height          =   285
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   4
      Top             =   2857
      Width           =   2532
   End
   Begin VB.TextBox textCP06 
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3181
      Width           =   2532
   End
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2857
      Width           =   1092
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1566
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1566
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1883
      Width           =   2532
   End
   Begin VB.TextBox textCP10S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1883
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1350
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2200
      Width           =   2385
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   1116
      Left            =   1200
      TabIndex        =   59
      Top             =   4536
      Width           =   7872
      _ExtentX        =   13885
      _ExtentY        =   1969
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
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   255
      Left            =   6600
      TabIndex        =   58
      Top             =   3495
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   255
      Left            =   4800
      TabIndex        =   57
      Top             =   3495
      Width           =   975
   End
   Begin MSForms.TextBox textCP14_2 
      Height          =   285
      Left            =   2040
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   3498
      Width           =   2295
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4048;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1200
      TabIndex        =   55
      Top             =   930
      Width           =   7605
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13414;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   1230
      Width           =   7125
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "12568;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5760
      TabIndex        =   53
      Top             =   2216
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
   Begin MSForms.TextBox textCP64 
      Height          =   612
      Left            =   1200
      TabIndex        =   13
      Top             =   5688
      Width           =   7872
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13885;1080"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
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
      Height          =   180
      Left            =   3780
      TabIndex        =   52
      Top             =   630
      Width           =   645
   End
   Begin VB.Label Label12 
      Caption         =   "點數 :"
      Height          =   255
      Left            =   7140
      TabIndex        =   51
      Top             =   4155
      Width           =   915
   End
   Begin VB.Label Label11 
      Caption         =   "費用 :"
      Height          =   255
      Left            =   4770
      TabIndex        =   50
      Top             =   4155
      Width           =   915
   End
   Begin VB.Label Label10 
      Caption         =   "專用權消滅日 :"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   4155
      Width           =   1335
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   61
      TabIndex        =   47
      Top             =   5688
      Width           =   972
   End
   Begin VB.Label Label9 
      Caption         =   "是否閉卷 :"
      Height          =   252
      Left            =   120
      TabIndex        =   46
      Top             =   3823
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "(Y:閉卷)"
      Height          =   252
      Left            =   2040
      TabIndex        =   45
      Top             =   3831
      Width           =   1452
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   44
      Top             =   3501
      Width           =   852
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   252
      Left            =   4800
      TabIndex        =   43
      Top             =   3197
      Width           =   852
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不算)"
      Height          =   252
      Left            =   6840
      TabIndex        =   42
      Top             =   3831
      Width           =   972
   End
   Begin VB.Label Label16 
      Caption         =   "是否算案件數 :"
      Height          =   252
      Left            =   4800
      TabIndex        =   41
      Top             =   3831
      Width           =   1212
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   252
      Left            =   4800
      TabIndex        =   40
      Top             =   2541
      Width           =   972
   End
   Begin VB.Label Label25 
      Caption         =   "法定期限 :"
      Height          =   252
      Left            =   4800
      TabIndex        =   39
      Top             =   2873
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "本所期限 :"
      Height          =   252
      Left            =   120
      TabIndex        =   38
      Top             =   3179
      Width           =   852
   End
   Begin VB.Label Label4 
      Caption         =   "下一程序 :"
      Height          =   252
      Left            =   120
      TabIndex        =   37
      Top             =   2857
      Width           =   852
   End
   Begin VB.Label Label7 
      Caption         =   "本案期限 :"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   4500
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "來函性質 :"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   34
      Top             =   2535
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   252
      Index           =   7
      Left            =   4800
      TabIndex        =   33
      Top             =   1582
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   1566
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   30
      Top             =   922
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   29
      Top             =   1244
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   3
      Left            =   4800
      TabIndex        =   28
      Top             =   1899
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   27
      Top             =   1888
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   26
      Top             =   2205
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4800
      TabIndex        =   25
      Top             =   2216
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   4800
      TabIndex        =   24
      Top             =   600
      Width           =   852
   End
End
Attribute VB_Name = "frm03010307_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/08/13 改成Form2.0 ;textTM23、cmbTM05、textCP13、textCP14_2、textCP64、grd1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 來函收文日
Dim m_CP05 As String
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
' 原業務區
Dim m_CP12 As String
' 原智權人員代號
Dim m_CP13 As String
' 對照名稱(中)
Dim m_cp40 As String
' 對照名稱(英)
Dim m_CP41 As String
' 對照名稱(日)
Dim m_CP42 As String
' 國家代碼
Dim m_TM10 As String
'
Dim m_CurrSel As Integer

Dim m_TM22 As String
Dim m_TM21 As String
'add by nickc 2005/05/31
Dim IsAppNpData As Boolean
Dim SeekNewCp09 As String
Dim oClsPrtForm001 As New ClsPrtForm001
'Add By Sindy 2012/3/5
Dim m_bolClose As String
Dim m_strCloseDT As String
Dim m_strCloseReason As String
'2012/3/5 End
'Add By Sindy 2023/4/27
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2023/4/27 END


'Add By Sindy 2023/4/27
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm03010307_02.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm03010307_02
   Unload frm03010307_01
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 存檔
      'edit by nick 2004/11/03
      'OnSaveData
      'add by nickc 2005/05/31
      IsAppNpData = False
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      'add by nickc 2005/05/31
      If IsAppNpData Then
         'add by nickc 2005/09/27
         If MsgBox("準備列印回覆單!!!", vbExclamation + vbOKCancel) = vbOK Then
            Call oClsPrtForm001.PrintReturnSheet(SeekNewCp09, textCF15, DBDATE(textCP07), False)
         End If
      End If
      
      'Add By Sindy 2021/12/8
      ' 列印定稿
      If textPrint <> "N" Then
         PrintLetter
      End If
      '2021/12/8 END
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Add By Sindy 2023/4/27
      If Me.m_strIR01 <> "" Then
         Unload frm03010307_02
         Unload frm03010307_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
      Else
      '2023/4/27 END
         Unload Me
         Unload frm03010307_02
         frm03010307_01.Show
      End If
   End If
End Sub

'Add By Sindy 2021/12/8
' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   If m_TM01 = "CFT" Then
      ' 清除定稿例外欄位檔原有資料(定稿別 /收文號或本所案號000 /處理狀況 /使用者編號)
      EndLetter "04", m_CP09, "00", strUserNum
   End If
End Sub

'Add By Sindy 2021/12/8
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2021/12/8 CFT 沒設定都出通用定稿(不列印)
   If m_TM01 = "CFT" Then
      NowPrint m_CP09, "04", "00", False, strUserNum, 0, , , , , , , , , , , , , , , , , True
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10S.BackColor = &H8000000F
   textCP10_2.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
      
   textCF15_2.BackColor = &H8000000F
   
   ' 下一程序
   textCF15.AddItem "補正"
   textCF15.AddItem "申請意見書"
   textCF15.AddItem "異議答辯"
   textCF15.AddItem "評定答辯"
   textCF15.AddItem "廢止答辯"
   textCF15.AddItem "補充答辯"
   textCF15.AddItem "補充理由"
   textCF15.AddItem "變更"
   textCF15.AddItem "領證"
   'Added by Lydia 2017/02/17 +特定案件性質
   textCF15.AddItem "使用宣誓"
   'textCF15.AddItem "異議"  'cancel by sonia 2019/8/13 開拓異議取消下一程序異議期限
   
   'Added by Lydia 2017/02/17 來函性質
   textCP10.AddItem "其他"
   textCP10.AddItem "通知使用宣誓"
   textCP10.AddItem "通知開拓異議"
   textCP10.AddItem "官方查名報告"    'add by sonia 2022/3/9
   'end 2017/02/17
   
   MoveFormToCenter Me
   
   'Add by Sindy 2021/12/8 CFT預設出通用定稿
   If m_TM01 = "CFT" Then
      textPrint = ""
   Else
      textPrint = "N"
   End If
   '2021/12/8 END
   
   'Add By Sindy 2023/4/27
   m_strIR01 = frm03010307_01.m_strIR01
   m_strIR02 = frm03010307_01.m_strIR02
   m_strIR03 = frm03010307_01.m_strIR03
   m_strIR04 = frm03010307_01.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2023/4/27 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      m_CP09 = Empty
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

' 讀取商標基本檔
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSub As ADODB.Recordset
   
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
      textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
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
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      
      If IsNull(rsTmp.Fields("TM21")) = False Then
        m_TM21 = rsTmp.Fields("TM21")
      End If
      If IsNull(rsTmp.Fields("TM22")) = False Then
        m_TM22 = rsTmp.Fields("TM22")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      m_bolClose = "" 'Add By Sindy 2012/3/5
      If IsNull(rsTmp.Fields("tm29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
         'Add By Sindy 2012/3/5
         textTM29 = rsTmp.Fields("TM29")
         m_bolClose = rsTmp.Fields("TM29")
         '2012/3/5 End
      End If
      
      'Add By Sindy 2012/3/5
      m_strCloseDT = ""
      If IsNull(rsTmp.Fields("TM30")) = False Then
         m_strCloseDT = rsTmp.Fields("TM30")
      End If
      m_strCloseReason = ""
      If IsNull(rsTmp.Fields("TM31")) = False Then
         m_strCloseReason = rsTmp.Fields("TM31")
      End If
      '2012/3/5 End
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得服務頁務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
      End If
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      m_bolClose = "" 'Add By Sindy 2012/3/5
      If IsNull(rsTmp.Fields("sp15")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
         'Add By Sindy 2012/3/5
         textTM29 = rsTmp.Fields("SP15")
         m_bolClose = rsTmp.Fields("SP15")
         '2012/3/5 End
      End If
      
      'Add By Sindy 2012/3/5
      m_strCloseDT = ""
      If IsNull(rsTmp.Fields("SP16")) = False Then
         m_strCloseDT = rsTmp.Fields("SP16")
      End If
      m_strCloseReason = ""
      If IsNull(rsTmp.Fields("SP17")) = False Then
         m_strCloseReason = rsTmp.Fields("SP17")
      End If
      '2012/3/5 End
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 來函收文日
   textCP05S = m_CP05
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = DBDATE(rsTmp.Fields("CP05"))
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         textCP09 = rsTmp.Fields("CP09")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10S = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10S = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True) 'Modified by Lydia 2016/03/25 離職人員也顯示
      End If
      
      '910722 Sieg
      ' 承辦人
      'Added by Lydia 2016/03/11 CFT案改成模組判斷
      'Modified by Lydia 2016/03/25 全部套用
      'If m_TM01 = "CFT" Then
        Dim strNA69 As String
        'Modified by Lydia 2016/08/17 改抓NA69
        'Call GetNP69("", "", "", strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        'Modified by Lydia 2016/11/21 傳入智權人員代號
        'Call GetNP69("", m_TM10, "", strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
        Call GetNA69("", m_TM10, "" & rsTmp.Fields("CP13"), strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
        textCP14 = strNA69
        textCP14_2 = GetStaffName(textCP14)
      'Else
      'end 2016/03/11
'        If IsNull(rsTmp.Fields("CP14")) = False Then
'           textCP14 = rsTmp.Fields("CP14")
'           textCP14_2 = GetStaffName(rsTmp.Fields("CP14"), True)
'        End If
'      End If
      'end 2016/03/25
      
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
   
   Set rsTmp = Nothing
End Sub

Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   
   ' 讀取基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
      ' 系統類別為CFC的為讀取服務業務基本檔
      Case Else:
         QueryServicePractice
   End Select
   
   ' 來函性質為專用權消滅時, 是否閉卷的預設值為Y
   If m_CP10 = "1704" Then
      textTM29 = "Y"
'2013/12/20 CANCEL BY SONIA CFT-012412 原已閉卷會被清除
'   Else
'      textTM29 = Empty
   End If
   
   ' 來函性質為專用權消滅時, 才可輸入專用權消滅日
   If m_CP10 = "1704" Then
      textCP25.BackColor = &H80000005
      textCP25.Locked = False
      textCP25.TabStop = True
   Else
      textCP25.BackColor = &H8000000F
      textCP25.Locked = True
      textCP25.TabStop = False
   End If
   
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
      'Added by Lydia 2023/10/17
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/17
   End If
   rsTmp.Close
   
   ' 讀取案件進度檔
   QueryCaseProgress
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
   
   Set rsTmp = Nothing
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim nIndex As Integer
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP14 As String
   Dim strCP27 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strNP22 As String
    'Add By Cheng 2002/12/02
    Dim strCP17 As String
   
 '911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質為來函性質
   strCP10 = textCP10
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   '92.6.14 MODIFY BY SONIA
   ' 發文日
   'If IsEmptyText(textCF15) = False Then
   '   strCP27 = "NULL"
   'Else
   '   strCP27 = DBDATE(SystemDate())
   'End If
   '2008/12/11 MODIFY BY SONIA 有期限則不上發文日
   'strCP27 = DBDATE(SystemDate())
   If IsEmptyText(textCP06) = False Then
      strCP27 = ""
   Else
      strCP27 = DBDATE(SystemDate())
   End If
   '2008/12/11 END
   '92.6.14 END
    'Add By Cheng 2002/12/02
    If Me.textCP16.Text <> "" Then
        strCP17 = Val(Me.textCP16.Text) - (Val(Me.textCP18.Text) * 1000)
    Else
        strCP17 = ""
    End If
   ' 新增案件進度資料
    'Modify By Cheng 2002/12/02
'   ' 91.03.25 modify by louis (單引號)
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "'," & CNULL(textCP14) & "," & _
'                    "'N','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
'   cnnConnection.Execute strSQL
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2004/02/04
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP17,CP18,CP20,CP26,CP27,CP32,CP43,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & CNULL(textCP14) & "," & Replace(CNULL(Me.textCP16.Text), "'", "") & "," & _
'                    Replace(CNULL(strCP17), "'", "") & "," & Replace(CNULL(Me.textCP18.Text), "'", "") & "," & "'N','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
   '2005/4/12 MODIFY BY SONIA
   'StrSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP17,CP18,CP20,CP26,CP27,CP32,CP43,CP64) " & _
   '         "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
   '                 "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & CNULL(textCP14) & "," & Replace(CNULL(Me.textCP16.Text), "'", "") & "," & _
   '                 Replace(CNULL(strCP17), "'", "") & "," & Replace(CNULL(Me.textCP18.Text), "'", "") & "," & "'N','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
   If textCP10 = "1712" Then
      '2008/12/11 modify by sonia 合併CP48,CP06,CP07
      'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP17,CP18,CP26,CP27,CP43,CP64) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                       "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & CNULL(textCP14) & "," & Replace(CNULL(Me.textCP16.Text), "'", "") & "," & _
                       Replace(CNULL(strCP17), "'", "") & "," & Replace(CNULL(Me.textCP18.Text), "'", "") & ",'" & textCP26 & "'," & strCP27 & ",'" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP17,CP18,CP26,CP27,CP43,CP64,CP48,CP06,CP07) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                       "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & CNULL(textCP14) & "," & Replace(CNULL(Me.textCP16.Text), "'", "") & "," & _
                       Replace(CNULL(strCP17), "'", "") & "," & Replace(CNULL(Me.textCP18.Text), "'", "") & ",'" & textCP26 & "'," & CNULL(DBDATE(strCP27)) & ",'" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & CNULL(DBDATE(textCP48)) & "," & CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & ") "
   Else
      '2008/12/11 modify by sonia 合併CP48,CP06,CP07
      'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP17,CP18,CP20,CP26,CP27,CP32,CP43,CP64) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                       "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & CNULL(textCP14) & "," & Replace(CNULL(Me.textCP16.Text), "'", "") & "," & _
                       Replace(CNULL(strCP17), "'", "") & "," & Replace(CNULL(Me.textCP18.Text), "'", "") & "," & "'N','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP17,CP18,CP20,CP26,CP27,CP32,CP43,CP64,CP48,CP06,CP07) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                       "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & CNULL(textCP14) & "," & Replace(CNULL(Me.textCP16.Text), "'", "") & "," & _
                       Replace(CNULL(strCP17), "'", "") & "," & Replace(CNULL(Me.textCP18.Text), "'", "") & "," & "'N','" & textCP26 & "'," & CNULL(DBDATE(strCP27)) & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & CNULL(DBDATE(textCP48)) & "," & CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & ") "
   End If
   '2005/4/12 END
    'End
   cnnConnection.Execute strSql
   
        'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
        Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   '2008/12/11 CANCEL BY SONIA 移至INSERT INTO CaseProgress
   '' 若有輸入承辦期限時
   'If IsEmptyText(textCP48) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 有輸入本所期限時
   'If IsEmptyText(textCP06) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP06 = " & DBDATE(textCP06) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 有輸入法定期限時
   'If IsEmptyText(textCP07) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP07 = " & DBDATE(textCP07) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '2008/12/11 END
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新基本檔 (是否閉卷欄位)
   ' 讀取基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為商標基本檔
      Case "CFT":
         'Modify By Sindy 2012/3/5 增加判斷及更新閉卷日期及閉卷原因
'         strSql = "UPDATE TradeMark SET TM29 = '" & textTM29 & "' " & _
'            "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                  "TM02 = '" & m_TM02 & "' AND " & _
'                  "TM03 = '" & m_TM03 & "' AND " & _
'                  "TM04 = '" & m_TM04 & "'"
         If textTM29 = "" Then
            strSql = "UPDATE TradeMark SET TM29=null,TM30=null,TM31=null " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "'"
         Else
            '原基本檔非閉卷時才更新
            If textTM29 = "Y" And m_bolClose <> "Y" Then
               strSql = "UPDATE TradeMark SET TM29='" & textTM29 & "',TM30=" & strSrvDate(1) & ",TM31='99' " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "'"
            End If
         End If
      ' 系統類別為CFC的為服務業務基本檔
      Case Else:
         'Modify By Sindy 2012/3/5 增加判斷及更新閉卷日期及閉卷原因
'         strSql = "UPDATE ServicePractice SET SP15 = '" & textTM29 & "' " & _
'            "WHERE SP01 = '" & m_TM01 & "' AND " & _
'                  "SP02 = '" & m_TM02 & "' AND " & _
'                  "SP03 = '" & m_TM03 & "' AND " & _
'                  "SP04 = '" & m_TM04 & "'"
         If textTM29 = "" Then
            strSql = "UPDATE ServicePractice SET SP15=null,SP16=null,SP17=null " & _
               "WHERE SP01 = '" & m_TM01 & "' AND " & _
                     "SP02 = '" & m_TM02 & "' AND " & _
                     "SP03 = '" & m_TM03 & "' AND " & _
                     "SP04 = '" & m_TM04 & "'"
         Else
            '原基本檔非閉卷時才更新
            If textTM29 = "Y" And m_bolClose <> "Y" Then
               strSql = "UPDATE ServicePractice SET SP15='" & textTM29 & "',SP16=" & strSrvDate(1) & ",SP17='99' " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "'"
            End If
         End If
   End Select
   ' 執行更新
   cnnConnection.Execute strSql
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若來函性質為專用權消滅時, 更新基本檔中專用權是否存在的欄位
   If m_CP10 = "1704" Then
      If m_TM01 = "CFT" Then
         strSql = "UPDATE TradeMark SET TM17 = '" & "N" & "' " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "'"
         cnnConnection.Execute strSql
      End If
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入下一程序時, 新增資料到下一程序檔
   If IsEmptyText(textCF15) = False Then
      strNP22 = GetNextProgressNo()
      strNP14 = Empty
      ' 相關人設為對造名稱
      If IsEmptyText(m_cp40) = False Then
         strNP14 = m_cp40
      ElseIf IsEmptyText(m_CP41) = False Then
         strNP14 = m_CP41
      ElseIf IsEmptyText(m_CP42) = False Then
         strNP14 = m_CP42
      End If
      'Modify By Cheng 2002/09/25
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "','" & textCP64 & "'," & strNP22 & ")"
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
      'modify by sonia 2022/2/18 ChgSQL(strNP14)-->Left(ChgSQL(strNP14), 60) (CFT-020677對造名稱太長)
      'Modified by Lydia 2024/07/02 +ChgSQL
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP08 & "','" & Left(ChgSQL(strNP14), 60) & "','" & ChgSQL(textCP64) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      Select Case textCF15
            'Modify By Cheng 2003/07/17
            '拿掉使用宣誓(105)
'         Case "102", "105", "702", "708", "305", "998", "997":
         '2013/9/24 MODIFY BY SONIA 取消702刊登廣告, 708繳年費
         'Case "102", "702", "708", "305", "998", "997":
         Case "102", "305", "998", "997":
         Case Else:
            'add by nickc 2005/05/31
            '2008/4/29 cancel by sonia 使用宣誓也要印
            'If textCF15 <> "105" Then
               IsAppNpData = True
               SeekNewCp09 = strCP09
            'End If
'            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2003/06/23
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新使用者所選取的本案期限資料
   For nIndex = 1 To grdList.Rows - 1
      ' 判斷該列是否有被選取
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         strNP07 = grdList.TextMatrix(nIndex, 8)
         strNP22 = grdList.TextMatrix(nIndex, 9)
         strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND " & _
                        "NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND " & _
                        "NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = " & strNP07 & " AND " & _
                        "NP22 = " & strNP22 & " "
         cnnConnection.Execute strSql
      End If
   Next nIndex
   
   'Add By Sindy 2009/08/17 CFT故將收達及提申期限一併上Y
   If m_TM01 = "CFT" Then
      strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
                     "WHERE NP01 = '" & m_CP09 & "' AND " & _
                           "NP02 = '" & m_TM01 & "' AND " & _
                           "NP03 = '" & m_TM02 & "' AND " & _
                           "NP04 = '" & m_TM03 & "' AND " & _
                           "NP05 = '" & m_TM04 & "' AND " & _
                           "NP07 in (997,998) AND " & _
                           "(NP06 IS NULL OR NP06 <> 'Y') "
      cnnConnection.Execute strSql
   End If
   '2009/08/17 End
   
   'Add by Sindy 2023/4/27
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm03010307_01", strCP09
   End If
   '2023/4/27 END
   
   '911106 nick transation
   cnnConnection.CommitTrans
   Exit Function
   
CheckingErr:
   MsgBox (Err.Description)
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Add By Cheng 2002/07/19
   Set frm03010307_03 = Nothing
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCF15_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCF15.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCF15_2 = Empty
   If IsEmptyText(textCF15) = False Then
      Select Case textCF15.Text
         Case "補正":
            textCF15 = "201"
         Case "申請意見書":
            textCF15 = "202"
         Case "異議答辯":
            textCF15 = "602"
         Case "評定答辯":
            textCF15 = "604"
         Case "廢止答辯":
            textCF15 = "606"
         Case "補充答辯":
            textCF15 = "613"
         Case "補充理由":
            textCF15 = "612"
         Case "變更":
            textCF15 = "301"
         Case "領證":
            textCF15 = "701"
         'Added by Lydia 2017/02/17
         Case "使用宣誓"
            textCF15 = "105"
         'cancel by sonia 2019/8/13 開拓異議取消下一程序異議期限
         'Case "異議"
         '   textCF15 = "601"
         'end 2019/8/13
         'end 2017/02/17
      End Select
   
      'Add By Cheng 2002/01/10
      If Len(Me.textCF15.Text) <> 3 Then
         Cancel = True
         MsgBox "下一程序欄位值必須為三碼!!!", vbExclamation
         textCF15_GotFocus
         Exit Sub
      End If
   
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
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      ' 檢查是否為西元日期
      If CheckIsDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = PUB_GetWorkDay1(textCP06, True)
      'end 2020/07/09
      End If
      'Add By Cheng 2002/03/11
      'Modify By Sindy 2009/09/18
      'If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
      If Val(Me.textCP06.Text) < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Cancel = True
         textCP06_GotFocus
         GoTo EXITSUB
      End If
      
      ' 以下一程序代號計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''      strDay = Empty
''''      strDay = GetWorkDays(m_TM01, m_TM10, textCP10)
''''      If IsEmptyText(strDay) = False Then
''''         ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''         'textCP48 = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''         If IsEmptyText(textCP06) = False Then
''''            If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''               textCP48 = DBDATE(textCP06)
''''            End If
''''         End If
''''      End If
        textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, textCP10, DBDATE(m_CP05), DBDATE(textCP06), textCP09)
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
      ' 檢查是否為民國年
      If CheckIsDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
      End If
   End If
End Sub
'92.6.8 ADD BY SONIA
Private Sub textCP10_LostFocus()
    '若來函性質為專用權消滅(1704), 預設為閉卷
    If Me.textCP10.Text = "1704" Then
       Me.textTM29.Text = "Y"
    End If
    
   'Added by Lydia 2017/02/17 預設下一程序
   Select Case Me.textCP10.Text
        Case "1711"
            Me.textCF15.Text = "105"
            textCF15_Validate True
        'cancel by sonia 2019/8/13 開拓異議取消下一程序異議期限
        'Case "1714"
        '    Me.textCF15.Text = "601"
        '    textCF15_Validate True
        'end 2019/8/13
   End Select
   'end 2017/02/17
End Sub
'92.6.8 END

'Added by Lydia 2017/02/17
' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP10_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP10.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

' 來函性質
Private Sub textCP10_Validate(Cancel As Boolean)
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim i As Integer
   
   Cancel = False
   
   textCP10_2 = Empty
   '若有輸入來函性質
   If IsEmptyText(textCP10) = False Then
      'Added by Lydia 2017/02/17 下拉選單
      Select Case textCP10.Text
         Case "其他":
            textCP10 = "1706"
         Case "通知使用宣誓":
            textCP10 = "1711"
         Case "通知開拓異議":
            textCP10 = "1714"
         'add by sonia 2022/3/9
         Case "官方查名報告":
            textCP10 = "1731"
         'end 2022/3/9
      End Select
      'end 2017/02/17
      
      'add by sonia 2019/8/13 開拓異議取消下一程序異議期限
      textCF15.Enabled = True
      textCF15.Locked = False
      textCP06.Enabled = True
      textCP06.Locked = False
      textCP07.Enabled = True
      textCP07.Locked = False
      If textCP10 = "1714" Then
         textCF15 = ""
         textCF15.Enabled = False
         textCF15.Locked = True
         textCP06 = ""
         textCP06.Enabled = False
         textCP06.Locked = True
         textCP07 = ""
         textCP07.Enabled = False
         textCP07.Locked = True
      End If
      'end 2019/8/13
   
     'Add By Cheng 2002/01/10
      If Len(Me.textCP10.Text) <> 4 Then
         Cancel = True
         MsgBox "來函性質欄位值必須為四碼!!!", vbExclamation
         textCP10_GotFocus
         Exit Sub
      End If
   
      ' 只取得國內的案件性質名稱
      textCP10_2 = GetCaseTypeName(m_TM01, textCP10, 0)
      If IsEmptyText(textCP10_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "來函性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP10_GotFocus
      End If
      
      '2009/10/22 ADD BY SONIA
      'modify by sonia 2019/10/9 +1602,1723,1799
      If InStr("1001,1002,1003,1004,1005,1006,1102,1403,1602,1723,1799", textCP10) > 0 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "此來函性質不可由此畫面輸入資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP10_GotFocus
      End If
      '2009/10/22 END
   
      ' 以下一程序代號計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''      strDay = Empty
''''      strDay = GetWorkDays(m_TM01, m_TM10, textCP10)
''''      If IsEmptyText(strDay) = False Then
''''        'Modify By Cheng 2003/09/02
'''''         textCP48 = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         textCP48 = DBDATE(DateAdd("d", Val(strDay), ChangeWStringToWDateString(DBDATE(m_CP05))))
''''         If IsEmptyText(textCP06) = False Then
''''            If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''               textCP48 = DBDATE(textCP06)
''''            End If
''''         End If
''''      End If
        textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, textCP10, DBDATE(m_CP05), DBDATE(textCP06), textCP09)
      'Add By Sindy 2013/6/4
      If m_TM01 = "CFT" And m_TM10 = "101" And textCP10 = "1711" Then
         For i = 1 To grdList.Rows - 1
            If grdList.TextMatrix(i, 1) = "通知使用宣誓" Then
               grdList.TextMatrix(i, 0) = "V"
               Call grdList_ShowSelection
            End If
         Next i
      End If
      '2013/6/4 End
   'Add By Cheng 2002/01/10
   '若無輸入來函性質
   Else
      Cancel = True
      MsgBox "來函性質欄位值必須輸入且為四碼!!!", vbExclamation
      textCP10_GotFocus
      Exit Sub
   End If
End Sub

'Add By Sindy 2010/11/29
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
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

Private Sub textCP16_GotFocus()
    'Add By Cheng 2002/12/02
    TextInverse Me.textCP16
End Sub

Private Sub textCP16_Validate(Cancel As Boolean)
    'Add By Cheng 2002/12/02
    If Me.textCP16.Text <> "" Then
        If IsNumeric(Me.textCP16.Text) = False Then
            MsgBox "費用項目輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            textCP16_GotFocus
        End If
    End If
End Sub

Private Sub textCP18_GotFocus()
    'Add By Cheng 2002/12/02
    TextInverse Me.textCP18
End Sub

Private Sub textCP18_Validate(Cancel As Boolean)
    'Add By Cheng 2002/12/02
    If Me.textCP18.Text <> "" Then
        If IsNumeric(Me.textCP18.Text) = False Then
            MsgBox "點數項目輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            textCP18_GotFocus
        End If
    End If
End Sub

' 專用權消滅日
Private Sub textCP25_Validate(Cancel As Boolean)
   Dim SysDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP25) = False Then
      ' 檢查是否為西元日期
      If CheckIsDate(textCP25, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的專用權消滅日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP25_GotFocus
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

' 承辦人期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為民國日期
      If CheckIsDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
         GoTo EXITSUB
      End If
   End If
   'Add By Cheng 2002/05/07
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         Cancel = True
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         textCP48_GotFocus
         GoTo EXITSUB
      End If
   End If

EXITSUB:
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

Private Sub textTM29_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否閉卷
Private Sub textTM29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textTM29) = False Then
      Select Case textTM29
         Case "", " ":
         Case "Y":
            strTit = "閉卷"
            strMsg = "請確認是否閉卷"
            nResponse = MsgBox(strMsg, vbYesNo, strTit)
            If nResponse = vbNo Then
               textTM29 = Empty
            End If
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入Y或空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM29_GotFocus
      End Select
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   ' 下一程序
   If IsEmptyText(textCF15) = True Then
      If IsEmptyText(textCP06) = False Or IsEmptyText(textCP07) = False Then
         strTit = "資料檢核"
         strMsg = "無下一程序時, 本所期限及法定期限不可輸入"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   Else
      If IsEmptyText(textCP06) = True Or IsEmptyText(textCP07) = True Then
         strTit = "資料檢核"
         strMsg = "有下一程序時, 本所期限及法定期限不可為空白"
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
         Else
            If Val(Me.textCP06.Text) > Val(Me.textCP07.Text) Then
               MsgBox "本所期限不可大於法定期限!!!", vbExclamation
               Me.textCP06.SetFocus
               textCP06_GotFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
   ' 承辦期限不可超過本所期限
   If IsEmptyText(textCP48) = False And IsEmptyText(textCP06) = False Then
      If Val(textCP48) > Val(textCP06) Then
         strTit = "資料檢核"
         strMsg = "承辦期限的日期不可超過本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 來函性質為專用權消滅時, 專用權消滅日不可為空白
   If m_CP10 = "1704" Then
      If IsEmptyText(textCP25) = True Then
         strTit = "資料檢核"
         strMsg = "請輸入專用權消滅日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP25.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

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

Private Sub textTM29_GotFocus()
   InverseTextBox textTM29
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP08_GotFocus()
   InverseTextBox textCP08
End Sub

Private Sub textCP10_GotFocus()
   InverseTextBox textCP10
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP25_GotFocus()
   InverseTextBox textCP25
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

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

If Me.textCP10.Enabled = True Then
   Cancel = False
   textCP10_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP25.Enabled = True Then
   Cancel = False
   textCP25_Validate Cancel
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

If Me.textCP48.Enabled = True Then
   Cancel = False
   textCP48_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM29.Enabled = True Then
   Cancel = False
   textTM29_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Cheng 2002/12/02
If Me.textCP16.Enabled = True Then
   Cancel = False
   textCP16_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textCP18.Enabled = True Then
   Cancel = False
   textCP18_Validate Cancel
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

'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If

TxtValidate = True
End Function

