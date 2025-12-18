VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03010306_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "延期受理"
   ClientHeight    =   5760
   ClientLeft      =   -3156
   ClientTop       =   3264
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9144
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   7
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5880
      TabIndex        =   5
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6840
      TabIndex        =   6
      Top             =   60
      Width           =   1212
   End
   Begin VB.TextBox textCF15 
      Height          =   285
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   1
      Top             =   3330
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3330
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.TextBox textCP06 
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3640
      Width           =   2532
   End
   Begin VB.TextBox textCP07 
      Height          =   285
      Left            =   5700
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3960
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1470
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3960
      Width           =   732
   End
   Begin VB.TextBox textCP08 
      Height          =   285
      Left            =   1200
      MaxLength       =   40
      TabIndex        =   0
      Top             =   3020
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2355
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1780
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1780
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2710
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1470
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2710
      Width           =   2532
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   1368
      Left            =   1200
      TabIndex        =   45
      Top             =   4344
      Width           =   7812
      _ExtentX        =   13780
      _ExtentY        =   2413
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
   Begin MSForms.TextBox textCP40 
      Height          =   285
      Left            =   1200
      TabIndex        =   44
      Top             =   2100
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
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   5700
      TabIndex        =   43
      Top             =   2090
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
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5700
      TabIndex        =   42
      Top             =   2400
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
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1160
      Width           =   7545
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "13309;503"
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
      TabIndex        =   40
      Top             =   850
      Width           =   7545
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13309;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
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
      TabIndex        =   39
      Top             =   570
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "本案期限 :"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   4350
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "下一程序 :"
      Height          =   252
      Left            =   120
      TabIndex        =   37
      Top             =   3348
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "本所期限 :"
      Height          =   252
      Left            =   120
      TabIndex        =   36
      Top             =   3660
      Width           =   852
   End
   Begin VB.Label Label25 
      Caption         =   "法定期限 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   35
      Top             =   3976
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   252
      Left            =   4740
      TabIndex        =   33
      Top             =   1486
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "對造名稱 :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   31
      Top             =   2100
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   4740
      TabIndex        =   30
      Top             =   556
      Width           =   852
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   120
      TabIndex        =   29
      Top             =   3976
      Width           =   972
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   252
      Left            =   2040
      TabIndex        =   28
      Top             =   3976
      Width           =   852
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   27
      Top             =   3036
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4740
      TabIndex        =   26
      Top             =   2415
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   25
      Top             =   2415
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   1788
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   3
      Left            =   4740
      TabIndex        =   23
      Top             =   1796
      Width           =   732
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   22
      Top             =   1164
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   21
      Top             =   852
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   540
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   252
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   2724
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   1476
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   252
      Index           =   7
      Left            =   4740
      TabIndex        =   17
      Top             =   2726
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人 :"
      Height          =   252
      Index           =   2
      Left            =   4740
      TabIndex        =   16
      Top             =   2106
      Width           =   852
   End
End
Attribute VB_Name = "frm03010306_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/08/13 改成Form2.0 ;textTM23、cmbTM05、textCP13、textCP14、textCP40
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 來函收文日
Dim m_CP05 As String
' 原法定期限
Dim m_CP07 As String
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
' 原業務區
Dim m_CP12 As String
' 原智權人員代號
Dim m_CP13 As String
' 國家代碼
Dim m_TM10 As String
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
Dim m_CP43 As String
'
Dim m_CurrSel As Integer
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
   frm03010306_02.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03010306_02
   Unload frm03010306_01
   Unload Me
End Sub

Private Sub cmdok_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 存檔
      'edit by nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
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
         Unload frm03010306_02
         Unload frm03010306_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
      Else
      '2023/4/27 END
         Unload Me
         Unload frm03010306_02
         frm03010306_01.Show
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
   textTM10.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP40.BackColor = &H8000000F
   
   'Modify By Cheng 2002/02/01
'   textCF15_2.BackColor = &H8000000F
  
   MoveFormToCenter Me
   
   'Add by Sindy 2021/12/8 CFT預設出通用定稿
   If m_TM01 = "CFT" Then
      textPrint = ""
   Else
      textPrint = "N"
   End If
   '2021/12/8 END
   
   'Add By Sindy 2023/4/27
   m_strIR01 = frm03010306_01.m_strIR01
   m_strIR02 = frm03010306_01.m_strIR02
   m_strIR03 = frm03010306_01.m_strIR03
   m_strIR04 = frm03010306_01.m_strIR04
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
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 審定號數
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
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("tm29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bCP40 As Boolean
   
   ' 來函收文日
   textCP05S = m_CP05
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = DBDATE(rsTmp.Fields("CP05"))
      End If
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then
         'Modify By Sindy 2012/5/31 Mark
         'textCP08 = rsTmp.Fields("CP08")
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         textCP09 = rsTmp.Fields("CP09")
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
      ' 業務區
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True)
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = GetStaffName(rsTmp.Fields("CP14"), True)
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
      Else
         m_CP43 = Empty
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 讀取檔案
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   
   ' 讀取商標基本檔
   QueryTradeMark
   ' 讀取案件進度檔
   QueryCaseProgress
   
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
   Dim strCP27 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strNP22 As String
   
   'Add By Cheng 2002/02/01
   Dim blnCheck As Boolean '判斷是否勾選本案期限
   
 '911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans
   
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質為延期受理
   strCP10 = "1005"
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   ' 發文日
   strCP27 = DBDATE(SystemDate())
   ' 新增案件進度資料
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2004/02/04
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
'                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "'," & _
'                    "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                    "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "') "
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "'," & _
                    "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
                    "'" & m_cp40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "') "
    'End
   cnnConnection.Execute strSql
   
   'Added by Lydia 2023/06/28 商標之延期受理都請存放畫面上之本所期限及法定期限。
   If (m_TM01 = "T" Or m_TM01 = "FCT" Or m_TM01 = "CFT") And strCP10 = "1005" And (Trim(textCP06) <> "" Or Trim(textCP07) <> "") Then
      strSql = "UPDATE CaseProgress SET CP06=" & TransDate(textCP06, 2) & ",CP07=" & TransDate(textCP07, 2) & " " & _
               "WHERE CP09='" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   'end 2023/06/28
   
        'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
        Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Modify By Cheng 2002/02/01
   '取消下一程序的輸入
'   ' 若有輸入下一程序時, 新增資料到下一程序檔
'   If IsEmptyText(textCF15) = False Then
'      strNP22 = GetNextProgressNo()
'      strNP14 = Empty
'      strNP14 = GetRelatedPerson(m_CP09)
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "'," & strNP22 & ")"
'      cnnConnection.Execute strSQL
'      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
'      Select Case textCF15
'         Case "102", "105", "702", "708", "305", "998", "997":
'         Case Else:
'            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
'      End Select
'   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Add By Cheng 2002/02/01
   blnCheck = False
   ' 更新使用者所選取的本案期限資料
   For nIndex = 1 To grdList.Rows - 1
      ' 判斷該列是否有被選取
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         'Add By Cheng 2002/02/01
         blnCheck = True
         strNP07 = grdList.TextMatrix(nIndex, 8)
         strNP22 = grdList.TextMatrix(nIndex, 9)
         'Modify By Cheng 2002/02/01
         '更新下一程序的期限
'         strSQL = "UPDATE NextProgress SET NP06 = 'Y' " & _
'                  "WHERE NP02 = '" & m_TM01 & "' AND " & _
'                        "NP03 = '" & m_TM02 & "' AND " & _
'                        "NP04 = '" & m_TM03 & "' AND " & _
'                        "NP05 = '" & m_TM04 & "' AND " & _
'                        "NP07 = " & strNP07 & " AND " & _
'                        "NP22 = " & strNP22 & " "
         '91.10.28 MODIFY BY SONIA 不可更新 NP06 = 'Y'
         'strSQL = " UPDATE NextProgress SET NP06 = 'Y' ,NP08 = " & DBDATE(Me.textCP06.Text) & " ,NP09 = " & DBDATE(Me.textCP07.Text) & _
         '         " WHERE NP02 = '" & m_TM01 & "' AND " & _
         '               "NP03 = '" & m_TM02 & "' AND " & _
         '               "NP04 = '" & m_TM03 & "' AND " & _
         '               "NP05 = '" & m_TM04 & "' AND " & _
         '               "NP07 = " & strNP07 & " AND " & _
         '               "NP22 = " & strNP22 & " "
         strSql = " UPDATE NextProgress SET NP08 = " & DBDATE(Me.textCP06.Text) & " ,NP09 = " & DBDATE(Me.textCP07.Text) & _
                  " WHERE NP02 = '" & m_TM01 & "' AND " & _
                        "NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND " & _
                        "NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = " & strNP07 & " AND " & _
                        "NP22 = " & strNP22 & " "
         '91.10.28 END
         cnnConnection.Execute strSql
      End If
   Next nIndex
   'Add By Cheng 2002/02/01
   '若沒勾選本案期限, 則依本案相關總收文號去更新案件進度檔的期限
   If blnCheck = False Then
      'Modified by Morgan 2025/10/15 未發文的才要更新，且要考慮受理前已收文的情形，改用函數判斷
      'strSql = " Update CaseProgress Set CP06 = " & DBDATE(Me.textCP06.Text) & " ,CP07 = " & DBDATE(Me.textCP07.Text) & _
               " Where CP09 = '" & m_CP43 & "'"
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
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm03010306_01", strCP09
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
   'Add By Cheng 2002/07/19
   Set frm03010306_03 = Nothing
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

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
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
         Exit Sub
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
         Exit Sub
      End If
   End If
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

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Modify By Cheng 2002/02/01
   '取消輸入下一程序
'   ' 下一程序
'   If IsEmptyText(textCF15) = True Then
'      If IsEmptyText(textCP06) = False Or IsEmptyText(textCP07) = False Then
'         strTit = "資料檢核"
'         strMsg = "無下一程序時, 本所期限及法定期限不可輸入"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP06.SetFocus
'         GoTo EXITSUB
'      End If
'   Else
      If IsEmptyText(textCP06) = True Or IsEmptyText(textCP07) = True Then
         strTit = "資料檢核"
         strMsg = "本所期限及法定期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
'   End If
   'Add By Cheng 2002/03/11
   If Me.textCP06.Text <> "" Then
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Me.textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
   End If

   ' 本所期限及法定期限的範圍不正確
   If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
      If Val(textCP06) > Val(textCP07) Then
         strTit = "資料檢核"
         strMsg = "本所期限的日期不可超過法定期限的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   End If
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP08_GotFocus()
   InverseTextBox textCP08
End Sub

Private Sub textCF15_GotFocus()
'Modify By Cheng 2002/02/01
'   InverseTextBox textCF15
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

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

TxtValidate = True
End Function

