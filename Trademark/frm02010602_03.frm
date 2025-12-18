VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010602_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人通知修正"
   ClientHeight    =   4836
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4836
   ScaleWidth      =   9336
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   5760
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3510
      Width           =   732
   End
   Begin VB.TextBox textCP05 
      Height          =   285
      Left            =   1500
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2541
      Width           =   1212
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8220
      TabIndex        =   11
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5940
      TabIndex        =   9
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6960
      TabIndex        =   10
      Top             =   60
      Width           =   1212
   End
   Begin VB.TextBox textCP48 
      Height          =   285
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   5
      Top             =   3168
      Width           =   1212
   End
   Begin VB.TextBox textCP14 
      Height          =   285
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   4
      Top             =   3168
      Width           =   732
   End
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   6600
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2541
      Width           =   2532
   End
   Begin VB.TextBox textCF15 
      Height          =   285
      Left            =   5760
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2541
      Width           =   732
   End
   Begin VB.TextBox textCP07 
      Height          =   285
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2862
      Width           =   1212
   End
   Begin VB.TextBox textCP06 
      Height          =   285
      Left            =   1500
      MaxLength       =   8
      TabIndex        =   2
      Top             =   2862
      Width           =   1212
   End
   Begin VB.TextBox textCP26 
      Height          =   285
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3510
      Width           =   372
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   921
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1578
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2220
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textCP27 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1899
      Width           =   2532
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   255
      Left            =   6690
      TabIndex        =   45
      Top             =   3510
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   255
      Left            =   4770
      TabIndex        =   44
      Top             =   3510
      Width           =   975
   End
   Begin MSForms.TextBox textCP64 
      Height          =   870
      Left            =   1260
      TabIndex        =   8
      Top             =   3900
      Width           =   7815
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13785;1535"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_2 
      Height          =   285
      Left            =   2100
      TabIndex        =   42
      Top             =   3168
      Width           =   1695
      VariousPropertyBits=   671105051
      Size            =   "2990;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_S 
      Height          =   285
      Left            =   1260
      TabIndex        =   43
      Top             =   1899
      Width           =   2535
      VariousPropertyBits=   671105051
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP44 
      Height          =   285
      Left            =   1260
      TabIndex        =   41
      Top             =   2220
      Width           =   2535
      VariousPropertyBits=   671105051
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
      Left            =   5760
      TabIndex        =   40
      Top             =   1590
      Width           =   2535
      VariousPropertyBits=   671105051
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1260
      TabIndex        =   39
      Top             =   1230
      Width           =   7485
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13203;503"
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
      Left            =   3840
      TabIndex        =   38
      Top             =   630
      Width           =   645
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   255
      Left            =   180
      TabIndex        =   37
      Top             =   3900
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "(N:不算)"
      Height          =   255
      Left            =   2100
      TabIndex        =   35
      Top             =   3510
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "是否算案件數 :"
      Height          =   255
      Left            =   180
      TabIndex        =   34
      Top             =   3510
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "本所期限 :"
      Height          =   252
      Left            =   180
      TabIndex        =   33
      Top             =   2878
      Width           =   852
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   180
      TabIndex        =   32
      Top             =   3184
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   10
      Left            =   180
      TabIndex        =   31
      Top             =   2557
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "承辦期限 :"
      Height          =   252
      Index           =   7
      Left            =   4800
      TabIndex        =   30
      Top             =   3184
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "法定期限 :"
      Height          =   252
      Index           =   5
      Left            =   4800
      TabIndex        =   29
      Top             =   2878
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "下一程序 :"
      Height          =   252
      Index           =   4
      Left            =   4800
      TabIndex        =   28
      Top             =   2557
      Width           =   852
   End
   Begin VB.Label Label11 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   180
      TabIndex        =   27
      Top             =   923
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   26
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   180
      TabIndex        =   25
      Top             =   1246
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   180
      TabIndex        =   24
      Top             =   1594
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   252
      Index           =   8
      Left            =   4800
      TabIndex        =   23
      Top             =   2236
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4800
      TabIndex        =   22
      Top             =   1594
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   615
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "發文日 :"
      Height          =   252
      Index           =   1
      Left            =   4800
      TabIndex        =   20
      Top             =   1915
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人 :"
      Height          =   252
      Index           =   2
      Left            =   180
      TabIndex        =   19
      Top             =   1915
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "代理人 :"
      Height          =   252
      Index           =   3
      Left            =   180
      TabIndex        =   18
      Top             =   2236
      Width           =   852
   End
End
Attribute VB_Name = "frm02010602_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/16 改成Form2.0 ;textTM23、cmbTM05、textCP13、textCP14_S、textCP14_2、textCP44、textCP64
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/9 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 申請國家
Dim m_TM10 As String
' 機關文號
Dim m_CP08 As String
' 所選取的收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 智權人員
Dim m_CP13 As String
' 業務區
Dim m_CP12 As String
' 申請人
Dim m_TM23 As String
' 代理人
Dim m_CP44 As String
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END
Dim strLD18 As String 'Add By Sindy 2020/1/7 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2020/1/7 FC代理人


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010602_02.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010602_02
   Unload frm02010602_01
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
      
      'Add By Sindy 2021/12/8
      ' 列印定稿
      If textPrint <> "N" Then
         PrintLetter
      End If
      '2021/12/8 END
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      Unload frm02010602_02
      'Add By Sindy 2019/5/22
      If Me.m_strIR01 <> "" Then
         Unload frm02010602_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
      '2019/5/22 END
      Else
         frm02010602_01.Show
      End If
      Unload Me
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
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_S.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP27.BackColor = &H8000000F
   textCP44.BackColor = &H8000000F
   textCF15_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010602_01.m_strIR01
   m_strIR02 = frm02010602_01.m_strIR02
   m_strIR03 = frm02010602_01.m_strIR03
   m_strIR04 = frm02010602_01.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
   
   'Add by Sindy 2021/12/8 CFT預設出通用定稿
   If m_TM01 = "CFT" Then
      textPrint = ""
   Else
      textPrint = "N"
   End If
   '2021/12/8 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
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
      ' 收文號
      Case 4: m_CP09 = strData
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
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      'Add By Cheng 2002/07/18
      m_TM10 = Empty
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
      'Add By Cheng 2002/07/18
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
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

' 取得服務業務基本檔
Private Sub QueryServicePractice()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      'Add By Cheng 2002/07/18
      m_TM10 = Empty
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      'Add By Cheng 2002/07/18
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
      End If
      
      'Add By Sindy 2020/1/7
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("SP26")) = False Then
         m_TM44 = rsTmp.Fields("SP26")
      End If
      '2020/1/7 END
      
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("sp15")) Then
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
   Dim bCP40 As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   'Add By Cheng 2002/07/09
   Dim strTempName As String
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 取得案件進度檔A類資料的最後一筆
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件性質
      'Add By Cheng 2002/07/18
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 = "000" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 智權人員
      'Add By Cheng 2002/07/18
      m_CP13 = Empty
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True)
      End If
      ' 業務區
      ' ADD BY SONIA 91.8.24
      m_CP12 = Empty
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = rsTmp.Fields("CP14")
         textCP14_2 = GetStaffName(rsTmp.Fields("CP14"), True)
      End If
      ' 發文日
      If IsNull(rsTmp.Fields("CP27")) = False Then
         Select Case m_TM01
            Case "CFT", "CFC", "S":
               textCP27 = DBDATE(rsTmp.Fields("CP27"))
            Case Else:
               textCP27 = TAIWANDATE(rsTmp.Fields("CP27"))
         End Select
      End If
      ' 代理人
      'Add By Cheng 2002/07/18
      m_CP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         m_CP44 = rsTmp.Fields("CP44")
         'Modify By Cheng 2002/07/09
         If PUB_GetAgentName(m_TM01, rsTmp.Fields("CP44"), strTempName) Then
            textCP44 = strTempName
         Else
            textCP44 = ""
         End If
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim strDay As String
   
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         ' 取得商標基本檔的相關項目
         QueryTradeMark
      Case Else
         QueryServicePractice
   End Select
   
   QueryCaseProgress
   
   ' 來函收文日預設為系統日
   Select Case m_TM01
      Case "CFT", "CFC", "S":
         textCP05 = DBDATE(SystemDate())
      Case Else:
         textCP05 = TAIWANDATE(SystemDate())
   End Select
   
   ' 承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''   strDay = GetWorkDays(m_TM01, m_TM10, "1702")
''''   If IsEmptyText(strDay) = False Then
      Select Case m_TM01
         Case "CFT", "CFC", "S":
            ' 90.07.03 承辦期限以工作天計算
            'textCP48 = DBDATE(DateSerial(Val(DBYEAR(textCP05)), Val(DBMONTH(textCP05)), Val(DBDAY(textCP05)) + Val(strDay)))
''''            textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(textCP05), 0))
                textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1702", DBDATE(textCP05), DBDATE(textCP06))
         Case Else:
            ' 90.07.03 承辦期限以工作天計算
            'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(textCP05)), Val(DBMONTH(textCP05)), Val(DBDAY(textCP05)) + Val(strDay)))
''''            textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(textCP05), 0))
                textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1702", DBDATE(textCP05), DBDATE(textCP06)))
      End Select
''''      If IsEmptyText(textCP06) = False Then
''''         If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''            Select Case m_TM01
''''               Case "CFT", "CFC", "S":
''''                  textCP48 = DBDATE(textCP06)
''''               Case Else:
''''                  textCP48 = TAIWANDATE(textCP06)
''''            End Select
''''         End If
''''      End If
''''   End If
   
   ' 非A類收文其預設為不可算案件數
   If Mid(m_CP09, 1, 1) <> "A" Then
      textCP26 = "N"
   End If
   
End Sub

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP27 As String
   'Dim strCP12 As String
   Dim strNP22 As String
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler

OnSaveData = True
cnnConnection.BeginTrans
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質為通知修正
   strCP10 = "1702"
   ' 業務區別 MODIFY BY SONIA 91.8.24
   'strCP12 = GetStaffDepartment(m_CP13)
   '92.6.14 ADD BY SONIA
   If m_TM01 = "CFT" Then
      '2008/12/11 MODIFY BY SONIA 有期限則不上發文日
      'strCP27 = DBDATE(SystemDate())
      If IsEmptyText(textCP06) = False Then
         strCP27 = ""
      Else
         strCP27 = DBDATE(SystemDate())
      End If
      '2008/12/11 END
   Else
      strCP27 = ""
   End If
   '92.6.14 END
   
   ' 先新增一筆案件進度記錄再更新其本所期限及法定期限
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
   '92.6.14 MODIFY BY SONIA 加發文日
   'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP64) " & _
   '               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(textCP05) & "," & _
   '                       "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
   '                       "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
   '                       "'" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
    'Modify By Cheng 2004/02/04
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(textCP05) & "," & _
'                          "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
'                          "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "'," & _
'                          "'" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
   '2008/12/11 modify by sonia 合併CP48,CP06,CP07
   'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(textCP05) & "," & _
                          "'" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
                          "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "'," & _
                          "'" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64,CP48,CP06,CP07) " & _
                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(textCP05) & "," & _
                          "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
                          "'" & "N" & "','" & textCP26 & "'," & CNULL(DBDATE(strCP27)) & ",'" & "N" & "'," & _
                          "'" & m_CP09 & "','" & ChgSQL(textCP64) & "'," & CNULL(DBDATE(textCP48)) & "," & CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & ")"
   '2008/12/11 END
    'End
   '92.6.14 END
   cnnConnection.Execute strSql
   
   'Add By Sindy 2019/12/25 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(m_TM01, 1) = "T" Then
      strLD18 = strCP09
      PUB_AddLetterProgress strLD18, 0, False, "", False, m_TM23, strCP10, m_TM44
   End If
   '2019/12/25 END
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   '2008/12/11 CANCEL BY SONIA 移至INSERT INTO CaseProgress
   '' 本所期限
   'If IsEmptyText(textCP06) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP06 = " & DBDATE(textCP06) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 法定期限
   'If IsEmptyText(textCP07) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP07 = " & DBDATE(textCP07) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 承辦期限
   'If IsEmptyText(textCP48) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '2008/12/11 END
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入下一程序時, 新增資料到下一程序檔
   If IsEmptyText(textCF15) = False Then
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
      strNP22 = GetNextProgressNo()
      'Modified by Lydia 2024/07/02 +ChgSql
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & ChgSQL(textCP64) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
'      '92.6.8 SONIA 加 言詞辯論, 準備程序
      Select Case textCF15
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         Case "102", "105", "702", "708", "305", "998", "997"
         Case Else:
            'Modify By Cheng 2003/07/09
            '改成整批列印
'            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新原案件進度檔的資料(原代理人收達日), 但若原代理人收達日有資料則不更新
   strSql = "UPDATE CaseProgress SET CP46 = " & DBDATE(textCP05) & " " & _
            "WHERE CP09 = '" & m_CP09 & "' AND " & _
                  "(CP46 IS NULL OR CP46 = 0)"
   cnnConnection.Execute strSql
   
   '2011/5/11 add by sonia 下一程序收達期限也要更新
   strSql = "UPDATE nextProgress SET NP06 = 'Y' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND NP06 IS NULL AND NP07='997' "
   cnnConnection.Execute strSql
   '2011/5/11 end
   
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
   
   'Add by Sindy 2019/5/22
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010602_01", strCP09
   End If
   '2019/5/22 END
   
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
   ' 來函收文日不可空白
   If IsEmptyText(textCP05) = True Then
      strTit = "檢核資料"
      strMsg = "來函收文日不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   ' 下一程序
   If IsEmptyText(textCF15) = True Then
      strTit = "檢核資料"
      strMsg = "下一程序不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   ' 本所期限
   If IsEmptyText(textCP06) = True Then
      strTit = "檢核資料"
      strMsg = "本所期限不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   'Add By Cheng 2002/03/11
   If Len(Me.textCP06.Text) = 8 Then
      If Val(Me.textCP06.Text) < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Me.textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
   ElseIf Len(Me.textCP06.Text) = 7 Or Len(Me.textCP06.Text) = 6 Then
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Me.textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 法定期限
   If IsEmptyText(textCP07) = True Then
      strTit = "檢核資料"
      strMsg = "法定期限不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   ' 本所期限不可超過法定期限
   If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
      If Val(textCP06) > Val(textCP07) Then
         strTit = "檢核資料"
         strMsg = "本所期限的日期不可超過法定期限的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   ' 承辦期限
   If IsEmptyText(textCP48) = True Then
      strTit = "檢核資料"
      strMsg = "承辦期限不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP48.SetFocus
      GoTo EXITSUB
   End If
'   ' 承辦期限不可超過本所期限
'   If IsEmptyText(textCP48) = False And IsEmptyText(textCP06) = False Then
'      If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
'         strTit = "檢核資料"
'         strMsg = "承辦期限不可超過本所期限"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         Me.textCP48.SetFocus
'         GoTo EXITSUB
'      End If
'   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010602_03 = Nothing
End Sub

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCF15_2 = Empty
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
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

' 來函收文日
Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDay As String
   Dim strDate As String
   Dim strTemp As String
   Cancel = False
   
   If IsEmptyText(textCP05) = False Then
      Select Case m_TM01
         Case "CFT", "CFC", "S":
            ' 來函收文日不正確
            If CheckIsDate(textCP05, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的來函收文日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP05_GotFocus
               GoTo EXITSUB
            End If
         Case Else:
            ' 來函收文日不正確
            If CheckIsTaiwanDate(textCP05, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的來函收文日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP05_GotFocus
               GoTo EXITSUB
            End If
      End Select
      
      ' 來函收文日不可超過系統日
      If Val(DBDATE(textCP05)) > Val(DBDATE(SystemDate())) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "來函收文日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If

      ' 承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''      strDay = GetWorkDays(m_TM01, m_TM10, "1702")
''''      If IsEmptyText(strDay) = False Then
         If Trim(textCP48) = "" Then 'Add By Sindy 2011/2/10 空白時才須計算承辦期限
            Select Case m_TM01
               Case "CFT", "CFC", "S":
                  ' 90.07.03 modify by louis (承辦期限以實際工作天數計算)
                  'textCP48 = DBDATE(DateSerial(Val(DBYEAR(textCP05)), Val(DBMONTH(textCP05)), Val(DBDAY(textCP05)) + Val(strDay)))
   ''''               textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(textCP05), 0))
                   textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1702", DBDATE(textCP05), DBDATE(textCP06))
               Case Else:
                  ' 90.07.03 modify by louis (承辦期限以實際工作天數計算)
                  'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(textCP05)), Val(DBMONTH(textCP05)), Val(DBDAY(textCP05)) + Val(strDay)))
   ''''               textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(textCP05), 0))
                   textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1702", DBDATE(textCP05), DBDATE(textCP06)))
            End Select
         End If
''''         If IsEmptyText(textCP06) = False Then
''''            If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''               Select Case m_TM01
''''                  Case "CFT", "CFC", "S":
''''                     textCP48 = DBDATE(textCP06)
''''                  Case Else:
''''                     textCP48 = TAIWANDATE(textCP06)
''''               End Select
''''            End If
''''         End If
''''      End If
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
      Select Case m_TM01
         Case "CFT", "CFC", "S":
            If CheckIsDate(textCP06, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的日期"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP06.SetFocus
               textCP06_GotFocus
               Exit Sub
            'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
            Else
                textCP06.Text = PUB_GetWorkDay1(textCP06, True)
            'end 2020/07/07
            End If
            'Add By Cheng 2002/03/11
            If Val(Me.textCP06.Text) < ServerDate Then
               MsgBox "本所期限不可小於系統日期!!!", vbExclamation
               Cancel = True
               textCP06.SetFocus
               textCP06_GotFocus
               Exit Sub
            End If
         
         Case Else:
            If CheckIsTaiwanDate(textCP06, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的日期"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP06.SetFocus
               textCP06_GotFocus
               Exit Sub
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
               Exit Sub
            End If
            
      End Select
      
      ' 承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''      strDay = GetWorkDays(m_TM01, m_TM10, "1702")
''''      If IsEmptyText(strDay) = False Then
         If Trim(textCP48) = "" Then 'Add By Sindy 2011/2/10 空白時才須計算承辦期限
            Select Case m_TM01
               Case "CFT", "CFC", "S":
                  ' 90.07.03 modify by louis (承辦期限以實際工作天數計算)
                  'textCP48 = DBDATE(DateSerial(Val(DBYEAR(textCP05)), Val(DBMONTH(textCP05)), Val(DBDAY(textCP05)) + Val(strDay)))
   ''''               textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(textCP05), 0))
                   textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1702", DBDATE(textCP05), DBDATE(textCP06))
               Case Else:
                  ' 90.07.03 modify by louis (承辦期限以實際工作天數計算)
                  'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(textCP05)), Val(DBMONTH(textCP05)), Val(DBDAY(textCP05)) + Val(strDay)))
   ''''               textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(textCP05), 0))
                   textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1702", DBDATE(textCP05), DBDATE(textCP06)))
            End Select
         End If
''''         If IsEmptyText(textCP06) = False Then
''''            If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''               Select Case m_TM01
''''                  Case "CFT", "CFC", "S":
''''                     textCP48 = DBDATE(textCP06)
''''                  Case Else:
''''                     textCP48 = TAIWANDATE(textCP06)
''''               End Select
''''            End If
''''         End If
''''      End If
   End If
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      Select Case m_TM01
         Case "CFT", "CFC", "S":
            If CheckIsDate(textCP07, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的日期"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP07_GotFocus
            End If
         Case Else:
            If CheckIsTaiwanDate(textCP07, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的日期"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP07_GotFocus
            End If
      End Select
   End If
End Sub

'Add By Sindy 2010/11/26
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
         Case Else
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
   Dim strDate As String
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim strCF15 As String
   Dim nResponse
   
   Cancel = False
   
   If IsEmptyText(textCP48) = False Then
      Select Case m_TM01
         Case "CFT", "CFC", "S":
            ' 檢查是否為西元日期
            If CheckIsDate(textCP48, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的承辦期限"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP48_GotFocus
               GoTo EXITSUB
            End If
         Case Else:
            ' 檢查是否為西元日期
            If CheckIsTaiwanDate(textCP48, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的承辦期限"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP48_GotFocus
               GoTo EXITSUB
            End If
      End Select
      
      ' 承辦期限不可超過本所期限
      'Modify By Sindy 2011/2/10
      If IsEmptyText(textCP06) = False Then
         If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "承辦期限不可超過本所期限"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Me.textCP48.SetFocus
            GoTo EXITSUB
         End If
      End If
'      If IsEmptyText(textCF15) = False And IsEmptyText(textCP05) = False Then
'''''edit by nickc 2007/10/11 改抓有時效性的
'''''         strDay = GetWorkDays(m_TM01, m_TM10, textCF15)
'        strDate = Pub_GetHandleDay(m_TM01, m_TM10, textCF15, DBDATE(textCP05), DBDATE(textCP06))
'''''         If IsEmptyText(strDay) = False Then
'        If IsEmptyText(strDate) = False Then
'            ' 90.07.03 modify by louis (承辦期限以實際工作天數計算)
'            'strDate = DBDATE(DateSerial(Val(DBYEAR(textCP05)), Val(DBMONTH(textCP05)), Val(DBDAY(textCP05)) + Val(strDay)))
'''''            strDate = DBDATE(CompWorkDay(Val(strDay), DBDATE(textCP05), 0))
'            If TAIWANDATE(textCP48) <> TAIWANDATE(strDate) Then
'               Cancel = True
'               strTit = "資料檢核"
'               strMsg = "承辦期限日期應為<" & TAIWANDATE(strDate) & ">"
'               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'               textCP48_GotFocus
'               GoTo EXITSUB
'            End If
'         End If
'      End If
   End If
EXITSUB:
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
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
If Me.textCF15.Enabled = True Then
   Cancel = False
   textCF15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP05.Enabled = True Then
   Cancel = False
   textCP05_Validate Cancel
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

If Me.textCP48.Enabled = True Then
   Cancel = False
   textCP48_Validate Cancel
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
