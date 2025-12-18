VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010506_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "受理"
   ClientHeight    =   5208
   ClientLeft      =   96
   ClientTop       =   996
   ClientWidth     =   9324
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5208
   ScaleWidth      =   9324
   Begin VB.TextBox textCP08 
      Height          =   300
      Left            =   1170
      MaxLength       =   40
      TabIndex        =   37
      Top             =   2910
      Width           =   2532
   End
   Begin VB.TextBox textPrint 
      Height          =   300
      Left            =   1170
      MaxLength       =   1
      TabIndex        =   36
      Top             =   3210
      Width           =   372
   End
   Begin VB.TextBox TextCP64_1 
      Height          =   300
      Left            =   5790
      MaxLength       =   40
      TabIndex        =   35
      Top             =   2880
      Width           =   2532
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&R)"
      Height          =   400
      Left            =   4650
      TabIndex        =   0
      Top             =   30
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6930
      TabIndex        =   2
      Top             =   30
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5910
      TabIndex        =   1
      Top             =   30
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8190
      TabIndex        =   3
      Top             =   30
      Width           =   972
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5790
      Locked          =   -1  'True
      TabIndex        =   15
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5790
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5790
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2340
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5790
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   570
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   870
      Width           =   2532
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   1488
      Left            =   1176
      TabIndex        =   38
      Top             =   3552
      Width           =   8088
      _ExtentX        =   14266
      _ExtentY        =   2625
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
   Begin MSForms.ComboBox cmbTM05 
      Height          =   315
      Left            =   1170
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1140
      Width           =   7155
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "12621;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5790
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "7223;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1170
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_Src 
      Height          =   285
      Left            =   1170
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP40_S 
      Height          =   285
      Left            =   1170
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label31 
      Caption         =   "收文文號 : "
      Height          =   255
      Left            =   4830
      TabIndex        =   34
      Top             =   2910
      Width           =   915
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   255
      Left            =   90
      TabIndex        =   33
      Top             =   3210
      Width           =   975
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
      Height          =   180
      Left            =   1560
      TabIndex        =   32
      Top             =   3270
      Width           =   2745
   End
   Begin VB.Label Label9 
      Caption         =   "本案期限 :"
      Height          =   255
      Left            =   90
      TabIndex        =   31
      Top             =   3510
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   255
      Left            =   90
      TabIndex        =   30
      Top             =   2910
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   255
      Left            =   4830
      TabIndex        =   29
      Top             =   570
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4830
      TabIndex        =   28
      Top             =   1740
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   90
      TabIndex        =   27
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   255
      Index           =   9
      Left            =   4830
      TabIndex        =   26
      Top             =   2010
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   255
      Index           =   8
      Left            =   4830
      TabIndex        =   25
      Top             =   2340
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   4830
      TabIndex        =   24
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   23
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   90
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   90
      TabIndex        =   21
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   20
      Top             =   570
      Width           =   855
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   90
      TabIndex        =   19
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "對照名稱 :"
      Height          =   255
      Left            =   90
      TabIndex        =   18
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   90
      TabIndex        =   17
      Top             =   870
      Width           =   975
   End
End
Attribute VB_Name = "frm02010506_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/19 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo By Sindy 2022/2/21 Form2.0已修改
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
' 申請人
Dim m_TM23 As String
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
' 卷宗性質  93.9.10 ADD BY SONIA
Dim m_TM28 As String
Dim m_CurrSel As Integer
Dim m_CP80 As String '對造商品類別
Dim m_CP43 As String '相關總收文號 Add By Sindy 2012/4/17
Dim strRvType As String 'Add By Sindy 2012/4/26
'Added by Morgan 2017/4/26 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/4/26
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END
Dim strLD18 As String 'Add By Sindy 2019/12/20 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2019/12/20 FC代理人


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010506_3.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010506_3
   Unload frm02010506_2
   Unload frm02010506_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 列印定稿
      If textPrint <> "N" Then
         PrintLetter
      End If
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010506_3
      Unload frm02010506_2
      'Add By Sindy 2019/5/22
      If Me.m_strIR01 <> "" Then
         Unload frm02010506_1
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
         '2019/5/22 END
      'Modified by Morgan 2017/4/26 電子公文
      'frm02010506_1.Show
      ElseIf m_DocNo <> "" Then
         Unload Me
         Unload frm02010506_1
         frm02010412.GoNext
      Else
         frm02010506_1.Show
         Unload Me
      End If
      'end 2017/4/26
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
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
   textCP40_S.BackColor = &H8000000F
   
'   textCF15_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010506_1.m_strIR01
   m_strIR02 = frm02010506_1.m_strIR02
   m_strIR03 = frm02010506_1.m_strIR03
   m_strIR04 = frm02010506_1.m_strIR04
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
   '2011/6/7 ADD BY SONIA 加服務業務TD
   Select Case m_TM01
      Case "TD":
         ' 設定SQL語法
         strSql = "SELECT SP01 AS TM01,SP02 AS TM02,SP03 AS TM03,SP04 AS TM04,SP05 AS TM05,SP06 AS TM06,SP07 AS TM07,SP09 AS TM10 " & _
            ",'' AS TM12,'' AS TM15,'' AS TM28,SP08 AS TM23,SP27 AS TM45,SP72 AS TM77,SP26 AS TM44 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
      Case Else
   '2011/6/7 END
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "'"
   End Select  '2011/6/7 ADD BY SONIA
                        
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
      'Add By Cheng 2002/07/18
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      
      'Add By Sindy 2019/12/20
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("TM44")) = False Then
         m_TM44 = rsTmp.Fields("TM44")
      End If
      '2019/12/20 END
      
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then
         textTM45 = rsTmp.Fields("TM45")
      End If
      ' 卷宗性質  93.9.10 ADD BY SONIA
      If IsNull(rsTmp.Fields("TM28")) = False Then
         m_TM28 = rsTmp.Fields("TM28")
      End If
      'add by nickc 2006/11/21
      textPrint = CheckStr(rsTmp.Fields("TM77"))
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
   m_CP80 = ""
   ' 卷宗性質  93.9.10 ADD BY SONIA
   m_TM28 = Empty

          
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
      'Add By Cheng 2002/07/18
      m_CP08 = Empty
      If IsNull(rsTmp.Fields("CP08")) = False Then
         m_CP08 = rsTmp.Fields("CP08")
      End If
      ' 案件性質
      'Add By Cheng 2002/07/18
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
         textCP14_Src = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 對照名稱 (無中文取英文, 無英文取日文)
      bCP40 = False
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               textCP40_S = rsTmp.Fields("CP40")
               bCP40 = True
            End If
         End If
      End If
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               textCP40_S = rsTmp.Fields("CP41")
               bCP40 = True
            End If
         End If
      End If
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               textCP40_S = rsTmp.Fields("CP42")
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
        'Add By Cheng 2004/04/01
      ' 對造商品類別
      If IsNull(rsTmp.Fields("CP80")) = False Then
         m_CP80 = rsTmp.Fields("CP80")
      End If
        'End
      'Add By Sindy 2012/4/17 相關總收文號
      m_CP43 = Empty
      If IsNull(rsTmp.Fields("CP43")) = False Then
         m_CP43 = rsTmp.Fields("CP43")
      End If
      '2012/4/17 End
   End If
   rsTmp.Close
   
'   Call ChgType 'Add By Sindy 2012/4/17 讀取來函期限
   
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
      'Added by Lydia 2023/10/19
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/19
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
   '2011/6/7 MODIFY BY SONIA
   'If m_TM10 < "010" Then
   If m_TM10 < "010" And m_TM01 <> "TD" Then
      If textCP08 = "" Then
         textCP08 = "（" & strTmp & "）慧商字第號"
      End If
   End If
    'Marked By Cheng 2004/04/08
'    'Add By Cheng 2004/03/16
'    '預設來文字號
'    TextCP64_1 = "（" & strTmp & "）智商字第號"
'    'End

   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by  nickc 2006/11/21
   If textPrint = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   
   '2011/6/8 add by sonia
   If m_TM01 = "TD" Then
      Label31 = "主管機關案號 : "
      Label31.Left = 4480
      Label31.Width = 1200
   Else
      Label31 = "收文文號 : "
      Label31.Left = 4830
      Label31.Width = 975
   End If
   '2011/6/8 end
   
   'Added by Morgan 2017/4/26 電子公文
   If m_DocWord <> "" Then
      textCP08 = m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號"
   ElseIf m_DocNo <> "" Then
      textCP08 = Replace(textCP08, "第號", "第" & PUB_GetEDocNo(m_DocNo) & "號")
   End If
   'end 2017/4/26
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
   Set frm02010506_4 = Nothing
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

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim strSql As String
   Dim bUpdate As Boolean
   Dim strCP06 As String
   Dim strCP07 As String
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP27 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strNP22 As String
    Dim strCP64 As String
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   
   ' 案件性質為爭議受理
   strCP10 = "1607"
   
   'Add By Cheng 2002/01/03
   '若來函性質屬於爭議程序(16XX), 應更新商標基本檔是否有爭議程序欄(TM19)為"Y"
   If Left(strCP10, 2) = "16" Then
      strSql = "UPDATE TradeMark SET TM19='Y'" & _
               " WHERE TM01 = '" & m_TM01 & "'" & _
               " And TM02 = '" & m_TM02 & "'" & _
               " And TM03 = '" & m_TM03 & "'" & _
               " And TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   End If
   
   'add by nickc 2006/11/21
   If textPrint <> "N" Then
      strSql = "UPDATE TradeMark SET TM77='" & textPrint & "'" & _
               " WHERE TM01 = '" & m_TM01 & "'" & _
               " And TM02 = '" & m_TM02 & "'" & _
               " And TM03 = '" & m_TM03 & "'" & _
               " And TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   End If
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 發文日為系統日
   strCP27 = DBDATE(SystemDate())
   
   strCP06 = Empty
   strCP07 = Empty
'   If IsEmptyText(textCP06) = False Then: strCP06 = DBDATE(textCP06)
'   If IsEmptyText(textCP07) = False Then: strCP07 = DBDATE(textCP07)
   
   'Add By Cheng 2004/03/16
    strCP64 = ""
    '2011/6/8 modify by sonia
    'If Me.TextCP64_1.Text <> "" Then strCP64 = "收文文號：" & Trim(TextCP64_1)
    If Me.TextCP64_1.Text <> "" Then strCP64 = Label31 & Trim(TextCP64_1)
    '2011/6/8 end
    'End
   
   ' 先新增一筆案件進度記錄再更新其本所期限及法定期限
   '承辦人為使用者, 發文日為系統日
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2004/02/03
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & strUserNum & "'," & _
'                          "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "'," & _
'                          "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                          "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "')"
    '業務區為最近收文A類接洽記錄單智權人員的業務區
    'Modify By Cheng 2004/04/01
    '加存欄位對造商品類別
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43, CP64) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & strUserNum & "'," & _
'                          "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "'," & _
'                          "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
'                          "'" & m_CP40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "','" & ChgSQL(strCP64) & "')"
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43, CP64, CP80) " & _
                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                          "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & strUserNum & "'," & _
                          "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "'," & _
                          "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
                          "'" & m_cp40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "','" & ChgSQL(strCP64) & "', '" & m_CP80 & "')"
    'End
    'End
   cnnConnection.Execute strSql
   
   'Add By Sindy 2019/12/20 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(m_TM01, 1) = "T" Then
      strLD18 = strCP09
      strExc(1) = Pub_GetSpecMan("內商程序客戶函發後補看人員")
      PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), , False, m_TM23, strCP10, m_TM44, , , , , strExc(1)
   End If
   '2019/12/20 END
   
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
   
'   'Add By Sindy 2012/4/26 儲存官方發文日及官方期限月數
'   If Trim(Text11) <> "" Then
'      strSql = "UPDATE CaseProgress SET CP133=" & DBDATE(m_CP05) & ",CP134=" & Text11 & " " & _
'               "WHERE CP09='" & strCP09 & "' "
'      cnnConnection.Execute strSql
'   End If
   
   'add by nickc 2008/01/10 FCT 加判斷，有期限用期限判斷(第三或第五個工作天)，無期限以第三個工作日(當日不算)，寫入承辦期限
   If m_TM01 = "FCT" Then
'        If Trim(textCP07) = "" Then
            strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(4, DBDATE(m_CP05), 0)) & " " & _
                     "WHERE CP09 = '" & strCP09 & "' "
            cnnConnection.Execute strSql
'        Else
'            If DateDiff("d", ChangeWStringToWDateString(DBDATE(m_CP05)), ChangeWStringToWDateString(DBDATE(textCP07))) <= 30 Then    '無法與上句合併，因為沒有日期時，datediff  會發生  型態不符 的錯誤
'                strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(4, DBDATE(m_CP05), 0)) & " " & _
'                         "WHERE CP09 = '" & strCP09 & "' "
'                cnnConnection.Execute strSql
'            Else
'                strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(6, DBDATE(m_CP05), 0)) & " " & _
'                         "WHERE CP09 = '" & strCP09 & "' "
'                cnnConnection.Execute strSql
'            End If
'        End If
    End If
   
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ' 若有輸入下一程序時, 新增資料到下一程序檔
'   strNP22 = GetNextProgressNo()
'   If IsEmptyText(textCF15) = False Then
'      strNP14 = Empty
'      strNP14 = GetRelatedPerson(m_CP09)
'      'Modify By Cheng 2002/09/25
''      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
''                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
''                          strCP06 & "," & strCP07 & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "'," & strNP22 & ")"
'        'Modify By Cheng 2003/04/03
'        '智權人員存最近收文A類接洽記錄單的智權人員
'      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          strCP06 & "," & strCP07 & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "'," & strNP22 & ")"
'      cnnConnection.Execute strSql
'      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
''      '92.6.8 SONIA 加 言詞辯論, 準備程序
'      Select Case textCF15
''         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
'         Case "102", "105", "702", "708", "305", "998", "997"
'         Case Else:
'            ' 列印國內案件接洽及結案記錄單
''            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
'            'Add By Cheng 2004/04/08
'            '新增列印接洽結案單資料
'            pub_AddressListSN = pub_AddressListSN + 1
'            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
'      End Select
'   End If
   
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
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If
   
   'Added by Morgan 2017/4/26 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
   End If
   'end 2017/4/26
   
   'Add by Sindy 2019/5/22
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010506_1"
   End If
   '2019/5/22 END
   
   Set rsTmp = Nothing
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

'' 下一程序
'Private Sub textCF15_Validate(Cancel As Boolean)
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
'End Sub

'' 本所期限
'Private Sub textCP06_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'
'   Cancel = False
'   If IsEmptyText(textCP06) = False Then
'      If CheckIsTaiwanDate(textCP06, False) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "請輸入正確的日期"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP06_GotFocus
'         GoTo EXITSUB
'      End If
'      'Add By Cheng 2002/03/11
'      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
'         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
'         Cancel = True
'         textCP06_GotFocus
'         GoTo EXITSUB
'      End If
'        'Modify By Cheng 2002/11/19
'        '按下確定才檢查
''      ' 申請國家為台灣時需檢查來函記錄檔
''      If m_TM10 < "010" Then
''         strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR16")
''         If IsEmptyText(strDate) = False Then
''            If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
''               strTit = "資料檢核"
''               strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
''               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
''               If nResponse = vbCancel Then
''                  Cancel = True
''                  textCP06_GotFocus
''                  GoTo EXITSUB
''               End If
''            End If
''         Else
''            strTit = "資料檢核"
''            strMsg = "來函記錄中無該筆記錄"
''            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
''            If nResponse = vbCancel Then
''               Cancel = True
''               textCP06_GotFocus
''               GoTo EXITSUB
''            End If
''         End If
''      End If
'   End If
'EXITSUB:
'End Sub

'' 法定期限
'Private Sub textCP07_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'
'   Cancel = False
'   If IsEmptyText(textCP07) = False Then
'      If CheckIsTaiwanDate(textCP07, False) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "請輸入正確的日期"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP07_GotFocus
'         GoTo EXITSUB
'      End If
'        'Modify By Cheng 2002/11/19
'        '按下確定才檢查
''      ' 申請國家為台灣時需檢查來函記錄檔
''      If m_TM10 < "010" Then
''         strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR17")
''         If IsEmptyText(strDate) = False Then
''            If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
''               strTit = "資料檢核"
''               strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
''               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
''               If nResponse = vbCancel Then
''                  Cancel = True
''                  textCP07_GotFocus
''                  GoTo EXITSUB
''               End If
''            End If
''         Else
''            strTit = "資料檢核"
''            strMsg = "來函記錄中無該筆記錄"
''            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
''            If nResponse = vbCancel Then
''               Cancel = True
''               textCP07_GotFocus
''               GoTo EXITSUB
''            End If
''         End If
''      End If
'   End If
'EXITSUB:
'End Sub

Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   If m_TM10 < "010" Then
      ' 申請國家為台灣時, 機關文號不可為空白
      '2011/6/8 modify by sonia
      'If IsEmptyText(textCP08) = True Then
      If IsEmptyText(textCP08) = True And m_TM01 <> "TD" Then
         strTit = "檢核資料"
         strMsg = "申請國家為台灣時, 機關文號不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP08.SetFocus
         GoTo EXITSUB
      End If
      
      ' 申請國家為台灣時, 下一程序不可為空白
      '90.08.16 modify by sonia
      'If IsEmptyText(textCF15) = True Then
      '   strTit = "檢核資料"
      '   strMsg = "申請國家為台灣時, 下一程序不可為空白"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCF15.SetFocus
      '   GoTo EXITSUB
      'End If
   End If
   
'   'Add By Sindy 2012/4/17
'   '檢查來函期限--日期
'   If m_TM10 = 台灣國家代號 Then
'      If Me.Option4(2).Value = True Then
'         If Me.Text12.Text = "" Then
'            MsgBox "請輸入來函期限!!!", vbExclamation + vbOKOnly
'            Me.Text12.SetFocus
'            GoTo EXITSUB
'         End If
'      End If
'   End If
   
'   ' 有輸入下一程序時, 本所期限與法定期限不可為空白
'   If IsEmptyText(textCF15) = False Then
'      If IsEmptyText(textCP06) = True Or IsEmptyText(textCP07) = True Then
'         strTit = "檢核資料"
'         strMsg = "有下一程序時, 本所期限與法定期限不可為空白"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP06.SetFocus
'         GoTo EXITSUB
'      End If
'      'Add By Cheng 2002/03/11
'      If Me.textCP06.Text <> "" Then
'         If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
'            MsgBox "本所期限不可小於系統日期!!!", vbExclamation
'            Me.textCP06.SetFocus
'            textCP06_GotFocus
'            GoTo EXITSUB
'         End If
'      End If
'      ' 本所期限必須小於法定期限
'      If Val(textCP06) > Val(textCP07) Then
'         strTit = "檢核資料"
'         strMsg = "本所期限的日期不可超過法定期限的日期"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP06.SetFocus
'         GoTo EXITSUB
'      End If
'   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub TextCP64_1_GotFocus()
    TextInverse Me.TextCP64_1
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         'edit by nickc 2006/06/29
         'Case " ", "N":
         Case "N", "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            'edit by nickc 2006/06/29
            'strMsg = "只可輸入空白或N"
            strMsg = "只可輸入 N 或 1-3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

'Private Sub textCF15_GotFocus()
'   InverseTextBox textCF15
'End Sub
'
'Private Sub textCP06_GotFocus()
'   InverseTextBox textCP06
'End Sub
'
'Private Sub textCP07_GotFocus()
'   InverseTextBox textCP07
'End Sub

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

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField(ET03 As String)
   Dim strTM23Nation As String
   Dim strSql As String
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 申請國家為台灣
   If m_TM10 < "010" Then
      ' 申請人國籍
      'edit by nickc 2006/06/30
      'If strTM23Nation < "010" Then
      If textPrint = "1" Then
         ' 清除定稿例外欄位檔原有資料
         EndLetter "07", m_CP09, ET03, strUserNum
         '2011/6/8 add by sonia TD定稿
         If m_TM01 = "TD" Then
            ' 機關文號
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "07" & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "'," & _
                     "'" & "機關文號" & "','" & Label31 & Trim(TextCP64_1) & "')"
         Else
            ' 機關文號
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "07" & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "'," & _
                     "'" & "機關文號" & "','" & textCP08 & "')"
         End If
         cnnConnection.Execute strSql
      'edit by nickc 2006/06/30
      'Else
      ElseIf textPrint = "2" Then
         ' 清除定稿例外欄位檔原有資料
         EndLetter "07", m_CP09, ET03, strUserNum
         ' 機關文號
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "07" & "','" & m_CP09 & "','" & ET03 & "','" & strUserNum & "'," & _
                  "'" & "機關文號" & "','" & textCP08 & "')"
         cnnConnection.Execute strSql
      End If
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
Dim ET03 As String
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   'add by nickc 2006/06/30
   ET03 = ""
   
   'Add By Sindy 2012/1/13
   ET01 = "07"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/13 End
   
   ' 申請國家為台灣
   If m_TM10 < "010" Then
      ' 申請人國籍為台灣
      'edit by nickc 2006/06/30
      'If strTM23Nation < "010" Then
      If textPrint = "1" Then
         ' 列印定稿
         '92.12.9 modify by sonia
         'NowPrint m_CP09, "07", "01", False, strUserNum, 0
         '93.9.10 MODIFY SONIA 改以卷宗性質判斷
         'Select Case m_CP10
         '   'Modify By Cheng 2003/12/25
         '   '補充答辯(613)
'        '    Case "602", "604", "606"
         '   'Modify By Cheng 2004/03/11
         '   '加補充理由(612)
'        '    Case "602", "604", "606", "613"
         '   Case "602", "604", "606", "612", "613"
         '      ET03 = "03"
         '   Case Else
         '      ET03 = "01"
         'End Select
         Select Case m_TM28
            Case "1"
               ET03 = "03"
            Case Else
               ET03 = "01"
         End Select
         '93.9.10 END
         '92.12.9 end
      '申請人國籍非台灣
      'edit by nickc 2006/06/30
      'Else
      ElseIf textPrint = "2" Then
         ' 列印定稿
         '93.9.10 MODIFY SONIA 改以卷宗性質判斷
         'ET03 = "02"
         Select Case m_TM28
            Case "1"
               ET03 = "04"
            Case Else
               ET03 = "02"
         End Select
         '93.9.10 END
      End If
   End If
   
   'add by nickc 2006/06/30
   If ET03 <> "" Then
       
       ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
       InsExpField (ET03)
       
'       NowPrint m_CP09, "07", ET03, False, strUserNum, 0
      'Add By Sindy 2012/1/13
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(m_TM01, 1) = "T" Then
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
         End If
      Else
         'Add By Sindy 2019/12/20 + strLD18.信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
      End If
      '2012/1/13 End
   'Add By Sindy 2021/1/5 沒有系統產出的定稿
   Else
      'Add By Sindy 2021/2/1 詢問有沒有客戶函
      If strLD18 <> "" Then
         Call PUB_TCaseAskIsPost_C(strLD18)
      End If
   '2021/1/5 EMD
   End If
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
'If Me.textCF15.Enabled = True Then
'   Cancel = False
'   textCF15_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If
'
'If Me.textCP06.Enabled = True Then
'   Cancel = False
'   textCP06_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If
'
'If Me.textCP07.Enabled = True Then
'   Cancel = False
'   textCP07_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If

If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

''Add By Cheng 2002/11/19
'' 申請國家為台灣時需檢查來函記錄檔
'If m_TM10 < "010" Then
'   strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR16")
'   If IsEmptyText(strDate) = False Then
'      If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
'         strTit = "資料檢核"
'         strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
'         nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'         If nResponse = vbCancel Then
'            Cancel = True
'            textCP06_GotFocus
'            Exit Function
'         End If
'      End If
'   '2008/11/27 CANCEL BY SONIA
'   'Else
'   '   strTit = "資料檢核"
'   '   strMsg = "來函記錄中無該筆記錄"
'   '   nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'   '   If nResponse = vbCancel Then
'   '      Cancel = True
'   '      textCP06_GotFocus
'   '     Exit Function
'   '   End If
'   '2008/11/27 END
'   '2011/6/15 ADD BY SONIA
'   Else
'     If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then
'     Else
'        If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
'           strTit = "資料檢核"
'           strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
'           nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'           If nResponse = vbCancel Then
'              Cancel = True
'              textCP06_GotFocus
'              Exit Function
'           End If
'        End If
'     End If
'     '2011/6/15 END
'   End If
'End If
'' 申請國家為台灣時需檢查來函記錄檔
'If m_TM10 < "010" Then
'   strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR17")
'   If IsEmptyText(strDate) = False Then
'      If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
'         strTit = "資料檢核"
'         strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
'         nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'         If nResponse = vbCancel Then
'            Cancel = True
'            textCP07_GotFocus
'            Exit Function
'         End If
'      End If
'   Else
'     If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then  '2011/6/15 ADD BY SONIA
'        strTit = "資料檢核"
'        strMsg = "來函記錄中無該筆記錄"
'        nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'        If nResponse = vbCancel Then
'           Cancel = True
'           textCP07_GotFocus
'           Exit Function
'        End If
'     '2011/6/15 ADD BY SONIA
'     Else
'        If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
'           strTit = "資料檢核"
'           strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
'           nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'           If nResponse = vbCancel Then
'              Cancel = True
'              textCP07_GotFocus
'              Exit Function
'           End If
'        End If
'     End If
'     '2011/6/15 END
'   End If
'End If

TxtValidate = True
End Function

''Add By Sindy 2012/4/17
'Private Sub Option1_Click(Index As Integer)
'   If Me.Option4(0).Value Then
'      Text10_Validate False
'   ElseIf Me.Option4(1).Value Then
'      Text11_Validate False
'   ElseIf Me.Option4(2).Value Then
'      Text12_Validate False
'   End If
'End Sub
'
'Private Sub Text10_GotFocus()
'   TextInverse Text10
'   CloseIme
'End Sub
'
'Private Sub Text10_LostFocus()
'   '非台灣"天"跳離時到"本所期限"欄位
'   If m_TM10 <> 台灣國家代號 Then
'      If textCP06.Enabled = True Then textCP06.SetFocus
'   End If
'End Sub
'
'Private Sub Text10_Validate(Cancel As Boolean)
'   If Text10 <> "" Then GetTime
'End Sub
'
'Private Sub Text11_GotFocus()
'   TextInverse Text11
'   CloseIme
'End Sub
'
'Private Sub Text11_LostFocus()
'   '非台灣"月"跳離時到"本所期限"欄位
'   'If m_TM10 <> 台灣國家代號 Then
'   '   If textCP06.Enabled = True Then textCP06.SetFocus
'   'End If
'End Sub
'
'Private Sub Text11_Validate(Cancel As Boolean)
'   If Text11 <> "" Then GetTime
'End Sub
'
'Private Sub Text12_GotFocus()
'   TextInverse Text12
'End Sub
'
'Private Sub Text12_LostFocus()
'   '非台灣"日"跳離時到"本所期限"欄位
'   If m_TM10 <> 台灣國家代號 Then
'      If textCP06.Enabled = True Then textCP06.SetFocus
'   End If
'End Sub
'
'Private Sub Text12_Validate(Cancel As Boolean)
'   If Option4(2).Value = False Then Exit Sub
'   If Text12 = "" Then
'   Else
'      If ChkDate(Text12) Then
'         If m_TM10 = 台灣國家代號 Then
'            If Val(Text12) < Val(strSrvDate(2)) Then
'               MsgBox "來函期限不可小於系統日 !", vbCritical
'               Cancel = True
'            Else
'               textCP07 = Text12
'               textCP06 = TransDate(CompDate(2, -2, TransDate(textCP07, 2)), 1)
'               '本所期限若非工作天則抓最近工作天
''               Me.textCP06.Text = TransDate(PUB_GetWorkDay1(Me.textCP06.Text, True), 1)
'            End If
'         End If
'      Else
'         Cancel = True
'      End If
'   End If
'   If Cancel = True Then TextInverse Text12
'End Sub
'
'Private Sub GetTime()
'   Dim i As Integer
'   Dim strFromDate As String '期限起算日
'
'   'strFromDate = DBDATE(textCP05)
'   strFromDate = DBDATE(frm02010506_1.textCP05)
'
'   If m_TM10 = 台灣國家代號 Then
'      '文到天數
'      If Option4(0).Value = True Then
'         textCP07 = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
'         If Option1(0).Value = True Then textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
'         If Val(Text10) >= 60 Then
'            i = -4
'         Else
'            i = -2
'         End If
'      '文到月數
'      ElseIf Option4(1).Value = True Then
'         textCP07 = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
'         If Option1(0).Value = True Then textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
'         If Val(Text11) >= 2 Then
'            i = -4
'         Else
'            i = -2
'         End If
'      End If
'      If textCP07 <> "" Then textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
'      '本所期限若非工作天則抓最近工作天
''      Me.textCP06.Text = TransDate(PUB_GetWorkDay1(Me.textCP06.Text, True), 1)
'   End If
'End Sub
'
''讀取來函期限
'Private Function ChgType() As Boolean
'Dim strTempName As String, BolTmp As Boolean
'Dim i As Integer
'Dim strFromDate As String '期限起算日
'
'   'strFromDate = DBDATE(textCP05)
'   strFromDate = DBDATE(frm02010506_1.textCP05)
'
'   ChgType = False
'   If m_TM10 = 台灣國家代號 Then
'      BolTmp = False
'   Else
'      BolTmp = True
'   End If
'
'   strRvType = ""
'   '相關總收文號的案件性質
'   strSql = "SELECT * FROM caseprogress WHERE cp09='" & m_CP43 & "' "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      strRvType = "" & RsTemp.Fields("CP10")
'   End If
'   If strRvType = "" Then Exit Function
'
'   If ClsPDGetCaseProperty(m_TM01, strRvType, strTempName, BolTmp) Then
'      textCP06 = ""
'      textCP07 = ""
'
'      If m_TM10 = 台灣國家代號 Then
'         strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & strRvType & "'"
'         If strExc(0) <> "" Then
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            With RsTemp
'               If intI = 1 Then
'                  If Not IsNull(.Fields(1)) Then
'                     '文到天數
'                     Option4(0).Value = True
'                     Text10 = .Fields(1)
'                     textCP07 = TransDate(CompDate(2, Text10, TransDate(strFromDate, 2)), 1)
'                  ElseIf Not IsNull(.Fields(2)) Then
'                     '文到月數
'                     Option4(1).Value = True
'                     Text11 = .Fields(2)
'                     textCP07 = TransDate(CompDate(1, .Fields(2), TransDate(strFromDate, 2)), 1)
'                  Else
'                     '文到天數
'                     Option4(0).Value = True
'                     Text10 = ""
'                     Text11 = ""
'                  End If
'                  If textCP07 <> "" And Not IsNull(.Fields(0)) Then
'                     '文到當日
'                     If .Fields(0) = "1" Then
'                        Option1(0).Value = True
'                        textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
'                     '文到次日
'                     Else
'                        Option1(1).Value = True
'                     End If
'                  End If
'                  '文到天數
'                  If Text10 <> "" Then
'                     If Val(Text10) >= 60 Then
'                        i = -4
'                     Else
'                        i = -2
'                     End If
'                  '文到月數
'                  ElseIf Not IsNull(.Fields(2)) Then
'                     If Val(.Fields(2)) >= 2 Then
'                        i = -4
'                     Else
'                        i = -2
'                     End If
'                  End If
'                  If textCP07 <> "" Then textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
'                  '本所期限若非工作天則抓最近工作天
''                  Me.textCP06.Text = TransDate(PUB_GetWorkDay1(Me.textCP06.Text, True), 1)
'               End If
'            End With
'         End If
'      End If
'      ChgType = True
'   End If
'End Function
