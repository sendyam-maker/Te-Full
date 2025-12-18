VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010409_7 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務結果輸入(其它)"
   ClientHeight    =   5340
   ClientLeft      =   80
   ClientTop       =   990
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9330
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   5400
      TabIndex        =   1
      Text            =   "第號(第次發佈)"
      Top             =   3720
      Width           =   2520
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   0
      Top             =   3720
      Width           =   400
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   5136
      TabIndex        =   5
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox textCP14 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3300
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2580
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2532
   End
   Begin VB.TextBox textSP06 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Width           =   6492
   End
   Begin VB.TextBox textSPKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3240
      Width           =   2532
   End
   Begin VB.TextBox textCP08 
      Height          =   264
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   2
      Top             =   4080
      Width           =   2532
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7188
      TabIndex        =   7
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6360
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8412
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textPrint 
      Height          =   270
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   3
      Top             =   4440
      Width           =   400
   End
   Begin VB.Label Label5 
      Caption         =   "公告公報 :"
      Height          =   250
      Left            =   4440
      TabIndex        =   36
      Top             =   3720
      Width           =   970
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否著作公告：            (Y/N)"
      Height          =   180
      Left            =   120
      TabIndex        =   35
      Top             =   3720
      Width           =   2270
   End
   Begin MSForms.TextBox textSP07 
      Height          =   264
      Left            =   1440
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1800
      Width           =   6492
      VariousPropertyBits=   679493663
      Size            =   "11451;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP05 
      Height          =   264
      Left            =   1440
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1080
      Width           =   6492
      VariousPropertyBits=   679493663
      Size            =   "11451;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP08 
      Height          =   264
      Left            =   1440
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2160
      Width           =   6492
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "11451;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5400
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2940
      Width           =   2532
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "13250;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   300
      Left            =   1440
      TabIndex        =   4
      Top             =   4800
      Width           =   6490
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "11451;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   4440
      TabIndex        =   34
      Top             =   3300
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   2520
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   3
      Left            =   4440
      TabIndex        =   31
      Top             =   2580
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   30
      Top             =   2880
      Width           =   972
   End
   Begin VB.Label Label10 
      Caption         =   "案件日文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   29
      Top             =   1800
      Width           =   1452
   End
   Begin VB.Label Label9 
      Caption         =   "案件英文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   28
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   27
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   720
      Width           =   852
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   25
      Top             =   2160
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   10
      Left            =   120
      TabIndex        =   24
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4440
      TabIndex        =   23
      Top             =   2940
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "機關文號 :"
      Height          =   250
      Index           =   4
      Left            =   120
      TabIndex        =   22
      Top             =   4080
      Width           =   1090
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2280
      TabIndex        =   21
      Top             =   4470
      Width           =   50
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "列印定稿 :                     (N:不印;1:台->各國;2:外->台;3:英文)"
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   4640
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   250
      Left            =   120
      TabIndex        =   19
      Top             =   4800
      Width           =   970
   End
End
Attribute VB_Name = "frm02010409_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0已修改 textSP05/textSP07/textSP08/textCP13/textCP64
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_SP01 As String
Dim m_SP02 As String
Dim m_SP03 As String
Dim m_SP04 As String
' 申請國家
Dim m_SP09 As String
' 商標審定號
Dim m_SP32 As String
' 來函收文日
Dim m_CP05 As String
' 所選取的收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
' 申請人
Dim m_TM23 As String
'Add By Cheng 2002/11/28
'原承辦人
Dim m_CP14 As String
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
   frm02010409_2.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010409_2
   Unload frm02010409_1
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
          'add by nickc 2005/04/22
          Pub_EndModCashMsg m_SP09
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
        'Modify By Cheng 2002/11/07
'      'OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
        'Add By Cheng 2002/11/08
        ' 列印定稿
        If textPrint <> "N" Then
           PrintLetter
        End If
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010409_2
      'Add By Sindy 2019/5/22
      If Me.m_strIR01 <> "" Then
         Unload frm02010409_1
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
      '2019/5/22 END
      Else
         frm02010409_1.Show
      End If
      Unload Me
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_SP01, m_SP02, m_SP03, m_SP04
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textSPKey.BackColor = &H8000000F
   textSP05.BackColor = &H8000000F
   textSP06.BackColor = &H8000000F
   textSP07.BackColor = &H8000000F
   textSP08.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010409_1.m_strIR01
   m_strIR02 = frm02010409_1.m_strIR02
   m_strIR03 = frm02010409_1.m_strIR03
   m_strIR04 = frm02010409_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
   
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_SP01 = Empty
      m_SP02 = Empty
      m_SP03 = Empty
      m_SP04 = Empty
      m_CP05 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_SP01 = strData
      ' 本所案號 欄位2
      Case 1: m_SP02 = strData
      ' 本所案號 欄位3
      Case 2: m_SP03 = strData
      ' 本所案號 欄位4
      Case 3: m_SP04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
   End Select
End Sub

' 取得服務業務基本檔
Private Sub QueryServicePractice()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   m_SP32 = Empty
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      'Add By Cheng 2002/07/17
      m_SP09 = Empty
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_SP09 = rsTmp.Fields("SP09")
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textSP05 = rsTmp.Fields("SP05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textSP06 = rsTmp.Fields("SP06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textSP07 = rsTmp.Fields("SP07")
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
         textSP08 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      
      'Add By Sindy 2019/12/20
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("SP26")) = False Then
         m_TM44 = rsTmp.Fields("SP26")
      End If
      '2019/12/20 END
      
      'add by nickc 2006/11/21
      textPrint = CheckStr(rsTmp.Fields("SP72"))
   End If

   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strCP53 As String
   Dim strCP54 As String
   m_CP13 = Empty
   m_CP12 = Empty
    'Add By Cheng 2002/11/28
    m_CP14 = Empty
   
   ' 來函收文日
   textCP05S = m_CP05
   ' 收文號
   textCP09 = m_CP09
   ' 讀取服務業務基本檔檔案
   QueryServicePractice
   
   ' 本所案號
   textSPKey = m_SP01 & m_SP02 & m_SP03 & m_SP04
   
   ' 取得案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_SP01 & "' AND " & _
                  "CP02 = '" & m_SP02 & "' AND " & _
                  "CP03 = '" & m_SP03 & "' AND " & _
                  "CP04 = '" & m_SP04 & "' AND " & _
                  "CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         textCP09 = rsTmp.Fields("CP09")
      End If
      ' 案件性質
      'Add By Cheng 2002/07/17
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_SP09 < "010" Then
            textCP10 = GetCaseTypeName(m_SP01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_SP01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True)
      End If
      '業務區   nick 91.08.22
      If IsNull(rsTmp.Fields("cp12")) = False Then
        m_CP12 = rsTmp.Fields("cp12")
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
        m_CP14 = "" & rsTmp.Fields("CP14").Value
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/21
   If textPrint = "" Then
        textPrint = GetTWordLng(m_SP01, m_SP02, m_SP03, m_SP04)
   End If
   
   'add by sonia 2023/10/6
   If m_SP01 = "TC" And m_SP09 = "000" And m_CP10 = "806" Then
      Text1.Visible = True
      Text1.Enabled = True
      Text1.Locked = False
      Text1.Text = "Y"
      Text2.Visible = True
      Text2.Enabled = True
      Text2.Locked = False
      Text2.Text = "第號(第次發佈)"
   Else
      Text1.Visible = False
      Text1.Enabled = False
      Text1.Locked = True
      Text1.Text = ""
      Text2.Visible = False
      Text2.Enabled = False
      Text2.Locked = True
      Text2.Text = ""
   End If
   'end 2023/10/6
End Sub

' 儲存資料
'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP27 As String
   Dim strSP20 As String
   Dim strSP21 As String
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質為服務業務結果
   strCP10 = "1801"
   'add by sonia 2023/10/6
   If m_SP01 = "TC" And m_SP09 = "000" Then
      If Text1 = "Y" Then
         strCP10 = "1707"
         textCP64 = "公告公報" & Text2 & "；" & textCP64
      End If
   End If
   'end 2023/10/6
   
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 發文日為系統日
   strCP27 = DBDATE(SystemDate())
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2002/11/28
    '若系統類別為查名"TS"
'    If m_SP01 = "TS" Then   '2015/1/14 CANCEL by sonia 所有服務業務結果的承辦人改放原承辦人TM-000067(宋若蘭),否則期限表帶出之承辦人會是程序
         '承辦人為原程序承辦人, 不上發文日
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modify By Cheng 2004/02/03
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'        strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
'                 "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
'                         "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
'                         "'" & "N" & "','" & "N" & "',Null,'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
        strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
                 "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                         "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
                         "'" & "N" & "','" & "N" & "',Null,'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
        'End
'2015/1/14 CANCEL by sonia 所有服務業務結果的承辦人改放原承辦人TM-000067(宋若蘭),否則期限表帶出之承辦人會是程序
'    Else
'         '承辦人為使用者, 發文日為系統日
'        'Modify By Cheng 2003/04/03
'        '智權人員存最近收文A類接洽記錄單的智權人員
'        'Modify By Cheng 2004/02/03
'        '業務區為最近收文A類接洽記錄單智權人員的業務區
''        strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
''                 "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
''                         "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
''                         "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
'        strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
'                 "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
'                         "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
'                         "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
'2015/1/14 END
        'End
'    End If
   cnnConnection.Execute strSql
   
   'Add By Sindy 2019/12/20 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", False, m_TM23, strCP10, m_TM44
   End If
   '2019/12/20 END
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_SP01, m_SP02, m_SP03, m_SP04
                       
   'add by nickc 2006/11/21
   If textPrint <> "N" Then
      strSql = "UPDATE ServicePractice SET SP72 = '" & textPrint & "' " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "' "
      cnnConnection.Execute strSql
   End If
                       
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   If strCP10 = "1801" Then   'add by sonia 2023/10/6
      ' 更新案件進度檔所選取收文資料的實際結果為 1
      strSql = "UPDATE CaseProgress SET CP24 = '1' " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   'add by sonia 2023/10/6
   Else
      'TC著作公告新進度承辦人上操作人員同時上發文日
      strSql = "UPDATE CaseProgress SET CP14='" & strUserNum & "',CP27 = " & strSrvDate(1) & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   'end 2023/10/6
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If
          'add by nickc 2005/04/22
          Pub_UpdateEndModCash m_SP01, m_SP02, m_SP03, m_SP04
   
   'Add by Sindy 2019/5/22
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010409_1"
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
   Set frm02010409_7 = Nothing
End Sub

'add by sonia 2023/10/6
Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   Select Case Text1
      Case "Y", "N":
      Case Else
         Cancel = True
         strTit = "檢核資料"
         strMsg = "著作公告來函欄位只可輸入Y或N"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
   End Select
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If Text1 = "Y" And (Text2 = "" Or Text2 = "第號(第次發佈)") Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "著作公告來函，請輸入公告公報欄！"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Text2_GotFocus
   End If
End Sub
'end 2023/10/6

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

' 檢查是否列印定稿欄位
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

' 檢查該輸入的資料是否已完成
Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/29檢查畫面的 TextBox是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP08_GotFocus()
   InverseTextBox textCP08
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
Dim strTM23Nation As String
Dim strSql As String
Dim strTmp As String
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 案件性質為刊登廣告
   If m_CP10 = "702" Then
      'add by nickc 2006/06/30
      If textPrint = "1" Then
        ' 清除定稿例外欄位檔原有資料
        EndLetter "06", m_CP09, "09", strUserNum
        ' 刊登廣告備註
        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 "VALUES ('" & "06" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
                 "'" & "刊登廣告備註" & "','" & textCP64 & "')"
        cnnConnection.Execute strSql
       End If
   End If
   
   'add by sonia 2023/10/6
   If m_SP01 = "TC" And m_SP09 = "000" Then
      If Text1 = "Y" Then
         ' 清除定稿例外欄位檔原有資料
         EndLetter "06", m_CP09, "00", strUserNum
         ' 刊登廣告備註
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "06" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                  "'" & "附件" & "','" & Text2 & "')"
         cnnConnection.Execute strSql
      End If
   End If
   'end 2023/10/6
  
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/13
   ET01 = "06"
   ET02 = m_CP09
   bolEdit = True
   '2012/1/13 End
   
   ' 案件性質為刊登廣告
   If m_CP10 = "702" Then
      'add by nickc 2006/06/30
      If textPrint = "1" Then
        ' 列印定稿
'        NowPrint m_CP09, "06", "09", True, strUserNum, 0
         ET03 = "09" 'Modify By Sindy 2012/1/13
      End If
   End If
   
   'add by sonia 2023/10/6
   If m_SP01 = "TC" And m_SP09 = "000" Then
      If Text1 = "Y" And textPrint = "1" Then
        ' 列印定稿
         ET03 = "00"
         bolEdit = False    'add by sonia 2024/3/20 整批出定稿
      End If
   End If
   'end 2023/10/6
   
   'Add By Sindy 2012/1/13
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_SP01 & m_SP02 & m_SP03 & m_SP04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_SP01) & " ]！"
         End If
      Else
         'Add By Sindy 2019/12/20 + strLD18.信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
      End If
   'Add By Sindy 2021/1/5 沒有系統產出的定稿
   Else
      'Add By Sindy 2021/2/1 詢問有沒有客戶函
      If strLD18 <> "" Then
         Call PUB_TCaseAskIsPost_C(strLD18)
      End If
   '2021/1/5 EMD
   End If
   '2012/1/13 End
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   'add by sonia 2023/10/6
   If Me.Text1.Enabled = True Then
      Cancel = False
      Text1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.Text2.Enabled = True Then
      Cancel = False
      Text2_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'end 2023/10/6
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
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

