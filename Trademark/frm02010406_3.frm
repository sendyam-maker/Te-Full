VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010406_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標案被禁止處分"
   ClientHeight    =   4956
   ClientLeft      =   120
   ClientTop       =   996
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4956
   ScaleWidth      =   9156
   Begin VB.TextBox textRvType 
      Height          =   264
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   0
      Top             =   2460
      Width           =   372
   End
   Begin VB.TextBox textCP41 
      Height          =   264
      Left            =   1440
      MaxLength       =   600
      TabIndex        =   2
      Top             =   3060
      Width           =   7572
   End
   Begin VB.TextBox textCP14 
      Height          =   264
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   5
      Top             =   3960
      Width           =   732
   End
   Begin VB.TextBox textCP48 
      Height          =   264
      Left            =   5820
      MaxLength       =   7
      TabIndex        =   6
      Top             =   3960
      Width           =   2292
   End
   Begin VB.TextBox textCP26 
      Height          =   264
      Left            =   5820
      MaxLength       =   1
      TabIndex        =   8
      Top             =   4260
      Width           =   372
   End
   Begin VB.TextBox textCP08 
      Height          =   264
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   4
      Top             =   3660
      Width           =   2832
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6996
      TabIndex        =   11
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6168
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8220
      TabIndex        =   12
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4260
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5580
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   660
      Width           =   3432
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   660
      Width           =   2532
   End
   Begin VB.TextBox textTM06 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1260
      Width           =   7572
   End
   Begin MSForms.TextBox textCP42 
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   3360
      Width           =   7572
      VariousPropertyBits=   679493659
      MaxLength       =   600
      Size            =   "13356;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP40 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   2760
      Width           =   7572
      VariousPropertyBits=   679493659
      MaxLength       =   600
      Size            =   "13356;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_2 
      Height          =   264
      Left            =   2280
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1992
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "3514;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5580
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3432
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6054;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPS 
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   4560
      Width           =   7572
      VariousPropertyBits=   -1467989989
      MaxLength       =   128
      ScrollBars      =   2
      Size            =   "13356;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1860
      Width           =   7572
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13356;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   264
      Left            =   1440
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   960
      Width           =   7572
      VariousPropertyBits=   679493663
      Size            =   "13356;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM07 
      Height          =   264
      Left            =   1440
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1560
      Width           =   7572
      VariousPropertyBits=   679493663
      Size            =   "13356;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   42
      Top             =   3960
      Width           =   852
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   252
      Left            =   4500
      TabIndex        =   41
      Top             =   3960
      Width           =   852
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不算)"
      Height          =   252
      Left            =   6540
      TabIndex        =   40
      Top             =   4260
      Width           =   972
   End
   Begin VB.Label Label16 
      Caption         =   "是否算案件數 :"
      Height          =   252
      Left            =   4500
      TabIndex        =   39
      Top             =   4260
      Width           =   1212
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   37
      Top             =   3660
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "對照名稱(日) :"
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   36
      Top             =   3360
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "對照名稱(英) :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   35
      Top             =   3060
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "對照名稱(中) :"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   34
      Top             =   2760
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "(1.被禁止處分 2.取消被禁止處分)"
      Height          =   252
      Index           =   2
      Left            =   2040
      TabIndex        =   33
      Top             =   2460
      Width           =   3132
   End
   Begin VB.Label Label1 
      Caption         =   "來函性質 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   2460
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4500
      TabIndex        =   31
      Top             =   2160
      Width           =   972
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印)"
      Height          =   180
      Left            =   1920
      TabIndex        =   29
      Top             =   4260
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   120
      TabIndex        =   28
      Top             =   4260
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label Label21 
      Caption         =   "列印備註 :"
      Height          =   252
      Left            =   120
      TabIndex        =   27
      Top             =   4560
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   10
      Left            =   120
      TabIndex        =   26
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   252
      Index           =   7
      Left            =   4500
      TabIndex        =   25
      Top             =   660
      Width           =   852
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   24
      Top             =   1860
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   660
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件中文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   22
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label9 
      Caption         =   "案件英文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   21
      Top             =   1260
      Width           =   1212
   End
   Begin VB.Label Label10 
      Caption         =   "案件日文名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   1452
   End
End
Attribute VB_Name = "frm02010406_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0已修改 textTM05/textTM07/textTM23/textCP13/textCP14_2/textPS/textCP40/textCP42
'Memo by Amy 2021/12/29 Form2.0已修改
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
' 來源畫面
Dim strPrevForm As String
'Added by Morgan 2017/4/24 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_RegNo As String
Public m_DeadLine As String
Public m_NewCP10 As String
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
   Select Case strPrevForm
      Case "2"
         frm02010406_2.Show
         Unload Me
      Case Else
         frm02010406_1.Show
         Unload Me
         Unload frm02010404_2
   End Select
End Sub

Private Sub cmdExit_Click()
   Unload frm02010406_2
   Unload frm02010406_1
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
      Unload frm02010406_2
      'Add By Sindy 2019/5/10
      If Me.m_strIR01 <> "" Then
        Unload frm02010406_1
        If Not m_PrevForm Is Nothing Then
           Call m_PrevForm.GoNext
        End If
        Unload Me
      '2019/5/10 END
      'Modified by Morgan 2017/4/24 電子公文
      'frm02010406_1.Show
      ElseIf m_DocNo <> "" Then
         Unload Me
         Unload frm02010406_1
         frm02010412.GoNext
      Else
         frm02010406_1.Show
         Unload Me
      End If
      'end 2017/4/24
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM05.BackColor = &H8000000F
   textTM06.BackColor = &H8000000F
   textTM07.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010406_1.m_strIR01
   m_strIR02 = frm02010406_1.m_strIR02
   m_strIR03 = frm02010406_1.m_strIR03
   m_strIR04 = frm02010406_1.m_strIR04
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
      strPrevForm = Empty
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
      ' 來源畫面
      Case 5: strPrevForm = strData
   End Select
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   m_TM10 = Empty
   m_CP13 = Empty
   m_CP12 = Empty
   'Add By Cheng 2002/07/17
   m_CP08 = Empty
   m_CP10 = Empty
   
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   ' 來函收文日
   textCP05S = m_CP05
   
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
      End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05 = rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         textTM06 = rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         textTM07 = rsTmp.Fields("TM07")
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
   End If
   rsTmp.Close
   
   ' 取得案件進度檔A類資料的最後一筆
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' AND " & _
                  "CP09 LIKE 'A%' " & _
                  "ORDER BY CP05 "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveLast
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then
         m_CP08 = rsTmp.Fields("CP08")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
      End If
      '業務區   nick 91.08.22
      If IsNull(rsTmp.Fields("Cp12")) = False Then
          m_CP12 = rsTmp.Fields("cp12")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
   
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
   '來函性質
   If m_NewCP10 = "1614" Then
      textRvType = "1"
   ElseIf m_NewCP10 = "1615" Then
      textRvType = "2"
   End If
   'end 2017/4/24
   
End Sub

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP27 As String
      
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans

   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   
   ' 案件性質為被禁止處分或取消禁止處分
   strCP10 = "1614"
   Select Case textRvType
      Case "1": strCP10 = "1614"
      Case "2": strCP10 = "1615"
   End Select
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 發文日為系統日
   'strCP27 = DBDATE(Date)
   '承辦人為原程序承辦人, 不上發文日
   If IsEmptyText(textCP26) = True Then
      'StrSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP27,CP32,CP40,CP41,CP42,CP48,CP64) " & _
      '         "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
      '                 "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & strCP12 & "','" & m_CP13 & "','" & textCP14 & "'," & _
      '                 "'" & "N" & "'," & strCP27 & ",'" & "N" & "','" & chgsql(textCP40) & "','" & chgsql(textCP41) & "','" & textCP42 & "'," & DBDATE(textCP48) & ",'" & textPS & "')"
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modify By Cheng 2004/02/03
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP32,CP40,CP41,CP42,CP48,CP64) " & _
'               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                       "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
'                       "'" & "N" & "','" & "N" & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & DBDATE(textCP48) & ",'" & textPS & "')"
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP32,CP40,CP41,CP42,CP48,CP64) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                       "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
                       "'" & "N" & "','" & "N" & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & DBDATE(textCP48) & ",'" & textPS & "')"
        'End
   Else
      'StrSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP40,CP41,CP42,CP48,CP64) " & _
      '         "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
      '                 "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & strCP12 & "','" & m_CP13 & "','" & textCP14 & "'," & _
      '                 "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & chgsql(textCP40) & "','" & chgsql(textCP41) & "','" & textCP42 & "'," & DBDATE(textCP48) & ",'" & textPS & "')"
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modify By Cheng 2004/02/03
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP40,CP41,CP42,CP48,CP64) " & _
'               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                       "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
'                       "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & DBDATE(textCP48) & ",'" & textPS & "')"
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP40,CP41,CP42,CP48,CP64) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                       "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
                       "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & DBDATE(textCP48) & ",'" & textPS & "')"
        'End
   End If
   cnnConnection.Execute strSql
   
   'Add By Sindy 2019/12/19 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      PUB_AddLetterProgress strLD18, 1, False, "", False, m_TM23, strCP10, m_TM44
   End If
   '2019/12/19 END
      
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
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
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010406_1"
   End If
   '2019/5/10 END
   
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010406_3 = Nothing
End Sub

'Add By Sindy 2010/11/26
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
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

' 對造名稱(中)
Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP40, 600) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "對造名稱(中)欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP40_GotFocus
   End If
End Sub

' 對造名稱(英)
Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP41, 600) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "對造名稱(英)欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP41_GotFocus
   End If
End Sub

' 對造名稱(日)
Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP42, 600) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "對造名稱(日)欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP42_GotFocus
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
   
   ' 案件性質依還函性質欄位來區分為"被禁止處分"或"被取消禁止處分"
   strCF15 = "1614"
   If textRvType = "2" Then
      strCF15 = "1615"
   End If

   ' 承辦期限的日期應為來函收文日加上工作天數
   ' 工作天數由系統別+國家代碼+案件性質(下一程序代號)搜尋案件收費表的工作天數
   ' 若有值才做檢查
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為西元日期
      If CheckIsTaiwanDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
         GoTo EXITSUB
      End If
      
      ' 承辦期限不可超過本所期限
''''edit by nickc 2007/10/12 改抓有時效的
''''      strDay = GetWorkDays(m_TM01, m_TM10, strCF15)
''''      If IsEmptyText(strDay) = False Then
''''         ' 90.07.03 modify by louis (承辦期限以實際工作天數計算)
''''         'strDate = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         strDate = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
    strDate = Pub_GetHandleDay(m_TM01, m_TM10, strCF15, DBDATE(m_CP05))
    If IsEmptyText(strDate) = False Then
         If TAIWANDATE(textCP48) <> TAIWANDATE(strDate) Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "承辦期限日期應為<" & TAIWANDATE(strDate) & ">"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP48_GotFocus
         End If
      End If
   End If
EXITSUB:
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

' 列印備註
Private Sub textPS_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textPS, 128) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "列印備註欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textPS_GotFocus
   End If
End Sub

'來函性質
Private Sub textRvType_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textRvType) = False Then
      Select Case textRvType
         Case "1", "2":
         Case Else:
            Cancel = True
            strTit = "檢核資料"
            strMsg = "來函性質只可輸入1或2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textRvType_GotFocus
      End Select
   End If
End Sub

Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/29檢查畫面的 TextBox是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
   End If

   ' 檢核對照中英日文名稱
   If IsEmptyText(textCP40) = True And IsEmptyText(textCP41) = True And IsEmptyText(textCP42) = True Then
      strTit = "檢核資料"
      strMsg = "對照中英日文名稱不可全都空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP40.SetFocus
      GoTo EXITSUB
   End If
   ' 承辦期限不可為空白
   If IsEmptyText(textCP48) = True Then
      strTit = "檢核資料"
      strMsg = "承辦期限不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP48.SetFocus
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
   ' 承辦人不可為空白
   'If IsEmptyText(textCP14) = True Then
   '   strTit = "檢核資料"
   '   strMsg = "承辦人不可為空白"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   GoTo ExitSub
   'End If
   ' 來函性質不可為空白
   If IsEmptyText(textRvType) = True Then
      strTit = "檢核資料"
      strMsg = "來函性質不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textRvType.SetFocus
      GoTo EXITSUB
   End If

   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPS_GotFocus()
   InverseTextBox textPS
End Sub

Private Sub textRvType_GotFocus()
   InverseTextBox textRvType
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

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP40_GotFocus()
   InverseTextBox textCP40
End Sub

Private Sub textCP41_GotFocus()
   InverseTextBox textCP41
End Sub

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
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

If Me.textCP40.Enabled = True Then
   Cancel = False
   textCP40_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP41.Enabled = True Then
   Cancel = False
   textCP41_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP42.Enabled = True Then
   Cancel = False
   textCP42_Validate Cancel
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

If Me.textPS.Enabled = True Then
   Cancel = False
   textPS_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textRvType.Enabled = True Then
   Cancel = False
   textRvType_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

