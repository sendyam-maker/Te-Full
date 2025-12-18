VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010301_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標申請案號輸入"
   ClientHeight    =   5736
   ClientLeft      =   192
   ClientTop       =   996
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9144
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   4140
      TabIndex        =   52
      Top             =   4590
      Width           =   4215
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   16
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   840
         MaxLength       =   2
         TabIndex        =   14
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text12 
         Height          =   252
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   18
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                      日"
         Height          =   225
         Index           =   2
         Left            =   2520
         TabIndex        =   17
         Top             =   180
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "        月"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   15
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "文到          天"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   1260
      TabIndex        =   51
      Top             =   4590
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "文到次日"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   12
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "文到當日"
         Height          =   180
         Index           =   0
         Left            =   144
         TabIndex        =   11
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox textCP64 
      Height          =   264
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   7
      Top             =   3600
      Width           =   2292
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "商品及服務資料查詢(&I)"
      Height          =   400
      Index           =   6
      Left            =   4230
      TabIndex        =   19
      Top             =   70
      Width           =   1935
   End
   Begin VB.TextBox textNP08 
      Height          =   264
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   9
      Top             =   4290
      Width           =   2292
   End
   Begin VB.TextBox textNP09 
      Height          =   264
      Left            =   5970
      MaxLength       =   7
      TabIndex        =   10
      Top             =   4290
      Width           =   2292
   End
   Begin VB.TextBox textCF09 
      Height          =   264
      Left            =   5112
      MaxLength       =   12
      TabIndex        =   5
      Top             =   3240
      Width           =   804
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2292
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   720
      Width           =   2292
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3240
      Width           =   372
   End
   Begin VB.TextBox textCP05 
      Height          =   264
      Left            =   5760
      MaxLength       =   7
      TabIndex        =   3
      Top             =   2880
      Width           =   2292
   End
   Begin VB.TextBox textCP45 
      Height          =   264
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   6
      Top             =   3600
      Width           =   2292
   End
   Begin VB.TextBox textCP30 
      Height          =   264
      Left            =   5760
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2520
      Width           =   2292
   End
   Begin VB.TextBox textTM27 
      Height          =   264
      Left            =   1380
      MaxLength       =   20
      TabIndex        =   2
      Top             =   2880
      Width           =   2292
   End
   Begin VB.TextBox textCP47 
      Height          =   264
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2520
      Width           =   2292
   End
   Begin VB.TextBox textTM32 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2160
      Width           =   7572
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1800
      Width           =   7572
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2292
   End
   Begin VB.TextBox textTM01 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   720
      Width           =   2292
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7020
      TabIndex        =   21
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8244
      TabIndex        =   22
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6192
      TabIndex        =   20
      Top             =   70
      Width           =   800
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5760
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2292
      VariousPropertyBits=   679493663
      Size            =   "4043;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPS 
      Height          =   300
      Left            =   1380
      TabIndex        =   8
      Top             =   3930
      Width           =   7575
      VariousPropertyBits=   -1467989989
      MaxLength       =   128
      Size            =   "13356;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   264
      Left            =   1380
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2292
      VariousPropertyBits=   679493663
      Size            =   "4043;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label32 
      Caption         =   "來函期限:"
      Height          =   255
      Left            =   180
      TabIndex        =   54
      Top             =   4770
      Width           =   855
   End
   Begin VB.Label LabNP07 
      Height          =   255
      Left            =   8400
      TabIndex        =   53
      Top             =   4740
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "大陸案受理函發文日 :                                                     (西元)"
      Height          =   255
      Index           =   19
      Left            =   3990
      TabIndex        =   50
      Top             =   3600
      Width           =   4785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "子案新本所期限 :"
      Height          =   180
      Index           =   18
      Left            =   180
      TabIndex        =   49
      Top             =   4320
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "子案新法定期限 :"
      Height          =   180
      Index           =   17
      Left            =   4560
      TabIndex        =   48
      Top             =   4320
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "大約                        可接獲審定公告"
      Height          =   180
      Index           =   16
      Left            =   4560
      TabIndex        =   47
      Top             =   3270
      Width           =   2700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
      Height          =   180
      Index           =   15
      Left            =   1800
      TabIndex        =   43
      Top             =   3270
      Width           =   2745
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   14
      Left            =   4560
      TabIndex        =   37
      Top             =   2880
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   13
      Left            =   4560
      TabIndex        =   36
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "列印備註 :"
      Height          =   252
      Index           =   12
      Left            =   180
      TabIndex        =   35
      Top             =   3960
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   11
      Left            =   180
      TabIndex        =   34
      Top             =   3600
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "列印定稿 :"
      Height          =   252
      Index           =   10
      Left            =   180
      TabIndex        =   33
      Top             =   3240
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "正商標號數 :"
      Height          =   252
      Index           =   9
      Left            =   180
      TabIndex        =   32
      Top             =   2880
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "申請日 :"
      Height          =   252
      Index           =   8
      Left            =   180
      TabIndex        =   31
      Top             =   2520
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   7
      Left            =   4560
      TabIndex        =   30
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   4560
      TabIndex        =   29
      Top             =   1080
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   252
      Index           =   5
      Left            =   4560
      TabIndex        =   28
      Top             =   720
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "商品群組 :"
      Height          =   252
      Index           =   4
      Left            =   180
      TabIndex        =   27
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   252
      Index           =   3
      Left            =   180
      TabIndex        =   26
      Top             =   1800
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   252
      Index           =   2
      Left            =   180
      TabIndex        =   25
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱 :"
      Height          =   252
      Index           =   1
      Left            =   180
      TabIndex        =   24
      Top             =   1080
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   23
      Top             =   720
      Width           =   1212
   End
End
Attribute VB_Name = "frm02010301_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/28 Form2.0已修改 textTM05/textCP13/textPS
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 商標種類
Dim m_TM08 As String
' 總收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 彼所案號
Dim m_CP45 As String
' 申請國家
Dim m_TM10 As String
' 智權人員代號
Dim m_CP13 As String
Dim m_CP12 As String
Dim m_CP14 As String 'Add By Sindy 2010/10/25
Dim m_CP27 As String 'add by sonia 2018/2/6
' 申請人
Dim m_TM23 As String
'Add By Cheng 2003/12/25
Dim m_strLanguage As String '定稿語文
Dim m_blnPriDate As Boolean '是否有優先權
Dim m_TM67 As String '放棄專用權
'End
'add by nickc 2006/07/24
Public UpForm As Form
Dim m_MonTM01 As String     '紀錄分割母案案號
Dim m_MonTM02 As String
Dim m_MonTM03 As String
Dim m_MonTM04 As String
Public m_MonCP09 As String  '傳入分割母案收文號
Dim m_MonNP08 As String
Dim m_MonNP09 As String
Public oStrCDate As String
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'add by nick 2004/10/05 檢查是否已經有商品及服務
Public ChkTG As Boolean
Dim strRvType As String 'Add By Sindy 2012/5/18
'Added by Morgan 2017/6/14 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_DocPdf As String
Public m_DocPdfDate As String
Public m_DocPdfTime As String
'end 2017/6/14
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END
Dim strLD18 As String 'Add By Sindy 2019/11/20 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2019/11/22 FC代理人


'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub Form_Load()
   ' 設定只顯示不輸入的控制項其背景顏色
   textTM01.BackColor = &H8000000F
   textTM05.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM32.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010301_1.m_strIR01
   m_strIR02 = frm02010301_1.m_strIR02
   m_strIR03 = frm02010301_1.m_strIR03
   m_strIR04 = frm02010301_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/10 END
End Sub

Public Sub SetData(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String, strCP09 As String)
   m_TM01 = strTM01
   m_TM02 = strTM02
   m_TM03 = strTM03
   m_TM04 = strTM04
   m_CP09 = strCP09
End Sub

' 取得員工姓名
Private Function GetStaffName(ByVal strKey As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetStaffName = Empty
   strSql = "SELECT * FROM Staff " & _
            "WHERE ST01 = '" & strKey & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("ST02")) = False Then
         GetStaffName = rsTmp.Fields("ST02")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 取得商標種類名稱
' Input : strKey ==> 商標代碼
'         nCountry ==> 0 表取國內名稱
'                      1 表取得大陸名稱
Private Function GetTradeName(ByVal strKey As String, ByVal nCountry As Integer) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetTradeName = Empty
   strSql = "SELECT * FROM PatentTradeMarkMap " & _
            "WHERE PTM01 = 'T' AND " & _
                  "PTM02 = '" & strKey & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Select Case nCountry
         Case 0:
            If IsNull(rsTmp.Fields("PTM03")) = False Then
               GetTradeName = rsTmp.Fields("PTM03")
            End If
         Case Else:
            If IsNull(rsTmp.Fields("PTM04")) = False Then
               GetTradeName = rsTmp.Fields("PTM04")
            End If
      End Select
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 由案件性質代碼取得案件性質名稱
Private Function GetCaseType(ByVal strKey1 As String, ByVal StrKey2 As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCaseType = Empty
   If IsEmptyText(strKey1) = False And IsEmptyText(StrKey2) = False Then
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM CasePropertyMap " & _
               "WHERE CPM01 = '" & strKey1 & "' AND " & _
                     "CPM02 = '" & StrKey2 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CPM03")) = False Then
            GetCaseType = rsTmp.Fields("CPM03")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 由客戶代碼取得客戶名稱
Private Function GetCustomer(ByVal strData As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCustomer = Empty
   If IsEmptyText(strData) = False Then
      Set rsTmp = New ADODB.Recordset
      If Len(strData) > 8 Then
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                        "CU02 = '" & Mid(strData, 9, 1) & "'"
      Else
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CU04")) = False Then
            GetCustomer = rsTmp.Fields("CU04")
         ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
            GetCustomer = rsTmp.Fields("CU05")
         ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
            GetCustomer = rsTmp.Fields("CU06")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

' 取得國家的名稱
Private Function GetNation(ByVal strNation As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetNation = Empty
   If IsEmptyText(strNation) = False Then
      strSql = "SELECT * FROM NATION " & _
               "WHERE NA01 = '" & strNation & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("NA03")) = False Then
            GetNation = rsTmp.Fields("NA03")
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

Public Function UpdateCtrl()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
'Add By Cheng 2003/03/03
Dim strTM23Nation As String '申請人國籍
Dim Cancel As Boolean 'Add By Sindy 2009/09/03
   
   m_TM10 = Empty
    m_TM67 = ""
    'add by nickc 2006/08/03
    If UpForm Is Nothing Then
        textCP05 = TAIWANDATE(SystemDate())
    Else
        textCP05 = TAIWANDATE(UpForm.oStrCDate)
    End If

   ' 本所案號
   textTM01 = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
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
      ' 智權人員
      'Add By Cheng 2002/07/17
      m_CP13 = Empty
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
         m_CP13 = rsTmp.Fields("CP13")
      End If
      m_CP12 = Empty
      '業務區   91.0822 nick   用原來的業務區
      If IsNull(rsTmp.Fields("cp12")) = False Then
         m_CP12 = rsTmp.Fields("cp12")
      End If
      'Add By Sindy 2010/10/25 承辦人
      m_CP14 = Empty
      If IsNull(rsTmp.Fields("CP14")) = False Then
         m_CP14 = rsTmp.Fields("CP14")
      End If
      ' 案件性質
      'Add By Cheng 2002/07/17
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
      End If
      'add by sonia 2018/2/6
      ' 發文日
      m_CP27 = 0
      If IsNull(rsTmp.Fields("CP27")) = False Then
         m_CP27 = rsTmp.Fields("CP27")
      End If
      'end 2018/2/6
      
      'edit by nick 2004/12/23 加入分割和申請相同
      'If m_CP10 <> "806" And m_CP10 <> "101" Then
      If m_CP10 <> "806" And m_CP10 <> "101" And m_CP10 <> "308" Then
         ' 申請日
         If IsNull(rsTmp.Fields("CP47")) = False Then
            textCP47 = TAIWANDATE(rsTmp.Fields("CP47"))
         End If
      End If
      
      ' 申請案號
      If IsNull(rsTmp.Fields("CP30")) = False Then
         textCP30 = rsTmp.Fields("CP30")
      End If
      
      ' 彼所案號
      'Add By Cheng 2002/07/17
      m_CP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         textCP45 = rsTmp.Fields("CP45")
         m_CP45 = rsTmp.Fields("CP45")
      End If
      
   End If
   rsTmp.Close
   
   Select Case m_TM01
      Case "T", "TF", "FCT", "CFT":
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            ' 商標名稱(中)
            If IsNull(rsTmp.Fields("TM05")) = False Then
               textTM05 = rsTmp.Fields("TM05")
            ElseIf IsNull(rsTmp.Fields("TM06")) = False Then
               textTM05 = rsTmp.Fields("TM06")
            ElseIf IsNull(rsTmp.Fields("TM07")) = False Then
               textTM05 = rsTmp.Fields("TM07")
            End If
            ' 申請國家
            If IsNull(rsTmp.Fields("TM10")) = False Then
               textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
               m_TM10 = rsTmp.Fields("TM10")
            End If
            ' 商標種類
            'Add By Cheng 2002/07/17
            m_TM08 = Empty
            If IsNull(rsTmp.Fields("TM08")) = False Then
               m_TM08 = rsTmp.Fields("TM08")
               If m_TM10 < "010" Then
                  textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
               Else
                  textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
               End If
            End If
            ' 商品類別
            If IsNull(rsTmp.Fields("TM09")) = False Then
               textTM09 = rsTmp.Fields("TM09")
            End If
            ' 申請人
            m_TM23 = Empty
            If IsNull(rsTmp.Fields("TM23")) = False Then
               m_TM23 = rsTmp.Fields("TM23")
            End If
            
            'Add By Sindy 2019/11/22
            ' FC代理人
            m_TM44 = Empty
            If IsNull(rsTmp.Fields("TM44")) = False Then
               m_TM44 = rsTmp.Fields("TM44")
            End If
            '2019/11/22 END
            
            ' 正商標號數
            If IsNull(rsTmp.Fields("TM27")) = False Then
               textTM27 = rsTmp.Fields("TM27")
            End If
            ' 商品群組
            If IsNull(rsTmp.Fields("TM32")) = False Then
               textTM32 = rsTmp.Fields("TM32")
            End If
            ' 若案件性質為申請時帶出申請日及申請案號
            'edit by nick 2004/12/23 加入分割等於申請
            'If m_CP10 = "101" Then
            If m_CP10 = "101" Or m_CP10 = "308" Then
               ' 申請日
               If IsNull(rsTmp.Fields("TM11")) = False Then
                  textCP47 = TAIWANDATE(rsTmp.Fields("TM11"))
               Else
                  textCP47 = ""
               End If
               ' 申請案號
               If IsNull(rsTmp.Fields("TM12")) = False Then
                  textCP30 = rsTmp.Fields("TM12")
                  textCP30.Tag = textCP30 'Added by Morgan 2025/9/12
               Else
                  textCP30 = ""
               End If
            End If
            'Add By Cheng 2003/12/25
            '放棄專用權
            m_TM67 = "" & rsTmp("TM67").Value
            'add by nickc 2006/11/17
            textPrint = CheckStr(rsTmp.Fields("TM77"))
            'End
         End If
         rsTmp.Close
      Case Else
         strSql = "SELECT * FROM ServicePractice " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            ' 商標名稱(中)
            If IsNull(rsTmp.Fields("SP05")) = False Then
               textTM05 = rsTmp.Fields("SP05")
            End If
            ' 申請國家
            If IsNull(rsTmp.Fields("SP09")) = False Then
               textTM10 = GetNationName(rsTmp.Fields("SP09"), 0)
               m_TM10 = rsTmp.Fields("SP09")
            End If
            ' 申請人
            m_TM23 = Empty
            If IsNull(rsTmp.Fields("SP08")) = False Then
               m_TM23 = rsTmp.Fields("SP08")
            End If
            
            'Add By Sindy 2019/12/25
            ' FC代理人
            m_TM44 = Empty
            If IsNull(rsTmp.Fields("SP26")) = False Then
               m_TM44 = rsTmp.Fields("SP26")
            End If
            '2019/12/25 END
            
            ' 若案件性質為申請時帶出申請日及申請案號
            If m_CP10 = "806" Then
               ' 申請日
               If IsNull(rsTmp.Fields("SP10")) = False Then
                  textCP47 = TAIWANDATE(rsTmp.Fields("SP10"))
               Else
                  textCP47 = ""
               End If
               ' 申請案號
               If IsNull(rsTmp.Fields("SP11")) = False Then
                  textCP30 = rsTmp.Fields("SP11")
               Else
                  textCP30 = ""
               End If
            End If
            'add by nickc 2006/11/17
            textPrint = CheckStr(rsTmp.Fields("SP72"))
         End If
         rsTmp.Close
   End Select
   
   ' 案件性質
   If m_TM10 < "010" Then
      textCP10 = GetCaseTypeName(m_TM01, m_CP10, 0)
   Else
      textCP10 = GetCaseTypeName(m_TM01, m_CP10, 1)
   End If
   
   '2016/3/21 add by sonia 若已輸過通知申請案號則提醒
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' AND " & _
                  "CP10='1101' AND " & _
                  "CP43 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      MsgBox "此筆進度已輸過通知申請案號!!!", vbExclamation + vbOKOnly
   End If
   rsTmp.Close
   '2016/3/21 end
   
   'edit by nick 2004/12/23 加入分割等於申請，但因為不預設所以取消此段
   If m_CP10 = "101" Then
      ' 大約可接獲審定公告
      textCF09 = Empty
      strSql = "SELECT * FROM CaseFee " & _
               "WHERE CF01 = '" & m_TM01 & "' AND " & _
                     "CF02 = '" & m_TM10 & "' AND " & _
                     "CF03 = '1101' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CF09")) = False Then
            textCF09 = rsTmp.Fields("CF09")
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
'2013/12/3 cancel by sonia 全部設在cf09
'    'Add By Cheng 2003/01/07
'    '大陸案的申請, 回音預設為24個月
'   '92.1.10 CANCEL BY SONIA
'   ' If m_TM10 = 大陸國家代號 And m_CP10 = "101" Then
'   '     textCF09 = "24個月"
'   ' End If
'   '92.1.10 END
'    'Add By Cheng 2003/03/03
'    '大對台的案件預設回音12個月
'    strTM23Nation = Empty
'    If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
'    'Modify By Cheng 2002/06/12
'    '申請國家為台灣
'    If m_TM10 < "010" Then
'        ' 申請人國籍不為台灣
'        'If strTM23Nation >= "010" Then
'            If m_CP10 = "101" Then Me.textCF09.Text = "12個月"
'        'End If
'    End If
'2013/12/3 end
   '2006/5/18 ADD BY SONIA 預設定稿語文
    'edit by nickc 2006/07/05 帶列印定稿預設值 已經在上面做過了
    'textPrint = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
   '2006/5/18 END
   
   'add by nickc 2006/07/24  若是由分割選子案畫面來，分割子案的申請日要同母案申請日
   If UpForm Is frm02010401_6 Then   '能進來代表是子案
         strSql = "SELECT * FROM TradeMark,divisioncase " & _
                  "WHERE dc01 = '" & m_TM01 & "' AND " & _
                        "dc02 = '" & m_TM02 & "' AND " & _
                        "dc03 = '" & m_TM03 & "' AND " & _
                        "dc04 = '" & m_TM04 & "' and dc05=tm01(+) and dc06=tm02(+) and dc07=tm03(+) and dc08=tm04(+) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
        If IsNull(rsTmp.Fields("tm11")) = False Then
           textCP47 = TAIWANDATE(rsTmp.Fields("tm11"))
        Else
           textCP47 = ""
        End If
        textTM27 = (CheckStr(rsTmp.Fields("TM27")))
        
        m_MonTM01 = CheckStr(rsTmp.Fields("tm01"))
        m_MonTM02 = CheckStr(rsTmp.Fields("tm02"))
        m_MonTM03 = CheckStr(rsTmp.Fields("tm03"))
        m_MonTM04 = CheckStr(rsTmp.Fields("tm04"))
        If textNP08.Enabled = True And textNP09.Enabled = True Then
             strSql = "SELECT * FROM nextprogress " & _
                      "WHERE np02 = '" & m_MonTM01 & "' AND " & _
                           " np03 = '" & m_MonTM02 & "' AND " & _
                           " np04 = '" & m_MonTM03 & "' AND " & _
                           " np05 = '" & m_MonTM04 & "' and np06 is null and np07=202 "
            rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
                m_MonNP08 = CheckStr(rsTmp.Fields("np08"))
                m_MonNP09 = CheckStr(rsTmp.Fields("np09"))
            End If
        End If
      End If
      rsTmp.Close
   End If
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/17
   If textPrint = "" Then
       textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   
   'add by sonia 2018/12/7 若已有過「核准通知」，則提醒『已通知核准，此程序將不函知客戶』T-212696
   'modify by sonia 2019/1/4 再加核駁 T-216082, 同時加核准1001
   'modify by sonia 2019/8/19 +1205部分核駁
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                 " CP02 = '" & m_TM02 & "' AND " & _
                 " CP03 = '" & m_TM03 & "' AND " & _
                 " CP04 = '" & m_TM04 & "' AND CP10 IN ('1102','1001','1002','1205') AND CP43='" & m_CP09 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      MsgBox "已通知審定，此程序將不函知客戶！"
      textPrint = "N"
   End If
   rsTmp.Close
   'end 2018/12/7
   
   'Add By Sindy 2009/09/03
   If m_CP10 <> "308" Then
      Cancel = False
      textPrint_Validate Cancel
   End If
   
   Call ChgType 'Add By Sindy 2012/5/18 讀取來函期限
End Function

Private Sub cmdExit_Click()
   Unload frm02010301_1
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
'add by nickc 2008/01/23 加入可以取消
If UpForm Is Nothing Or Me.Visible = False Then
   strTit = "回前畫面"
   strMsg = "你並未存檔，確定離開嗎?"
   nResponse = MsgBox(strMsg, vbYesNo, strTit)
   If nResponse = vbYes Then
      Unload Me
      frm02010301_1.Show
   End If
Else
    'add by nickc 2008/01/23 加入可以取消
    If UpForm Is frm02010401_6 Then
        frm02010401_6.m_IsCancal = True
        Unload Me
    End If
End If
End Sub

Public Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
Exit Sub
End Sub

'Add By Sindy 2009/05/14
Public Sub PubShowNextData()
Select Case cmdState
Case 0
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      
      'add by nickc 2006/07/27
      If UpForm Is Nothing Or Me.Visible = False Then
            ' 設定滑鼠游標為等待狀態
            Screen.MousePointer = vbHourglass
            ' 更新資料
              'Modify By Cheng 2002/11/07
      '      OnWork
              If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
              'Add By Cheng 2002/11/08
              ' 列印定稿
              
              If textPrint <> "N" Then
                 PrintLetter
              End If
      
            ' 設定滑鼠游標為預設
            Screen.MousePointer = vbDefault
      End If
      
      If UpForm Is Nothing Then
          '******* 90.11.23 nick    清畫面
    '      frm02010301_1.textTM01.Text = ""
          frm02010301_1.textTM02.Text = ""
          frm02010301_1.textTM02_2.Text = ""
          frm02010301_1.textTM03.Text = ""
          frm02010301_1.textTM04.Text = ""
          frm02010301_1.textTM05.Text = ""
          frm02010301_1.textTM06.Text = ""
          frm02010301_1.textTM07.Text = ""
          frm02010301_1.textTM10.Text = ""
          frm02010301_1.textTM23.Text = ""
          'Modify By Cheng 2002/07/11
    '      frm02010301_1.Init
          frm02010301_1.InitialGrdList
          '***********************************
          'Add By Sindy 2019/5/10
          If Me.m_strIR01 <> "" Then
            Unload frm02010301_1
            If Not m_PrevForm Is Nothing Then
               Call m_PrevForm.GoNext
            End If
            Unload Me
          Else
          '2019/5/10 END
            Unload Me
            frm02010301_1.Show
            frm02010301_1.textTM02.SetFocus
          End If
       ElseIf UpForm Is frm02010401_6 Then
          '若是畫面有出現可以輸資料，要將資料丟回前面存
          If Me.Visible = True Then
            frm02010401_6.PutSeekData01 = textCP47
            frm02010401_6.PutSeekData02 = textCP30
            frm02010401_6.PutSeekData03 = textTM27
            frm02010401_6.PutSeekData04 = textCP05
            frm02010401_6.PutSeekData05 = textPrint
            frm02010401_6.PutSeekData06 = textCF09
            frm02010401_6.PutSeekData07 = textCP45
            frm02010401_6.PutSeekData08 = textPS
            frm02010401_6.PutSeekData09 = textNP08
            frm02010401_6.PutSeekData10 = textNP09
          End If
          Unload Me
       End If
   End If
   
'add by nick 2004/10/05
Case 6
    'frm03010303_04.Hide 'Modify By Sindy 2009/09/17
    Set frm03010303_04.UpForm = Me
    frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 'textTM01 'lbl1(0).Caption
    frm03010303_04.AllClass = textTM09 'txt1(0).Text
    frm03010303_04.cmdok(0).Visible = False
    frm03010303_04.cmd.Visible = False
    frm03010303_04.cmd2.Visible = False
    frm03010303_04.txt2(0).Visible = False
    frm03010303_04.Line1.Visible = False
    frm03010303_04.txt2(1).Visible = False
    frm03010303_04.txt2(2).Visible = False
    frm03010303_04.txt2(3).Visible = False
    frm03010303_04.Caption = "商品及服務資料"
    'edit by nickc 2008/02/12 改成可以複製
    'frm03010303_04.TXT1(0).Enabled = False
    'frm03010303_04.TXT1(1).Enabled = False
    'frm03010303_04.TXT1(2).Enabled = False
    frm03010303_04.txt1(0).Locked = True
    frm03010303_04.txt1(1).Locked = True
    frm03010303_04.txt1(2).Locked = True
    frm03010303_04.Label2.Visible = False
    'Me.Hide 'Modify By Sindy 2009/09/17
    frm03010303_04.QueryData
    frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
End Select
End Sub

'Modify By Cheng 2002/11/07
'Private Sub OnWork()
Private Function OnWork() As Boolean
Dim strSql As String
Dim strSubSQL As String
Dim strTemp As String
Dim strCP05 As String
Dim strCP09 As String
Dim strCP12 As String
Dim strCP27 As String
Dim rsTmp As New ADODB.Recordset
'add by nickc 2006/07/25
Dim ii As Integer
Dim strNP08 As String 'Add By Sindy 2010/10/25
'add by sonia 2020/12/24
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim strNP07 As String
Dim strNP09 As String
Dim bInsert As Boolean
'end 2020/12/24
   
'add by nickc 2006/07/25
If Me.Visible = True Then
'Add By Cheng 2002/11/07
    On Error GoTo ErrorHandler
End If
OnWork = True
'add by nickc 2006/07/25
If Me.Visible = True Then
    cnnConnection.BeginTrans
End If
   
   ' 申請國家為大陸時, 更新案件進度檔
   '92.9.10 MODIFY BY SONIA
   'If m_TM10 = "020" Then
   If m_TM10 <> "000" Then
   '92.9.10 END
      strCP09 = m_CP09
      strSql = "UPDATE CaseProgress "
      ' 代理人提申日 (申請日)
      strTemp = "00000000"
      If IsEmptyText(textCP47) = False Then: strTemp = ChangeTStringToWString(textCP47)
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & ", "
      strSubSQL = strSubSQL & "CP47=" & strTemp & " "
      ' 大陸申請案號
      strTemp = Empty
      If IsEmptyText(textCP30) = False Then: strTemp = textCP30
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & ", "
      strSubSQL = strSubSQL & "CP30='" & strTemp & "' "
      ' 彼所案號
      strTemp = Empty
      If IsEmptyText(textCP45) = False Then: strTemp = textCP45
      If strSubSQL <> Empty Then: strSubSQL = strSubSQL & ", "
      ' 91.03.25 (單引號)
      'strSubSQL = strSubSQL & "CP45='" & strTemp & "' "
      strSubSQL = strSubSQL & "CP45=" & CNULL(ChgSQL(strTemp)) & " "
      ' 組成SQL語法
      If strSubSQL <> Empty Then: strSql = strSql & " SET " & strSubSQL
      strSql = strSql & "WHERE CP01 = '" & m_TM01 & "' AND " & _
                              "CP02 = '" & m_TM02 & "' AND " & _
                              "CP03 = '" & m_TM03 & "' AND " & _
                              "CP04 = '" & m_TM04 & "' AND " & _
                              "CP09 = '" & m_CP09 & "' "
      ' 更新資料庫
      cnnConnection.Execute strSql
      'add by nick 2005/01/04 更新相同本所案號之相同代理人的彼所案號，若是彼所案號空的話
      strSql = "update caseprogress set cp45=" & CNULL(ChgSQL(strTemp)) & " where cp09 in (select cp09 from caseprogress where cp45 is null and CP01 = '" & m_TM01 & "' AND  CP02 = '" & m_TM02 & "' AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' and cp09<'C' AND cp44 in (select cp44 from caseprogress where cp09='" & m_CP09 & "' ))"
      cnnConnection.Execute strSql
    End If
   
   
   ' 更新基本檔
   Select Case m_TM01
      ' 更新商標基本檔
      Case "T", "TF", "FCT", "CFT":
         strSql = "UPDATE TradeMark "
         strSubSQL = Empty
         'edit by nick 2004/12/23 加入分割等於申請
         'If m_CP10 = "101" Then
         If m_CP10 = "101" Or m_CP10 = "308" Then
            ' 申請日
            If IsEmptyText(textCP47) = False Then
               If strSubSQL <> Empty Then: strSubSQL = strSubSQL & ", "
               strSubSQL = strSubSQL & "TM11=" & DBDATE(textCP47) & " "
            End If
            ' 申請案號
            If IsEmptyText(textCP30) = False Then
               If strSubSQL <> Empty Then: strSubSQL = strSubSQL & ", "
               strSubSQL = strSubSQL & "TM12='" & textCP30 & "' "
            End If
            ' 正商標號數
            If IsEmptyText(textTM27) = False Then
               If strSubSQL <> Empty Then: strSubSQL = strSubSQL & ", "
               strSubSQL = strSubSQL & "TM27='" & textTM27 & "' "
            End If
            'add by nickc 2006/11/17
            If textPrint <> "N" Then
               If strSubSQL <> Empty Then: strSubSQL = strSubSQL & ", "
               strSubSQL = strSubSQL & "TM77='" & textPrint & "' "
            End If
            ' 組成SQL語法
            If strSubSQL <> Empty Then: strSql = strSql & " SET " & strSubSQL
            strSql = strSql & "WHERE TM01 = '" & m_TM01 & "' AND " & _
                                    "TM02 = '" & m_TM02 & "' AND " & _
                                    "TM03 = '" & m_TM03 & "' AND " & _
                                    "TM04 = '" & m_TM04 & "' "
            ' 更新資料庫
            cnnConnection.Execute strSql
            
            UpdateT308EDoc 'Added by Morgan 2025/9/12
         End If
      ' 更新服務業務基本檔
      Case Else
         strSql = "UPDATE ServicePractice "
         strSubSQL = Empty
         If m_CP10 = "806" Then
            ' 申請日
            If IsEmptyText(textCP47) = False Then
               If strSubSQL <> Empty Then: strSubSQL = strSubSQL & ", "
               strSubSQL = strSubSQL & "SP10=" & DBDATE(textCP47) & " "
            End If
            ' 申請案號
            If IsEmptyText(textCP30) = False Then
               If strSubSQL <> Empty Then: strSubSQL = strSubSQL & ", "
               strSubSQL = strSubSQL & "SP11='" & textCP30 & "' "
            End If
            'add by nickc 2006/11/17
            If textPrint <> "N" Then
               If strSubSQL <> Empty Then: strSubSQL = strSubSQL & ", "
               strSubSQL = strSubSQL & "SP72='" & textPrint & "' "
            End If
            ' 組成SQL語法
            If strSubSQL <> Empty Then: strSql = strSql & " SET " & strSubSQL
            strSql = strSql & "WHERE SP01 = '" & m_TM01 & "' AND " & _
                                    "SP02 = '" & m_TM02 & "' AND " & _
                                    "SP03 = '" & m_TM03 & "' AND " & _
                                    "SP04 = '" & m_TM04 & "' "
            ' 更新資料庫
            cnnConnection.Execute strSql
         End If
   End Select
   
   'Modify By Cheng 2002/06/18
   ' 當案件性質為申請時, 另外新增一筆資料到案件進度檔
'   If m_CP10 = "101" Then
'   ' 當案件性質為"申請"(101) 或 "著作權申請"(806)時, 另外新增一筆資料到案件進度檔
   'edit by nick 2004/12/23 加入分割等於申請
   'If m_CP10 = "101" Or m_CP10 = "806" Then
   'edit by nickc 2006/07/25 308 不做
   'If m_CP10 = "101" Or m_CP10 = "806" Or m_CP10 = "308" Then
   'Modified by Lydia 2016/03/10 T、TF非台灣案增加通知申請案進度
   'If m_CP10 = "101" Or m_CP10 = "806" Then
   'modify by sonia 2021/3/8 +TC TC-010755
   If m_CP10 = "101" Or m_CP10 = "806" Or (m_TM10 <> "000" And (m_TM01 = "T" Or m_TM01 = "TF" Or m_TM01 = "TC")) Then
      ' 收文號
      strCP09 = AutoNo("C", 6)
      ' 收文日
      strCP05 = "00000000"
      If IsEmptyText(textCP05) = False Then: strCP05 = DBDATE(textCP05)
      ' 業務區別    '91.0822    改成原來的  nick
      'strCP12 = GetST15(m_CP13)
      ' 發文日
      strCP27 = DBDATE(SystemDate())
      ' 組成SQL語法
      'modify by sonia 2017/3/27 智權人員m_CP13改依規則
      strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14, CP20, CP26, CP27, CP32, CP43) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "1101" & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "','" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "')"
      ' 新增資料到資料庫
      cnnConnection.Execute strSql
      
      'Add By Sindy 2019/6/28 輸入申請的申請案號時,要沖銷1101期限NP06=Y
      If m_CP10 = "101" And IsEmptyText(textCP30) = False Then
         strSql = "update nextprogress set np06='Y' where np01='" & m_CP09 & "' and np07='1101'"
         cnnConnection.Execute strSql
      End If
      
      'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
      Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
      
      'Add By Sindy 2019/11/22 商標電子化
      If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
         strLD18 = strCP09
         'Add By Sindy 2020/11/6
         If m_TM10 = "020" Then '受理通知書(彩色OA)
            PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", False, m_TM23, "1101", m_TM44
         Else
         '2020/11/6 END
            'Modify By Sindy 2021/1/5 附件數1,等申請案的收據或回執
            PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", False, m_TM23, "1101", m_TM44
         End If
      End If
      '2019/11/22 END
      
      Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   
      'add by sonia 2020/12/24 TF子案有菲律賓030時, 新增下一程序為"使用宣誓", 法定期限為申請日 + 3年, 本所期限 = 法定期限 - 2個月
      If m_TM01 = "TF" And m_TM03 = "0" And m_TM04 = "00" Then
         Set rsA = New ADODB.Recordset
         StrSQLa = "select * from trademark where tm01='" & m_TM01 & "' and tm04<>'00' and tm29 is null and tm10='030' " & IIf(Mid(m_TM02, 6, 1) = "0", " and substr(tm02,1,5)='" & Mid(m_TM02, 1, 5) & "' ", " and tm02='" & m_TM02 & "' ")
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            strNP07 = "105" '使用宣誓
            '法定期限
            strNP09 = DBDATE(DateAdd("yyyy", 3, ChangeWStringToWDateString(DBDATE(textCP47.Text))))
            '本所期限,非工作天則直接調整至最近的工作天
            strNP08 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
            strNP08 = PUB_GetWorkDay1(strNP08, True)
            rsA.MoveFirst
            Do While Not rsA.EOF
               '判斷是否已掛使用宣誓期限
               bInsert = True
               strSql = "SELECT * FROM NextProgress " & _
                        "WHERE NP02 = '" & rsA.Fields("TM01") & "' AND " & _
                              "NP03 = '" & rsA.Fields("TM02") & "' AND " & _
                              "NP04 = '" & rsA.Fields("TM03") & "' AND " & _
                              "NP05 = '" & rsA.Fields("TM04") & "' AND " & _
                              "NP06 IS NULL AND NP07 = '" & strNP07 & "' "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount > 0 Then
                  bInsert = False
               End If
               If bInsert = False Then
                  strSql = "Update NextProgress Set NP01='" & strCP09 & "',NP08=" & strNP08 & ", NP09=" & strNP09 & " " & _
                            "WHERE NP02 = '" & rsA.Fields("TM01") & "' AND " & _
                              "NP03 = '" & rsA.Fields("TM02") & "' AND " & _
                              "NP04 = '" & rsA.Fields("TM03") & "' AND " & _
                              "NP05 = '" & rsA.Fields("TM04") & "' AND " & _
                              "NP06 IS NULL AND NP07 = '" & strNP07 & "' "
               Else
                  strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                           "VALUES ('" & strCP09 & "','" & rsA.Fields("TM01") & "','" & rsA.Fields("TM02") & "','" & rsA.Fields("TM03") & "','" & rsA.Fields("TM04") & "','" & strNP07 & "'," & _
                                       strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
               End If
               cnnConnection.Execute strSql
               If rsTmp.State <> adStateClosed Then rsTmp.Close
               Set rsTmp = Nothing
               rsA.MoveNext
            Loop
         End If
      End If
      'end 2020/12/24
   End If
   
    'add by nickc 2006/07/24
    If m_CP10 = "308" Then
      '新增子案核准來文
      strCP09 = AutoNo("C", 6)
      strCP05 = DBDATE(UpForm.oStrCDate)
      strCP27 = DBDATE(SystemDate())
      ' 組成SQL語法
      'modify by sonia 2017/3/27 智權人員m_CP13改依規則
      strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26, cp27,  CP43) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "1001" & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "','" & "N" & "'," & strCP27 & ",'" & m_CP09 & "')"
      ' 新增資料到資料庫
      cnnConnection.Execute strSql
      
      'Add By Sindy 2019/11/22 商標電子化
      If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
         strLD18 = strCP09
         PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", False, m_TM23, "1001", m_TM44
      End If
      '2019/11/22 END
      
      Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
      
      'Added by Morgan 2017/6/14 電子公文
      If m_DocNo <> "" Then
         '更新機關文號
         strSql = "update caseprogress set cp08='" & m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號' where cp09='" & strCP09 & "'"
         cnnConnection.Execute strSql, intI
         '複製母案公文電子檔
         strExc(0) = PUB_GetEDocFileName(m_TM01, m_TM02, m_TM03, m_TM04, "1001")
         SaveAttFile_PDF strCP09, m_DocPdf, strExc(0), Format(m_DocPdfDate), Format(m_DocPdfTime), False, , , True
      End If
      'end 2017/6/14
      
      '新增子案申請，自動發文
      strCP09 = AutoNo("B", 6)
      strCP05 = DBDATE("111111")
      strCP27 = DBDATE("111111")   '2010/11/3 MODIFY BY SONIA 原放系統日
      ' 組成SQL語法
      'modify by sonia 2017/3/27 智權人員m_CP13改依規則
      strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26, cp27,  CP43) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "101" & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "','" & "N" & "'," & strCP27 & ",'" & m_CP09 & "')"
      ' 新增資料到資料庫
      cnnConnection.Execute strSql
      'add by nickc 2006/11/09 將子按下一程序分割催審，改成申請催審
'      strSql = "update nextprogress set np01='" & strCP09 & "' where np01='" & m_CP09 & "' and np07=305 "
'      cnnConnection.Execute strSql
      'Add By Sindy 2010/10/25 改為更新子案分割催審之NP06=Y
      strSql = "update nextprogress set np06='Y' where np01='" & m_CP09 & "' and np07=305 "
      cnnConnection.Execute strSql
      'Add By Sindy 2010/10/25 另新增B類申請之催審期限
      strNP08 = GetUrgeDate(m_TM01, m_TM10, "101", DBDATE(UpForm.oStrCDate))
      'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',305," & _
                          strNP08 & "," & strNP08 & ",'" & m_CP14 & "'," & GetNextProgressNo & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',305," & _
                          PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & m_CP14 & "'," & GetNextProgressNo & ")"
      cnnConnection.Execute strSql
      '更新子案核准及結果日
      strCP05 = DBDATE(UpForm.oStrCDate)
      strSql = "update caseprogress set cp24='1',cp25=" & strCP05 & " where cp09='" & m_CP09 & "' "
      cnnConnection.Execute strSql
         '有期限時
         If textNP08.Enabled = True And textNP09.Enabled = True Then
                '若畫面有輸入新期限以新期限為主，沒有的話將繼承母案期限
                If Trim(textNP08) <> "" And Trim(textNP09) <> "" Then
                   If UpForm.IsHaveNp202 Then
                         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                             "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',202," & _
                             DBDATE(textNP08) & "," & DBDATE(textNP09) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo & ")"
                         cnnConnection.Execute strSql
                   ElseIf UpForm.IsHaveCp202 Then
'                         strCP09 = AutoNo("B", 6)
'                         strCP05 = DBDATE("111111")
'                         strCP27 = "null"
'                         strSQL = "insert into caseprogress select "
'                         For ii = 1 To TF_CP
'                             Select Case ii
'                             Case 1
'                                 strSQL = strSQL & "'" & m_TM01 & "',"
'                             Case 2
'                                strSQL = strSQL & "'" & m_TM02 & "',"
'                             Case 3
'                                strSQL = strSQL & "'" & m_TM03 & "',"
'                             Case 4
'                                strSQL = strSQL & "'" & m_TM04 & "',"
'                             Case 9
'                                 strSQL = strSQL & "'" & strCP09 & "',"
'                             Case 27
'                                 strSQL = strSQL & strCP27 & ","
'                             Case 5
'                                 strSQL = strSQL & strCP05 & ","
'                             Case 6
'                                 strSQL = strSQL & DBDATE(textNP08) & ","
'                             Case 7
'                                 strSQL = strSQL & DBDATE(textNP09) & ","
'                             Case Else
'                                 If ii < 100 Then
'                                     strSQL = strSQL & "CP" & Format(ii, "00") & ","
'                                 Else
'                                     strSQL = strSQL & "CP" & Format(ii, "000") & ","
'                                 End If
'                             End Select
'                         Next ii
'                         strSQL = Left(strSQL, Len(strSQL) - 1) & " from caseprogress where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202' and cp27 is null  "
'                         cnnConnection.Execute strSQL
                        If Trim(textNP08) <> "" Then
                            strSql = "update caseprogress set cp06=" & DBDATE(textNP08) & ",cp07=" & DBDATE(textNP09) & ",cp64=cp64||'從母案轉入，案號：" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";原相關收文號：'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        Else
                            strSql = "update caseprogress set cp64=cp64||'從母案轉入，案號：" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";原相關收文號：'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        End If
                        cnnConnection.Execute strSql
                        '2010/11/17 modify by sonia cp43改掛分割案之B類申請strCP09
                        'strSql = "update caseprogress set cp43='" & m_CP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        strSql = "update caseprogress set cp43='" & strCP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        cnnConnection.Execute strSql
                        strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        cnnConnection.Execute strSql
                        'Add by Sonia 2013/8/8 同時更正ACC0J0的T-184230,不可更新已收款傳票的案號,因為分割與申請意見書的案號因上述語法而不同
                        strSql = "update acc0j0 set a0j02='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "' where a0j01 in (select cp09 from caseprogress where cp27 is null and cp57 is null and cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='202') "
                        cnnConnection.Execute strSql
                        'end 2013/8/8

                   End If
                Else
                   If UpForm.IsHaveNp202 Then
                         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                             "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',202," & _
                             m_MonNP08 & "," & m_MonNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo & ")"
                         cnnConnection.Execute strSql
                   ElseIf UpForm.IsHaveCp202 Then
'直接轉子案
'                         strCP09 = AutoNo("B", 6)
'                         strCP05 = DBDATE("111111")
'                         strCP27 = "null"
'                         strSQL = "insert into caseprogress select "
'                         For ii = 1 To TF_CP
'                             Select Case ii
'                             Case 1
'                                 strSQL = strSQL & "'" & m_TM01 & "',"
'                             Case 2
'                                strSQL = strSQL & "'" & m_TM02 & "',"
'                             Case 3
'                                strSQL = strSQL & "'" & m_TM03 & "',"
'                             Case 4
'                                strSQL = strSQL & "'" & m_TM04 & "',"
'                             Case 9
'                                 strSQL = strSQL & "'" & strCP09 & "',"
'                             Case 27
'                                 strSQL = strSQL & strCP09 & ","
'                             Case 5
'                                 strSQL = strSQL & strCP05 & ","
'                             Case Else
'                                 If ii < 100 Then
'                                     strSQL = strSQL & "CP" & Format(ii, "00") & ","
'                                 Else
'                                     strSQL = strSQL & "CP" & Format(ii, "000") & ","
'                                 End If
'                             End Select
'                         Next ii
'                         strSQL = Left(strSQL, Len(strSQL) - 1) & " from caseprogress where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202' and cp27 is null  "
'                         cnnConnection.Execute strSQL
                        If Trim(textNP08) <> "" Then
                            strSql = "update caseprogress set cp06=" & DBDATE(textNP08) & ",cp07=" & DBDATE(textNP09) & ",cp64=cp64||'從母案轉入，案號：" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";原相關收文號：'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        Else
                            strSql = "update caseprogress set cp64=cp64||'從母案轉入，案號：" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";原相關收文號：'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        End If
                        cnnConnection.Execute strSql
                        '2010/11/17 modify by sonia cp43改掛分割案之B類申請
                        'strSql = "update caseprogress set cp43='" & m_CP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        strSql = "update caseprogress set cp43='" & strCP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        cnnConnection.Execute strSql
                        strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        cnnConnection.Execute strSql
                        'Add by Sonia 2013/8/8 同時更正ACC0J0的T-184230,不可更新已收款傳票的案號,因為分割與申請意見書的案號因上述語法而不同
                        strSql = "update acc0j0 set a0j02='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "' where a0j01 in (select cp09 from caseprogress where cp27 is null and cp57 is null and cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='202') "
                        cnnConnection.Execute strSql
                        'end 2013/8/8

                   End If
                End If
                If UpForm.IsHaveNp202 Then
                     strSql = "update nextprogress set np06='N',np15=np15||'轉入子案，子案案號：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' where np02='" & m_MonTM01 & "' and np03='" & m_MonTM02 & "' and np04='" & m_MonTM03 & "' and np05='" & m_MonTM04 & "' and np06 is null and np07=202 "
                     cnnConnection.Execute strSql
                ElseIf UpForm.IsHaveCp202 Then
                     strSql = "update caseprogress set cp57=to_number(to_char(sysdate,'YYYYMMDD')),cp64=cp64||'轉入子案，子案案號：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202' and cp27 is null "
                     cnnConnection.Execute strSql
                End If
                '母案分割發文後的收文及發文案件皆轉入有期限的子案
                Dim m_MonCP27 As String
                strSql = "select cp27 from caseprogress where cp09='" & m_MonCP09 & "' "
                m_MonCP27 = ""
                Set rsTmp = New ADODB.Recordset
                If rsTmp.State = 1 Then rsTmp.Close
                rsTmp.CursorLocation = adUseClient
                rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
                If rsTmp.RecordCount > 0 Then
                    m_MonCP27 = CheckStr(rsTmp.Fields("cp27"))
                End If
                If m_MonCP27 <> "" Then
                   strSql = "update caseprogress set cp64=cp64||'從母案轉入，案號：" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & "' where cp05>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10<>'1001' "
                   cnnConnection.Execute strSql
                   strSql = "update caseprogress set cp64=cp64||'從母案轉入，案號：" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & "' where cp27>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10<>'1001'  "
                   cnnConnection.Execute strSql
                   strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp05>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10<>'1001'  "
                   cnnConnection.Execute strSql
                   'Add by Sonia 2013/8/8 同時更正ACC0J0的T-184230,不可更新已收款傳票的案號,因為分割與申請意見書的案號因上述語法而不同
                   strSql = "update acc0j0 set a0j02='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "' where a0j01 in (select cp09 from caseprogress where cp05>" & m_MonCP27 & " and cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10<>'1001') "
                   cnnConnection.Execute strSql
                   'end 2013/8/8
                   
                   strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10<>'1001'  "
                   cnnConnection.Execute strSql
                   'Add by Sonia 2013/8/8 同時更正ACC0J0的T-184230,不可更新已收款傳票的案號,因為分割與申請意見書的案號因上述語法而不同
                   strSql = "update acc0j0 set a0j02='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "' where a0j01 in (select cp09 from caseprogress where cp27>" & m_MonCP27 & " and cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10<>'1001') "
                   cnnConnection.Execute strSql
                   'end 2013/8/8
                End If
         End If
    End If
   
   'Add By Sindy 2009/09/03
   '存檔時(商標局發函日)此欄有輸入值則存入cp64
   'Modify By Sindy 2010/3/10 增加CP133
   If m_CP10 <> "308" Then
      If Trim(m_TM10) = "020" And Trim(textPrint.Text) = "3" And textCP64.Visible = True Then
         If m_CP10 <> "101" Then '爭議案
            strSql = "UPDATE caseprogress SET CP64 = decode(CP64,null,'主管機關受理函發文日：" & Trim(textCP64.Text) & "',CP64||" & "';'" & "||'主管機關受理函發文日：" & Trim(textCP64.Text) & "'),CP133= " & Trim(textCP64.Text) & _
                     " WHERE CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql
         Else '申請案
            strSql = "UPDATE caseprogress SET CP64 = decode(CP64,null,'主管機關發函日：" & Trim(textCP64.Text) & "',CP64||" & "';'" & "||'主管機關發函日：" & Trim(textCP64.Text) & "'),CP133= " & Trim(textCP64.Text) & _
                     " WHERE CP09 = '" & strCP09 & "' "
            cnnConnection.Execute strSql
         End If
      End If
   End If
   
   ' 更新下一程序檔
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 = " & "997"
   cnnConnection.Execute strSql
   ' 更新下一程序檔
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 = " & "998"
   cnnConnection.Execute strSql
   
   '2006/11/3 ADD BY SONIA 非台灣之申請案以申請日重新算催審期限
   Dim NP08305 As String
   If m_TM10 <> "000" And m_CP10 = "101" Then
      NP08305 = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP47)
      'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "UPDATE NextProgress SET NP08 = '" & NP08305 & "', NP09 = '" & NP08305 & "' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = " & "305"
      strSql = "UPDATE NextProgress SET NP08 = '" & PUB_GetWorkDay1(NP08305, True) & "', NP09 = '" & NP08305 & "' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = " & "305"
   cnnConnection.Execute strSql
   End If
   '2006/11/3 END
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'm_CP09 = strCP09
   'add by nick 2004/12/23 區分分割定稿及申請案號定稿
   If m_CP10 <> "308" Then
        m_CP10 = "1101"
   End If
   'add by nickc 2006/07/28
   If Not (UpForm Is Nothing) Then
       m_CP10 = "1101"
   End If
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If

   'Add by Sindy 2019/5/10
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010301_1", strCP09
   End If
   '2019/5/10 END

'add by nickc 2006/07/25
If Me.Visible = True Then
    'Add By Cheng 2002/11/07
    cnnConnection.CommitTrans
End If
Exit Function
ErrorHandler:
    'add by nickc 2006/07/25
    If Me.Visible = True Then
        cnnConnection.RollbackTrans
    End If
    OnWork = False
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010301_2 = Nothing
End Sub

' 來函收文日
Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If CheckIsTaiwanDate(textCP05, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "來函收文日不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If
      ' 檢查來函記錄檔
      'strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05, "MR16")
      'If IsEmptyText(strDate) = False Then
      '   If TAIWANDATE(textCP05) <> TAIWANDATE(strDate) Then
      '      strTit = "檢核資料"
      '      strMsg = "與櫃台之來函收文記錄不符, 起確認"
      '      nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
      '      If nResponse = vbNo Then
      '         Cancel = True
      '         textCP05_GotFocus
      '      End If
      '   End If
      'End If
   End If
EXITSUB:
End Sub

'Add By Sindy 2010/8/31
Private Sub textCP30_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If IsEmptyText(textCP30) = False Then
      'Modify By Sindy 2010/12/27
      If m_TM10 = "020" And m_CP10 <> "101" Then
         '申請國家為大陸並且不是申請案時, 不須檢查申請案號
      '2010/12/27 End
      Else
         '檢查申請案號所輸入的長度是否正確
         'Add By Sindy 2017/5/17 + strRetrunText
         If PUB_ChkTm12Tm15Length("1", textCP30, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, , , strRetrunText) = False Then
            Cancel = True
            textCP30_GotFocus
            Exit Sub
         'Add By Sindy 2017/5/17
         Else
            textCP30 = strRetrunText
         '2017/5/17 END
         End If
      End If
   End If
End Sub

' 申請日
Private Sub textCP47_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   
   Cancel = False
   If IsEmptyText(textCP47) = False Then
      If CheckIsTaiwanDate(textCP47, False) = False Then
         Cancel = True
         strMsg = "日期不正確"
         strTit = "申請日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP47_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2003/09/08
        '所有系統的系統日一律抓Server的日期
'      If Val(TAIWANDATE(textCP47)) > Val(TAIWANDATE(Date)) Then
      If Val(TAIWANDATE(textCP47)) > Val(TAIWANDATE(strSrvDate(1))) Then
         Cancel = True
         strMsg = "申請日不可超過系統日"
         strTit = "申請日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP47_GotFocus
         GoTo EXITSUB
      End If
   
      'add by sonia 2018/2/6
      'modify by sonia 2018/4/18 分割案要剔除
      If Val(DBDATE(textCP47)) < Val(m_CP27) And m_CP10 <> "308" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "申請日不可小於發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP47_GotFocus
         GoTo EXITSUB
      End If
      'end 2018/2/6
   End If
EXITSUB:
End Sub

' 檢核資料是否正確輸入
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/28檢查畫面的 TextBox是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
   End If

   ' 申請日不可為空白
   If IsEmptyText(textCP47) = True Then
      strMsg = "申請日不可為空白"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP47.SetFocus
      GoTo EXITSUB
   End If
   ' 申請日不可超過系統日
   'edit by nickc 2006/03/17
   'If Val(textCP47) > Val(ChangeWDateStringToWString(Date)) Then
   If Val(textCP47) > Val(strSrvDate(1)) Then
      strMsg = "申請日不可超過系統日"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP47.SetFocus
      GoTo EXITSUB
   End If
   ' 申請案號不可為空白
   'modify by sonia 2014/12/27 大陸商爭4字頭及6字頭之案件無申請案號,定稿也要修改T-192499
   'If IsEmptyText(textCP30) = True Then
   If IsEmptyText(textCP30) = True And Left(m_CP10, 1) <> "4" And Left(m_CP10, 1) <> "6" Then
      strMsg = "申請案號不可為空白"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP30.SetFocus
      GoTo EXITSUB
   End If
   
   'Add By Sindy 2011/6/3
   ' 申請國家為台灣時，檢查申請案號的前二(三)碼必須為申請年度
   '2011/9/22 modify by sonia 分割案不檢查
   'If m_TM10 = "000" Then
   If m_TM10 = "000" And m_CP10 <> "308" Then
      If IsEmptyText(textCP47) = False And IsEmptyText(textCP30) = False Then
         If Val(Left(textCP30, 1)) > "1" Then
            strExc(1) = Val(Left(textCP30, 2))
         Else
            strExc(1) = Val(Left(textCP30, 3))
         End If
         strExc(2) = Trim(Val(textCP47) \ 10000)
         If strExc(1) <> strExc(2) Then
            strTit = "資料檢核"
            strMsg = "申請案號的前二(三)碼必須為申請年度!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP30.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   ' 商標種類為聯合商標, 防護商標, 聯合服務標章, 防護服務標章時, 正商標號數不可為空白
   Select Case m_TM08
      Case "2", "3", "5", "6":
         If IsEmptyText(textTM27) = True Then
            strMsg = "正商標號數不可為空白"
            strTit = "資料檢核"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM27.SetFocus
            GoTo EXITSUB
         End If
   End Select
   ' 來函收文日不可為空白
   If IsEmptyText(textCP05) = True Then
      strMsg = "來函收文日不可為空白"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP05.SetFocus
      GoTo EXITSUB
   End If
   ' 列印定稿
   Select Case textPrint
      Case "N", "1", "2", "3":
      Case Else
         strMsg = "只可輸入N,1,2或3"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPrint.SetFocus
         GoTo EXITSUB
   End Select
   
   'Add By Sindy 2012/5/18
   If LabNP07.Caption <> "" Then
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
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

'Add By Sindy 2009/09/03
Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

'Add By Sindy 2009/09/03
'商標局發函日
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP64) = False Then
      If CheckIsDate(textCP64, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "大陸案發函日不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP64_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textNP08_GotFocus()
InverseTextBox textNP08
End Sub

Private Sub textNP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   
   Cancel = False
   If IsEmptyText(textNP08) = False Then
      If CheckIsTaiwanDate(textNP08, False) = False Then
         Cancel = True
         strMsg = "日期不正確"
         strTit = "子案新本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP08.SetFocus
         textNP08_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08, True), 1)
      'end 2020/07/07
      End If
   End If
EXITSUB:
End Sub

Private Sub textNP09_GotFocus()
    InverseTextBox textNP09
End Sub

Private Sub textNP09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   
   Cancel = False
   If IsEmptyText(textNP09) = False Then
      If CheckIsTaiwanDate(textNP09, False) = False Then
         Cancel = True
         strMsg = "日期不正確"
         strTit = "子案新法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP09_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   Select Case textPrint
      Case "N", "1", "2", "3":
      Case Else
         Cancel = True
         strMsg = "只可輸入 N 或 1-3"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPrint_GotFocus
   End Select
   
   'Add By Sindy 2009/09/03
   '列印定稿為3(英文)時, 申請國家為大陸(020)者, 商標局發函日一定要輸入
   If m_CP10 <> "308" Then
      If Trim(m_TM10) = "020" And Trim(textPrint.Text) = "3" Then
         Label1(19).Visible = True
         textCP64.Visible = True
         'If m_CP10 <> "101" Then '爭議案
         '   Label1(19).Caption = "商標局受理函發文日 :                                                     (西元)"
         'Else '申請案
            Label1(19).Caption = "大陸案發函日 :                                                     (西元)"
         'End If
      Else
         Label1(19).Visible = False
         textCP64.Visible = False
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
      strMsg = "備註欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textPS_GotFocus
   End If
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPS_GotFocus()
   InverseTextBox textPS
End Sub

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
End Sub

Private Sub textTM27_GotFocus()
   InverseTextBox textTM27
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub

Private Sub textCP30_GotFocus()
   InverseTextBox textCP30
End Sub

Private Sub textCP45_GotFocus()
   InverseTextBox textCP45
End Sub

Private Sub textCP47_GotFocus()
   InverseTextBox textCP47
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
Dim strTM23Nation As String
Dim strSql As String
Dim strTmp As String
Dim StrSQLa As String
Dim rsA  As New ADODB.Recordset
Dim strTemp As String
Dim arrTM09 As Variant, strGoodsKind As String 'Add By Sindy 2010/11/12
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
    'Add  By Cheng 2003/12/25
    '判斷是否有優先權資料
    StrSQLa = "Select Count(*) From PriDate Where PD01='" & m_TM01 & "' And PD02='" & m_TM02 & "' And PD03='" & m_TM03 & "' And PD04='" & m_TM04 & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.Fields(0).Value > 0 Then
        m_blnPriDate = True
    Else
        m_blnPriDate = False
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    'End
   ' 案件性質為申請案號
   If m_CP10 = "1101" Then
      ' 系統類別TD
      Select Case m_TM01
         Case "TC":
            ' 清除定稿例外欄位檔原有資料
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                EndLetter "02", m_CP09, "04", strUserNum
                '91.12.5 add by sonia
                ' 列印備註
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "02" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                         "'" & "列印備註" & "','" & textPS & "')"
                cnnConnection.Execute strSql
            End If
            '91.12.5 end
         Case "T", "TF":
            '若定稿語文為中文
            'edit by nickc 2006/06/29
            'If m_strLanguage = "1" Then
            If textPrint = "1" Or textPrint = "2" Then
                '910723 Sieg
                If textCF09 <> "" Then
                   '91.11.13 modify by sonia
                   'strTmp = "大約" & textCF09 & "可接獲審定公告"
                   strTmp = textCF09
                   '91.11.13 end
                Else
                   strTmp = ""
                End If
                If m_TM10 < "010" Then
                   ' 申請人國籍為台灣
                   'edit by nickc 2006/06/29
                   'If strTM23Nation < "010" Then
                   If textPrint = "1" Then
                        '若申請日小於20031128
                        If DBDATE(Me.textCP47.Text) < 20031128 Then
                            ' 清除定稿例外欄位檔原有資料
                            EndLetter "02", m_CP09, "01", strUserNum
                        '若申請日大於等於20031128
                        Else
                            '判斷商標種類
                            Select Case m_TM08
                            Case "6", "7", "8"
                                ' 清除定稿例外欄位檔原有資料
                                EndLetter "02", m_CP09, "07", strUserNum
                               ' 回音
                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "02" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                                        "'回音'," & CNULL(strTmp) & ")"
                               cnnConnection.Execute strSql
                            Case Else
                                ' 清除定稿例外欄位檔原有資料
                                EndLetter "02", m_CP09, "06", strUserNum
                               ' 回音
                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "02" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                                        "'回音'," & CNULL(strTmp) & ")"
                               cnnConnection.Execute strSql
                            End Select
                        End If
                   ' 申請人國籍非台灣
                   'edit by nickc 2006/06/29
                   'Else
                   ElseIf textPrint = "2" Then
                        '若申請日小於20031128
                        If DBDATE(Me.textCP47.Text) < 20031128 Then
                            ' 清除定稿例外欄位檔原有資料
                            EndLetter "02", m_CP09, "02", strUserNum
                            ' 列印備註
                            If textPS <> "" Then
                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "02" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                                        "'" & "列印備註" & "','" & textPS & vbCrLf & "')"
                               cnnConnection.Execute strSql
                            End If
                            ' 回音
                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & "02" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                                     "'回音'," & CNULL(strTmp) & ")"
                            cnnConnection.Execute strSql
                        '若申請日大於等於20031128
                        Else
                            ' 清除定稿例外欄位檔原有資料
                            EndLetter "02", m_CP09, "08", strUserNum
                            ' 列印備註
                            If textPS <> "" Then
                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "02" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                                        "'" & "列印備註" & "','" & textPS & vbCrLf & "')"
                               cnnConnection.Execute strSql
                            End If
                            ' 回音
                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & "02" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                                     "'回音'," & CNULL(strTmp) & ")"
                            cnnConnection.Execute strSql
                        End If
                   End If
                '2009/2/26 MODIFY BY SONIA
                'ElseIf m_TM10 = "020" Then
                ElseIf m_TM10 <> "000" Then
                    'add by nickc 2006/06/29
                    If textPrint = "1" Then
                        ' 清除定稿例外欄位檔原有資料
                        EndLetter "02", m_CP09, "03", strUserNum
                        ' 回音
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "02" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & "'," & _
                                 "'回音'," & CNULL(strTmp) & ")"
                        cnnConnection.Execute strSql
                        '2014/12/27 ADD BY SONIA 大陸商爭4字頭及6字頭之案件無申請案號,定稿也要修改T-192499
                        If textCP30 <> "" Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "02" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & "'," & _
                                    "'大陸申請案號','♀')"
                           cnnConnection.Execute strSql
                        End If
                        '2014/12/27 END
                    End If
                End If
            '若定稿語文為英文
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "3" Then
               'Add By Sindy 2009/09/03
               If m_TM10 = "020" Then
                  'Add By Sindy 2012/1/18
                  If Trim(textCP10) = "轉讓" Then
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "02", m_CP09, "00", strUserNum
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "02", m_CP09, "01", strUserNum
                     ' 商標局發函日
                     If Trim(textCP64) <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "02" & "','" & m_CP09 & "','01','" & strUserNum & _
                                 "','商標局發函日','" & ChgSQL(textCP64) & "')"
                        cnnConnection.Execute strSql
                     End If
                  '2012/1/18 End
                  Else
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "02", m_CP09, "15", strUserNum
                     ' 商標局發函日
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "02" & "','" & m_CP09 & "','15','" & strUserNum & _
                              "','商標局發函日','" & ChgSQL(textCP64) & "')"
                     cnnConnection.Execute strSql
                  End If
               '2009/09/03 End
               Else
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "02", m_CP09, IIf(m_blnPriDate, "10", "12"), strUserNum
                   ' 是否補件
                   If IsEmptyText(strTemp) = False Then
                      ' 是否補件
                      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "10", "12") & "','" & strUserNum & _
                               "','是否補件','" & strTemp & "')"
                      cnnConnection.Execute strSql
                   End If
                   ' 是否列印翻譯函
   '                If textPrtTrans <> "N" Then
                      ' 清除定稿例外欄位檔原有資料
                        EndLetter "02", m_CP09, IIf(m_blnPriDate, "11", "13"), strUserNum
                        'Add By Cheng 2003/02/26
                        '若有放棄專用權
                        If m_TM67 <> "" Then
                            ' 放棄專用權
                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "11", "13") & "','" & strUserNum & _
                                     "','放棄專用權','" & vbCrLf & "The following part disclaimed : " & ChgSQL(m_TM67) & "')"
                            cnnConnection.Execute strSql
                        End If
   '                End If
               End If
            End If
      End Select
   'add by nick 2004/12/23 加入分割定稿
   ElseIf m_CP10 = "308" Then
        'Add By Sindy 2010/11/12
        '1-34商品 35-45服務
        strGoodsKind = "商品"
        If Trim(textTM09.Text) > "" Then
            arrTM09 = Split(textTM09.Text, ",")
            If Val(arrTM09(0)) >= 35 And Val(arrTM09(0)) <= 45 Then
               strGoodsKind = "服務"
            End If
        End If
        If textPrint = "1" Then
            EndLetter "02", m_CP09, "01", strUserNum
            'Add By Sindy 2010/11/12
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "02" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                        "'商品或服務','" & strGoodsKind & "')"
            cnnConnection.Execute strSql
            '2010/11/12 End
        'Add By Sindy 2010/10/29
        ElseIf textPrint = "2" Then '大->台
            EndLetter "02", m_CP09, "02", strUserNum
            'Add By Sindy 2010/11/12
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "02" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                        "'商品或服務','" & strGoodsKind & "')"
            cnnConnection.Execute strSql
            '2010/11/12 End
        End If
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
'Add By Sindy 2012/1/12
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean, ET03_1 As String
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/12 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
    'Add By Cheng 2003/12/25
    '取得定稿語文
    '2005/11/9 MODIFY BY SONIA
    'm_strLanguage = GetLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
    '2006/5/18 MODIFY BY SONIA 改於UpdateCtrl時先預設定稿語文至textPrint,列印定稿時以textPrint判斷
    'm_strLanguage = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
    m_strLanguage = textPrint
    '2006/5/18 END
    '2005/11/9 END
    'End
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/12
   ET01 = "02"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/12 End
   
   ' 案件性質為申請案號
   If m_CP10 = "1101" Then
      ' 系統類別TD
      Select Case m_TM01
         Case "TC":
            ' 列印定稿
            'add by nickc 2006/06/29
            If textPrint = "1" Then
'                NowPrint m_CP09, "02", "04", False, strUserNum, 0
               ET03 = "04" 'Modify By Sindy 2012/1/12
            End If
         Case "T", "TF":
            '若定稿語文為中文
            'edit by nickc 2006/06/29
            'If m_strLanguage = "1" Then
            If textPrint = "1" Or textPrint = "2" Then
                 'Modify By Cheng 2002/06/12
                If m_TM10 < "010" Then
                   ' 申請人國籍為台灣
                   
                   'edit by nickc 2006/06/29
                   'If strTM23Nation < "010" Then
                   If textPrint = "1" Then
                        '若申請日小於20031128
                        If DBDATE(Me.textCP47.Text) < 20031128 Then
                          ' 列印定稿
'                          NowPrint m_CP09, "02", "01", False, strUserNum, 0
                           ET03 = "01" 'Modify By Sindy 2012/1/12
                        '若申請日大於等於20031128
                        Else
                          ' 列印定稿
                          Select Case m_TM08
                          Case "6", "7", "8" '標章
'                             NowPrint m_CP09, "02", "07", False, strUserNum, 0
                              ET03 = "07" 'Modify By Sindy 2012/1/12
                          Case Else
'                             NowPrint m_CP09, "02", "06", False, strUserNum, 0
                              ET03 = "06" 'Modify By Sindy 2012/1/12
                          End Select
                        End If
                   ' 申請人國籍非台灣
                   'edit by nickc 2006/06/29
                   'Else
                   ElseIf textPrint = "2" Then
                        '若申請日小於20031128
                        If DBDATE(Me.textCP47.Text) < 20031128 Then
                            ' 列印定稿
'                            NowPrint m_CP09, "02", "02", False, strUserNum, 0
                           ET03 = "02" 'Modify By Sindy 2012/1/12
                        '若申請日大於等於20031128
                        Else
                          Select Case m_TM08
                          Case "6", "7", "8" '標章   '2009/4/22 ADD BY SONIA 新增定稿
'                             NowPrint m_CP09, "02", "09", False, strUserNum, 0
                              ET03 = "09" 'Modify By Sindy 2012/1/12
                          Case Else
                            ' 列印定稿
'                             NowPrint m_CP09, "02", "08", False, strUserNum, 0
                              ET03 = "08" 'Modify By Sindy 2012/1/12
                          End Select
                        End If
                   End If
                '92.9.10 MODIFY BY SONIA
                'ElseIf m_TM10 = "020" Then
                ElseIf m_TM10 <> "000" Then
                '92.9.10 END
                   ' 列印定稿
                   'add by nickc 2006/06/29
                   If textPrint = "1" Then
'                        NowPrint m_CP09, "02", "03", False, strUserNum, 0
                        ET03 = "03" 'Modify By Sindy 2012/1/12
                   End If
                   'Added by Lydia 2016/03/10 +TF續展定稿
                   If m_TM01 = "TF" And Trim(textCP10) = "續展" Then
                        ET03 = "16"
                   End If
                   
                End If
            '若定稿語文為英文
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "3" Then
               'Add By Sindy 2009/09/03
               If m_TM10 = "020" Then
                  'Add By Sindy 2012/1/18
                  If Trim(textCP10) = "轉讓" Then
                     ET03 = "00"
                     ET03_1 = "01" '英譯本
                  '2012/1/18 End
                  Else
'                     NowPrint m_CP09, "02", "14", False, strUserNum, 0
'                     NowPrint m_CP09, "02", "15", False, strUserNum, 0
                     ET03 = "14" 'Modify By Sindy 2012/1/12
                     ET03_1 = "15" 'Modify By Sindy 2012/1/12
                  End If
               '2009/09/03 End
               Else
                   ' 列印定稿
'                   NowPrint m_CP09, "02", IIf(m_blnPriDate, "10", "12"), False, strUserNum, 0
                   ET03 = IIf(m_blnPriDate, "10", "12") 'Modify By Sindy 2012/1/12
                   ' 是否列印翻譯函
   '                If textPrtTrans <> "N" Then
                      ' 列印定稿
'                       NowPrint m_CP09, "02", IIf(m_blnPriDate, "11", "13"), False, strUserNum, 0
                       ET03_1 = IIf(m_blnPriDate, "11", "13") 'Modify By Sindy 2012/1/12
   '                End If
               End If
            End If
      End Select
   'add by nick 2004/12/23 加入分割定稿
   ElseIf m_CP10 = "308" Then
        'add by nickc 2006/06/29
        If textPrint = "1" Then
'            NowPrint m_CP09, "02", "01", True, strUserNum, 0
            bolEdit = True 'Add By Sindy 2012/1/12
            ET03 = "01" 'Modify By Sindy 2012/1/12
        'Add By Sindy 2010/10/29
        ElseIf textPrint = "2" Then '大->台
'            NowPrint m_CP09, "02", "02", True, strUserNum, 0
            bolEdit = True 'Add By Sindy 2012/1/12
            ET03 = "02" 'Modify By Sindy 2012/1/12
        End If
   End If
   
   'Add By Sindy 2012/1/12
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
            If ET03_1 <> "" Then
               NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
            End If
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            If ET03_1 <> "" Then
               NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            End If
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
         End If
      Else
         'Add By Sindy 2019/11/20 + 信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
         If ET03_1 <> "" Then
            'Add By Sindy 2019/11/20 + 信函總收文號
            NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
         End If
      End If
   'Add By Sindy 2021/1/5 沒有系統產出的定稿
   Else
      'Add By Sindy 2021/2/1 詢問有沒有客戶函
      If strLD18 <> "" Then
         Call PUB_TCaseAskIsPost_C(strLD18)
      End If
   '2021/1/5 EMD
   End If
   '2012/1/12 End
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCP05.Enabled = True Then
   Cancel = False
   textCP05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2009/09/03
If Me.textCP64.Enabled = True Then
   Cancel = False
   textCP64_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   '列印定稿為3(英文)時, 申請國家為大陸(020)者, 商標局發函日一定要輸入
   If m_CP10 <> "308" And Trim(m_TM10) = "020" And Trim(textPrint.Text) = "3" Then
      If Trim(textCP64) = "" Then
         MsgBox "大陸案發函日不可空白！", vbExclamation, "資料檢核"
         textCP64.SetFocus
         Exit Function
      End If
   End If
End If

'Add By Sindy 2010/12/24
If Me.textCP30.Enabled = True Then
   Cancel = False
   textCP30_Validate Cancel
   If Cancel = True Then
      textCP30.SetFocus
      Exit Function
   End If
End If

If Me.textCP47.Enabled = True Then
   Cancel = False
   textCP47_Validate Cancel
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

'add by nickc 2006/07/24 若是有輸入，要檢查一下
If Me.textNP08.Visible = True Then
   Cancel = False
   textNP08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textNP09.Visible = True Then
   Cancel = False
   textNP09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Trim(textNP08) <> "" Or Trim(textNP09) <> "" Then
    If Trim(textNP08) = "" Then
        MsgBox "子案新法定期限有輸入，所以子案新本所期限不許空白！", vbExclamation, "資料檢核"
        textNP08.SetFocus
        Exit Function
    End If
    If Trim(textNP09) = "" Then
        MsgBox "子案新本所期限有輸入，所以子案新法定期限不許空白！", vbExclamation, "資料檢核"
        textNP09.SetFocus
        Exit Function
    End If
    If Not Val(textNP08) <= Val(textNP09) Then
        MsgBox "子案新本所期限應小於等於新法定期限！", vbExclamation, "資料檢核"
        textNP08.SetFocus
        Exit Function
    End If
End If
TxtValidate = True
End Function

'Add By Sindy 2012/5/18
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
      If textNP08.Enabled = True Then textNP08.SetFocus
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
   '   If textNP08.Enabled = True Then textNP08.SetFocus
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
      If textNP08.Enabled = True Then textNP08.SetFocus
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
               textNP09 = Text12
               'Modify By Sindy 2014/10/6 台灣案之本所期限設定
               If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                  textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
               Else
               '2014/10/6 END
                  textNP08 = TransDate(CompDate(2, -2, TransDate(textNP09, 2)), 1)
               End If
               textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
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
   
   strFromDate = DBDATE(textCP05)
   
   If m_TM10 = 台灣國家代號 Then
      '文到天數
      If Option4(0).Value = True Then
         textNP09 = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
         If Option1(0).Value = True Then textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
         If Val(Text10) >= 60 Then
            i = -4
         Else
            i = -2
         End If
      '文到月數
      ElseIf Option4(1).Value = True Then
         textNP09 = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
         If Option1(0).Value = True Then textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
      
      If textNP09 <> "" Then
         'Modify By Sindy 2014/10/6 台灣案之本所期限設定
         If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
            textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
         Else
         '2014/10/6 END
            textNP08 = TransDate(CompDate(2, i, TransDate(textNP09, 2)), 1)
         End If
      End If
      textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
   End If
End Sub

'讀取來函期限
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
Dim strFromDate As String '期限起算日
   
   strFromDate = DBDATE(textCP05)
   
   ChgType = False
   If m_TM10 = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   
   ' 案件性質
   strRvType = LabNP07.Caption '202.申請意見書
   If strRvType = "" Then Exit Function
   
   If ClsPDGetCaseProperty(m_TM01, strRvType, strTempName, bolTmp) Then
      textNP08 = ""
      textNP09 = ""
      
      If m_TM10 = 台灣國家代號 Then
         strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & strRvType & "'"
         If strExc(0) <> "" Then
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            With RsTemp
               If intI = 1 Then
                  If Not IsNull(.Fields(1)) Then
                     '文到天數
                     Option4(0).Value = True
                     Text10 = .Fields(1)
                     textNP09 = TransDate(CompDate(2, Text10, TransDate(strFromDate, 2)), 1)
                  ElseIf Not IsNull(.Fields(2)) Then
                     '文到月數
                     Option4(1).Value = True
                     Text11 = .Fields(2)
                     textNP09 = TransDate(CompDate(1, .Fields(2), TransDate(strFromDate, 2)), 1)
                  Else
                     '文到天數
                     Option4(0).Value = True
                     Text10 = ""
                     Text11 = ""
                  End If
                  If textNP09 <> "" And Not IsNull(.Fields(0)) Then
                     '文到當日
                     If .Fields(0) = "1" Then
                        Option1(0).Value = True
                        textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
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
                  If textNP09 <> "" Then
                     'Modify By Sindy 2014/10/6 台灣案之本所期限設定
                     If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                        textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
                     Else
                     '2014/10/6 END
                        textNP08 = TransDate(CompDate(2, i, TransDate(textNP09, 2)), 1)
                     End If
                  End If
                   textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
               End If
            End With
         End If
      End If
      ChgType = True
   End If
End Function

'Added by Morgan 2025/9/12
'電子公文的分割核准輸入子案申請號若同時有發核准審定書，要一併更新該電子公文的本所案號及繳費單的檔名(User沒權限,先不做)
Private Sub UpdateT308EDoc()
   Dim stSQL As String, intR As Integer
'   Dim rsQuery As ADODB.Recordset
'   Dim arrFile() As String
'   Dim strFile As String, strNewFile As String

   If m_DocNo <> "" And m_CP10 = "308" And textCP30 <> textCP30.Tag Then
'      stSQL = "select * from edocument where ed02='" & textCP30 & "' and ed10='T' and ed11='C' and ed27 is null"
'      intR = 1
'      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
'      If intR = 1 Then
'         With rsQuery
'         Do While Not .EOF
'            If Not IsNull(.Fields("ed09")) Then
'               If Dir(strTFeeForm, vbDirectory) <> "" Then
'                  arrFile() = Split(.Fields("ed09"), ";")
'                  For intR = LBound(arrFile) To UBound(arrFile)
'                     If arrFile(intR) <> "" Then
'                        strFile = strTFeeForm & "\" & .Fields("ed02") & "." & arrFile(intR)
'                        If Dir(strFile) <> "" Then
'                           strNewFile = strTFeeForm & "\" & m_TM01 & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & "." & arrFile(intR)
'                           If Dir(strNewFile) = "" Then
'                              Name strFile As strNewFile
'                           End If
'                        End If
'                     End If
'                  Next
'               End If
'            End If
'            .MoveNext
'         Loop
'         End With
         stSQL = "update edocument set ed27='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "' where ed02='" & textCP30 & "' and ed10='T' and ed11='C' and ed27 is null"
         cnnConnection.Execute stSQL, intR
'      End If
   End If
'   Set rsQuery = Nothing
End Sub
