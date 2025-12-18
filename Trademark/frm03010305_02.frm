VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03010305_02 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標案被禁止處分"
   ClientHeight    =   5250
   ClientLeft      =   -3525
   ClientTop       =   3810
   ClientWidth     =   9135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9135
   Begin VB.TextBox textCP26 
      Height          =   285
      Left            =   5940
      MaxLength       =   1
      TabIndex        =   8
      Top             =   4124
      Width           =   372
   End
   Begin VB.TextBox textCP48 
      Height          =   285
      Left            =   5940
      MaxLength       =   8
      TabIndex        =   6
      Top             =   3764
      Width           =   2172
   End
   Begin VB.TextBox textCP14 
      Height          =   285
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   5
      Top             =   3764
      Width           =   732
   End
   Begin VB.TextBox textCP08 
      Height          =   285
      Left            =   1320
      MaxLength       =   40
      TabIndex        =   4
      Top             =   3420
      Width           =   2532
   End
   Begin VB.TextBox textCPSel 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   1980
      Width           =   372
   End
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4124
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6912
      TabIndex        =   11
      Top             =   60
      Width           =   1152
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5940
      TabIndex        =   10
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   12
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1470
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1620
      Width           =   2385
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5580
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   540
      Width           =   2532
   End
   Begin MSForms.TextBox textCP42 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   3044
      Width           =   7605
      VariousPropertyBits=   671105051
      MaxLength       =   600
      Size            =   "13414;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP41 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   2684
      Width           =   7605
      VariousPropertyBits=   671105051
      MaxLength       =   600
      Size            =   "13414;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP40 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   2324
      Width           =   7605
      VariousPropertyBits=   671105051
      MaxLength       =   600
      Size            =   "13414;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   615
      Left            =   1320
      TabIndex        =   9
      Top             =   4500
      Width           =   7605
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13414;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_2 
      Height          =   285
      Left            =   2160
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3764
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
      Left            =   1320
      TabIndex        =   38
      Top             =   900
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
      Left            =   1320
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1260
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
      Left            =   5580
      TabIndex        =   36
      Top             =   1620
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
      Left            =   3870
      TabIndex        =   35
      Top             =   570
      Width           =   645
   End
   Begin VB.Label Label16 
      Caption         =   "是否算案件數 :"
      Height          =   252
      Left            =   4620
      TabIndex        =   34
      Top             =   4140
      Width           =   1212
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不算)"
      Height          =   252
      Left            =   6660
      TabIndex        =   33
      Top             =   4140
      Width           =   972
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   252
      Left            =   4620
      TabIndex        =   32
      Top             =   3780
      Width           =   852
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   31
      Top             =   3780
      Width           =   852
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   30
      Top             =   3420
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "(1:被禁止處分  2:取消被禁止處分)"
      Height          =   252
      Left            =   1980
      TabIndex        =   29
      Top             =   1980
      Width           =   2772
   End
   Begin VB.Label Label1 
      Caption         =   "對造名稱(日) :"
      Height          =   252
      Index           =   4
      Left            =   120
      TabIndex        =   28
      Top             =   3060
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "對造名稱(英) :"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   27
      Top             =   2700
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "對造名稱(中) :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   2340
      Width           =   1332
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   120
      TabIndex        =   25
      Top             =   4500
      Width           =   972
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   120
      TabIndex        =   24
      Top             =   4140
      Width           =   972
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   252
      Left            =   2160
      TabIndex        =   23
      Top             =   4140
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4620
      TabIndex        =   22
      Top             =   1620
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   21
      Top             =   1620
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   1980
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   19
      Top             =   1260
      Width           =   1092
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   18
      Top             =   900
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   540
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   252
      Index           =   7
      Left            =   4620
      TabIndex        =   16
      Top             =   540
      Width           =   852
   End
End
Attribute VB_Name = "frm03010305_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/13 改成Form2.0 ;textTM23、cmbTM05、textCP13、textCP14_2、textCP64、textCP40、textCP41、textCP42
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
' 原業務區
Dim m_CP12 As String
' 原智權人員代號
Dim m_CP13 As String
' 國家代碼
Dim m_TM10 As String

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm03010305_01.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03010305_01
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
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      Unload Me
      frm03010305_01.Show
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   
   textCP05S.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
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
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("TM29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Public Sub QueryData()
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   '讀取基本檔
   Select Case m_TM01
      '系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
   End Select
   
   '以本所案號取得案件進度檔中最後一筆A類收文資料之智權人員
   'Modify By Sindy 2018/5/15 + AND CP10 not in('107','001') : 踢除107.跨類 001.查名
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' AND " & _
                  "CP09 LIKE 'A%' AND CP10 not in('107','001') AND " & _
                  "CP05 IN (SELECT MAX(CP05) FROM CaseProgress " & _
                           "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                 "CP02 = '" & m_TM02 & "' AND " & _
                                 "CP03 = '" & m_TM03 & "' AND " & _
                                 "CP04 = '" & m_TM04 & "' AND " & _
                                 "CP09 LIKE 'A%' AND CP10 not in('107','001')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True) 'Modified by Lydia 2016/03/25 離職人員也顯示
      End If
   End If
   rsTmp.Close
   'Added by Lydia 2016/03/11 CFT承辦人
   'Modified by Lydia 2016/03/25 全部套用
   'If m_TM01 = "CFT" Then
       Dim strNA69 As String
       'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
       Call GetNA69("", m_TM10, m_CP13, strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
       textCP14 = strNA69
       textCP14_2 = GetStaffName(textCP14)
   'End If
   'end 2016/03/11
   'end 2016/03/25

   
   ' 以下一程序代號計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''   strDay = Empty
   Select Case textCPSel
      Case "1":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1614")
            textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1614", DBDATE(m_CP05))
      Case "2":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1615")
            textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1615", DBDATE(m_CP05))
   End Select
''''   If IsEmptyText(strDay) = False Then
''''      ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''      'textCP48 = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''      textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''   End If
   
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
   Dim strCP27 As String
   
 '911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans
   
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質為(被禁止處分或取消被禁止處分)
   strCP10 = "1614"
   If textCPSel = "2" Then: strCP10 = "1615"
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   '92.6.14 MODIFY BY SONIA
   ' 發文日
   'strCP27 = DBDATE(Date)
   strCP27 = DBDATE(SystemDate())
   '92.6.14 END
   ' 新增案件進度資料
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
   '92.6.14 MODIFY BY SONIA 加發文日
   'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP40,CP41,CP42,CP48,CP64) " & _
   '         "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
   '                 "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
   '                 "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
   '                 "'" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & DBDATE(textCP48) & ",'" & ChgSQL(textCP64) & "') "
    'Modify By Cheng 2004/02/04
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP40,CP41,CP42,CP48,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
'                    "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "'," & _
'                    "'" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & DBDATE(textCP48) & ",'" & ChgSQL(textCP64) & "') "
'Modified by Lydia 2018/02/21 因為以案件代號計算承辦期限,所以承辦期限有可能空白DBDATE(textCP48)=> CNULL(DBDATE(textCP48),True)
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP40,CP41,CP42,CP48,CP64) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
                    "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "'," & _
                    "'" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & CNULL(DBDATE(textCP48), True) & ",'" & ChgSQL(textCP64) & "') "
    'End
   '92.6.14  END
   cnnConnection.Execute strSql
   
   'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
   Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
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
   Set frm03010305_02 = Nothing
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
         strMsg = "承辦期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
      End If
   End If
End Sub

' 案件性質
Private Sub textCPSel_Validate(Cancel As Boolean)
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCPSel) = False Then
      Select Case textCPSel
         Case "1", "2":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入1或2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCPSel_GotFocus
            GoTo EXITSUB
      End Select
      
      ' 以案件代號計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''      strDay = Empty
      Select Case textCPSel
         Case "1":
''''            strDay = GetWorkDays(m_TM01, m_TM10, "1614")
                textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1614", DBDATE(m_CP05))
         Case "2":
''''            strDay = GetWorkDays(m_TM01, m_TM10, "1615")
                textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1615", DBDATE(m_CP05))
      End Select
''''      If IsEmptyText(strDay) = False Then
''''         ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''         'textCP48 = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''      End If
   End If
EXITSUB:
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
   
   ' 案件性質不可為空白
   If IsEmptyText(textCPSel) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入案件性質"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCPSel.SetFocus
      GoTo EXITSUB
   End If
   
   ' 承辦期限不可為空白
   If IsEmptyText(textCP48) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入承辦期限"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP48.SetFocus
      GoTo EXITSUB
   End If
   
   ' 對造名稱不可同時為空白
   If IsEmptyText(textCP40) = True And IsEmptyText(textCP41) = True And IsEmptyText(textCP42) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入對造名稱"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP40.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCPSel_GotFocus()
   InverseTextBox textCPSel
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

Private Sub textCP08_GotFocus()
   InverseTextBox textCP08
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
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

If Me.textCP48.Enabled = True Then
   Cancel = False
   textCP48_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCPSel.Enabled = True Then
   Cancel = False
   textCPSel_Validate Cancel
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

