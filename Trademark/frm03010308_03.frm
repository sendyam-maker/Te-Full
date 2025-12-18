VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03010308_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務結果輸入"
   ClientHeight    =   4290
   ClientLeft      =   -2840
   ClientTop       =   4290
   ClientWidth     =   9160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   9160
   Begin VB.TextBox textCP16 
      Height          =   285
      Left            =   5820
      TabIndex        =   1
      Top             =   2760
      Width           =   1125
   End
   Begin VB.TextBox textCP18 
      Height          =   285
      Left            =   8040
      TabIndex        =   2
      Top             =   2760
      Width           =   945
   End
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   3
      Top             =   3120
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6900
      TabIndex        =   20
      Top             =   60
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5940
      TabIndex        =   18
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8160
      TabIndex        =   19
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textCP08 
      Height          =   285
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   0
      Top             =   2760
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   8
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2385
   End
   Begin MSForms.TextBox textCP64 
      Height          =   705
      Left            =   1260
      TabIndex        =   4
      Top             =   3480
      Width           =   7725
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13626;1244"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5760
      TabIndex        =   32
      Top             =   2024
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   5760
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2384
      Width           =   2532
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1260
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1304
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
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1260
      TabIndex        =   29
      Top             =   960
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
      Left            =   3930
      TabIndex        =   28
      Top             =   630
      Width           =   645
   End
   Begin VB.Label Label11 
      Caption         =   "費用 :"
      Height          =   255
      Left            =   4800
      TabIndex        =   27
      Top             =   2760
      Width           =   915
   End
   Begin VB.Label Label12 
      Caption         =   "點數 :"
      Height          =   255
      Left            =   7170
      TabIndex        =   26
      Top             =   2775
      Width           =   915
   End
   Begin VB.Label Label21 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   180
      TabIndex        =   25
      Top             =   3480
      Width           =   972
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   252
      Left            =   180
      TabIndex        =   24
      Top             =   3120
      Width           =   972
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   252
      Left            =   2100
      TabIndex        =   23
      Top             =   3120
      Width           =   852
   End
   Begin VB.Label Label8 
      Caption         =   "機關文號 :"
      Height          =   252
      Left            =   180
      TabIndex        =   22
      Top             =   2760
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人 :"
      Height          =   252
      Index           =   2
      Left            =   4800
      TabIndex        =   21
      Top             =   2400
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   180
      TabIndex        =   17
      Top             =   1680
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   16
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   180
      TabIndex        =   15
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   180
      TabIndex        =   14
      Top             =   1320
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   3
      Left            =   4800
      TabIndex        =   13
      Top             =   1680
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   180
      TabIndex        =   12
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   180
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4800
      TabIndex        =   10
      Top             =   2040
      Width           =   972
   End
End
Attribute VB_Name = "frm03010308_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/13 改成Form2.0 ;textTM23、cmbTM05、textCP13、textCP14、textCP64
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
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
' 原業務區
Dim m_CP12 As String
' 原智權人員代號
Dim m_CP13 As String
' 原承辦人員代號
Dim m_CP14 As String
' 國家代碼
Dim m_TM10 As String
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
   frm03010308_02.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03010308_02
   Unload frm03010308_01
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
          'add by nickc 2005/04/22
          Pub_EndModCashMsg m_TM10
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 存檔
      'edit by  nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Add By Sindy 2023/4/27
      If Me.m_strIR01 <> "" Then
         Unload frm03010308_02
         Unload frm03010308_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
      Else
      '2023/4/27 END
         Unload Me
         Unload frm03010308_02
         frm03010308_01.Show
      End If
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2023/4/27
   m_strIR01 = frm03010308_02.m_strIR01
   m_strIR02 = frm03010308_02.m_strIR02
   m_strIR03 = frm03010308_02.m_strIR03
   m_strIR04 = frm03010308_02.m_strIR04
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

' 取得服務業務基本檔的欄位內容
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
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"))
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
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
         textCP05 = rsTmp.Fields("CP05")
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
      ' 承辦人
      m_CP14 = ""
      'Added by Lydia 2016/03/11 CFT改成模組判斷
      'Modified by Lydia 2016/03/25 全部套用
      'If m_TM01 = "CFT" Then
         Dim strNA69 As String
         'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
         Call GetNA69("", m_TM10, "" & rsTmp.Fields("CP13"), strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
         m_CP14 = strNA69
         textCP14 = GetStaffName(m_CP14)
'      Else
'      'end 2016/03/11
'        If IsNull(rsTmp.Fields("CP14")) = False Then
'           textCP14 = GetStaffName(rsTmp.Fields("CP14"))
'           m_CP14 = rsTmp.Fields("CP14") '2009/6/18 ADD BY SONIA
'        End If
'      End If
      'end 2016/03/25
      
      ' 業務區
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True) 'Modified by Lydia 2016/03/25 離職人員也顯示
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   
   ' 本所案號
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   
   ' 讀取基本檔
   QueryServicePractice
   
   ' 讀取案件進度檔
   QueryCaseProgress
   
   Set rsTmp = Nothing
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP27 As String
    'Add By Cheng 2002/12/02
    Dim strCP17 As String
   
 '911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質為服務業務結果
   strCP10 = "1801"
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   ' 發文日
   strCP27 = DBDATE(SystemDate())
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
'                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
'                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2004/02/04
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP17,CP18,CP20,CP26,CP27,CP32,CP43,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & Replace(CNULL(Me.textCP16.Text), "'", "") & "," & _
'                     Replace(CNULL(strCP17), "'", "") & "," & Replace(CNULL(Me.textCP18.Text), "'", "") & ",'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
   '2009/6/18 MODIFY BY SONIA承辦人改點選收文號之原承辦人
   'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP17,CP18,CP20,CP26,CP27,CP32,CP43,CP64) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & StrCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & Replace(CNULL(Me.textCP16.Text), "'", "") & "," & _
                     Replace(CNULL(strCP17), "'", "") & "," & Replace(CNULL(Me.textCP18.Text), "'", "") & ",'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP16,CP17,CP18,CP20,CP26,CP27,CP32,CP43,CP64) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & m_CP14 & "'," & Replace(CNULL(Me.textCP16.Text), "'", "") & "," & _
                     Replace(CNULL(strCP17), "'", "") & "," & Replace(CNULL(Me.textCP18.Text), "'", "") & ",'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
   '2009/6/18 END
    'End
   cnnConnection.Execute strSql
   
   'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
   Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '  更新原案件進度資料的實際結果為1
   strSql = "UPDATE CaseProgress SET CP24 = '1' " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   cnnConnection.Execute strSql
   
   '2008/12/29 add by sonia
   ' 更新下一程序檔案件性質為催審的資料
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 = " & "305"
   cnnConnection.Execute strSql
   '2008/12/29 end
   
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
   
   'add by nickc 2005/04/22
   Pub_UpdateEndModCash m_TM01, m_TM02, m_TM03, m_TM04
   
   'Add by Sindy 2023/4/27
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm03010308_01", strCP09
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
   Set frm03010308_03 = Nothing
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
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCP08_GotFocus()
   InverseTextBox textCP08
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
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

If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
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

'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If

TxtValidate = True
End Function

