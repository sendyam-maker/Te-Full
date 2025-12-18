VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010601_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人已收達/已提申"
   ClientHeight    =   4608
   ClientLeft      =   -3384
   ClientTop       =   3636
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4608
   ScaleWidth      =   9360
   Begin VB.TextBox textCP46 
      Height          =   285
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   0
      Top             =   3120
      Width           =   2052
   End
   Begin VB.CommandButton Command2 
      Caption         =   "商品及服務(&I)"
      Height          =   400
      Left            =   4470
      TabIndex        =   10
      Top             =   60
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   4800
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   3465
      Begin VB.TextBox textTM12_2 
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   3
         Top             =   0
         Width           =   2292
      End
      Begin VB.Label Label1 
         Caption         =   "申請案號 :"
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2532
   End
   Begin VB.TextBox textTM11 
      Height          =   285
      Left            =   5760
      MaxLength       =   8
      TabIndex        =   1
      Top             =   3120
      Width           =   1185
   End
   Begin VB.CheckBox Check1 
      Caption         =   "申請收據"
      Height          =   255
      Index           =   1
      Left            =   6510
      TabIndex        =   7
      Top             =   4200
      Width           =   1065
   End
   Begin VB.CheckBox Check1 
      Caption         =   "申請書"
      Height          =   255
      Index           =   0
      Left            =   5430
      TabIndex        =   6
      Top             =   4200
      Width           =   1065
   End
   Begin VB.TextBox textCP47 
      Height          =   285
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3480
      Width           =   2052
   End
   Begin VB.TextBox textCP45 
      Height          =   285
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3840
      Width           =   2052
   End
   Begin VB.TextBox textCP27 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   5
      Top             =   4200
      Width           =   372
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6991
      TabIndex        =   9
      Top             =   60
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5942
      TabIndex        =   8
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8280
      TabIndex        =   11
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   960
      Width           =   2532
   End
   Begin MSForms.TextBox textCP44 
      Height          =   285
      Left            =   1200
      TabIndex        =   43
      Top             =   2400
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
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1200
      TabIndex        =   42
      Top             =   2040
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
      Left            =   1170
      TabIndex        =   41
      Top             =   1320
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
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5730
      TabIndex        =   40
      Top             =   1680
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
   Begin VB.Label Label9 
      Caption         =   "商品類別 :"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   2760
      Width           =   975
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
      TabIndex        =   36
      Top             =   630
      Width           =   645
   End
   Begin VB.Label Label8 
      Caption         =   "申請日 :"
      Height          =   255
      Left            =   4800
      TabIndex        =   35
      Top             =   3120
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   "附件 :"
      Height          =   255
      Left            =   4800
      TabIndex        =   34
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "彼所案號 :"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "代理人提申日 :"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "代理人 :"
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   31
      Top             =   2400
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人 :"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   2040
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "發文日 :"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   28
      Top             =   2040
      Width           =   885
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印)"
      Height          =   180
      Left            =   1680
      TabIndex        =   26
      Top             =   4200
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4800
      TabIndex        =   24
      Top             =   1680
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   255
      Index           =   8
      Left            =   4800
      TabIndex        =   23
      Top             =   2400
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   972
   End
   Begin VB.Label Label11 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   "代理人收達日 :"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "frm02010601_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/16 改成Form2.0 ;textTM23、cmbTM05、textCP13、textCP14、textCP44
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
' 來函收文日
Dim m_CP05 As String
' 機關文號
Dim m_CP08 As String
' 所選取的收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
'Add By Sindy 2009/09/24
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
' 申請人
Dim m_TM23 As String
' 代理人
Dim m_CP44 As String
'Add By Cheng 2004/05/13
Dim m_TM11 As String '申請日
'End
'Add By Sindy 2009/06/25
Dim m_CP54 As String
Dim m_TM22 As String
Dim m_CP43 As String '相關總收文號 Add By Sindy 2010/10/5
Dim m_CP47 As String 'Add By Sindy 2010/12/23
Dim m_strLanguage As String 'Add By Sindy 2012/4/26 定稿語文
Dim m_CP27 As String 'add by sonia 2018/2/6
Dim m_CP07 As String 'add by sonia 2022/1/21
Public ChkTG As Boolean '檢查是否已經有商品及服務 Add By Sindy 2018/1/12
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END
Dim m_CP14 As String
Dim strLD18 As String 'Add By Sindy 2019/12/25 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2019/12/25 FC代理人


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010601_02.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm02010601_02
   Unload frm02010601_01
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
'        'OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
        'Modify By Cheng 2003/01/07
        '若為已提申, 才有定稿
        If frm02010601_02.textResult.Text = "2" Then
            'Add By Cheng 2002/11/08
            ' 列印定稿
            If textPrint <> "N" Then
               PrintLetter
            End If
        End If
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010601_02
      'Add By Sindy 2019/5/22
      If Me.m_strIR01 <> "" Then
         Unload frm02010601_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
      '2019/5/22 END
      Else
         frm02010601_01.Show
      End If
      Unload Me
   End If
End Sub

'Add By Sindy 2018/1/12
Private Sub Command2_Click()
   frm03010303_04.Hide
   Set frm03010303_04.UpForm = Me
   frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   frm03010303_04.AllClass = textTM09.Text
   frm03010303_04.cmdOK(2).Visible = True
   
'   frm03010303_04.Label2.Visible = False
'   frm03010303_04.cmdok(0).Visible = False
'   frm03010303_04.cmdok(2).Visible = False
'   frm03010303_04.cmd.Visible = False
'   frm03010303_04.cmd2.Visible = False
'   frm03010303_04.txt2(0).Visible = False
'   frm03010303_04.txt2(1).Visible = False
'   frm03010303_04.txt2(2).Visible = False
'   frm03010303_04.txt2(3).Visible = False
'   frm03010303_04.Line1.Visible = False
   
   If textTM09 <> "" Then '有商品類別才可進入 T-113511團體標章
      Me.Hide
      frm03010303_04.QueryData
      frm03010303_04.Show vbModal '強制回應表單
   Else
      MsgBox ("無商品類別，不可使用此按鈕 !")
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   
   'Add By Sindy 2009/05/12
   textTM09.BackColor = &H8000000F
   
   textTM15.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP27.BackColor = &H8000000F
   textCP44.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   'Add By Sindy 2018/1/12
   If frm02010601_02.textResult.Text = "2" Then '2:已提申
      Command2.Visible = True
   Else
      Command2.Visible = False
   End If
   '2018/1/12 END
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010601_01.m_strIR01
   m_strIR02 = frm02010601_01.m_strIR02
   m_strIR03 = frm02010601_01.m_strIR03
   m_strIR04 = frm02010601_01.m_strIR04
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
        'Add By Cheng 2004/05/13
        ' 申請日
        If IsNull(rsTmp.Fields("TM11")) = False Then
            Select Case m_TM01
            Case "CFT", "CFC", "S":
                textTM11 = DBDATE(rsTmp.Fields("TM11"))
            Case Else:
                textTM11 = TAIWANDATE(rsTmp.Fields("TM11"))
            End Select
        End If
        m_TM11 = Me.textTM11.Text
        'End
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
         textTM12_2 = rsTmp.Fields("TM12") 'Add By Sindy 2016/10/14
      End If
      
      'Add By Sindy 2009/05/12
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
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
         m_TM23 = rsTmp.Fields("TM23")
      End If
      
      'Add By Sindy 2019/12/25
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("TM44")) = False Then
         m_TM44 = rsTmp.Fields("TM44")
      End If
      '2019/12/25 END
      
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
      
      'Add By Sindy 2009/06/25
      ' 專用期間(迄)
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_TM22 = DBDATE(rsTmp.Fields("TM22"))
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
      ' 申請日
      'modify by sonia 2023/11/23
      'm_TM11 = ""
      If IsNull(rsTmp.Fields("SP10")) = False Then
          Select Case m_TM01
          Case "CFT", "CFC", "S":
              textTM11 = DBDATE(rsTmp.Fields("SP10"))
          Case Else:
              textTM11 = TAIWANDATE(rsTmp.Fields("SP10"))
          End Select
      End If
      m_TM11 = Me.textTM11.Text
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
         textTM12_2 = rsTmp.Fields("SP11")
      End If
      ' 審定號
      If IsNull(rsTmp.Fields("SP14")) = False Then
         textTM15 = rsTmp.Fields("SP14")
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("SP73")) = False Then
         textTM09 = rsTmp.Fields("SP73")
      End If
      'end 2023/11/23
      ' 申請人
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
'         If m_CP10 = "000" Then
'            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
'         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
'         End If
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True)
      End If
      '業務區
      m_CP12 = Empty
      If IsNull(rsTmp.Fields("cp12")) = False Then
          m_CP12 = rsTmp.Fields("cp12")
      End If
      ' 承辦人員
      m_CP14 = Empty
      If IsNull(rsTmp.Fields("CP14")) = False Then
         m_CP14 = rsTmp.Fields("CP14")
         textCP14 = GetStaffName(rsTmp.Fields("CP14"), True)
      End If
      'Add By Sindy 2010/10/5
      '相關總收文號
      m_CP43 = ""
      If IsNull(rsTmp.Fields("CP43")) = False Then
          m_CP43 = rsTmp.Fields("CP43")
      End If
      '2010/10/5
      ' 發文日
      If IsNull(rsTmp.Fields("CP27")) = False Then
         m_CP27 = "" & rsTmp.Fields("CP27")   'add by sonia 2018/2/6
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
'         textCP44 = GetFAgentName(rsTmp.Fields("CP44"))
         If PUB_GetAgentName(m_TM01, rsTmp.Fields("CP44").Value, strTempName) Then
            textCP44 = strTempName
         Else
            textCP44 = ""
         End If
      End If
      ' 彼所案號
      If IsNull(rsTmp.Fields("CP45")) = False Then
         textCP45 = rsTmp.Fields("CP45")
      End If
      ' 代理人收達日
      If IsNull(rsTmp.Fields("CP46")) = False Then
         Select Case m_TM01
            Case "CFT", "CFC", "S":
               textCP46 = DBDATE(rsTmp.Fields("CP46"))
            Case Else:
               textCP46 = TAIWANDATE(rsTmp.Fields("CP46"))
         End Select
      End If
      ' 代理人提申日
      m_CP47 = "" & rsTmp.Fields("CP47") 'Add By Sindy 2010/12/23
      'Modify By Sindy 2010/12/23
      '若操作人員為內商(部門P2X)則不帶出cp47
      If Left(GetStaffDepartment(strUserNum), 2) <> "P2" Then
      '2010/12/23 End
         If IsNull(rsTmp.Fields("CP47")) = False Then
            Select Case m_TM01
               Case "CFT", "CFC", "S":
                  textCP47 = DBDATE(rsTmp.Fields("CP47"))
               Case Else:
                  textCP47 = TAIWANDATE(rsTmp.Fields("CP47"))
            End Select
         End If
      End If
      
      'Add By Sindy 2009/06/25
      '授權期間(迄)
      m_CP54 = ""
      If IsNull(rsTmp.Fields("CP54")) = False Then
         m_CP54 = DBDATE(rsTmp.Fields("CP54"))
      End If
      'add by sonia 2022/1/21 原法定期限(延展原專用期止日)
      m_CP07 = ""
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = DBDATE(rsTmp.Fields("CP07"))
      End If
      'end 2022/1/21
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   'Add By Cheng 2002/07/18
   m_TM23 = Empty
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         ' 取得商標基本檔的相關項目
         QueryTradeMark
      Case Else
         QueryServicePractice
   End Select
   QueryCaseProgress
   ' 若前畫面輸入的是已收達時, 將代理人提申日Disable
   If frm02010601_02.GetTextResult = "1" Then
        EnableTextBox textCP47, False
        EnableTextBox textTM11, False
        Me.textTM11.Visible = False
        Me.Label8.Visible = False
   Else
        EnableTextBox textCP47, True
        'Add By Cheng 2003/07/22
        '游標預設在代理人提申日
        SendKeys "{Tab}"
        If m_TM01 = "TF" And m_CP10 = "101" Then
            EnableTextBox textTM11, True
            Me.textTM11.Visible = True
            Me.Label8.Visible = True
        Else
            EnableTextBox textTM11, False
            Me.textTM11.Visible = False
            Me.Label8.Visible = False
        End If
        'Add by Sindy 2016/10/14
        If m_CP10 = "101" Or m_CP10 = "308" Then
            Frame1.Visible = True
        End If
        '2016/10/14 END
   End If
   ' 若彼所案為空白時, 則以此本所案號同代理人找出以前收文資料的彼所案號
   If IsEmptyText(textCP45) = True And IsEmptyText(m_CP44) = False Then
      strSql = "SELECT CP45 FROM CaseProgress " & _
               "WHERE CP01 = '" & m_TM01 & "' AND " & _
                     "CP02 = '" & m_TM02 & "' AND " & _
                     "CP03 = '" & m_TM03 & "' AND " & _
                     "CP04 = '" & m_TM04 & "' AND " & _
                     "CP09 <> '" & m_CP09 & "' AND " & _
                     "CP44 = '" & m_CP44 & "' AND " & _
                     "CP45 <> NULL "
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            If IsNull(rsTmp.Fields("CP45")) = False Then
               If IsEmptyText(rsTmp.Fields("CP45")) = False Then
                  textCP45 = rsTmp.Fields("CP45")
                  Exit Do
               End If
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
   '91.7.14 modify by sonia
   If frm02010601_02.textResult = "1" Then
      textPrint = "N"
   Else
        'Modify By Cheng 2003/01/08
'      '91.10.28 MODIFY BY SONIA
'      If m_TM01 = "T" Then
'         textPrint = "N"
'      Else
'         textPrint = ""
'      End If
'      '91.10.28 END
   End If
   '91.7.14 end
End Sub

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
Dim strSql As String
'Add By Sindy 2009/06/25
Dim strNP07 As String
Dim strNP09 As String
Dim strNP08 As String
Dim strNA78 As String    'add by sonia 2015/12/4
Dim rsTmp As New ADODB.Recordset
Dim strCP09 As String 'Add By Sindy 2019/5/22
Dim strCP10 As String 'Add By Sindy 2019/12/25
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   ' 更新彼所案號
   strSql = "UPDATE CaseProgress SET CP45 = '" & textCP45 & "'"
   ' 更新代理人收達日
   If IsEmptyText(textCP46) = False Then
      strSql = strSql & ",CP46 = " & DBDATE(textCP46)
   End If
   ' 更新代理人提申日
   If IsEmptyText(textCP47) = False Then
      strSql = strSql & ",CP47=" & DBDATE(textCP47)
   End If
   strSql = strSql & " WHERE CP09 = '" & m_CP09 & "' "
   cnnConnection.Execute strSql
   strCP09 = m_CP09 'Add By Sindy 2019/5/22
   
   'add by nick 2005/01/04 更新相同本所案號之相同代理人的彼所案號，若是彼所案號空的話
   If textCP45 <> "" Then
      strSql = "update caseprogress set cp45=" & CNULL(ChgSQL(textCP45)) & " where cp09 in (select cp09 from caseprogress where cp45 is null and CP01 = '" & m_TM01 & "' AND  CP02 = '" & m_TM02 & "' AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' and cp09<'C' AND cp44 in (select cp44 from caseprogress where cp09='" & m_CP09 & "' ))"
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2009/06/25
   '有代理人提申日且案件性質為102.延展
   '更新基本檔的TM22.專用期間(迄)
   '新增下一程序102期限及105期限
   If IsEmptyText(textCP47) = False And m_CP10 = "102" Then
      'add by sonia 2023/11/23 TM案件要新服務業務
      If m_TM01 = "TM" Then
         strSql = "UPDATE servicepractice SET " & _
                              "SP21 = " & m_CP54 & " " & _
                              "WHERE " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
      Else
      'end 2023/11/23
         strSql = "UPDATE Trademark SET " & _
                              "TM22 = " & m_CP54 & " " & _
                              "WHERE " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
      End If
      cnnConnection.Execute strSql
      
      ' 新增延展記錄到下一程序檔
      strNP07 = "102"
      ' 法定期限為專用期限之截止日
      strNP09 = m_CP54
      ' 本所期限為法定期限-2天,TF為-1個月
      If m_TM01 = "TF" Then
         strNP08 = CompDate(1, -1, strNP09)
      Else
         'Modify By Sindy 2014/10/6 台灣案之本所期限設定
         'Modify By Sindy 2016/6/17 CFT為內對外案件,不會有申請國家為000台灣案
         'If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
'         If Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
'         '2016/6/17 END
'            strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
'         Else
'         '2014/10/6 END
            'Modify By Sindy 2016/6/17 非TF案本所期限=法定期限-2個月
            'strNP08 = CompDate(2, -2, strNP09)
            strNP08 = CompDate(1, -2, strNP09)
            '2016/6/17 END
'         End If
      End If
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      
      '2010/2/24 modify by sonia 沒有才新增
      If rsTmp.State <> adStateClosed Then rsTmp.Close
      Set rsTmp = Nothing
      'Modify By Sindy 2019/6/27
      'strSql = "select * from NEXTPROGRESS where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP09=" & strNP09 & " AND NP07='102' AND NP06 IS NULL"
      strSql = "select * from NEXTPROGRESS where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP07='102' AND NP06 IS NULL"
      '2019/6/27 END
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <> 0 Then
         '2013/2/25 同時更新業務員
         'Modify By Sindy 2019/6/27 + ,NP09=" & strNP09 & "
         strSql = "UPDATE NEXTPROGRESS SET NP08=" & strNP08 & ",NP09=" & strNP09 & ",NP10='" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "' where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP07='102' AND NP06 IS NULL"
      Else
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
      End If
      cnnConnection.Execute strSql
      '2010/2/24 END
      
'modify by sonia 2015/12/4 不限定申請國家,依國家設定為主故改寫
'      ' 申請國家為美國, 菲律賓, 柬埔寨
'      'modify by sonia 2014/10/31 加入莫三比克318
'      If (m_TM10 = "101" Or m_TM10 = "030" Or m_TM10 = "046" Or m_TM10 = "318") Then
'         ' 取得下次使用宣誓年度
'         strNA39 = 0
'         If rsTmp.State <> adStateClosed Then rsTmp.Close
'         Set rsTmp = Nothing
'         strSql = "SELECT * FROM Nation WHERE NA01 = '" & m_TM10 & "' AND NA39 IS NOT NULL "
'         rsTmp.CursorLocation = adUseClient
'         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsTmp.RecordCount > 0 Then
'            strNA39 = rsTmp.Fields("NA39")
'            strNP22 = GetNextProgressNo()
'            ' 案件性質為使用宣誓
'            strNP07 = "105"
'            ' 法定期限為該案件原來之專用期止日 + 下次使用宣誓年度
'            strNP09 = DBDATE(DateAdd("yyyy", Val(strNA39), ChangeWStringToWDateString(m_TM22)))
''Modify By Sindy 2012/3/16 業務說改成本所=法定-2個月 不管任何國家, 但此處沒改到
''            '若申請國家為"菲律賓"時, 本所期限 = 法定期限 - 半年
''            '                        其他國家則 本所期限 = 法定期限 - 1年
''            If m_TM10 = "030" Then
'               'strNP08 = DBDATE(DateAdd("m", -6, ChangeWStringToWDateString(DBDATE(strNP09))))
'               strNP08 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
''            Else
''               strNP08 = DBDATE(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(strNP09))))
''            End If
'            '2013/2/25 modify by sonia 先讀,不存在才可新增CFT-014924
'            'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                      "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                                strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
'            If rsTmp.State <> adStateClosed Then rsTmp.Close
'            Set rsTmp = Nothing
'            strSql = "select * from NEXTPROGRESS where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP09=" & strNP09 & " AND NP07='105' AND NP06 IS NULL"
'            rsTmp.CursorLocation = adUseClient
'            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsTmp.RecordCount <> 0 Then
'               strSql = "UPDATE NEXTPROGRESS SET NP08=" & strNP08 & ",NP10='" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "' where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP09=" & strNP09 & " AND NP07='105' AND NP06 IS NULL"
'            Else
'               strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                                   strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
'            End If
'            '2013/2/25 end
'            cnnConnection.Execute strSql
'         '若無下次使用宣誓年度則不新增下一程序檔
'         Else
'            'rsTmp.Close 'Modify By Sindy 2009/09/28
'         End If
'         rsTmp.Close
'      End If
      '取得國家之延展後使用宣誓年度NA78
      strNA78 = 0
      If rsTmp.State <> adStateClosed Then rsTmp.Close
      Set rsTmp = Nothing
      strSql = "SELECT * FROM Nation WHERE NA01 = '" & m_TM10 & "' AND NA78 IS NOT NULL "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         strNA78 = rsTmp.Fields("NA78")
         strNP07 = "105"
         '法定期限為該案件原來之專用期止日 + 延展後使用宣誓年度(柬埔寨以延展核准日計算故延展證書時再更新)
         'Modify By Sindy 2016/6/17 調整期限計算的函數 DBDATE ==> CompDate
         'strNP09 = DBDATE(DateAdd("yyyy", Val(strNA78), ChangeWStringToWDateString(m_TM22)))
         'modify by sonia 2022/1/21 m_TM22改用m_cp07
         strNP09 = CompDate(0, Val(strNA78), m_CP07)
         If m_TM10 = "110" Then strNP09 = CompDate(1, 3, Val(strNP09))   'add by sonia 2023/9/15 海地110法定期限要再加3個月CFT-023278
         '業務說改成本所=法定-2個月 不管任何國家
         'strNP08 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
         strNP08 = CompDate(1, -2, strNP09)
         '2016/6/17 END
         strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
         'Modify By Sindy 2019/6/27
         'strSql = "select * from NEXTPROGRESS where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP09=" & strNP09 & " AND NP07='105' AND NP06 IS NULL"
         'modify by sonia 2022/1/21 菲律賓原延展期限後一年有掛使用宣誓期限,不可蓋掉,故加入NP09>m_cp07+20000原延展法定期限2年的條件
         strSql = "select * from NEXTPROGRESS where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP07='105' AND NP06 IS NULL and NP09>" & m_CP07 + 20000
         '2019/6/27 END
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <> 0 Then
            'Modify By Sindy 2019/6/27 + ,NP09=" & strNP09 & "
            'modify by sonia 2022/1/21 菲律賓原延展期限後一年有掛使用宣誓期限,不可蓋掉,故加入NP09>m_cp07+20000原延展法定期限2年的條件
            strSql = "UPDATE NEXTPROGRESS SET NP08=" & strNP08 & ",NP09=" & strNP09 & ",NP10='" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "' where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP07='105' AND NP06 IS NULL and NP09>" & m_CP07 + 20000
         Else
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                      "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                                strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
         End If
         cnnConnection.Execute strSql
      End If
      rsTmp.Close
'end 2015/12/4

   End If
   '2009/06/25 End
   
   ' 若有輸入代理人收達日
   If IsEmptyText(textCP46) = False Then
      strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP07 = 997 "
      cnnConnection.Execute strSql
   End If
   ' 若有輸入代理人提申日
   If IsEmptyText(textCP47) = False Then
      '2005/8/26 MODIFY BY SONIA
      'strSQL = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
      '         "WHERE NP01 = '" & m_CP09 & "' AND " & _
      '               "NP07 = 998 "
      strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP07 IN (997,998) AND NP06 IS NULL "
      '2005/8/26 END
      cnnConnection.Execute strSql
      
      'Add By Sindy 2019/7/10 消催審期限
      If m_CP10 = "734" Then '734.代理人撰稿
         strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
                  "WHERE NP01 = '" & m_CP09 & "' AND " & _
                        "NP07 = '305' AND NP06 IS NULL "
         cnnConnection.Execute strSql
      End If
   End If
   
   '若為已提申
   If frm02010601_02.textResult.Text = "2" Then
      'Add By Sindy 2016/10/14 更新申請案號
      If Frame1.Visible = True And IsEmptyText(textTM12_2) = False Then
         strSql = "UPDATE Trademark SET TM12 = '" & textTM12_2 & "'" & _
                          "WHERE " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
         cnnConnection.Execute strSql
      End If
      '2016/10/14 END
      If m_TM01 = "TF" And m_CP10 = "101" Then
         strSql = "UPDATE Trademark SET TM11 = " & IIf(Trim(DBDATE(Me.textTM11.Text)) = "", "Null", Trim(DBDATE(Me.textTM11.Text))) & " " & _
                          "WHERE " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
         cnnConnection.Execute strSql
      End If

      'Added by Morgan 2018/11/19
      'Modify By Sindy 2020/1/9
      'If strSrvDate(1) >= e化客戶啟用日 And textPrint <> "N" Then
      If (strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(m_TM01, 1) = "T") Or _
         (strSrvDate(1) >= e化客戶啟用日 And textPrint <> "N") Then
         '新增1909已提申C類收文
         strExc(1) = AutoNo("C", 6)
         strCP09 = strExc(1) 'Add By Sindy 2019/5/22
         strCP10 = "1909" '已提申
         strExc(2) = PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
         strExc(3) = GetSalesArea(strExc(2))
         strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26," & _
            "cp32,cp27,cp43) values('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'" & _
            "," & strSrvDate(1) & ",'" & strExc(1) & "','" & strCP10 & "','" & strExc(3) & "','" & strExc(2) & "','" & strUserNum & "','N','N','N'," & strSrvDate(1) & _
            ",'" & m_CP09 & "')"
         cnnConnection.Execute strSql, intI
         
         'Add By Sindy 2019/12/25 商標電子化
         If Left(m_TM01, 1) = "T" Then
            strLD18 = strCP09
            PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", False, m_TM23, strCP10, m_TM44
         End If
         '2019/12/25 END
      End If
      'end 2018/11/19
      
      'Add by Sindy 2019/6/27 新增大陸商申提申後管制催通知申請案號
      If m_CP10 = "101" And IsEmptyText(textTM12_2) = True Then
         strNP07 = "1101"
         '代理人提申日+45個工作天
         strNP09 = PUB_GetWorkDayAfterSysDate(CDbl(DBDATE(textCP47.Text)), 45) + 19110000
         'Modified by Lydia 2020/07/07本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strNP08 = strNP09
         strNP08 = PUB_GetWorkDay1(strNP09, True)
         If rsTmp.State <> adStateClosed Then rsTmp.Close
         Set rsTmp = Nothing
         strSql = "select * from NEXTPROGRESS where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP07='1101' AND NP06 IS NULL"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <> 0 Then
            strSql = "UPDATE NEXTPROGRESS SET NP08=" & strNP08 & ",NP09=" & strNP09 & ",NP10='" & m_CP14 & "' where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP07='1101' AND NP06 IS NULL"
         Else
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                      "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                                strNP08 & "," & strNP09 & ",'" & m_CP14 & "'," & GetNextProgressNo() & ")"
         End If
         cnnConnection.Execute strSql
      End If
      '2019/6/27 END
   End If
   
'   'Add By Sindy 2009/09/24
'   '因為有些來函由內商輸入，內商有自行控管之承辦期限及發文日。改為內商輸入所有C類來函，
'   '若業務區為F字頭者，除爭議受理外，自動產生B類收文，案件性質為外商發文722，不上發文日，不向客戶請款
'   Dim strCP48 As String, strCP09B As String
'   If Left(GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)), 1) = "F" And _
'      ((m_TM01 = "T" And m_TM10 = "020") Or (m_TM01 = "FCT" And m_TM10 = "000")) Then
'      strCP09B = AutoNo("B", 6)
'      '承辦期限為系統日加4個工作天
'      strCP48 = DBDATE(Pub_GetHandleDay(m_TM01, m_TM10, "722", strSrvDate(1), , m_CP09))
'      strSQL = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp48,cp20,cp26,cp32,cp43) " & _
'                     "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
'                     "," & CNULL(m_TM04) & "," & CNULL(strSrvDate(1)) & "," & CNULL(strCP09B) & ",722," & _
'                     CNULL(m_CP12) & "," & CNULL(m_CP13) & "," & CNULL(m_CP13) & "," & CNULL(strCP48) & ",'N','N','N'," & CNULL(m_CP09) & ")"
'      cnnConnection.Execute strSQL
'   End If
'   '2009/09/24 End
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If
   
   'Add by Sindy 2019/5/22
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010601_01", strCP09
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
   ' 代理人收達日
   If IsEmptyText(textCP46) = True Then
      If frm02010601_02.GetTextResult() = "1" Then
         strTit = "檢核資料"
         strMsg = "代理人收達日不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If

   ' 代理人提申日
   If IsEmptyText(textCP47) = True Then
      If frm02010601_02.GetTextResult() = "2" Then
         strTit = "檢核資料"
         strMsg = "代理人提申日不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        Me.textCP47.SetFocus
         GoTo EXITSUB
      End If
   Else
      'Modify By Sindy 2010/12/23
      '若操作人員為內商(部門P2X), 檢查若與前次輸入不同則顯示提示訊息
      If Left(GetStaffDepartment(strUserNum), 2) = "P2" Then
         If m_CP47 <> "" Then
            If DBDATE(textCP47) <> m_CP47 Then
               strTit = "檢核資料"
               strMsg = "與前次輸入的提申日不同，請再確認！"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               Me.textCP47.SetFocus
               GoTo EXITSUB
            End If
         End If
      End If
      '2010/12/23 End
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010601_03 = Nothing
End Sub

'Add By Sindy 2016/9/1
Private Sub textTM12_2_GotFocus()
   InverseTextBox textTM12_2
End Sub
Private Sub textTM12_2_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If IsEmptyText(textTM12_2) = False Then
      '檢查申請案號所輸入的長度是否正確
      'Add By Sindy 2017/5/17 + strRetrunText
      If PUB_ChkTm12Tm15Length("1", textTM12_2, m_TM01, m_TM02, m_TM03, m_TM04, "020", , , strRetrunText) = False Then
         Cancel = True
         textTM12_2_GotFocus
         Exit Sub
      'Add By Sindy 2017/5/17
      Else
         textTM12_2 = strRetrunText
      '2017/5/17 END
      End If
   End If
End Sub
'2016/9/1 END

' 代理人收達日
Private Sub textCP46_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP46) = False Then
      Select Case m_TM01
         Case "CFT", "CFC", "S":
            ' 代理人收達日不正確
            If CheckIsDate(textCP46, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的代理人收達日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP46_GotFocus
               GoTo EXITSUB
            End If
         Case Else:
            ' 代理人收達日不正確
            If CheckIsTaiwanDate(textCP46, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的代理人收達日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP46_GotFocus
               GoTo EXITSUB
            End If
      End Select
      
      ' 代理人收達日不可超過系統日
      If Val(DBDATE(textCP46)) > Val(DBDATE(SystemDate())) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人收達日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP46_GotFocus
         GoTo EXITSUB
      End If
      
      'add by sonia 2018/2/6
      If Val(DBDATE(textCP46)) < Val(m_CP27) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人收達日不可小於發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP46_GotFocus
         GoTo EXITSUB
      End If
      'end 2018/2/6
      'add by sonia 2025/8/5 催審若輸收達日則同時更新提申日
      If m_CP10 = "305" Then
         textCP47 = textCP46
      End If
      'end 2025/8/5
   End If
EXITSUB:
End Sub

' 代理人提申日
Private Sub textCP47_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP47) = False Then
      Select Case m_TM01
         Case "CFT", "CFC", "S":
            ' 代理人提申日不正確
            If CheckIsDate(textCP47, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的代理人提申日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP47_GotFocus
               GoTo EXITSUB
            End If
         Case Else:
            ' 代理人提申日不正確
            If CheckIsTaiwanDate(textCP47, False) = False Then
               Cancel = True
               strTit = "資料檢核"
               strMsg = "請輸入正確的代理人提申日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP47_GotFocus
               GoTo EXITSUB
            End If
      End Select
      
      ' 代理人提申日不可超過系統日
      If Val(DBDATE(textCP47)) > Val(DBDATE(SystemDate())) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人提申日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP47_GotFocus
         GoTo EXITSUB
      End If
      
      'add by sonia 2018/2/6
      If Val(DBDATE(textCP47)) < Val(m_CP27) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人提申日不可小於發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP47_GotFocus
         GoTo EXITSUB
      End If
      'end 2018/2/6
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

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP45_GotFocus()
   InverseTextBox textCP45
End Sub

Private Sub textCP46_GotFocus()
   InverseTextBox textCP46
End Sub

Private Sub textCP47_GotFocus()
   InverseTextBox textCP47
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
    Dim strSql As String
    'Add By Cheng 2003/01/07
    Dim strExceptField As String
    Dim ii As Integer
    
    'Modify By Cheng 2003/01/07
    '與PrintLetter的程式一致
'   ' 申請國家非台灣
'   If m_TM10 > "010" Then
'      If m_TM01 = "TF" Then
'         ' 清除定稿例外欄位檔原有資料
'         EndLetter "09", m_CP09, "01", strUserNum
'      Else
'         ' 清除定稿例外欄位檔原有資料
'         EndLetter "09", m_CP09, "02", strUserNum
'      End If
'   Else
        If m_TM01 = "CFC" Then
            ' 清除定稿例外欄位檔原有資料
            EndLetter "09", m_CP09, "01", strUserNum
        'Add By Cheng 2003/01/08
        ElseIf m_TM01 = "T" Then
            '2010/9/20 modify by sonia 大陸復審401,異議601,裁定603,撤銷605定稿不同
            'EndLetter "09", m_CP09, "00", strUserNum
            Select Case m_CP10
             Case "401", "601", "603", "605"
                EndLetter "09", m_CP09, "01", strUserNum
             Case Else
                EndLetter "09", m_CP09, "00", strUserNum
                'Add By Sindy 2010/10/5 取得相關總收文號案件性質
                Dim strCP43CP10 As String
                If m_CP43 <> "" Then
                  strSql = "SELECT CPM03,CPM04 FROM CaseProgress,CASEPROPERTYMAP WHERE CP09='" & m_CP43 & "' and CPM01(+)=CP01 AND CPM02(+)=CP10 "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
'                     If m_TM10 = "000" Then
'                        strCP43CP10 = RsTemp("CPM03")
'                     Else
                        strCP43CP10 = RsTemp("CPM04")
'                     End If
                  End If
                  Select Case m_CP10
                    Case "306", "307", "313", "612", "613"
                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & "09" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                                "'" & "相關總收文號案件性質" & "','" & strCP43CP10 & "')"
                       cnnConnection.Execute strSql
                  End Select
                End If
                '2010/10/5 End
                'Add By Sindy 2016/9/1 + 台至大提申加入申請案號
                If Frame1.Visible = True And textTM12_2 <> "" Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "09" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                              "'" & "台至大提申加入申請案號" & "','申請案號編為" & textTM12_2 & "。大約6個月後可接獲進一步消息，屆時將立即轉知。')"
                     cnnConnection.Execute strSql
                End If
                '2016/9/1 END
            End Select
            '2010/9/20 end
'        'Add By Sindy 2010/5/18
'        ElseIf m_TM01 = "TF" Then
'            EndLetter "09", m_CP09, "03", strUserNum
        Else
            'Modify By Sindy 2012/4/26 增加定稿語文做判斷
            If m_strLanguage = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "09", m_CP09, "02", strUserNum
               'Add By Cheng 2003/01/07
               ' 附件
               strExceptField = ""
               For ii = 0 To Me.Check1.Count - 1
                   If Me.Check1(ii).Value = vbChecked Then strExceptField = strExceptField & Me.Check1(ii).Caption & "、"
               Next ii
               If strExceptField <> "" Then
                   strExceptField = "附件：" & Left(strExceptField, Len(strExceptField) - 1) & "。"
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "09" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                            "'" & "附件" & "','" & strExceptField & "')"
                   cnnConnection.Execute strSql
               End If
            ElseIf m_strLanguage = "2" Then
               EndLetter "09", m_CP09, "03", strUserNum 'Add By Sindy 2010/5/18
            End If
        End If
'   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
'Add By Sindy 2012/1/12
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/12 End
   
   '2012/4/26 ADD BY Sindy
   '取得定稿語文
   m_strLanguage = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'MODIFY BY SONIA 90.11.17申請國家非台灣才有此功能
   ' 申請國家非台灣
   'If m_TM10 > "010" Then
   '   If m_TM01 = "TF" Then
   '      ' 列印定稿
   '      NowPrint m_CP09, "09", "01", False, strUserNum, 0
   '   Else
   '      ' 列印定稿
   '      NowPrint m_CP09, "09", "02", False, strUserNum, 0
   '   End If
   'Else
   '   If m_TM01 = "CFC" Then
   '      ' 列印定稿
   '      NowPrint m_CP09, "09", "01", False, strUserNum, 0
   '   Else
   '      ' 列印定稿
   '      NowPrint m_CP09, "09", "02", False, strUserNum, 0
   '   End If
   'End If
   
   'Add By Sindy 2012/1/12
   ET01 = "09"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/12 End
   
    Select Case m_TM01
    Case "CFC"
'        NowPrint m_CP09, "09", "01", False, strUserNum, 0
         ET03 = "01" 'Modify By Sindy 2012/1/12
    'Add By Cheng 2003/01/08
    Case "T"
        '2010/9/20 modify by sonia 大陸復審401,異議601,裁定603,撤銷605定稿不同
        'NowPrint m_CP09, "09", "00", False, strUserNum, 0
        Select Case m_CP10
         Case "401", "601", "603", "605"
'            NowPrint m_CP09, "09", "01", False, strUserNum, 0
            ET03 = "01" 'Modify By Sindy 2012/1/12
         Case Else
'            NowPrint m_CP09, "09", "00", False, strUserNum, 0
            ET03 = "00" 'Modify By Sindy 2012/1/12
        End Select
        '2010/9/20 end
'    'Add By Sindy 2010/5/18
'    Case "TF"
''        NowPrint m_CP09, "09", "03", False, strUserNum, 0
'         ET03 = "03" 'Modify By Sindy 2012/1/12
    Case Else
         'Modify By Sindy 2012/4/26 增加定稿語文做判斷
         If m_strLanguage = "1" Then
   '        NowPrint m_CP09, "09", "02", False, strUserNum, 0
            ET03 = "02" 'Modify By Sindy 2012/1/12
         ElseIf m_strLanguage = "2" Then
            'Add By Sindy 2010/5/18
            ET03 = "03" 'Modify By Sindy 2012/1/12
         End If
    End Select
    
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
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
         End If
      Else
         'Add By Sindy 2019/12/25 + strLD18.信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
      End If
   End If
   '2012/1/12 End
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCP46.Enabled = True Then
   Cancel = False
   textCP46_Validate Cancel
   If Cancel = True Then
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
If Me.textTM11.Enabled = True Then
   Cancel = False
   textTM11_Validate Cancel
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

Private Sub textTM11_GotFocus()
    TextInverse Me.textTM11
End Sub

Private Sub textTM11_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM11) = False Then
      Select Case m_TM01
      Case "CFT", "CFC", "S":
          ' 申請日不正確
          If CheckIsDate(textTM11, False) = False Then
              Cancel = True
              strTit = "資料檢核"
              strMsg = "請輸入正確的申請日"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              textTM11.SetFocus
              textTM11_GotFocus
              GoTo EXITSUB
          End If
      Case Else:
          ' 申請日不正確
          If CheckIsTaiwanDate(textTM11, False) = False Then
              Cancel = True
              strTit = "資料檢核"
              strMsg = "請輸入正確的申請日"
              nResponse = MsgBox(strMsg, vbOKOnly, strTit)
              textTM11.SetFocus
              textTM11_GotFocus
              GoTo EXITSUB
          End If
      End Select
        
      'add by sonia 2018/2/6
      If Val(DBDATE(textTM11)) < Val(m_CP27) And Me.textTM11.Visible = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "申請日不可小於發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM11_GotFocus
         GoTo EXITSUB
      End If
      'end 2018/2/6
   End If
EXITSUB:
End Sub
