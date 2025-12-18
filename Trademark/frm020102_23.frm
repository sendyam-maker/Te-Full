VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_23 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件申請人地址資料"
   ClientHeight    =   6390
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdCopy 
      Caption         =   "複製"
      Default         =   -1  'True
      Height          =   300
      Index           =   4
      Left            =   8370
      TabIndex        =   53
      Top             =   5520
      Width           =   555
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "複製"
      Height          =   300
      Index           =   3
      Left            =   8370
      TabIndex        =   52
      Top             =   4315
      Width           =   555
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "複製"
      Height          =   300
      Index           =   2
      Left            =   8370
      TabIndex        =   51
      Top             =   3150
      Width           =   555
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "複製"
      Height          =   300
      Index           =   1
      Left            =   8370
      TabIndex        =   50
      Top             =   1920
      Width           =   555
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "複製"
      Height          =   300
      Index           =   0
      Left            =   8370
      TabIndex        =   49
      Top             =   660
      Width           =   555
   End
   Begin VB.TextBox textTM81 
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   47
      Top             =   5205
      Width           =   1212
   End
   Begin VB.TextBox textTM80 
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   44
      Top             =   4005
      Width           =   1212
   End
   Begin VB.TextBox textTM79 
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   41
      Top             =   2820
      Width           =   1212
   End
   Begin VB.TextBox textTM78 
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   38
      Top             =   1545
      Width           =   1212
   End
   Begin VB.TextBox textTM25 
      Height          =   300
      Left            =   1500
      MaxLength       =   185
      TabIndex        =   1
      Top             =   925
      Width           =   6825
   End
   Begin VB.TextBox textTM86 
      Height          =   300
      Left            =   1500
      MaxLength       =   185
      TabIndex        =   4
      Top             =   2175
      Width           =   6825
   End
   Begin VB.TextBox textTM87 
      Height          =   300
      Left            =   1500
      MaxLength       =   185
      TabIndex        =   7
      Top             =   3400
      Width           =   6825
   End
   Begin VB.TextBox textTM88 
      Height          =   300
      Left            =   1500
      MaxLength       =   185
      TabIndex        =   10
      Top             =   4600
      Width           =   6825
   End
   Begin VB.TextBox textTM89 
      Height          =   300
      Left            =   1500
      MaxLength       =   185
      TabIndex        =   13
      Top             =   5800
      Width           =   6825
   End
   Begin VB.TextBox textTM23 
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   20
      Top             =   315
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "確定(&O)"
      Height          =   360
      Left            =   7200
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8070
      TabIndex        =   16
      Top             =   10
      Width           =   800
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   15
      Width           =   2532
   End
   Begin MSForms.TextBox textTM24 
      Height          =   300
      Left            =   1500
      TabIndex        =   0
      Top             =   625
      Width           =   6825
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12039;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM26 
      Height          =   300
      Left            =   1500
      TabIndex        =   2
      Top             =   1230
      Width           =   6825
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12039;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM82 
      Height          =   300
      Left            =   1500
      TabIndex        =   3
      Top             =   1875
      Width           =   6825
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12039;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM90 
      Height          =   300
      Left            =   1500
      TabIndex        =   5
      Top             =   2475
      Width           =   6825
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12039;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM83 
      Height          =   300
      Left            =   1500
      TabIndex        =   6
      Top             =   3115
      Width           =   6825
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12039;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM91 
      Height          =   300
      Left            =   1500
      TabIndex        =   8
      Top             =   3690
      Width           =   6825
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12039;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM84 
      Height          =   300
      Left            =   1500
      TabIndex        =   9
      Top             =   4315
      Width           =   6825
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12039;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM92 
      Height          =   300
      Left            =   1500
      TabIndex        =   11
      Top             =   4890
      Width           =   6825
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12039;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM85 
      Height          =   300
      Left            =   1500
      TabIndex        =   12
      Top             =   5500
      Width           =   6825
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12039;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM93 
      Height          =   300
      Left            =   1500
      TabIndex        =   14
      Top             =   6090
      Width           =   6825
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "12039;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM81_NM 
      Height          =   300
      Left            =   2325
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   5205
      Width           =   6315
      VariousPropertyBits=   679493663
      Size            =   "11139;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80_NM 
      Height          =   300
      Left            =   2325
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   4005
      Width           =   6315
      VariousPropertyBits=   679493663
      Size            =   "11139;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79_NM 
      Height          =   300
      Left            =   2325
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2820
      Width           =   6315
      VariousPropertyBits=   679493663
      Size            =   "11139;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78_NM 
      Height          =   300
      Left            =   2325
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1545
      Width           =   6315
      VariousPropertyBits=   679493663
      Size            =   "11139;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23_NM 
      Height          =   300
      Left            =   2325
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   315
      Width           =   6315
      VariousPropertyBits=   679493663
      Size            =   "11139;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "申請人5："
      Height          =   255
      Left            =   150
      TabIndex        =   48
      Top             =   5205
      Width           =   810
   End
   Begin VB.Label Label3 
      Caption         =   "申請人4："
      Height          =   255
      Left            =   150
      TabIndex        =   45
      Top             =   4005
      Width           =   810
   End
   Begin VB.Label Label4 
      Caption         =   "申請人3："
      Height          =   255
      Left            =   150
      TabIndex        =   42
      Top             =   2820
      Width           =   810
   End
   Begin VB.Label Label2 
      Caption         =   "申請人2："
      Height          =   255
      Left            =   150
      TabIndex        =   39
      Top             =   1545
      Width           =   810
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "申請地址1(中)："
      Height          =   180
      Left            =   150
      TabIndex        =   36
      Top             =   625
      Width           =   1290
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "申請地址1(英)："
      Height          =   180
      Left            =   150
      TabIndex        =   35
      Top             =   925
      Width           =   1290
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "申請地址1(日)："
      Height          =   180
      Left            =   150
      TabIndex        =   34
      Top             =   1230
      Width           =   1290
   End
   Begin VB.Label Label64 
      AutoSize        =   -1  'True
      Caption         =   "申請地址2(中)："
      Height          =   180
      Left            =   150
      TabIndex        =   33
      Top             =   1875
      Width           =   1290
   End
   Begin VB.Label Label65 
      AutoSize        =   -1  'True
      Caption         =   "申請地址2(英)："
      Height          =   180
      Left            =   150
      TabIndex        =   32
      Top             =   2175
      Width           =   1290
   End
   Begin VB.Label Label66 
      AutoSize        =   -1  'True
      Caption         =   "申請地址2(日)："
      Height          =   180
      Left            =   150
      TabIndex        =   31
      Top             =   2475
      Width           =   1290
   End
   Begin VB.Label Label67 
      AutoSize        =   -1  'True
      Caption         =   "申請地址3(中)："
      Height          =   180
      Left            =   150
      TabIndex        =   30
      Top             =   3115
      Width           =   1290
   End
   Begin VB.Label Label68 
      AutoSize        =   -1  'True
      Caption         =   "申請地址3(英)："
      Height          =   180
      Left            =   150
      TabIndex        =   29
      Top             =   3400
      Width           =   1290
   End
   Begin VB.Label Label69 
      AutoSize        =   -1  'True
      Caption         =   "申請地址3(日)："
      Height          =   180
      Left            =   150
      TabIndex        =   28
      Top             =   3690
      Width           =   1290
   End
   Begin VB.Label Label70 
      AutoSize        =   -1  'True
      Caption         =   "申請地址4(中)："
      Height          =   180
      Left            =   150
      TabIndex        =   27
      Top             =   4315
      Width           =   1290
   End
   Begin VB.Label Label71 
      AutoSize        =   -1  'True
      Caption         =   "申請地址4(英)："
      Height          =   180
      Left            =   150
      TabIndex        =   26
      Top             =   4600
      Width           =   1290
   End
   Begin VB.Label Label72 
      AutoSize        =   -1  'True
      Caption         =   "申請地址4(日)："
      Height          =   180
      Left            =   150
      TabIndex        =   25
      Top             =   4890
      Width           =   1290
   End
   Begin VB.Label Label73 
      AutoSize        =   -1  'True
      Caption         =   "申請地址5(中)："
      Height          =   180
      Left            =   150
      TabIndex        =   24
      Top             =   5500
      Width           =   1290
   End
   Begin VB.Label Label74 
      AutoSize        =   -1  'True
      Caption         =   "申請地址5(英)："
      Height          =   180
      Left            =   150
      TabIndex        =   23
      Top             =   5800
      Width           =   1290
   End
   Begin VB.Label Label75 
      AutoSize        =   -1  'True
      Caption         =   "申請地址5(日)："
      Height          =   180
      Left            =   150
      TabIndex        =   22
      Top             =   6090
      Width           =   1290
   End
   Begin VB.Label Label84 
      Caption         =   "申請人1："
      Height          =   255
      Left            =   150
      TabIndex        =   21
      Top             =   315
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   18
      Top             =   45
      Width           =   900
   End
End
Attribute VB_Name = "frm020102_23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/23 Form2.0已修改 textTM23_NM(申請人名).../textTM24/textTM26/textTM82/textTM90/textTM83/textTM91/textTM84/textTM92/textTM85/textTM93
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Public m_CP09 As String
Public m_TM01 As String
Public m_TM02 As String
Public m_TM03 As String
Public m_TM04 As String

' 前畫面
Public UpForm As Form

Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer
Dim TField(), PField()  'Add by Amy 2018/09/14
'Add by Amy 2018/10/26
Dim i As Integer
Public stModApply As String '前畫面修改後之申請人
Dim strSeq As String

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiOldData = strFieldData
         m_CPList(nPos).fiNewData = strFieldData
         m_CPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_CPList(m_CPCount + 1)
      m_CPList(m_CPCount).fiName = strFieldName
      m_CPList(m_CPCount).fiOldData = strFieldData
      m_CPList(m_CPCount).fiNewData = strFieldData
      m_CPList(m_CPCount).fiType = nFieldType
      m_CPCount = m_CPCount + 1
   End If
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 清除欄位串列
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

'Add By Sindy 2018/2/9
Private Sub cmdCopy_Click(Index As Integer)
Dim strCAddr As String
Dim strEAddr As String
Dim strJAddr As String
   
   If Index = 0 Then '申請人1
      If textTM23 <> "" Then
         Call GetCustomer(textTM23, strCAddr, strEAddr, strJAddr)
         textTM24 = strCAddr
         textTM25 = strEAddr
         textTM26 = strJAddr
      End If
   ElseIf Index = 1 Then '申請人2
      If textTM78 <> "" Then
         Call GetCustomer(textTM78, strCAddr, strEAddr, strJAddr)
         textTM82 = strCAddr
         textTM86 = strEAddr
         textTM90 = strJAddr
      End If
   ElseIf Index = 2 Then '申請人3
      If textTM79 <> "" Then
         Call GetCustomer(textTM79, strCAddr, strEAddr, strJAddr)
         textTM83 = strCAddr
         textTM87 = strEAddr
         textTM91 = strJAddr
      End If
   ElseIf Index = 3 Then '申請人4
      If textTM80 <> "" Then
         Call GetCustomer(textTM80, strCAddr, strEAddr, strJAddr)
         textTM84 = strCAddr
         textTM88 = strEAddr
         textTM92 = strJAddr
      End If
   Else '申請人5
      If textTM81 <> "" Then
         Call GetCustomer(textTM81, strCAddr, strEAddr, strJAddr)
         textTM85 = strCAddr
         textTM89 = strEAddr
         textTM93 = strJAddr
      End If
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
   'UpForm.Show
End Sub

Private Sub cmdOK_Click()
   '檢查輸入資料的有效性
   If CheckDataValidate = False Then Exit Sub
   If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
   Unload Me
   UpForm.Show
End Sub

Private Sub Form_Load()
   textTMKey.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM23_NM.BackColor = &H8000000F
   textTM78.BackColor = &H8000000F
   textTM78_NM.BackColor = &H8000000F
   textTM79.BackColor = &H8000000F
   textTM79_NM.BackColor = &H8000000F
   textTM80.BackColor = &H8000000F
   textTM80_NM.BackColor = &H8000000F
   textTM81.BackColor = &H8000000F
   textTM81_NM.BackColor = &H8000000F
   
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   stModApply = "" 'Add by Amy 2018/10/26
   ClearCPFieldList
   Set frm020102_23 = Nothing
End Sub

' 由客戶代碼取得客戶名稱
'Modify By Sindy 2018/2/9 取得地址 Optional ByRef CAddr As String = "", _
   Optional ByRef EAddr As String = "", Optional ByRef JAddr As String = ""
Private Function GetCustomer(ByVal strData As String, Optional ByRef CAddr As String = "", _
   Optional ByRef EAddr As String = "", Optional ByRef JAddr As String = "") As String
Dim rsTmp As ADODB.Recordset
Dim strSql As String
   
   GetCustomer = Empty
   strData = Left(Trim(strData) & "000", 9)
   If IsEmptyText(strData) = False Then
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT Customer.*,cu23 as CAddr,cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 as EAddr,cu29 as JAddr FROM Customer " & _
               "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                     "CU02 = '" & Mid(strData, 9, 1) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("CU04")) = False Then
            GetCustomer = rsTmp.Fields("CU04")
            
         ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
            GetCustomer = rsTmp.Fields("CU05")
            If IsNull(rsTmp.Fields("CU88")) = False Then GetCustomer = GetCustomer & rsTmp.Fields("CU88")
            If IsNull(rsTmp.Fields("CU89")) = False Then GetCustomer = GetCustomer & rsTmp.Fields("CU89")
            If IsNull(rsTmp.Fields("CU90")) = False Then GetCustomer = GetCustomer & rsTmp.Fields("CU90")
            
         ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
            GetCustomer = rsTmp.Fields("CU06")
         End If
         'Add By Sindy 2018/2/9
         If IsNull(rsTmp.Fields("CAddr")) = False Then
            CAddr = rsTmp.Fields("CAddr")
            'ZipCode+中文地址
            If IsNull(rsTmp.Fields("CU112")) = False Then
               If InStr(CAddr, rsTmp.Fields("CU112")) = 0 Then
                  CAddr = rsTmp.Fields("CU112") & CAddr
               End If
            End If
         End If
         If IsNull(rsTmp.Fields("EAddr")) = False Then
            EAddr = Trim(rsTmp.Fields("EAddr"))
         End If
         If IsNull(rsTmp.Fields("JAddr")) = False Then
            JAddr = rsTmp.Fields("JAddr")
         '2018/2/9 END
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

Public Sub QueryData()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
'Add by Amy 2018/10/26
Dim strApply(1 To 5) As String
Dim strAddrC(1 To 5) As String, strAddrE(1 To 5) As String, strAddrJ(1 To 5) As String '地址(中/英/日)

   'Add by Amy 2018/09/14
   TField = Array("TM23", "TM24", "TM25", "TM26", "TM78", "TM82", "TM86", "TM90", "TM79", "TM83", "TM87", "TM91", "TM80", "TM84", "TM88", "TM92", "TM81", "TM85", "TM89", "TM93")
   PField = Array("PA26", "PA31", "PA36", "PA41", "PA27", "PA32", "PA37", "PA42", "PA28", "PA33", "PA38", "PA43", "PA29", "PA34", "PA39", "PA44", "PA30", "PA35", "PA40", "PA45")
   'end 2018/09/14
   
   textTM23 = Empty
   textTM23_NM = Empty
   textTM24 = Empty
   textTM25 = Empty
   textTM26 = Empty
   textTM78 = Empty
   textTM78_NM = Empty
   textTM82 = Empty
   textTM86 = Empty
   textTM90 = Empty
   textTM79 = Empty
   textTM79_NM = Empty
   textTM83 = Empty
   textTM87 = Empty
   textTM91 = Empty
   textTM80 = Empty
   textTM80_NM = Empty
   textTM84 = Empty
   textTM88 = Empty
   textTM92 = Empty
   textTM81 = Empty
   textTM81_NM = Empty
   textTM85 = Empty
   textTM89 = Empty
   textTM93 = Empty
   ClearCPFieldList
   
   'Add By Sindy 2018/2/2
   If m_CP09 <> "" Then
   '2018/2/2 END
      strSql = "SELECT * FROM CaseProgress WHERE CP09='" & m_CP09 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         m_TM01 = RsTemp("CP01")
         m_TM02 = RsTemp("CP02")
         m_TM03 = RsTemp("CP03")
         m_TM04 = RsTemp("CP04")
      End If
   End If
   Me.textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
  
   'Modify by Amy 2018/09/14 加入專利基本檔資料及ChgPField
   strSql = "SELECT TM23,TM24,TM25,TM26,TM78,TM82,TM86,TM90,TM79,TM83,TM87,TM91,TM80,TM84,TM88,TM92,TM81,TM85,TM89,TM93  FROM TradeMark " & _
                "WHERE TM01 = '" & m_TM01 & "' AND TM02 = '" & m_TM02 & "' AND TM03 = '" & m_TM03 & "' AND TM04 = '" & m_TM04 & "' " & _
    "Union Select PA26 tm23,PA31 tm24,PA36 tm25,PA41 tm26,PA27 tm78,PA32 tm82,PA37 tm86,PA42 tm90,PA28 tm79,PA33 tm83,PA38 tm87,PA43 tm91,PA29 tm80,PA34 tm84,PA39 tm88,PA44 tm92,PA30 tm81,PA35 tm85,PA40 tm89,PA45 tm93  From Patent " & _
                "Where PA01='" & m_TM01 & "' AND PA02 = '" & m_TM02 & "' AND PA03 = '" & m_TM03 & "' AND PA04 = '" & m_TM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'Modify by Amy 2018/10/26 前畫面有修改申請人需抓新申請人資料(將抓出資料先放於變數中)
'      textTM23 = Trim("" & rsTmp.Fields("TM23"))
'      If Trim(textTM23) <> "" Then textTM23_NM = GetCustomer(textTM23)
'      textTM24 = Trim("" & rsTmp.Fields("TM24")): SetCPFieldOldData ChgPField("TM24"), textTM24, 0
'      textTM25 = Trim("" & rsTmp.Fields("TM25")): SetCPFieldOldData ChgPField("TM25"), textTM25, 0
'      textTM26 = Trim("" & rsTmp.Fields("TM26")): SetCPFieldOldData ChgPField("TM26"), textTM26, 0
      strApply(1) = Trim("" & rsTmp.Fields("TM23")): SetCPFieldOldData ChgPField("TM23"), strApply(1), 0
      strAddrC(1) = Trim("" & rsTmp.Fields("TM24")): SetCPFieldOldData ChgPField("TM24"), strAddrC(1), 0
      strAddrE(1) = Trim("" & rsTmp.Fields("TM25")): SetCPFieldOldData ChgPField("TM25"), strAddrE(1), 0
      strAddrJ(1) = Trim("" & rsTmp.Fields("TM26")): SetCPFieldOldData ChgPField("TM26"), strAddrJ(1), 0
      
      strApply(2) = Trim("" & rsTmp.Fields("TM78")): SetCPFieldOldData ChgPField("TM78"), strApply(2), 0
      If Trim(textTM78) <> "" Then textTM78_NM = GetCustomer(textTM78)
      strAddrC(2) = Trim("" & rsTmp.Fields("TM82")): SetCPFieldOldData ChgPField("TM82"), strAddrC(2), 0
      strAddrE(2) = Trim("" & rsTmp.Fields("TM86")): SetCPFieldOldData ChgPField("TM86"), strAddrE(2), 0
      strAddrJ(2) = Trim("" & rsTmp.Fields("TM90")): SetCPFieldOldData ChgPField("TM90"), strAddrJ(2), 0
      
      strApply(3) = Trim("" & rsTmp.Fields("TM79")): SetCPFieldOldData ChgPField("TM79"), strApply(3), 0
      strAddrC(3) = Trim("" & rsTmp.Fields("TM83")): SetCPFieldOldData ChgPField("TM83"), strAddrC(3), 0
      strAddrE(3) = Trim("" & rsTmp.Fields("TM87")): SetCPFieldOldData ChgPField("TM87"), strAddrE(3), 0
      strAddrJ(3) = Trim("" & rsTmp.Fields("TM91")): SetCPFieldOldData ChgPField("TM91"), strAddrJ(3), 0
      
      strApply(4) = Trim("" & rsTmp.Fields("TM80")): SetCPFieldOldData ChgPField("TM80"), strApply(4), 0
      If Trim(textTM80) <> "" Then textTM80_NM = GetCustomer(textTM80)
      strAddrC(4) = Trim("" & rsTmp.Fields("TM84")): SetCPFieldOldData ChgPField("TM84"), strAddrC(4), 0
      strAddrE(4) = Trim("" & rsTmp.Fields("TM88")): SetCPFieldOldData ChgPField("TM88"), strAddrE(4), 0
      strAddrJ(4) = Trim("" & rsTmp.Fields("TM92")): SetCPFieldOldData ChgPField("TM92"), strAddrJ(4), 0
      
      strApply(5) = Trim("" & rsTmp.Fields("TM81")): SetCPFieldOldData ChgPField("TM81"), strApply(5), 0
      strAddrC(5) = Trim("" & rsTmp.Fields("TM85")): SetCPFieldOldData ChgPField("TM85"), strAddrC(5), 0
      strAddrE(5) = Trim("" & rsTmp.Fields("TM89")): SetCPFieldOldData ChgPField("TM89"), strAddrE(5), 0
      strAddrJ(5) = Trim("" & rsTmp.Fields("TM93")): SetCPFieldOldData ChgPField("TM93"), strAddrJ(5), 0
   End If
   'end 2018/09/14
   rsTmp.Close
   'Add by Amy 2018/10/26 申請人有修改抓客戶檔資料
   If stModApply <> MsgText(601) Then
        strSeq = Left(stModApply, InStr(stModApply, ";") - 1)
        stModApply = Replace(stModApply, strSeq & ";", "")
        If stModApply = "" Then
            strApply(strSeq) = ""
            strAddrC(strSeq) = ""
            strAddrE(strSeq) = ""
            strAddrJ(strSeq) = ""
        Else
            strSql = "Select CU23,CU24||Decode(cu25,null,'',' '||CU25)||Decode(cu26,null,'',' '||CU26)||Decode(cu27,null,'',' '||CU27)||Decode(cu28,null,'',' '||CU28)||Decode(cu120,null,'',' '||cu102) " & _
                        ",CU29 From Customer Where Cu01='" & Left(GetNewFagent(stModApply), 8) & "' And Cu02='0' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
                strApply(strSeq) = stModApply
                strAddrC(strSeq) = "" & RsTemp.Fields(0)
                strAddrE(strSeq) = "" & RsTemp.Fields(1)
                strAddrJ(strSeq) = "" & RsTemp.Fields(2)
            End If
        End If
   End If
  
   For i = LBound(strApply) To UBound(strApply)
      Select Case i
        Case 1
            textTM23 = strApply(i)
            If Trim(textTM23) <> "" Then textTM23_NM = GetCustomer(textTM23)
            textTM24 = strAddrC(i)
            textTM25 = strAddrE(i)
            textTM26 = strAddrJ(i)
        Case 2
            textTM78 = strApply(i)
            If Trim(textTM78) <> "" Then textTM78_NM = GetCustomer(textTM78)
            textTM82 = strAddrC(i)
            textTM86 = strAddrE(i)
            textTM90 = strAddrJ(i)
        Case 3
            textTM79 = strApply(i)
            If Trim(textTM79) <> "" Then textTM79_NM = GetCustomer(textTM79)
            textTM83 = strAddrC(i)
            textTM87 = strAddrE(i)
            textTM91 = strAddrJ(i)
        Case 4
            textTM80 = strApply(i)
            If Trim(textTM80) <> "" Then textTM80_NM = GetCustomer(textTM80)
            textTM84 = strAddrC(i)
            textTM88 = strAddrE(i)
            textTM92 = strAddrJ(i)
        Case 5
            textTM81 = strApply(i)
            If Trim(textTM81) <> "" Then textTM81_NM = GetCustomer(textTM81)
            textTM85 = strAddrC(i)
            textTM89 = strAddrE(i)
            textTM93 = strAddrJ(i)
      End Select
   Next i
   'end 2018/10/26
   Set rsTmp = Nothing
End Sub

Public Function OnSaveData() As Boolean
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   'Modify by Amy 2018/09/14 +專利
   'Add by Amy 2018/10/26
   SetCPFieldNewData ChgPField("TM23"), textTM23
   SetCPFieldNewData ChgPField("TM78"), textTM78
   SetCPFieldNewData ChgPField("TM79"), textTM79
   SetCPFieldNewData ChgPField("TM80"), textTM80
   SetCPFieldNewData ChgPField("TM81"), textTM81
   'end 2018/10/26
   SetCPFieldNewData ChgPField("TM24"), textTM24
   SetCPFieldNewData ChgPField("TM25"), textTM25
   SetCPFieldNewData ChgPField("TM26"), textTM26
   SetCPFieldNewData ChgPField("TM82"), textTM82
   SetCPFieldNewData ChgPField("TM86"), textTM86
   SetCPFieldNewData ChgPField("TM90"), textTM90
   SetCPFieldNewData ChgPField("TM83"), textTM83
   SetCPFieldNewData ChgPField("TM87"), textTM87
   SetCPFieldNewData ChgPField("TM91"), textTM91
   SetCPFieldNewData ChgPField("TM84"), textTM84
   SetCPFieldNewData ChgPField("TM88"), textTM88
   SetCPFieldNewData ChgPField("TM92"), textTM92
   SetCPFieldNewData ChgPField("TM85"), textTM85
   SetCPFieldNewData ChgPField("TM89"), textTM89
   SetCPFieldNewData ChgPField("TM93"), textTM93
   
   ' 更新商標基本資料檔
   strSql = "UPDATE " & IIf(m_TM01 = "P", "Patent", "TradeMark") & " SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_CPList(nIndex).fiName & " = " & m_CPList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   
   ' 設定SQL語法更新的條件
   'Modify by Amy 2018/09/14 +專利
   If m_TM01 = "P" Then
     strSql = strSql & " WHERE PA01 = '" & m_TM01 & "'  AND PA02 = '" & m_TM02 & "' AND PA03 = '" & m_TM03 & "' AND PA04 = '" & m_TM04 & "' "
   Else
     strSql = strSql & " WHERE TM01 = '" & m_TM01 & "'  AND TM02 = '" & m_TM02 & "' AND TM03 = '" & m_TM03 & "' AND TM04 = '" & m_TM04 & "' "
   End If
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

Private Function CheckDataValidate() As Boolean
Dim Cancel As Boolean
   
   CheckDataValidate = False
   Cancel = False
   
   'Add by Amy 2021/12/23檢查畫面的 TextBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
    End If
      
   If Trim(textTM23) <> "" And _
      (Trim(textTM24) = "" And Trim(textTM25) = "" And Trim(textTM26) = "") Then
      MsgBox "申請人1地址不可空白!", vbCritical
      textTM24.SetFocus
      Exit Function
   End If
   
   If Trim(textTM78) <> "" And _
      (Trim(textTM82) = "" And Trim(textTM86) = "" And Trim(textTM90) = "") Then
      MsgBox "申請人2地址不可空白!", vbCritical
      textTM82.SetFocus
      Exit Function
   End If
   
   If Trim(textTM79) <> "" And _
      (Trim(textTM83) = "" And Trim(textTM87) = "" And Trim(textTM91) = "") Then
      MsgBox "申請人3地址不可空白!", vbCritical
      textTM83.SetFocus
      Exit Function
   End If
   
   If Trim(textTM80) <> "" And _
      (Trim(textTM84) = "" And Trim(textTM88) = "" And Trim(textTM92) = "") Then
      MsgBox "申請人4地址不可空白!", vbCritical
      textTM84.SetFocus
      Exit Function
   End If
   
   If Trim(textTM81) <> "" And _
      (Trim(textTM85) = "" And Trim(textTM89) = "" And Trim(textTM93) = "") Then
      MsgBox "申請人5地址不可空白!", vbCritical
      textTM85.SetFocus
      Exit Function
   End If
   
   Call textTM24_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   Call textTM26_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   Call textTM82_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   Call textTM90_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   Call textTM83_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   Call textTM91_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   Call textTM84_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   Call textTM92_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   Call textTM85_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   Call textTM93_Validate(Cancel)
   If Cancel = True Then
      Exit Function
   End If
   
   CheckDataValidate = True
End Function

Private Sub textTM24_GotFocus()
   OpenIme
   InverseTextBox textTM24
End Sub

'Add by Amy 2018/09/14
'Modify by Amy 2021/12/23 原:Integer
Private Sub textTM24_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ChangeZIP(KeyAscii, textTM24)
End Sub

' 申請地址(中)
Private Sub textTM24_Validate(Cancel As Boolean)
   Cancel = False
   '長度不符
   If CheckLengthIsOK(textTM24, textTM24.MaxLength) = False Then
      Cancel = True
      textTM24_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM25_GotFocus()
   InverseTextBox textTM25
End Sub

Private Sub textTM26_GotFocus()
   OpenIme
   InverseTextBox textTM26
End Sub

'Add by Amy 2018/09/14
'Modify by Amy 2021/12/23
Private Sub textTM26_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ChangeZIP(KeyAscii, textTM26)
End Sub

' 申請地址(日)
Private Sub textTM26_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM26, textTM26.MaxLength) = False Then
      Cancel = True
      textTM26_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM82_GotFocus()
   OpenIme
   InverseTextBox textTM82
End Sub

'Add by Amy 2018/09/14
'Modify by Amy 2021/12/23
Private Sub textTM82_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ChangeZIP(KeyAscii, textTM82)
End Sub

Private Sub textTM82_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM82, textTM82.MaxLength) = False Then
      Cancel = True
      textTM82_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

'Add byAmy 2018/0914
'Modify by Amy 2021/12/23
Private Sub textTM83_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ChangeZIP(KeyAscii, textTM83)
End Sub

'Add by Amy 2018/09/14
'Modify by Amy 2021/12/23
Private Sub textTM84_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ChangeZIP(KeyAscii, textTM84)
End Sub

'Add by Amy 2018/09/14
'Modify by Amy 2021/12/23
Private Sub textTM85_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ChangeZIP(KeyAscii, textTM85)
End Sub

Private Sub textTM86_GotFocus()
   InverseTextBox textTM86
End Sub

Private Sub textTM90_GotFocus()
   OpenIme
   InverseTextBox textTM90
End Sub

'Add by Amy 2018/09/14
'Modify by Amy 2021/12/23
Private Sub textTM90_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ChangeZIP(KeyAscii, textTM90)
End Sub

Private Sub textTM90_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM90, textTM90.MaxLength) = False Then
      Cancel = True
      textTM90_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM83_GotFocus()
   OpenIme
   InverseTextBox textTM83
End Sub

Private Sub textTM83_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM83, textTM83.MaxLength) = False Then
      Cancel = True
      textTM83_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM87_GotFocus()
   InverseTextBox textTM87
End Sub

Private Sub textTM91_GotFocus()
   OpenIme
   InverseTextBox textTM91
End Sub

'Add by Amy 2018/09/14
'Modify by Amy 2021/12/23
Private Sub textTM91_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ChangeZIP(KeyAscii, textTM91)
End Sub

Private Sub textTM91_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM91, textTM91.MaxLength) = False Then
      Cancel = True
      textTM91_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM84_GotFocus()
   OpenIme
   InverseTextBox textTM84
End Sub

Private Sub textTM84_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM84, textTM84.MaxLength) = False Then
      Cancel = True
      textTM84_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM88_GotFocus()
   InverseTextBox textTM88
End Sub

Private Sub textTM92_GotFocus()
   OpenIme
   InverseTextBox textTM92
End Sub

'Add by Amy 2018/09/14
'Modify by Amy 2021/12/23
Private Sub textTM92_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ChangeZIP(KeyAscii, textTM92)
End Sub

Private Sub textTM92_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM92, textTM92.MaxLength) = False Then
      Cancel = True
      textTM92_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM85_GotFocus()
   OpenIme
   InverseTextBox textTM85
End Sub

Private Sub textTM85_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM85, textTM85.MaxLength) = False Then
      Cancel = True
      textTM85_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM89_GotFocus()
   InverseTextBox textTM89
End Sub

Private Sub textTM93_GotFocus()
   OpenIme
   InverseTextBox textTM93
End Sub

'Add by Amy 2018/09/14
'Modify by Amy 2021/12/23
Private Sub textTM93_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = ChangeZIP(KeyAscii, textTM93)
End Sub

Private Sub textTM93_Validate(Cancel As Boolean)
   Cancel = False
   If CheckLengthIsOK(textTM93, textTM93.MaxLength) = False Then
      Cancel = True
      textTM93_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

'Add by Amy 2018/09/14
Private Function ChgPField(ByVal TMField As String) As String
    Dim jj As Integer
    
    If m_TM01 <> "P" Then
        ChgPField = TMField
        Exit Function
    End If
    For jj = LBound(TField) To UBound(TField)
       If UCase(TField(jj)) = UCase(TMField) Then
            ChgPField = UCase(PField(jj))
            Exit For
       End If
    Next jj
End Function

