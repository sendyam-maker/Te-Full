VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010007_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "聯絡人資料"
   ClientHeight    =   4650
   ClientLeft      =   5580
   ClientTop       =   1860
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   8520
   Begin VB.Frame fraWindow1 
      BorderStyle     =   0  '沒有框線
      Height          =   3972
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   8532
      Begin VB.TextBox txtSystem 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   22
         Top             =   510
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         Height          =   288
         Index           =   0
         Left            =   2220
         MaxLength       =   6
         TabIndex        =   21
         Top             =   510
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         Height          =   288
         Index           =   1
         Left            =   3480
         MaxLength       =   1
         TabIndex        =   20
         Top             =   510
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Enabled         =   0   'False
         Height          =   288
         Index           =   2
         Left            =   3900
         MaxLength       =   2
         TabIndex        =   19
         Top             =   510
         Visible         =   0   'False
         Width           =   492
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   1170
         Width           =   3615
         VariousPropertyBits=   679493659
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   840
         Width           =   1635
         VariousPropertyBits=   679493659
         Size            =   "2884;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   7
         Left            =   1440
         TabIndex        =   7
         Top             =   1170
         Width           =   6825
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "12039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   6
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   6825
         VariousPropertyBits=   679493659
         MaxLength       =   60
         Size            =   "12039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   5
         Left            =   1440
         TabIndex        =   5
         Top             =   2490
         Width           =   6825
         VariousPropertyBits=   679493659
         Size            =   "12039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   4
         Left            =   1440
         TabIndex        =   4
         Top             =   2160
         Width           =   3615
         VariousPropertyBits=   679493659
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   1440
         TabIndex        =   3
         Top             =   1830
         Width           =   3615
         VariousPropertyBits=   679493659
         Size            =   "6376;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   2
         Left            =   1440
         TabIndex        =   2
         Top             =   1500
         Width           =   6825
         VariousPropertyBits=   679493659
         Size            =   "12039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "本 所 案 號："
         Height          =   255
         Left            =   330
         TabIndex        =   23
         Top             =   540
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LabelPA52 
         Caption         =   "聯絡人1(英)："
         Height          =   255
         Left            =   270
         TabIndex        =   12
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label LabelPA51 
         Caption         =   "聯絡人1(中)："
         Height          =   255
         Left            =   270
         TabIndex        =   11
         Top             =   870
         Width           =   1155
      End
      Begin VB.Label LabelSP75 
         Caption         =   "聯絡人2："
         Height          =   255
         Left            =   270
         TabIndex        =   18
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label LabelSP30 
         Caption         =   "聯絡人1："
         Height          =   255
         Left            =   270
         TabIndex        =   17
         Top             =   870
         Width           =   1155
      End
      Begin VB.Label LabelPA56 
         Caption         =   "聯絡人2(日)："
         Height          =   255
         Left            =   270
         TabIndex        =   16
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label LabelPA55 
         Caption         =   "聯絡人2(英)："
         Height          =   255
         Left            =   270
         TabIndex        =   15
         Top             =   2190
         Width           =   1155
      End
      Begin VB.Label LabelPA54 
         Caption         =   "聯絡人2(中)："
         Height          =   255
         Left            =   270
         TabIndex        =   14
         Top             =   1860
         Width           =   1155
      End
      Begin VB.Label LabelPA53 
         Caption         =   "聯絡人1(日)："
         Height          =   255
         Left            =   270
         TabIndex        =   13
         Top             =   1530
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5664
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6492
      TabIndex        =   9
      Top             =   70
      Width           =   1100
   End
End
Attribute VB_Name = "frm010007_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/16 Form2.0已修改 txt1()
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/22 日期欄已修改
Option Explicit

'回傳值
Public strPA51s As String, strPA52s As String, strPA53s As String
Public strPA54s As String, strPA55s As String, strPA56s As String
Public strSP30s As String, strSP75s As String
Public BolOk As Boolean     'True: 確定  False: 取消


Private Sub cmdOK_Click(Index As Integer)
Dim varSaveCursor, strAuto1 As String, strAuto2 As String, i As Integer

If Index = 0 Then
   'Add by Amy 2021/12/16檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Sub
    End If
 
   If LabelPA51.Visible = True Then
      strPA51s = Trim(txt1(0))
      strPA52s = Trim(txt1(1))
      strPA53s = Trim(txt1(2))
      strPA54s = Trim(txt1(3))
      strPA55s = Trim(txt1(4))
      strPA56s = Trim(txt1(5))
   ElseIf LabelSP30.Visible = True Then
      strSP30s = Trim(txt1(6))
      strSP75s = Trim(txt1(7))
   End If
   BolOk = True
Else
   BolOk = False
End If
Me.Hide
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   strPA51s = "": strPA52s = "": strPA53s = ""
   strPA54s = "": strPA55s = "": strPA56s = ""
   strSP30s = "": strSP75s = ""
End Sub

'讀取Patent資料
Public Function ReadPatent() As Boolean
Dim strSql As String, i As Integer
Dim rsRecordset As New ADODB.Recordset

On Error GoTo Err

LabelPA51.Visible = True
LabelPA52.Visible = True
LabelPA53.Visible = True
LabelPA54.Visible = True
LabelPA55.Visible = True
LabelPA56.Visible = True
LabelSP30.Visible = False
LabelSP75.Visible = False
For i = 0 To 7
   txt1(i).Text = ""
   txt1(i).Visible = True
Next i
txt1(6).Visible = False
txt1(7).Visible = False

'If txtSystem.Visible = True Then
'   strSql = "SELECT PA51,PA52,PA53,PA54,PA55,PA56" & _
'                       " From Patent" & _
'                  " WHERE PA01='" & txtSystem & "' AND PA02='" & txtCode(0) & "' AND PA03='" & txtCode(1) & "' AND PA04='" & txtCode(2) & "'"
'   rsRecordset.CursorLocation = adUseClient
'   rsRecordset.Open strSql, cnnConnection
'   If rsRecordset.RecordCount > 0 Then
'      txt1(0) = "" & rsRecordset.Fields(0)
'      txt1(1) = "" & rsRecordset.Fields(1)
'      txt1(2) = "" & rsRecordset.Fields(2)
'      txt1(3) = "" & rsRecordset.Fields(3)
'      txt1(4) = "" & rsRecordset.Fields(4)
'      txt1(5) = "" & rsRecordset.Fields(5)
'   End If
'   rsRecordset.Close
'End If
txt1(0).SetFocus
BolOk = True
Exit Function
Err:
End Function

'讀取ServicePractice資料
Public Function ReadServicePractice() As Boolean
Dim strSql As String, i As Integer
Dim rsRecordset As New ADODB.Recordset

On Error GoTo Err

LabelPA51.Visible = False
LabelPA52.Visible = False
LabelPA53.Visible = False
LabelPA54.Visible = False
LabelPA55.Visible = False
LabelPA56.Visible = False
LabelSP30.Visible = True
LabelSP75.Visible = True
For i = 0 To 7
   txt1(i).Text = ""
   txt1(i).Visible = False
Next i
txt1(6).Visible = True
txt1(7).Visible = True

'If txtSystem.Visible = True Then
'   strSql = "SELECT SP30,SP75" & _
'                       " From ServicePractice" & _
'                  " WHERE SP01='" & txtSystem & "' AND SP02='" & txtCode(0) & "' AND SP03='" & txtCode(1) & "' AND SP04='" & txtCode(2) & "'"
'   rsRecordset.CursorLocation = adUseClient
'   rsRecordset.Open strSql, cnnConnection
'   If rsRecordset.RecordCount > 0 Then
'      txt1(6) = "" & rsRecordset.Fields(0)
'      txt1(7) = "" & rsRecordset.Fields(1)
'   End If
'   rsRecordset.Close
'End If
txt1(6).SetFocus
BolOk = True
Exit Function
Err:
End Function

Private Sub Form_Unload(Cancel As Integer)
Set frm010007_1 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
InverseTextBox txt1(Index)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index).Text)
'切換輸入法
Select Case Index
   Case 0, 3, 6, 7
      OpenIme
   Case Else
      CloseIme
End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
'關閉輸入法
CloseIme
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   'Added by Lydia 2017/06/14 設欄位長度
    Dim iLen As Integer
    Select Case Index
    Case 0, 3 '專利-聯絡人中文
         iLen = 30
    Case 1, 4 '聯絡人英文
         iLen = 35
    Case 2, 5  '聯絡人日文
         iLen = 60
    Case Else  '服務業務
         iLen = txt1(Index).MaxLength
    End Select
    'end 2017/06/14
    
'Modified by Lydia 2017/06/14
'If CheckLengthIsOK(txt1(Index), txt1(Index).MaxLength) = False Then
If CheckLengthIsOK(txt1(Index), iLen) = False Then
   Call txt1_GotFocus(Index)
   Cancel = True
   Exit Sub
End If
CloseIme
End Sub
