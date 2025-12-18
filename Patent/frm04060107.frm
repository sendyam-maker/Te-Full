VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060107 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內公報代理人換事務所作業"
   ClientHeight    =   1965
   ClientLeft      =   30
   ClientTop       =   1935
   ClientWidth     =   6165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   6165
   Begin VB.CommandButton bottonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   4464
      TabIndex        =   2
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton buttonExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   5292
      TabIndex        =   3
      Top             =   36
      Width           =   800
   End
   Begin VB.TextBox text03 
      Height          =   270
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1560
      Width           =   1572
   End
   Begin VB.TextBox text01_01 
      Height          =   270
      Left            =   2040
      MaxLength       =   4
      TabIndex        =   0
      Top             =   720
      Width           =   1572
   End
   Begin MSForms.TextBox text02 
      Height          =   300
      Left            =   2040
      TabIndex        =   5
      Top             =   1140
      Width           =   1605
      VariousPropertyBits=   671105055
      Size            =   "2831;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox text01_02 
      Height          =   300
      Left            =   3720
      TabIndex        =   4
      Top             =   720
      Width           =   2295
      VariousPropertyBits=   671105055
      Size            =   "4048;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "換事務所起始日期 :"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "事務所名稱 :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "代理人 :"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frm04060107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/23 改成Form2.0 (text01_02,text02)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Private Sub Form_Load()
   MoveFormToCenter Me
   UpdateState
End Sub

Private Sub bottonOK_Click()
   If CheckDataValid() = True Then
      ExecuteTransf
      Unload Me
   End If
End Sub

Private Sub buttonExit_Click()
   Unload Me
End Sub

Public Sub ExecuteTransf()
Dim strSql As String
Dim strData As String
Dim rsTmp As ADODB.Recordset
   
On Error GoTo ErrorHandler
    cnnConnection.BeginTrans
    ' 讀取公報代理人檔以取得代理人的事務所名稱
    strSql = "SELECT * FROM Tagent WHERE TA01='P' AND TA02 = '" & text01_01 & "'"
    Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenDynamic
    If rsTmp.RecordCount <= 0 Then
       GoTo EXITSUB
    End If
    rsTmp.MoveFirst
    strData = "" & rsTmp.Fields("TA04")
    strSql = "UPDATE TPBulletin SET TPB08 = '" & strData & "' " & _
             "WHERE TPB07 = '" & text01_01 & "'"
    If IsEmpty(text03) = False Then
       strSql = strSql & " AND TPB03 >= " & ChangeTStringToWString(text03)
    End If
    cnnConnection.Execute strSql
    'Add By Cheng 2003/05/15
    '更新公開公報檔
    strData = "" & rsTmp.Fields("TA04")
    strSql = "UPDATE TPGazette SET TPG08 = '" & strData & "' " & _
             "WHERE TPG07 = '" & text01_01 & "'"
    If IsEmpty(text03) = False Then
       strSql = strSql & " AND TPG03 >= " & ChangeTStringToWString(text03)
    End If
    cnnConnection.Execute strSql
    cnnConnection.CommitTrans
EXITSUB:
   rsTmp.Close
   Set rsTmp = Nothing
Exit Sub
ErrorHandler:
   rsTmp.Close
   Set rsTmp = Nothing
    cnnConnection.RollbackTrans
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Add By Cheng 2002/07/18
Set frm04060107 = Nothing
End Sub

Private Sub text01_01_GotFocus()
  TextInverse text01_01
End Sub

Private Sub text01_01_KeyPress(KeyAscii As Integer)
   KeyAscii = UCase(KeyAscii)
End Sub

Private Sub text01_01_Validate(Cancel As Boolean)
   Dim rsTmp As ADODB.Recordset
   
   Set rsTmp = New ADODB.Recordset
   strSql = "Select * from TAGENT where TA01='P' AND TA02 = '" & text01_01 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      text01_02 = rsTmp.Fields("TA03")
      text02 = rsTmp.Fields("TA04")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub text03_GotFocus()
  TextInverse text03
End Sub

Private Sub text03_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmpty(text03) = False Then
      If CheckIsTaiwanDate(text03, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的公告日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

Public Sub UpdateState()
   text01_02.BackColor = &H8000000F
   text02.BackColor = &H8000000F
End Sub

Public Function IsEmpty(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmpty = False
   
   If Len(strData) <= 0 Then
      IsEmpty = True
   Else
      IsEmpty = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmpty = False
            Exit For
         End If
      Next nIndex
   End If
End Function

' 檢核資料是否正確
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   ' 代理人代號不可為空白
   If IsEmpty(text01_01) = True Then
      strMsg = "請輸入代理人代號"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   ' 換事務所起始日期不可為空白
   If IsEmpty(text03) = True Then
      strMsg = "請輸入換事務所起始日期"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function
