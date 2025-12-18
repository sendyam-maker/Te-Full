VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030604 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內公報代理人合併作業"
   ClientHeight    =   2145
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6045
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4980
      TabIndex        =   4
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   4020
      TabIndex        =   3
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTMBM06_New_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1020
      Width           =   1272
   End
   Begin VB.TextBox textTMBM06_Old_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   660
      Width           =   1272
   End
   Begin VB.TextBox textTMBM07 
      Height          =   264
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1380
      Width           =   975
   End
   Begin MSForms.TextBox textTMBM06_New_1 
      Height          =   300
      Left            =   2040
      TabIndex        =   1
      Top             =   1020
      Width           =   2472
      VariousPropertyBits=   679493659
      MaxLength       =   12
      Size            =   "4360;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM06_Old_1 
      Height          =   300
      Left            =   2040
      TabIndex        =   0
      Top             =   660
      Width           =   2472
      VariousPropertyBits=   679493659
      MaxLength       =   12
      Size            =   "4360;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "合併起始公報卷期 :"
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   1380
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "新代理人名稱 :"
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   1020
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "原代理人名稱 :"
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   660
      Width           =   1332
   End
End
Attribute VB_Name = "frm030604"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/10 Form2.0已修改 textTMBM06_Old_1/textTMBM06_New_1
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 使用者按下離開按紐
Private Sub cmdExit_Click()
   Unload Me
End Sub
' 使用者按下查詢按紐
Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nAffect As Long
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 執行作業
      nAffect = Process
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      If nAffect <= 0 Then
         strTit = "檢核資料"
         strMsg = "沒有符合條件的資料可更新!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Else
         strTit = "檢核資料"
         strMsg = "此公報資料已更新完畢!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         ' 清除畫面欄位
         textTMBM06_Old_1 = Empty
         textTMBM06_Old_2 = Empty
         textTMBM06_New_1 = Empty
         textTMBM06_New_2 = Empty
         textTMBM07 = Empty
      End If
      
      textTMBM06_Old_1.SetFocus
      
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   textTMBM06_Old_2.BackColor = &H8000000F
   textTMBM06_New_2.BackColor = &H8000000F
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030604 = Nothing
End Sub

' 原代理人名稱
Private Sub textTMBM06_Old_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Cancel = False
   If IsEmptyText(textTMBM06_Old_1) = False Then
      If StrLength(textTMBM06_Old_1) > 12 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "原代理人名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM06_Old_1_GotFocus
         GoTo EXITSUB
      End If
      
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM TAGENT " & _
               "WHERE TA01 = 'T' AND " & _
                     "TA03 = '" & textTMBM06_Old_1 & "'"
      ' 查詢資料庫
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         textTMBM06_Old_2 = rsTmp.Fields("TA02")
      Else
         Cancel = True
         strTit = "檢核資料"
         strMsg = "原代理人名稱不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM06_Old_1_GotFocus
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
EXITSUB:
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTMBM06_Old_1.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 新代理人名稱
Private Sub textTMBM06_New_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Cancel = False
   If IsEmptyText(textTMBM06_New_1) = False Then
      If StrLength(textTMBM06_Old_1) > 12 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "新代理人名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM06_New_1_GotFocus
         GoTo EXITSUB
      End If
      
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM TAGENT " & _
               "WHERE TA01 = 'T' AND " & _
                     "TA03 = '" & textTMBM06_New_1 & "'"
      ' 查詢資料庫
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         textTMBM06_New_2 = rsTmp.Fields("TA02")
      Else
         Cancel = True
         strTit = "檢核資料"
         strMsg = "新代理人名稱不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM06_New_1_GotFocus
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
EXITSUB:
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTMBM06_New_1.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 合併起始公報卷期
Private Sub textTMBM07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTMBM07) = False Then
      If IsNumeric(textTMBM07) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "合併起始公報卷期只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM07_GotFocus
      End If
   End If
End Sub

' 更新作業
Private Function Process() As Long
   Dim strSql As String
   Dim nAffect As Long
   
   strSql = "UPDATE TMBULLETIN SET TMBM06 = '" & textTMBM06_New_1 & "' " & _
            "WHERE TMBM07 >= '" & textTMBM07 & "' AND " & _
                  "TMBM06 = '" & textTMBM06_Old_1 & "' "
   
   cnnConnection.Execute strSql, nAffect
   Process = nAffect
End Function

' 檢查輸入資料是否正確
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   ' 原代理人名稱
   If IsEmptyText(textTMBM06_Old_1) = True Then
      strTit = "檢核資料"
      strMsg = "請輸原代理人名稱"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTMBM06_Old_1.SetFocus
      GoTo EXITSUB
   End If
   ' 新代理人名稱
   If IsEmptyText(textTMBM06_New_1) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入新代理人名稱"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTMBM06_New_1.SetFocus
      GoTo EXITSUB
   End If
   ' 原代理人名稱與新代理人名稱不可相同
   If textTMBM06_Old_1 = textTMBM06_New_1 Then
      strTit = "檢核資料"
      strMsg = "原代理人名稱與新代理人名稱不可相同"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTMBM06_Old_1.SetFocus
      GoTo EXITSUB
   End If
   ' 合併起始公報卷期不可空白
   If IsEmptyText(textTMBM07) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入合併起始公報卷期"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTMBM07.SetFocus
      GoTo EXITSUB
   End If
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textTMBM06_Old_1_GotFocus()
   InverseTextBox textTMBM06_Old_1
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTMBM06_Old_1.IMEMode = 1
   OpenIme
End Sub

Private Sub textTMBM06_New_1_GotFocus()
   InverseTextBox textTMBM06_New_1
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTMBM06_New_1.IMEMode = 1
   OpenIme
End Sub

Private Sub textTMBM07_GotFocus()
   InverseTextBox textTMBM07
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textTMBM06_New_1.Enabled = True Then
   Cancel = False
   textTMBM06_New_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTMBM06_Old_1.Enabled = True Then
   Cancel = False
   textTMBM06_Old_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

