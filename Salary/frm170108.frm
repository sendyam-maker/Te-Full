VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm170108 
   BorderStyle     =   1  '單線固定
   Caption         =   "薪資查詢密碼修改"
   ClientHeight    =   3240
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4092
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4092
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2010
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   105
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   2850
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   105
      Width           =   800
   End
   Begin VB.TextBox txtPassword 
      Height          =   270
      IMEMode         =   3  '暫止
      Left            =   1230
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Text            =   "password"
      Top             =   1170
      Width           =   1575
   End
   Begin VB.TextBox setSalesNo 
      Height          =   270
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   3
      Top             =   750
      Width           =   732
   End
   Begin MSForms.Label lblUserName 
      Height          =   300
      Left            =   2130
      TabIndex        =   7
      Top             =   795
      Width           =   1188
      Caption         =   "lblUserName"
      Size            =   "2096;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   $"frm170108.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   1080
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   3300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "密碼"
      Height          =   270
      Index           =   1
      Left            =   375
      TabIndex        =   5
      Top             =   1170
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號"
      Height          =   270
      Index           =   0
      Left            =   375
      TabIndex        =   4
      Top             =   795
      Width           =   720
   End
End
Attribute VB_Name = "frm170108"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/10 Form2.0已修改(lblUserName)
'2015/12/21 Create by sonia
Option Explicit
'呼叫表單
Dim frmCaller As Form
'下一表單名
Dim stNextForm As String

Public Sub setNextForm(stFromName As String)
   stNextForm = stFromName
End Sub

Public Sub setCaller(frmFrom As Form)
   Set frmCaller = frmFrom
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim frmTmp As Form
   
   Select Case Index
      Case 0   '確定
         '部門代碼
         If IsAuthorized(setSalesNo, txtPassword) = True Then
            Select Case stNextForm
               Case "frm170236"
                  Set frmTmp = frm170236
                     Case "frm170237"
                        Set frmTmp = frm170237
                     Case "frm170238"
                        Set frmTmp = frm170238
                     Case "frm170239"
                        Set frmTmp = frm170239
               Case "frm170108"
                  PUB_GetSalaryData  '讀取薪資查詢系統初始時相關變數值
                  Unload Me
                  Exit Sub
            End Select
            
            PUB_GetSalaryData  '讀取薪資查詢系統初始時相關變數值
            
            frmTmp.Show
            Unload Me
         Else
            txtPassword_GotFocus
         End If
         
      Case 1   '結束
      
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtPassword = ""
End Sub

Public Function IsAuthorized(ByVal strID As String, Optional ByVal strPWD As String) As Boolean
Dim strSql As String, rsQuery As New ADODB.Recordset, strSQLpwd As String
Dim oMailCount As String
      
   IsAuthorized = False
   strSQLpwd = ""
   
   If (strPWD = "" Or strID = strPWD) Then
      MsgBox "密碼不可空白或與員工代號相同！", vbCritical
      Exit Function
   End If
   
   strSql = "Select st26,sp02,sp03 From staff,staff_pwd where st01='" & strID & "' and sp01='" & strID & "'"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      If "" & rsQuery.Fields("St26") = strPWD Then
         MsgBox "密碼不可與身份證字號相同！", vbCritical
         If rsQuery.State <> adStateClosed Then rsQuery.Close
         Set rsQuery = Nothing
         Exit Function
      End If
      If "" & rsQuery.Fields("sp03") <> "" Then
         If Encrypt(rsQuery.Fields("sp03"), False) = strPWD Then
            MsgBox "密碼不可與案件系統密碼相同！", vbCritical
            If rsQuery.State <> adStateClosed Then rsQuery.Close
            Set rsQuery = Nothing
            Exit Function
         End If
      End If
      If "" & rsQuery.Fields("sp02") <> "" Then
         If Encrypt(rsQuery.Fields("sp02"), False) = strPWD Then
            MsgBox "密碼不可與Windows登入密碼號相同！", vbCritical
            If rsQuery.State <> adStateClosed Then rsQuery.Close
            Set rsQuery = Nothing
            Exit Function
         End If
      End If
         
      '新增薪資密碼檔
      strSql = "update staff_pwd set sp04='" & Encrypt(strPWD, True) & "' where sp01='" & strID & "'"
      cnnConnection.Execute strSql
      MsgBox "薪資查詢密碼己設定完成！"
      IsAuthorized = True
   Else
      MsgBox "員工代號資料或密碼資料有誤，請通知電腦中心！", vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing

End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm170108 = Nothing
End Sub

Private Sub txtPassword_GotFocus()
   TextInverse txtPassword
   CloseIme
End Sub

Private Sub setSalesNo_GotFocus()
   TextInverse setSalesNo
End Sub

Private Function getUserName(strID As String) As String
Dim strSql As String, rsQuery As New ADODB.Recordset
      
   strSql = "Select ST02 From STAFF where ST01='" & strID & "'"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      getUserName = "" & rsQuery.Fields(0)
   Else
      getUserName = ""
   End If
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
End Function

Private Sub setSalesNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub setSalesNo_LostFocus()
   lblUserName = getUserName(setSalesNo)
End Sub
