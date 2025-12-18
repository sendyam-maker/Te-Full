VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210101 
   BorderStyle     =   1  '單線固定
   Caption         =   "個人客戶資料修改-登入"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3885
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2010
      Style           =   1  '圖片外觀
      TabIndex        =   2
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
      TabIndex        =   3
      Top             =   105
      Width           =   800
   End
   Begin VB.TextBox txtPassword 
      Height          =   270
      IMEMode         =   3  '暫止
      Left            =   1230
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "password"
      Top             =   1050
      Width           =   1575
   End
   Begin VB.TextBox txtUserNo 
      Height          =   270
      Left            =   1230
      MaxLength       =   5
      TabIndex        =   0
      Top             =   675
      Width           =   732
   End
   Begin MSForms.Label lblUserName 
      Height          =   285
      Left            =   2130
      TabIndex        =   6
      Top             =   675
      Width           =   1350
      VariousPropertyBits=   27
      Caption         =   "lblUserName"
      Size            =   "2381;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "密碼"
      Height          =   270
      Index           =   1
      Left            =   375
      TabIndex        =   5
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號"
      Height          =   270
      Index           =   0
      Left            =   375
      TabIndex        =   4
      Top             =   675
      Width           =   720
   End
End
Attribute VB_Name = "frm210101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/14 Form2.0已修改 lblUserName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
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
   Select Case Index
      Case 0   '確定
         '部門代碼
         Dim strDept As String
         Dim frmTmp As Form
         If IsAuthorized(txtUserNo, txtPassword) = True And getSalesDept(txtUserNo, strDept) = True Then
            'Modify by Morgan 2004/4/21
            'frm210101_1.setSalesNo txtUserNo
            'frm210101_1.setDeptNo strDept
            'frm210101_1.Show
            
            Select Case stNextForm
               Case "frm210101_1"
                  Set frmTmp = frm210101_1
               Case "frm210102"
                  Set frmTmp = frm210102
            End Select
            
            frmTmp.setSalesNo txtUserNo
            frmTmp.setDeptNo strDept
            frmTmp.Show
            
            'Modify end
            
            Unload Me
         Else
            txtUserNo.SetFocus
         End If
         
      Case 1   '結束
      
'         Modify by Morgan 2003/12/16
'         If PUB_CheckFormExist(frmCaller.Name) = True Then
'            frmCaller.Show
'         End If

         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtUserNo = strUserNum
   lblUserName = strUserName
   txtPassword = ""
End Sub

Public Function getSalesDept(ByRef strID As String, ByRef strDept As String) As Boolean
   Dim strSql As String, rsQuery As New ADODB.Recordset
      
   strSql = "Select ST15 From STAFF where ST01='" & strID & "'"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      strDept = "" & rsQuery.Fields(0)
      getSalesDept = True
   Else
      MsgBox "無法取得所屬部門！", vbCritical
   End If
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
End Function

Public Function IsAuthorized(ByVal strID As String, Optional ByVal strPWD As String, Optional ByVal bolChkPwd As Boolean = True) As Boolean
   Dim strSql As String, rsQuery As New ADODB.Recordset, strSQLpwd As String
      
   IsAuthorized = False
   strSQLpwd = ""
   If bolChkPwd = True Then
      If (strPWD = "" Or strID = strPWD) Then
         MsgBox "密碼不可空白或與員工代號相同！", vbCritical
         Exit Function
      Else
         strSQLpwd = " and sp03='" & Encrypt(strPWD, True) & "'"
      End If
   End If
   
   strSql = "Select 1 as Msg From staff_pwd where sp03<>'" & Encrypt(strID, True) & "' and sp03 is not null" & _
      " and sp01='" & strID & "'" & strSQLpwd
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      IsAuthorized = True
   ElseIf (bolChkPwd = True) Then
      MsgBox "登入失敗，請重新輸入！"
   End If
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm210101 = Nothing
End Sub

Private Sub txtPassword_GotFocus()
   TextInverse txtPassword
End Sub

Private Sub txtUserNo_GotFocus()
   TextInverse txtUserNo
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

'Add By Sindy 2010/11/26
Private Sub txtUserNo_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_LostFocus()
   lblUserName = getUserName(txtUserNo)
End Sub
