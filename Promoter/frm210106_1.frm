VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210106_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "簽收資料查詢-智權人員登入"
   ClientHeight    =   1596
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3888
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1596
   ScaleWidth      =   3888
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1755
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   90
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   2790
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   90
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
      Top             =   1050
      Width           =   1575
   End
   Begin VB.TextBox txtUserNo 
      Height          =   270
      Left            =   1230
      MaxLength       =   5
      TabIndex        =   3
      Top             =   675
      Width           =   732
   End
   Begin MSForms.Label lblUserName 
      Height          =   300
      Left            =   2040
      TabIndex        =   6
      Top             =   675
      Width           =   1470
      VariousPropertyBits=   27
      Size            =   "2593;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
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
Attribute VB_Name = "frm210106_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/07/12 將表單設"MDI子表單呈現" => MDIChild = False ; 因為frm210115會呼叫來檢查權限
'Memo by Lydia 2022/01/03 改成Form2.0 ; lblUserName
'Memo by Lydia 2021/07/12 將表單設"MDI子表單呈現" => MDIChild = True
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

'呼叫表單
Dim frmCaller As Form
Dim m_PrevForm As Form 'Added by Lydia 2021/07/27 記錄呼叫的表單; 智權-調整財務系統(20200909) 'Memo by Lydia 2021/08/27 上線

'Added by Lydia 2017/01/26 是否需要輸入密碼
Public Function setNextForm(Optional ByVal pPrevFrm As String) As String
   If Pub_StrUserSt03 = "M31" Or Pub_strUserST05 = "CM" Or Pub_strUserST05 = "C1" Or Pub_strUserST05 = "NM" Or Pub_strUserST05 = "KM" Then
      setNextForm = ""
   Else
      setNextForm = Me.Name
   End If
End Function

'Modified by Lydia 2021/07/27  記錄呼叫的表單; 智權-調整財務系統(20200909)
Public Sub setCaller(frmFrom As Form, Optional ByVal frmPrev As Form)
   Set frmCaller = frmFrom
   'Added by Lydia 2021/07/27
   If TypeName(frmPrev) <> "Nothing" Then
       Set m_PrevForm = frmPrev
   End If
   'end 2021/07/27
End Sub

Private Sub cmdOK_Click(Index As Integer)

   Select Case Index
      Case 0   '確定
         If IsAuthorized(txtUserNo, txtPassword) = True Then
            frmCaller.Tag = txtUserNo
            'Added by Lydia 2023/12/28
            If UCase(TypeName(m_PrevForm)) = "FRM210115" Then
               m_PrevForm.Tag = txtUserNo
               m_PrevForm.bolRunPWD = True
               m_PrevForm.Show
               m_PrevForm.PubShowNextData
            Else
            'end 2023/12/28
               'Added by Lydia 2021/07/27 記錄呼叫的表單; 智權-調整財務系統(20200909)
               If TypeName(m_PrevForm) <> "Nothing" Then
                   Call frmCaller.SetParent(m_PrevForm)
               End If
               'end 2021/07/27
            End If
            Unload Me
         Else
            txtUserNo.SetFocus
         End If
      Case 1   '結束
        'Added by Lydia 2021/07/27 記錄呼叫的表單; 智權-調整財務系統(20200909)
        If TypeName(m_PrevForm) <> "Nothing" Then
            m_PrevForm.Show
        End If
        'end 2021/07/27
         Unload Me
   End Select
   
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   'edit by nickc 2007/09/28 加入預設
   'txtuserno=""
   'lblUserName = ""
   txtUserNo = strUserNum
   txtUserNo_LostFocus
   
   txtPassword = ""
   
End Sub

Public Function IsAuthorized(ByVal strID As String, Optional ByVal strPWD As String, Optional ByVal bolChkPwd As Boolean = True) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, strSQLpwd As String
      
   IsAuthorized = False
   strSQLpwd = ""
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
   MenuEnabled 'Added by Lydia 2017/01/26
   Set frm210106_1 = Nothing
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
