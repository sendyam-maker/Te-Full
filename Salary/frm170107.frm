VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm170107 
   BorderStyle     =   1  '單線固定
   Caption         =   "薪資查詢-登入"
   ClientHeight    =   2592
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3888
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2592
   ScaleWidth      =   3888
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "忘記密碼"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   360
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   120
      Width           =   1000
   End
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
   Begin VB.TextBox txtUserNo 
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
      TabIndex        =   9
      Top             =   795
      Width           =   1188
      Caption         =   "lblUserName"
      Size            =   "2096;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "　　按「忘記密碼」鈕會將密碼設成　　「身份證字號」！"
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   360
      TabIndex        =   8
      Top             =   2130
      Width           =   3000
   End
   Begin VB.Label Label2 
      Caption         =   "PS：第一次登錄時請使用身份證字號　　，隨後立即要設定新密碼！"
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   360
      TabIndex        =   6
      Top             =   1656
      Width           =   3000
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "密碼"
      Height          =   276
      Index           =   1
      Left            =   372
      TabIndex        =   5
      Top             =   1152
      Width           =   720
   End
   Begin VB.Label Label4 
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
Attribute VB_Name = "frm170107"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/10 Form2.0已修改(lblUserName)
'2015/12/18 Create by sonia
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
Dim strSql As String, rsQuery As New ADODB.Recordset
Dim frmTmp As Form
Dim intExe As Integer 'Add by Amy 2020/09/29
   
   Select Case Index
      Case 0   '確定
         '先檢查薪資密碼檔
         If (txtPassword = "" Or txtUserNo = txtPassword) Then
            MsgBox "密碼不可空白或與員工代號相同！", vbCritical
            Exit Sub
         End If
      
         strSql = "Select sp04,st26 From staff_pwd,staff where sp01='" & txtUserNo & "' and st01='" & txtUserNo & "'"
         rsQuery.CursorLocation = adUseClient
         rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsQuery.RecordCount > 0 Then
            '已設密碼
            If "" & rsQuery.Fields("sp04") <> "" Then
               If txtPassword = Encrypt(rsQuery.Fields("sp04"), False) Then
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
                        frm170108.setSalesNo = txtUserNo
                        frm170108.setNextForm stNextForm
                        Set frmTmp = frm170108
                        frm170108.lblUserName = getUserName(txtUserNo)  'add by sonia 2021/12/10
                  End Select
                  
                  If stNextForm <> "frm170108" Then PUB_GetSalaryData  '讀取薪資查詢系統初始時相關變數值
                  
                  frmTmp.Show
                  Unload Me
                  Exit Sub
               Else
                  MsgBox "登入失敗，請重新輸入！", vbCritical
               End If
            '未設密碼
            '若為身份證字號則進入密碼修改
            ElseIf rsQuery.Fields("ST26") <> "" And txtPassword = rsQuery.Fields("ST26") Then
               frm170108.setSalesNo = txtUserNo
               frm170108.setNextForm stNextForm
               frm170108.lblUserName = getUserName(txtUserNo)  'add by sonia 2021/12/10
               frm170108.Show
               Unload Me
               Exit Sub
            ElseIf "" & rsQuery.Fields("sp04") = "" Then
               MsgBox "此員工代號尚未設定薪資查詢密碼，請以身份證字號輸入後進行密碼修改！(英文字母要大寫)", vbCritical
            ElseIf rsQuery.Fields("ST26") = "" Then
               MsgBox "此員工代號未輸入身份證字號，請通知人事處輸入！", vbCritical
            Else
               MsgBox "登入失敗，請重新輸入！", vbCritical
            End If
         Else
            MsgBox "無此員工代號密碼資料，請通知電腦中心設定！"
         End If
         If rsQuery.State <> adStateClosed Then rsQuery.Close
         txtPassword_GotFocus
      Case 1   '結束
         Unload Me
      'Add by Amy 2020/09/29
      Case 2 '忘記密碼
        If txtUserNo = MsgText(601) Then
            MsgBox "請輸入員工號！"
            txtUserNo.SetFocus
            Exit Sub
        End If
        strSql = "Update Staff_Pwd Set sp04=null Where sp01='" & txtUserNo & "' "
        cnnConnection.Execute strSql, intExe
        If intExe > 0 Then
            MsgBox "密碼已還原為初始值，請用身份證號登入"
        End If
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txtUserNo = strUserNum
   lblUserName = strUserName
   txtPassword = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170107 = Nothing
End Sub

Private Sub txtPassword_GotFocus()
   TextInverse txtPassword
   CloseIme
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

Private Sub txtUserNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_LostFocus()
   lblUserName = getUserName(txtUserNo)
End Sub
