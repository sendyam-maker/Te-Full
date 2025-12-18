VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040209_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "CFP申請文件齊備維護"
   ClientHeight    =   4110
   ClientLeft      =   180
   ClientTop       =   840
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   9330
   Begin VB.TextBox TextCP143 
      Height          =   270
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   0
      Top             =   3645
      Width           =   1215
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   6300
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   7128
      TabIndex        =   2
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton CmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   8256
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1155
      TabIndex        =   18
      Top             =   1170
      Width           =   7830
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "13811;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblcp05 
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "收  文  日："
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請文件齊備日："
      Height          =   180
      Index           =   39
      Left            =   240
      TabIndex        =   21
      Top             =   3645
      Width           =   1440
   End
   Begin MSForms.Label LbeEngName 
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   2685
      Width           =   1815
      VariousPropertyBits=   27
      Size            =   "3201;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "智權人員："
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2685
      Width           =   975
   End
   Begin MSForms.Label LbeSalesName 
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   2685
      Width           =   1695
      VariousPropertyBits=   27
      Size            =   "2990;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "分所案號："
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   735
      Width           =   900
   End
   Begin VB.Label lbePA47 
      Height          =   255
      Left            =   5010
      TabIndex        =   14
      Top             =   735
      Width           =   2055
   End
   Begin MSForms.Label lbeCusName 
      Height          =   255
      Left            =   2430
      TabIndex        =   13
      Top             =   1680
      Width           =   6675
      VariousPropertyBits=   27
      Size            =   "11774;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbeCustomer 
      Height          =   255
      Left            =   1260
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "申  請  人："
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lbePropertyName 
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   2205
      Width           =   2745
   End
   Begin VB.Label Label13 
      Caption         =   "案件性質："
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   2205
      Width           =   975
   End
   Begin VB.Label lbeCaseNum 
      Height          =   255
      Left            =   1230
      TabIndex        =   8
      Top             =   735
      Width           =   2055
   End
   Begin VB.Label lbeNum 
      Height          =   255
      Left            =   1290
      TabIndex        =   7
      Top             =   2205
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "承  辦  人："
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   2685
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號：  "
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   735
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號：    "
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   2205
      Width           =   975
   End
End
Attribute VB_Name = "frm040209_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/3 改成Form2.0 (cboCaseName,lbeCusName,LbeSalesName,LbeEngName)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2011/5/20 CREATE BY SONIA
Option Explicit

Public UpForm As Form
Dim rs As New ADODB.Recordset, strCP09() As String, t As Integer
Dim blnIsSave As Boolean
Dim m_CP09 As String
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_CP05 As String

Private Sub cmdBack_Click()
Dim yn As Integer
   
   If blnIsSave = False Then
      yn = MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2)
      If yn = 7 Then
         Exit Sub
      End If
   End If
   Me.Hide
   UpForm.Show
   Unload Me
End Sub

Private Sub cmdEnd_Click()
Dim yn As Integer
   
   If blnIsSave = False Then
      yn = MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2)
      If yn = 7 Then
         Exit Sub
      End If
   End If
   
   Me.Hide
   Unload UpForm
   Unload Me

End Sub

Private Sub cmdSure_Click()
Dim strDay1 As String
Dim strDay2 As String
Dim strDate As String

   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   Screen.MousePointer = 11
   If Not SaveData Then
      MsgBox "存檔失敗,請洽系統管理者", vbCritical
   Else
      Unload frm040209_1
      Set frm040209_1 = Nothing
      'Unload frm040209
      frm040209.Show
      frm040209.cmdSearch_Click
   End If
   Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Dim i As Integer, n As Integer
  
   MoveFormToCenter Me
   blnIsSave = False
End Sub

Sub GetData(ByVal Init As Integer)
Dim i As Integer
Dim strName As String
 
   m_CP09 = Mid(frm040209_1.Tag, 1, 9)
   '收文號
   lbeNum = Mid(frm040209_1.Tag, 1, 9)
   cboCaseName.Clear
   'cboCaseName.Text = "" 'Removed by Morgan 2022/1/3
   lblcp05 = ""
   
   strExc(1) = "select * from caseprogress,patent where cp09='" + m_CP09 + _
       "' and CP01=pa01(+) AND CP02=pa02(+) AND CP03=pa03(+) AND CP04=pa04(+) "
   intI = 1
   Set rs = ClsLawReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      ' 本所案號
      m_CP01 = rs.Fields!cp01
      m_CP02 = rs.Fields!cp02
      m_CP03 = rs.Fields!cp03
      m_CP04 = rs.Fields!cp04
      lbeCaseNum = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
      ' 收文日
      m_CP05 = rs.Fields!cp05
      lblcp05 = ChangeWStringToTDateString(m_CP05)
      ' 分所案號
      lbePA47 = "" & rs.Fields!PA47
      ' 案件名稱
      AddCboName cboCaseName, "" & rs.Fields!pa05, "" & rs.Fields!pa06, "" & rs.Fields!pa07
      ' 申請人
      lbeCustomer = "" & rs.Fields!pa26
      lbeCusName = GetCustomerName("" & rs.Fields!pa26, 0)
      '智權人員
      strName = ""
      If ClsPDGetStaffN("" & rs.Fields!cp13, strName) Then
         LbeSalesName = strName
      End If
      '案件性質
      strName = ""
      If ClsPDGetCaseProperty(m_CP01, "" & rs.Fields!CP10, strName, False) Then
         lbePropertyName = strName
      End If
      '承辦人
      strName = "": LbeEngName = ""
      If Not IsNull(rs.Fields("CP14")) Then
         If ClsPDGetStaffN("" & rs.Fields("CP14"), strName) Then
            LbeEngName = strName
         End If
      End If
      ' 申請文件齊備日
      textCP143 = ChangeWStringToTString("" & rs.Fields("CP143"))
      
      If textCP143 = "" Then textCP143 = strSrvDate(2)  '2011/6/3 add by sonia 預設系統日
      
   End If
   
End Sub

Private Function SaveData() As Boolean
Dim strNewNum As String, strNum As String, strTemp As String
Dim i As Integer
Dim strNP22 As String
   
On Error GoTo ErrorHandler
   
   SaveData = True
   cnnConnection.BeginTrans
   'Modified by Morgan 2019/5/23 開放電腦中心可以取消齊備日
   'strExc(1) = "update caseprogress set cp143=" & DBDATE(textCP143) & " where cp09='" & m_CP09 & "' "
   strExc(1) = "update caseprogress set cp143=" & CNULL(DBDATE(textCP143)) & " where cp09='" & m_CP09 & "' "
   cnnConnection.Execute strExc(1)
    
   If SaveData Then blnIsSave = True Else blnIsSave = False

   cnnConnection.CommitTrans
   
   Exit Function

ErrorHandler:
   cnnConnection.RollbackTrans
   SaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm040209_1 = Nothing
End Sub

Private Sub textCP143_GotFocus()
   TextInverse textCP143
   CloseIme
End Sub

Private Sub TextCP143_Validate(Cancel As Boolean)
   
   If textCP143 = "" Then
      'Modified by Morgan 2019/5/23 開放電腦中心可以取消齊備日
      If Pub_StrUserSt03 <> "M51" Then
         MsgBox "申請文件齊備日不可空白!"
         Cancel = True
      End If
   Else
      If ChkDate(textCP143) = False Then
         Cancel = True
      ElseIf DBDATE(textCP143) < m_CP05 Then
         MsgBox "申請文件齊備日不可小於收文日!"
         Cancel = True
      End If
   End If
   
   If Cancel Then TextInverse textCP143
   
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   Cancel = False
   TextCP143_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   
   TxtValidate = True
End Function
