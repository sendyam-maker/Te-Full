VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071020 
   BorderStyle     =   1  '單線固定
   Caption         =   "信件退回"
   ClientHeight    =   4140
   ClientLeft      =   600
   ClientTop       =   960
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7935
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6900
      TabIndex        =   4
      Top             =   90
      Width           =   760
   End
   Begin VB.TextBox txtDay 
      Height          =   288
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2976
      Width           =   1092
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5772
      TabIndex        =   3
      Top             =   90
      Width           =   1100
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4944
      TabIndex        =   2
      Top             =   90
      Width           =   800
   End
   Begin MSForms.Label lbeAccpet 
      Height          =   285
      Left            =   1125
      TabIndex        =   22
      Top             =   2597
      Width           =   5220
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "9208;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCP64 
      Height          =   585
      Left            =   1140
      TabIndex        =   1
      Top             =   3360
      Width           =   6720
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "11853;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   285
      Left            =   1125
      TabIndex        =   21
      Top             =   1849
      Width           =   6465
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11404;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   285
      Left            =   2280
      TabIndex        =   20
      Top             =   1475
      Width           =   5565
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "9816;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "退件原因："
      Height          =   255
      Left            =   150
      TabIndex        =   19
      Top             =   3390
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "案件名稱："
      Height          =   255
      Left            =   150
      TabIndex        =   18
      Top             =   1862
      Width           =   930
   End
   Begin VB.Label Label14 
      Caption         =   "收  件  人："
      Height          =   255
      Left            =   150
      TabIndex        =   17
      Top             =   2612
      Width           =   975
   End
   Begin VB.Label lbeDay 
      Height          =   288
      Left            =   1125
      TabIndex        =   16
      Top             =   2222
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "發  文  日："
      Height          =   255
      Left            =   150
      TabIndex        =   15
      Top             =   2237
      Width           =   975
   End
   Begin VB.Label lbePropertyName 
      Height          =   285
      Left            =   1770
      TabIndex        =   14
      Top             =   1095
      Width           =   2340
   End
   Begin VB.Label lbeProperty 
      Height          =   288
      Left            =   1125
      TabIndex        =   13
      Top             =   1097
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "案件性質："
      Height          =   255
      Left            =   150
      TabIndex        =   12
      Top             =   1112
      Width           =   975
   End
   Begin VB.Label lbeCus 
      Height          =   288
      Left            =   1125
      TabIndex        =   11
      Top             =   1473
      Width           =   1095
   End
   Begin VB.Label lbeCaseNum 
      Height          =   288
      Left            =   1125
      TabIndex        =   10
      Top             =   721
      Width           =   1575
   End
   Begin VB.Label lbePaperNum 
      Height          =   288
      Left            =   1125
      TabIndex        =   9
      Top             =   345
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "當  事  人："
      Height          =   255
      Left            =   150
      TabIndex        =   8
      Top             =   1487
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號： "
      Height          =   255
      Left            =   150
      TabIndex        =   7
      Top             =   737
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號： "
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   362
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "退  件  日："
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   2993
      Width           =   975
   End
End
Attribute VB_Name = "frm071020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; lbeCusName、txtCP64、cboCaseName、lbeAccpet
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim strCP09() As String, rs As New ADODB.Recordset, t As Integer, LcTmp As String
Dim blnIsSave As Boolean
Dim m_count As Integer
Dim m_Nowindex As Integer
Dim m_CP64 As String


Private Sub cmdBack_Click()
Dim yn As Integer
 If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
     Exit Sub
 End If
 frm071019.Show
 Unload Me
End Sub

Private Sub cmdEnd_Click()
Dim yn As Integer
 If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
     Exit Sub
 End If
 Unload frm071019
 Unload Me
End Sub

Private Sub cmdSure_Click()
If txtDay = "" Or IsNull(txtDay) Then MsgBox "退件日不可空白", vbOKOnly, "警告": txtDay.SetFocus: Exit Sub
If txtCP64 = "" Or IsNull(txtCP64) Then MsgBox "退件原因不可空白", vbOKOnly, "警告": txtCP64.SetFocus: Exit Sub

'重新檢查欄位有效性
If TxtValidate = False Then Exit Sub

If SaveData Then blnIsSave = True Else DataErrorMessage (3)

If UBound(strCP09) = m_Nowindex Then
   cmdSure.Enabled = False
   Unload Me
   Unload frm071019
   frm071019.Show
   Exit Sub
End If
txtDay = ""
txtCP64 = ""
rs.Close
m_Nowindex = m_Nowindex + 1
GetData (m_Nowindex)
End Sub

Private Sub Form_Activate()
txtDay.SetFocus
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Dim i As Integer, n As Integer
With frm071019.MSHFlexGrid1
n = 0
For i = 1 To .Rows - 1
 .row = i
 .col = 0
    If .Text = "v" Then
  .col = 2
      ReDim Preserve strCP09(n)
      strCP09(n) = .Text
      n = n + 1
    End If
Next
End With
m_count = n
m_Nowindex = 0
LcTmp = frm071019.txtcp01
If frm071019.txtcp03 <> "" Then
  If frm071019.txtcp04 <> "" Then
     lbeCaseNum = frm071019.txtcp01 + "-" + frm071019.txtcp02 + "-" + frm071019.txtcp03 + "-" + frm071019.txtcp04
  Else
     lbeCaseNum = frm071019.txtcp01 + "-" + frm071019.txtcp02 + "-" + frm071019.txtcp03
  End If
Else
  lbeCaseNum = frm071019.txtcp01 + "-" + frm071019.txtcp02
End If
lbeCus = frm071019.lbeCusNum
lbeCusName = frm071019.lbeCusName
For i = 0 To frm071019.cboCaseName.ListCount - 1
 cboCaseName.AddItem frm071019.cboCaseName.List(i)
Next
cboCaseName.ListIndex = 0

Call GetData(0)
End Sub

Private Sub GetData(Init As Integer)
Dim strTemp As String
   strExc(1) = "select cp09,cp10,cp27,cp46,cp50,cp64 from caseprogress where cp09='" + strCP09(Init) + "'"
   intI = 0
   Set rs = ClsLawReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      lbePaperNum = strCP09(Init)
      If Not IsNull(rs.Fields!CP10) Then lbeProperty = rs.Fields!CP10
      If ClsPDGetCaseProperty(LcTmp, lbeProperty, strTemp, False) Then lbePropertyName = strTemp
      If Not IsNull(rs.Fields!Cp27) Then lbeDay = ChangeTStringToTDateString(ChangeWStringToTString(rs.Fields!Cp27))
      If Not IsNull(rs.Fields!cp50) Then lbeAccpet = rs.Fields!cp50
      If Not IsNull(rs.Fields!cp46) Then txtDay = ChangeWStringToTString(rs.Fields!cp46)
      m_CP64 = ""
      If Not IsNull(rs.Fields!CP64) Then
         'txtCP64 = rs.Fields!CP64
         m_CP64 = rs.Fields!CP64
      End If
   End If
End Sub

Private Function SaveData() As Boolean
Dim strCP64 As String
   
   strCP64 = ""
   If txtCP64 <> "" Then strCP64 = ChangeTStringToTDateString(txtDay) & "退件;" & txtCP64
   If m_CP64 <> "" Then strCP64 = strCP64 & ";" & m_CP64
   strExc(1) = "update caseprogress set cp46=19221111" & _
      ",cp64=" & CNULL(ChgSQL(strCP64)) & _
      " where cp09='" & lbePaperNum & "'"
   If ClsLawExecSQL(1, strExc) Then
      frm071019.SetDataComplete lbePaperNum
      SaveData = True
      blnIsSave = True
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm071020 = Nothing
End Sub

Private Sub txtCP64_GotFocus()
   TextInverse txtCP64
   OpenIme
End Sub

Private Sub txtCP64_Validate(Cancel As Boolean)
CloseIme
End Sub

Private Sub txtDay_GotFocus()
   TextInverse txtDay
End Sub

Private Sub txtDay_Validate(Cancel As Boolean)
If txtDay <> "" Or IsNull(txtDay) Then
    If CheckIsTaiwanDate(txtDay) Then
       If Val(GetTaiwanTodayDate) - Val(txtDay) < 0 Then
           MsgBox "輸入日期大於系統日", vbCritical
           Cancel = True
        End If
    Else
       Cancel = True
    End If
 End If
 If Cancel Then TextInverse txtDay
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.txtDay.Enabled = True Then
   Cancel = False
   txtDay_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2021/09/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If

TxtValidate = True
End Function
