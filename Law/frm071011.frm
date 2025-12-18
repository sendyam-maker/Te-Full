VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071011 
   BorderStyle     =   1  '單線固定
   Caption         =   "回執"
   ClientHeight    =   3420
   ClientLeft      =   600
   ClientTop       =   960
   ClientWidth     =   7935
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   7935
   Begin VB.TextBox txtInDay 
      Height          =   285
      Left            =   5280
      MaxLength       =   7
      TabIndex        =   1
      Top             =   2976
      Width           =   1092
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6900
      TabIndex        =   4
      Top             =   70
      Width           =   760
   End
   Begin VB.TextBox txtDay 
      Height          =   285
      Left            =   1248
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
      Top             =   70
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
      Left            =   1248
      TabIndex        =   22
      Top             =   2598
      Width           =   5220
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "9208;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   285
      Left            =   1230
      TabIndex        =   21
      Top             =   1848
      Width           =   6525
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11509;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   285
      Left            =   2370
      TabIndex        =   20
      Top             =   1473
      Width           =   5295
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "9340;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblInput 
      Caption         =   "退件日/回執未回郵局送達日："
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   2991
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "案件名稱："
      Height          =   252
      Left            =   228
      TabIndex        =   18
      Top             =   1864
      Width           =   924
   End
   Begin VB.Label Label14 
      Caption         =   "收  件  人："
      Height          =   252
      Left            =   228
      TabIndex        =   17
      Top             =   2614
      Width           =   972
   End
   Begin VB.Label lbeDay 
      Height          =   285
      Left            =   1248
      TabIndex        =   16
      Top             =   2223
      Width           =   1692
   End
   Begin VB.Label Label11 
      Caption         =   "發  文  日："
      Height          =   252
      Left            =   228
      TabIndex        =   15
      Top             =   2239
      Width           =   972
   End
   Begin VB.Label lbePropertyName 
      Height          =   285
      Left            =   1968
      TabIndex        =   14
      Top             =   1098
      Width           =   2340
   End
   Begin VB.Label lbeProperty 
      Height          =   285
      Left            =   1248
      TabIndex        =   13
      Top             =   1098
      Width           =   612
   End
   Begin VB.Label Label13 
      Caption         =   "案件性質："
      Height          =   252
      Left            =   228
      TabIndex        =   12
      Top             =   1114
      Width           =   972
   End
   Begin VB.Label lbeCus 
      Height          =   285
      Left            =   1248
      TabIndex        =   11
      Top             =   1473
      Width           =   1092
   End
   Begin VB.Label lbeCaseNum 
      Height          =   285
      Left            =   1248
      TabIndex        =   10
      Top             =   723
      Width           =   1572
   End
   Begin VB.Label lbePaperNum 
      Height          =   285
      Left            =   1248
      TabIndex        =   9
      Top             =   348
      Width           =   1692
   End
   Begin VB.Label Label3 
      Caption         =   "當  事  人："
      Height          =   252
      Left            =   228
      TabIndex        =   8
      Top             =   1489
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號： "
      Height          =   252
      Left            =   228
      TabIndex        =   7
      Top             =   739
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號： "
      Height          =   252
      Left            =   228
      TabIndex        =   6
      Top             =   364
      Width           =   972
   End
   Begin VB.Label Label21 
      Caption         =   "收  受  日："
      Height          =   252
      Left            =   228
      TabIndex        =   5
      Top             =   2992
      Width           =   972
   End
End
Attribute VB_Name = "frm071011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; lbeCusName、cboCaseName、lbeAccpet
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim strCP09() As String, rs As New ADODB.Recordset, t As Integer, LcTmp As String
Dim blnIsSave As Boolean
Dim m_count As Integer
Dim m_Nowindex As Integer


Private Sub cmdBack_Click()
Dim yn As Integer
'If blnIsSave = False Then
 If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
     Exit Sub
 End If
'End If
 frm071010.Show
 Unload Me
End Sub

Private Sub cmdEnd_Click()
Dim yn As Integer
 If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
     Exit Sub
 End If
 Unload frm071010
 Unload Me
End Sub

Private Sub cmdSure_Click()
If txtDay = "" Or IsNull(txtDay) Then MsgBox "收受日不可空白", vbOKOnly, "警告": txtDay.SetFocus: Exit Sub
'Add By Cheng 2002/05/24
'重新檢查欄位有效性
If TxtValidate = False Then Exit Sub

If SaveData Then blnIsSave = True Else DataErrorMessage (3)

If UBound(strCP09) = m_Nowindex Then
   cmdSure.Enabled = False
   Unload Me
   Unload frm071010
   frm071010.Show
   Exit Sub
End If
 txtDay = ""
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
With frm071010.MSHFlexGrid1
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
LcTmp = frm071010.txtcp01
If frm071010.txtcp03 <> "" Then
  If frm071010.txtcp04 <> "" Then
     lbeCaseNum = frm071010.txtcp01 + "-" + frm071010.txtcp02 + "-" + frm071010.txtcp03 + "-" + frm071010.txtcp04
  Else
     lbeCaseNum = frm071010.txtcp01 + "-" + frm071010.txtcp02 + "-" + frm071010.txtcp03
  End If
Else
  lbeCaseNum = frm071010.txtcp01 + "-" + frm071010.txtcp02
End If
lbeCus = frm071010.lbeCusNum
lbeCusName = frm071010.lbeCusName
For i = 0 To frm071010.cboCaseName.ListCount - 1
 cboCaseName.AddItem frm071010.cboCaseName.List(i)
Next
cboCaseName.ListIndex = 0

GetData (0)
End Sub

Private Sub GetData(Init As Integer)
Dim strTemp As String
   'Modified by Lydia 2016/05/30 +cp47
   strExc(1) = "select cp09,cp10,cp27,cp46,cp50,cp47 from caseprogress where cp09='" + strCP09(Init) + "'"
   intI = 0
   'edit by nickc 2007/02/07 不用 dll 了
   'Set rs = objLawDll.ReadRstMsg(intI, strExc(1))
   Set rs = ClsLawReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      lbePaperNum = strCP09(Init)
      If Not IsNull(rs.Fields!CP10) Then lbeProperty = rs.Fields!CP10
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetCaseProperty(LcTmp, lbeProperty, strTemp, False) Then lbePropertyName = strTemp
      If ClsPDGetCaseProperty(LcTmp, lbeProperty, strTemp, False) Then lbePropertyName = strTemp
      If Not IsNull(rs.Fields!Cp27) Then lbeDay = ChangeTStringToTDateString(ChangeWStringToTString(rs.Fields!Cp27))
      If Not IsNull(rs.Fields!cp50) Then lbeAccpet = rs.Fields!cp50
      If Not IsNull(rs.Fields!cp46) Then txtDay = ChangeWStringToTString(rs.Fields!cp46)
      If Not IsNull(rs.Fields!cp47) Then txtInDay = ChangeWStringToTString(rs.Fields!cp47) 'Added by Lydia 2016/05/30
      
   End If
End Sub

Private Function SaveData() As Boolean
   'Modified by Lydia 2016/05/30 + 退件日/回執未回郵局送達日(CP47)
   'strExc(1) = "update caseprogress set cp46=" & ChangeTStringToWString(txtDay) & _
      " where cp09='" & lbePaperNum & "'"
   strExc(1) = "update caseprogress set cp46=" & ChangeTStringToWString(txtDay) & IIf(txtInDay <> "", ", cp47=" & ChangeTStringToWString(txtInDay), "") & _
      " where cp09='" & lbePaperNum & "'"
   'edit by nickc 2007/02/07 不用 dll 了
   'If objLawDll.ExecSQL(1, strExc) Then
   If ClsLawExecSQL(1, strExc) Then
      frm071010.SetDataComplete lbePaperNum
      SaveData = True
      blnIsSave = True
   End If
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm071011 = Nothing
End Sub

Private Sub txtDay_GotFocus()
   TextInverse txtDay
End Sub

Private Sub txtDay_Validate(Cancel As Boolean)
If txtDay <> "" Then
   If CheckIsTaiwanDate(txtDay) Then
      If Val(GetTaiwanTodayDate) - Val(txtDay) < 0 Then
         MsgBox "輸入日期大於系統日", vbCritical
         Cancel = True
      Else
         'Remove by Lydia 2016/05/30
         'If Val(txtDay) = 111111 Then
         '   MsgBox "收受日不可輸入111111，若為退件，請改至信件退回作業輸入", vbCritical
         '   Cancel = True
         'End If
         'Added by Lydia 2016/05/30
         If Val(txtDay) = 111111 Or Val(txtDay) = 110101 Then
            txtInDay.Locked = False
         Else
            txtInDay.Locked = True
         End If
      End If
   Else
      Cancel = True
   End If
End If
If Cancel Then TextInverse txtDay
End Sub

'Add By Cheng 2002/05/24
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
'Added by Lydia 2016/05/30 法務或顧問案件的回執退件日/回執未回郵局送達日輸入檢查
If Val(txtDay) = 111111 Or Val(txtDay) = 110101 Then
   If Me.txtInDay = "" Then
      MsgBox "退件日/回執未回郵局送達日不可空白", vbCritical
      txtInDay.SetFocus
      txtInDay_GotFocus
      Exit Function
   Else
      txtInDay_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Else
   If Me.txtInDay <> "" Then
      MsgBox "收受日為111111或110101時才可輸入", vbCritical
      txtDay.SetFocus
      txtDay_GotFocus
      Exit Function
   End If
End If
'end 2016/05/30

'Added by Lydia 2021/09/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If


TxtValidate = True
End Function
'Added by Lydia 2016/05/30
Private Sub txtInDay_GotFocus()
   TextInverse txtInDay
End Sub

Private Sub txtInDay_Validate(Cancel As Boolean)
If txtInDay <> "" Then
   If CheckIsTaiwanDate(txtInDay) Then
      If Val(GetTaiwanTodayDate) - Val(txtInDay) < 0 Then
         MsgBox "輸入日期大於系統日", vbCritical
         Cancel = True
      Else
      End If
   Else
      Cancel = True
   End If
End If
If Cancel Then TextInverse txtInDay

End Sub
