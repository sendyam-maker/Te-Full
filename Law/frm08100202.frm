VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm08100202 
   BorderStyle     =   1  '單線固定
   Caption         =   "相對人資料"
   ClientHeight    =   5175
   ClientLeft      =   1140
   ClientTop       =   810
   ClientWidth     =   7245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7245
   Begin VB.CommandButton Command2 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5148
      TabIndex        =   0
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5976
      TabIndex        =   1
      Top             =   70
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   108
      TabIndex        =   2
      Top             =   876
      Width           =   6975
      Begin MSForms.Label lbeCusName 
         Height          =   285
         Left            =   2250
         TabIndex        =   21
         Top             =   960
         Width           =   4005
         BackColor       =   -2147483637
         VariousPropertyBits=   27
         Size            =   "7064;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboCaseName 
         Height          =   285
         Left            =   1230
         TabIndex        =   20
         Top             =   585
         Width           =   5055
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "8916;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblName 
         Caption         =   "案件名稱："
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Index           =   1
         Left            =   2400
         TabIndex        =   11
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "是否為智慧財產權案："
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1800
      End
      Begin VB.Label Label18 
         Caption         =   "分所案號："
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lbeNumber 
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "當  事  人："
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbeCustomer 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   1200
         TabIndex        =   5
         Top             =   960
         Width           =   765
      End
      Begin VB.Label lbeIsIP 
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lbeDisNum 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   4440
         TabIndex        =   3
         Top             =   1320
         Width           =   1725
      End
   End
   Begin MSForms.TextBox txtName 
      Height          =   585
      Index           =   0
      Left            =   1560
      TabIndex        =   24
      Top             =   2790
      Width           =   5175
      VariousPropertyBits=   -1466941413
      MaxLength       =   600
      ScrollBars      =   2
      Size            =   "9128;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtName 
      Height          =   585
      Index           =   1
      Left            =   1560
      TabIndex        =   23
      Top             =   3525
      Width           =   5175
      VariousPropertyBits=   -1466941413
      MaxLength       =   600
      ScrollBars      =   2
      Size            =   "9128;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtName 
      Height          =   585
      Index           =   2
      Left            =   1560
      TabIndex        =   22
      Top             =   4290
      Width           =   5175
      VariousPropertyBits=   -1466941413
      MaxLength       =   600
      ScrollBars      =   2
      Size            =   "9128;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbePaperNum 
      AutoSize        =   -1  'True
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """#-##-######"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
      Height          =   228
      Left            =   1428
      TabIndex        =   19
      Top             =   588
      Width           =   1344
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收  文  號： "
      Height          =   180
      Index           =   1
      Left            =   336
      TabIndex        =   18
      Top             =   636
      Width           =   1068
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收  文  日："
      Height          =   180
      Index           =   0
      Left            =   3468
      TabIndex        =   17
      Top             =   636
      Width           =   900
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "對造名稱(中)："
      Height          =   180
      Left            =   225
      TabIndex        =   16
      Top             =   2790
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "對造名稱(英)："
      Height          =   180
      Left            =   228
      TabIndex        =   15
      Top             =   3525
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "對造名稱(日)："
      Height          =   180
      Left            =   228
      TabIndex        =   14
      Top             =   4290
      Width           =   1200
   End
   Begin VB.Label lbeDate 
      AutoSize        =   -1  'True
      Height          =   204
      Left            =   4548
      TabIndex        =   13
      Top             =   612
      Width           =   1368
   End
End
Attribute VB_Name = "frm08100202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; cboCaseName、lbeCusName、txtName(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim LcTmp As String

Private Sub Command2_Click()

   'Added by Lydia 2021/09/22 修正畫面所有含跳行符號的文字框.
   PUB_FilterFormText Me

   'Added by Lydia 2021/09/22 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Sub
   End If
   
   If txtName(0) <> "" Or txtName(1) <> "" And txtName(2) <> "" Then
      strExc(1) = "update caseprogress set cp40=" & CNULL(ChgSQL(txtName(0))) & ",cp41=" & CNULL(ChgSQL(txtName(1))) & _
         ",cp42=" & CNULL(ChgSQL(txtName(2))) & " where " & ChgCaseprogress(LcTmp) + " and cp09=" + CNULL(lbePaperNum)
      'edit by nickc 2007/02/07 不用 dll 了
      'If Not objLawDll.ExecSQL(1, strExc) Then
      If Not ClsLawExecSQL(1, strExc) Then
         DataErrorMessage (3)
      End If
   End If
   Unload Me
   frm081002.Show
End Sub

Private Sub Command3_Click()
   Unload Me
   frm081002.Show
End Sub

Private Sub Form_Load()
 Dim i As Integer, temp(2 To 4) As String
   MoveFormToCenter Me
   lbeNumber = frm081002.lbeNumber
   lbePaperNum = frm081002.lbePaperNum
   lbeDate = frm081002.Text(0)
   lbeCustomer = frm081002.Text(1)
   temp(2) = "中:"
   temp(3) = "英:"
   temp(4) = "日:"
   For i = 2 To 4
      If frm081002.Text(i) <> "" Then
         cboCaseName.AddItem temp(i) + frm081002.Text(i)
      End If
   Next
   If cboCaseName.ListCount > 0 Then cboCaseName.ListIndex = 0
   lbeIsIP = frm081002.Text(5)
   lbeDisNum = frm081002.Text(6)
   LcTmp = frm081002.lbeNumber.Tag
   GetData
End Sub

Private Sub GetData()
   strExc(0) = "select cp40,cp41,cp42  from caseprogress where " & ChgCaseprogress(LcTmp) & " and cp09=" + CNULL(lbePaperNum)
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rstemp = objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      txtName(0) = IIf(IsNull(RsTemp.Fields!cp40), "", RsTemp.Fields!cp40)
      txtName(1) = IIf(IsNull(RsTemp.Fields!cp41), "", RsTemp.Fields!cp41)
      txtName(2) = IIf(IsNull(RsTemp.Fields!cp42), "", RsTemp.Fields!cp42)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm08100202 = Nothing
End Sub

Private Sub lbeCustomer_Change()
 Dim StrCusName As String, i As Integer
   If lbeCustomer <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetCustomer(lbeCustomer, StrCusName) Then lbeCusName = StrCusName
      If ClsPDGetCustomer(lbeCustomer, StrCusName) Then lbeCusName = StrCusName
   End If
End Sub

Private Sub txtName_GotFocus(Index As Integer)
   Select Case Index
      Case Index
         TextInverse txtName(Index)
   End Select
   Select Case Index
      Case 0, 2
          'edit by nickc 2007/06/11  切換輸入法改用API
          'txtName(Index).IMEMode = 1
          OpenIme
      Case Else
          'edit by nickc 2007/06/11  切換輸入法改用API
          'txtName(Index).IMEMode = 2
          CloseIme
   End Select
   
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub txtName_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtName_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case Index
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txtName_LostFocus(Index As Integer)
          'edit by nickc 2007/06/11  切換輸入法改用API
          'txtName(Index).IMEMode = 2
          CloseIme
End Sub

Private Sub txtName_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1, 2
         If txtName(Index) <> "" Then
           txtName(Index) = UCase(txtName(Index))
         End If
   End Select
End Sub
