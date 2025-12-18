VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071003 
   BorderStyle     =   1  '單線固定
   Caption         =   "法務－顧問－聘任"
   ClientHeight    =   5076
   ClientLeft      =   4572
   ClientTop       =   840
   ClientWidth     =   9060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5076
   ScaleWidth      =   9060
   Begin VB.TextBox textCP15 
      Height          =   285
      Left            =   5340
      MaxLength       =   6
      TabIndex        =   7
      Top             =   3420
      Width           =   732
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一筆(&N)"
      Height          =   400
      Left            =   4128
      TabIndex        =   32
      Top             =   70
      Width           =   900
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8136
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1740
      TabIndex        =   9
      Text            =   "專利、商標、著作權"
      Top             =   4620
      Width           =   2655
   End
   Begin VB.TextBox txtHire2 
      Height          =   285
      Left            =   6900
      MaxLength       =   7
      TabIndex        =   5
      Top             =   2940
      Width           =   1335
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   5100
      MaxLength       =   7
      TabIndex        =   0
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印證書(&P)"
      Height          =   400
      Left            =   5052
      TabIndex        =   10
      Top             =   70
      Width           =   1100
   End
   Begin VB.TextBox txtHire1 
      Height          =   285
      Left            =   5340
      MaxLength       =   7
      TabIndex        =   4
      Top             =   2940
      Width           =   1212
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   180
      TabIndex        =   14
      Top             =   900
      Width           =   8775
      Begin VB.TextBox txtDisNumber 
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtCustomer 
         Height          =   285
         Left            =   4920
         MaxLength       =   9
         TabIndex        =   1
         Top             =   240
         Width           =   1092
      End
      Begin MSForms.TextBox txtCaseName 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   705
         Width           =   7335
         VariousPropertyBits=   671105051
         BackColor       =   16777215
         MaxLength       =   40
         Size            =   "12938;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbeCusName 
         Height          =   285
         Left            =   6030
         TabIndex        =   33
         Top             =   240
         Width           =   2385
         BackColor       =   -2147483637
         VariousPropertyBits=   27
         Size            =   "4207;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         Caption         =   "分所案號："
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1215
         Width           =   975
      End
      Begin VB.Label lblName 
         Caption         =   "案件名稱："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "本所案號："
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "當  事  人："
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Top             =   255
         Width           =   975
      End
      Begin VB.Label lbeCaseNumber 
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6180
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7008
      TabIndex        =   12
      Top             =   70
      Width           =   1100
   End
   Begin VB.TextBox txtSale 
      Height          =   285
      Left            =   1380
      MaxLength       =   6
      TabIndex        =   6
      Top             =   3420
      Width           =   732
   End
   Begin VB.Label Label6 
      Caption         =   "簽約時數："
      Height          =   252
      Index           =   2
      Left            =   4380
      TabIndex        =   35
      Top             =   3436
      Width           =   972
   End
   Begin MSForms.TextBox txtMemo 
      Height          =   585
      Left            =   1380
      TabIndex        =   8
      Top             =   3870
      Width           =   7575
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13361;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeSaleName 
      Height          =   285
      Left            =   2190
      TabIndex        =   34
      Top             =   3420
      Width           =   1665
      VariousPropertyBits=   27
      Size            =   "2937;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6660
      X2              =   6840
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Label Label3 
      Caption         =   "證書事務所名稱："
      Height          =   255
      Left            =   300
      TabIndex        =   31
      Top             =   4620
      Width           =   1455
   End
   Begin VB.Label lbeHire2 
      Height          =   285
      Left            =   2820
      TabIndex        =   30
      Top             =   2940
      Width           =   972
   End
   Begin VB.Line Line2 
      X1              =   2580
      X2              =   2700
      Y1              =   3075
      Y2              =   3075
   End
   Begin VB.Label lbeHire1 
      Height          =   285
      Left            =   1740
      TabIndex        =   29
      Top             =   2940
      Width           =   732
   End
   Begin VB.Label Label6 
      Caption         =   "上次聘任期間："
      Height          =   252
      Index           =   1
      Left            =   300
      TabIndex        =   28
      Top             =   2956
      Width           =   1332
   End
   Begin VB.Label lbeNumber 
      Height          =   285
      Left            =   1260
      TabIndex        =   25
      Top             =   540
      Width           =   1932
   End
   Begin VB.Label lbeCost 
      Height          =   288
      Left            =   7680
      TabIndex        =   24
      Top             =   3418
      Width           =   1212
   End
   Begin VB.Label Label13 
      Caption         =   "案件備註："
      Height          =   252
      Left            =   300
      TabIndex        =   23
      Top             =   3900
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "聘任期間："
      Height          =   252
      Index           =   0
      Left            =   4380
      TabIndex        =   22
      Top             =   2956
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "費       用："
      Height          =   252
      Left            =   6744
      TabIndex        =   21
      Top             =   3436
      Width           =   876
   End
   Begin VB.Label Label24 
      Caption         =   "智權人員："
      Height          =   252
      Left            =   300
      TabIndex        =   20
      Top             =   3436
      Width           =   972
   End
   Begin VB.Label Label7 
      Caption         =   "收  文  日："
      Height          =   252
      Left            =   4140
      TabIndex        =   19
      Top             =   556
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號： "
      Height          =   252
      Left            =   288
      TabIndex        =   18
      Top             =   556
      Width           =   972
   End
End
Attribute VB_Name = "frm071003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; lbeCusName、txtCaseName、txtMemo、lbeSaleName
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim strCP09() As String, t As Integer, blnIsSave As Boolean
Dim strDate As String, LcTmp As String, strPubcp10() As String, lC() As String
Dim m_CPCount As Integer
Dim m_Cpindex As Integer
Dim strCon(0 To 2) As String 'Added by Lydia 2024/04/15 取代strExc
Dim intA As Integer, rsAD As New ADODB.Recordset 'Added by Lydia 2024/04/15 取代intI, RsTemp

Private Sub cmdBack_Click()
   If Not blnIsSave Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   intForm = 0
   intNowRec = 0
   blnIsFormBack = True
   Unload Me
   Set frm071003 = Nothing
   frm071001.Show
End Sub

Private Sub cmdEnd_Click()
 Dim yn As Integer
   If Not blnIsSave Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   Unload frm071001
   Unload Me
   Set frm071003 = Nothing
End Sub

Private Sub cmdNext_Click()
  Dim i As Integer
  ClearForm
  m_Cpindex = m_Cpindex + 1
  If m_Cpindex = m_CPCount - 1 Then
     CmdNext.Enabled = False
  ElseIf m_Cpindex = m_CPCount Then
     Exit Sub
  End If
  If UCase(Left(lC(m_Cpindex), 2)) = "LA" And strPubcp10(m_Cpindex) = "顧問聘任" Then
      GetData (m_Cpindex)
  End If

End Sub

Private Sub cmdok_Click()
 Dim i As Integer
   If AllTextBeforeSaveCheck Then Exit Sub
   'Add By Cheng 2002/05/24
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   If Not SaveData Then
      DataErrorMessage (3)
      Exit Sub
   End If
   'If UBound(strCP09) = t Then
    If t = intRecount - 1 Then
      cmdOK.Enabled = False
      intForm = 0
      intNowRec = 0
      blnIsFormBack = True
      Unload Me
      Set frm071003 = Nothing
      frm071001.Show
      Exit Sub
   End If
   t = t + 1
   If Left(lC(t), 2) = "LA" And strPubcp10(t) = "顧問聘任" Then
      GetData (t)
   Else
      intForm = 3
      intNowRec = t
      t = 0
      For i = 0 To UBound(strCP09)
         ReDim Preserve strArryCP09(i)
         strArryCP09(i) = strCP09(i)
         ReDim Preserve strCP10(i)
         strCP10(i) = strPubcp10(i)
         ReDim Preserve strCaseKind(i)
         strCaseKind(i) = lC(i)
      Next
      frm071002.Show
      Unload Me
   End If
   
End Sub

Private Sub Form_Load()
 Dim i As Integer, n As Integer
   MoveFormToCenter Me
   blnIsSave = False
   m_CPCount = 0
   If intForm = 2 Then
      For i = 0 To UBound(strArryCP09)
         ReDim Preserve strCP09(n)
         strCP09(n) = strArryCP09(n)
         ReDim Preserve strPubcp10(n)
         strPubcp10(n) = strCP10(n)
         ReDim Preserve lC(n)
         lC(n) = strCaseKind(n)
        ' If Left(lc(n), 2) = "LA" And strPubcp10(n) = "顧問聘任" Then
            m_CPCount = m_CPCount + 1
        ' End If
         n = n + 1
      Next
      t = intNowRec
      m_Cpindex = t
      If m_Cpindex = m_CPCount Then
         CmdNext.Enabled = False
      End If
      GetData (t)
   Else
      With frm071001.MSHFlexGrid1
         n = 0
         For i = 1 To .Rows - 1
            .row = i
            .col = 0
            If .Text = "v" Then
               .col = 2
               ReDim Preserve strCP09(n)
               strCP09(n) = .Text
               .col = 3
               ReDim Preserve strPubcp10(n)
               strPubcp10(n) = .Text
               .col = 4
               ReDim Preserve lC(n)
               lC(n) = .Text
               n = n + 1
               m_CPCount = m_CPCount + 1
            End If
         Next
         GetData (0)
      End With
   End If
    'Add By Cheng 2003/06/02
   intRecount = m_CPCount
End Sub

Private Sub GetData(ByVal intai As Integer)
 Dim cp01 As String, cp02 As String, cp03 As String, cp04 As String, cp05 As String, _
    CP09 As String, CP10 As String, cp13 As String, cp16 As String, cp18 As String, _
    cp53 As String, cp54 As String, hc05 As String, hc06 As String, hc07 As String, hc12 As String, yn As Boolean
 Dim strTemp As String
   'Modified by Lydia 2024/04/15 +CP15
   strCon(0) = "select cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp13,cp16,cp18,cp53," & _
      "cp54,hc05,hc06,hc07,hc12,CP15 from caseprogress,hirecase where " & _
      "cp09=" + CNULL(strCP09(intai)) + " and CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04"
   intA = 0
   Set rsAD = ClsLawReadRstMsg(intA, strCon(0))    'edit by nickc 2007/02/07 不用 dll 了 Set rsad = objLawDll.ReadRstMsg(inta, strcon(0))
   If intA = 1 Then
      With rsAD
         cp01 = IIf(IsNull(.Fields!cp01), "", .Fields!cp01)
         cp02 = IIf(IsNull(.Fields!cp02), "", .Fields!cp02)
         cp03 = IIf(IsNull(.Fields!cp03), "", .Fields!cp03)
         cp04 = IIf(IsNull(.Fields!cp04), "", .Fields!cp04)
         cp05 = IIf(IsNull(.Fields!cp05), "", .Fields!cp05)
         CP10 = IIf(IsNull(.Fields!CP10), "", .Fields!CP10)
         cp13 = IIf(IsNull(.Fields!cp13), "", .Fields!cp13)
         cp16 = IIf(IsNull(.Fields!cp16), "", .Fields!cp16)
         cp18 = IIf(IsNull(.Fields!cp18), "", .Fields!cp18)
         cp53 = IIf(IsNull(.Fields!cp53), "", .Fields!cp53)
         cp54 = IIf(IsNull(.Fields!cp54), "", .Fields!cp54)
         hc05 = IIf(IsNull(.Fields!hc05), "", ChangeCustomerS(.Fields!hc05))
         hc06 = IIf(IsNull(.Fields!hc06), "", .Fields!hc06)
         hc07 = IIf(IsNull(.Fields!hc07), "", .Fields!hc07)
         hc12 = IIf(IsNull(.Fields!hc12), "", .Fields!hc12)
         lbeNumber = strCP09(intai)
         lbeCaseNumber = GiveSymbol(cp01, cp02, cp03, cp04, LcTmp)
         txtDate = IIf(IsNull(cp05), "", ChangeWStringToTString(cp05))
         txtCustomer = IIf(IsNull(hc05), "", hc05)
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetCustomer(txtCustomer, strTemp) Then lbeCusName = strTemp Else lbeCusName = ""
         If ClsPDGetCustomer(txtCustomer, strTemp) Then lbeCusName = strTemp Else lbeCusName = ""
         txtCaseName = IIf(IsNull(hc06), "", hc06)
         txtDisNumber = IIf(IsNull(hc07), "", hc07)
         txtMemo = IIf(IsNull(hc12), "", hc12)
         txtSale = IIf(IsNull(cp13), "", cp13)
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(txtSale, strTemp) Then lbeSaleName = strTemp
         If ClsPDGetStaff(txtSale, strTemp) Then lbeSaleName = strTemp
         lbeCost = IIf(IsNull(cp16), "", cp16)
         txtHire1 = IIf(IsNull(cp53), "", ChangeWStringToTString(cp53))
         txtHire2 = IIf(IsNull(cp54), "", ChangeWStringToTString(cp54))
         'Added by Lydia 2024/04/15 簽約時數
         textCP15.Text = "" & .Fields("cp15")
         textCP15.Tag = textCP15.Text
         'end 2024/04/15
      End With
   End If
   GetPreHire
   'Add by Morgan 2003/12/07
   Call PUB_CheckSales(Left(LcTmp, InStr(LcTmp, Right(LcTmp, 9)) - 1), Left(Right(LcTmp, 9), 6), Left(Right(LcTmp, 3), 1), Right(LcTmp, 2), txtDate, txtSale, lbeSaleName)
   'End 2003/12/07
End Sub

Private Sub lbeHire_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set rsAD = Nothing 'Added by Lydia 2024/04/15
   'Add By Cheng 2002/07/18
   Set frm071003 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text1.IMEMode = 1
   OpenIme
End Sub

Private Sub Text1_LostFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'Text1.IMEMode = 2
   CloseIme
End Sub

Private Sub txtCaseName_GotFocus()
   TextInverse txtCaseName
End Sub

Private Sub txtCaseName_Validate(Cancel As Boolean)

      If CheckLengthIsOK(txtCaseName, 40) = False Then
          Cancel = True
          txtCaseName.SetFocus
      End If

End Sub

Private Sub txtCustomer_Change()
 Dim StrCusName As String, i As Integer
   If txtCustomer <> "" Then
      txtCustomer = UCase(txtCustomer)
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetCustomer(txtCustomer, StrCusName) Then lbeCusName = StrCusName
      If ClsPDGetCustomer(txtCustomer, StrCusName) Then lbeCusName = StrCusName
   End If
   If txtCustomer = "" Then lbeCusName = ""
End Sub

Private Sub txtCustomer_GotFocus()
   TextInverse txtCustomer
End Sub

Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
 Dim StrCusName As String, i As Integer
   If txtCustomer <> "" Then
      txtCustomer = UCase(txtCustomer)
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetCustomer(txtCustomer, StrCusName) Then lbeCusName = StrCusName Else Cancel = True
      If ClsPDGetCustomer(txtCustomer, StrCusName) Then lbeCusName = StrCusName Else Cancel = True
   End If
   If txtCustomer = "" Then lbeCusName = ""
   If Cancel Then TextInverse txtCustomer
End Sub

Private Sub txtDate_GotFocus()
   TextInverse txtDate
End Sub

Private Sub txtDate_Validate(keepfocus As Boolean)
   If Len(txtDate) > 5 Then
       If CheckIsTaiwanDate(txtDate) Then
          If Val(GetTaiwanTodayDate) - Val(txtDate) < 0 Then
             MsgBox "輸入日期大於系統日", vbCritical
             keepfocus = True
          End If
       Else
          'MsgBox "輸入日期非民國日期", vbCritical
          keepfocus = True
       End If
   Else
      MsgBox "日期格式不正確", vbCritical
      keepfocus = True
   End If
   If keepfocus Then TextInverse txtDate
End Sub

Private Sub txtDisNumber_GotFocus()
   TextInverse txtDisNumber
End Sub

Private Sub txtDisNumber_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDisNumber_Validate(Cancel As Boolean)
   If txtDisNumber <> "" Then txtDisNumber = UCase(txtDisNumber)
End Sub

Private Sub txtHire1_GotFocus()
   TextInverse txtHire1
End Sub

Private Sub txtHire1_Validate(keepfocus As Boolean)
   If txtHire1 <> "" Then
      If Not CheckIsTaiwanDate(txtHire1) Then
         keepfocus = True
      End If
   Else
      DataErrorMessage 1, "聘任期間－起日"
      keepfocus = True
   End If
   If keepfocus Then TextInverse txtHire1
End Sub

Private Sub txtHire2_GotFocus()
   TextInverse txtHire2
End Sub

Private Sub txtHire2_Validate(keepfocus As Boolean)
   If txtHire2 <> "" Then
      If CheckIsTaiwanDate(txtHire2) Then
         If Val(GetTaiwanTodayDate) - Val(txtHire2) > 0 Then
            MsgBox "輸入日期小於系統日", vbCritical
            keepfocus = True
         Else
            If Val(txtHire2) <= Val(txtHire1) Then
               MsgBox "聘任期限－止日小於或等於起日", vbCritical
               keepfocus = True
            End If
         End If
      Else
         keepfocus = True
      End If
   Else
      DataErrorMessage 2, "聘任期限－止日"
      keepfocus = True
   End If
   If keepfocus Then TextInverse txtHire2
End Sub

Private Sub txtMemo_GotFocus()
   TextInverse txtMemo
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtMemo.IMEMode = 1
   OpenIme
End Sub

Private Sub GetPreHire()
 Dim strDate1 As String, StrDate2 As String, i As Integer
  strCon(1) = "select cp53,cp54,cp09 from caseprogress, casepropertymap where " & _
     ChgCaseprogress(LcTmp) + " and cpm03='顧問聘任' and cp01=cpm01(+) AND cp10=cpm02(+) order by cp05 "
   intA = 0
   Set rsAD = ClsLawReadRstMsg(intA, strCon(1))    'edit by nickc 2007/02/07 不用 dll 了 Set rsad = objLawDll.ReadRstMsg(inta, strcon(1))
   If intA = 1 Then
      rsAD.Find "cp09='" + lbeNumber + "'", 0, adSearchForward
      i = rsAD.AbsolutePosition
   End If
   If i > 1 Then
      rsAD.AbsolutePosition = i - 1
      strDate1 = IIf(IsNull(rsAD.Fields!cp53), "", rsAD.Fields!cp53)
      StrDate2 = IIf(IsNull(rsAD.Fields!cp54), "", rsAD.Fields!cp54)
      lbeHire1 = ChangeWStringToTString(strDate1): lbeHire2 = ChangeWStringToTString(StrDate2)
   End If
End Sub

Private Function SaveData() As Boolean
 Dim blnYN As Boolean
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
SaveData = True
cnnConnection.BeginTrans
   
   LcTmp = Replace(lbeCaseNumber, "-", "")
   If Len(LcTmp) = 7 Then
      LcTmp = LcTmp + String(10 - Len(LcTmp), "0")
   ElseIf Len(LcTmp) = 8 Then
      LcTmp = LcTmp + String(11 - Len(LcTmp), "0")
   ElseIf Len(LcTmp) = 9 Then
      LcTmp = LcTmp + String(12 - Len(LcTmp), "0")
   End If
   ' 91.04.04 modify by louis (修改單引號)
   strCon(1) = "update hirecase set hc05=" + CNULL(ChangeCustomerL(txtCustomer)) + _
   ",hc06=" + CNULL(txtCaseName) + " ,hc07=" + CNULL(txtDisNumber) + _
   " ,hc12=" + CNULL(ChgSQL(txtMemo)) + " where " & ChgHirecase(LcTmp)
    'Add By Cheng 2002/11/07
    Pub_SeekTbLog strCon(1) 'Added by Lydia 2024/04/15
    cnnConnection.Execute strCon(1)
    
   'Modified by Lydia 2024/04/15 +簽約時數CP15
   strCon(2) = "update caseprogress set cp05=" + CNULL(ChangeTStringToWString(txtDate)) + _
   ",cp13=" + CNULL(txtSale) + ",cp53=" + CNULL(ChangeTStringToWString(txtHire1)) + _
   ",cp54=" + CNULL(ChangeTStringToWString(txtHire2)) + ",cp15=" + CNULL(textCP15, True) + _
   " where cp09=" + CNULL(lbeNumber) + ""
    'Add By Cheng 2002/11/07
    Pub_SeekTbLog strCon(2) 'Added by Lydia 2024/04/15
    cnnConnection.Execute strCon(2)
    'Modify By Cheng 2002/11/07
'   blnYN = objLawDll.ExecSQL(2, strExc)
'   If blnYN Then SaveData = True: blnIsSave = True Else SaveData = False
   If SaveData = True Then blnIsSave = True
   frm071001.SetDataComplete lbeNumber.Caption

'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    SaveData = False
End Function

Private Function AllTextBeforeSaveCheck() As Boolean
 Dim i As Integer
   AllTextBeforeSaveCheck = True
   If txtDate = "" Then
      MsgBox "收文日不可空白", vbCritical
      txtDate.SetFocus
      Exit Function
   End If

   If txtCaseName = "" Then
        MsgBox " 案件名稱不可空白", vbCritical
        txtCaseName.SetFocus
        Exit Function
   End If
   
   If txtCustomer = "" Or IsNull(txtCustomer) Then
      MsgBox "當事人代號不可空白", vbCritical
      txtCustomer.SetFocus
      Exit Function
   End If
   If txtHire1 = "" Or txtHire2 = "" Then
      MsgBox "聘任期間不可空白", vbCritical
      If txtHire1 = "" Then
         txtHire1.SetFocus
         Exit Function
      Else
         txtHire2 = ""
         Exit Function
      End If
   End If
   AllTextBeforeSaveCheck = False
End Function

Private Sub txtMemo_LostFocus()
   'edit by nickc 2007/06/11  切換輸入法改用API
   'txtMemo.IMEMode = 2
   CloseIme
End Sub

Private Sub txtMemo_Validate(Cancel As Boolean)
   If txtMemo <> "" Then
      If CheckLengthIsOK(txtMemo.Text, 2000) = False Then
          Cancel = True
          txtMemo.SetFocus
      End If
   End If
End Sub

Private Sub txtSale_GotFocus()
   TextInverse txtSale
End Sub

'Add By Sindy 2010/11/26
Private Sub txtSale_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
 Dim strSaleName As String
   If txtSale <> "" Then
      txtSale = UCase(txtSale)
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetStaff(txtSale, strSaleName) Then
      If ClsPDGetStaff(txtSale, strSaleName) Then
         lbeSaleName = strSaleName
      Else
         Cancel = True
      End If
   Else
      DataErrorMessage 1, "智權人員"
      Cancel = True
   End If
   If txtSale = "" Then lbeSaleName = ""
   If Cancel Then TextInverse txtSale
End Sub
Private Sub ClearForm()
  txtDate.Text = ""
  txtCustomer.Text = ""
  txtCaseName.Text = ""
  txtDisNumber.Text = ""
  txtHire1.Text = ""
  txtHire2.Text = ""
  txtSale.Text = ""
  txtMemo.Text = ""
  lbeNumber.Caption = ""
  lbeCaseNumber.Caption = ""
  lbeHire1.Caption = ""
  lbeHire2.Caption = ""
  lbeSaleName.Caption = ""
  lbeCost.Caption = ""
  'Added by Lydia 2024/04/15
  textCP15.Text = ""
  textCP15.Tag = ""
End Sub

'Add By Cheng 2002/05/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.txtCaseName.Enabled = True Then
   Cancel = False
   txtCaseName_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtCustomer.Enabled = True Then
   Cancel = False
   txtCustomer_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtDate.Enabled = True Then
   Cancel = False
   txtDate_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtDisNumber.Enabled = True Then
   Cancel = False
   txtDisNumber_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtHire1.Enabled = True Then
   Cancel = False
   txtHire1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtHire2.Enabled = True Then
   Cancel = False
   txtHire2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.txtSale.Enabled = True Then
   Cancel = False
   txtSale_Validate Cancel
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

'Added by Lydia 2024/04/15
Private Sub textCP15_GotFocus()
   TextInverse textCP15
End Sub

'Added by Lydia 2024/04/15
Private Sub textCP15_KeyPress(KeyAscii As Integer)
   'Modified by Lydia 2025/03/05 可以輸入小數點True
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
