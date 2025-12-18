VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010015 
   BorderStyle     =   1  '單線固定
   Caption         =   "顧問案件電話諮詢"
   ClientHeight    =   4320
   ClientLeft      =   180
   ClientTop       =   840
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9330
   Begin VB.TextBox txtCP113 
      Height          =   288
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   2
      Top             =   2712
      Width           =   840
   End
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6300
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7128
      TabIndex        =   5
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton CmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8256
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin MSForms.TextBox Text 
      Height          =   288
      Index           =   1
      Left            =   5040
      TabIndex        =   1
      Top             =   2340
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1508;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   1095
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   3090
      Width           =   7995
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14102;1931"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   288
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   2340
      Width           =   840
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1482;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   285
      Left            =   2340
      TabIndex        =   30
      Top             =   1224
      Width           =   6435
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11351;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作時數："
      Height          =   288
      Index           =   12
      Left            =   225
      TabIndex        =   29
      Top             =   2712
      Width           =   1050
   End
   Begin VB.Label Label8 
      Caption         =   "客戶智權人員："
      Height          =   255
      Left            =   3690
      TabIndex        =   28
      Top             =   1596
      Width           =   1305
   End
   Begin VB.Label lbeSalesNo 
      Height          =   288
      Left            =   5040
      TabIndex        =   27
      Top             =   1596
      Width           =   855
   End
   Begin VB.Label lbeSalesName 
      Height          =   288
      Left            =   6000
      TabIndex        =   26
      Top             =   1596
      Width           =   1425
   End
   Begin VB.Label Label3 
      Caption         =   "(電話諮詢人員)"
      Height          =   255
      Left            =   7425
      TabIndex        =   25
      Top             =   2340
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "分所案號："
      Height          =   255
      Left            =   4080
      TabIndex        =   24
      Top             =   869
      Width           =   900
   End
   Begin VB.Label lbeHC07 
      Height          =   288
      Left            =   5010
      TabIndex        =   23
      Top             =   869
      Width           =   2055
   End
   Begin VB.Label Label26 
      Caption         =   "協辦人員："
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   2340
      Width           =   900
   End
   Begin MSForms.Label lbe 
      Height          =   288
      Index           =   1
      Left            =   6000
      TabIndex        =   21
      Top             =   2340
      Width           =   1335
      VariousPropertyBits=   27
      Size            =   "2355;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbeAccept 
      Height          =   288
      Left            =   1200
      TabIndex        =   20
      Top             =   1596
      Width           =   1452
   End
   Begin VB.Label lbeCustomer 
      Height          =   288
      Left            =   1200
      TabIndex        =   19
      Top             =   1224
      Width           =   1092
   End
   Begin VB.Label Label16 
      Caption         =   "當  事  人："
      Height          =   252
      Left            =   240
      TabIndex        =   18
      Top             =   1224
      Width           =   972
   End
   Begin VB.Label lbePropertyName 
      Height          =   288
      Left            =   2040
      TabIndex        =   17
      Top             =   1985
      Width           =   3105
   End
   Begin VB.Label lbeProperty 
      Height          =   288
      Left            =   1200
      TabIndex        =   16
      Top             =   1985
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "案件性質："
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1985
      Width           =   975
   End
   Begin MSForms.Label lbe 
      Height          =   288
      Index           =   0
      Left            =   2160
      TabIndex        =   14
      Top             =   2340
      Width           =   1575
      VariousPropertyBits=   27
      Size            =   "2778;508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbeCaseNum 
      Height          =   288
      Left            =   1224
      TabIndex        =   13
      Top             =   869
      Width           =   2052
   End
   Begin VB.Label lbeNum 
      Height          =   288
      Left            =   1200
      TabIndex        =   12
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label Label10 
      Caption         =   "進度備註："
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "承辦人："
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2340
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號：  "
      Height          =   252
      Left            =   240
      TabIndex        =   9
      Top             =   869
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號：    "
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label21 
      Caption         =   "來電日期："
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   1596
      Width           =   972
   End
End
Attribute VB_Name = "frm010015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ; lbeCusName、Text(index)、lbe(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
Option Explicit

Public UpForm As Form
Dim rs As New ADODB.Recordset, strCP09() As String, t As Integer
Dim blnIsSave As Boolean
Dim m_CP09 As String
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String


Private Sub cmdBack_Click()
Dim yn As Integer
   
   If blnIsSave = False Then
      yn = MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2)
      If yn = 7 Then
         Exit Sub
      End If
   End If
   '2011/5/26 MODIFY BY SONIA
   'Me.Hide
   'UpForm.Show
   'Unload Me
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   '2011/5/26 END
End Sub

Private Sub cmdEnd_Click()
Dim yn As Integer
   
   If blnIsSave = False Then
      yn = MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2)
      If yn = 7 Then
         Exit Sub
      End If
   End If
   '2011/5/26 MODIFY BY SONIA
   'If Left(UpForm.Name, 6) = "frm100" Then
   '   Me.Hide
   '   Unload Me
   '   fnCloseAllFrm100
   'Else
   '   Me.Hide
   '   Unload UpForm
   '   Unload Me
   'End If
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   If UpForm.Name = "frm010014" Then
      Unload UpForm
   Else
      fnCloseAllFrm100
   End If
   '2011/5/26 END
End Sub

Private Sub cmdSure_Click()
Dim strDay1 As String
Dim strDay2 As String
Dim strDate As String

   If AllTextBeforeSaveCheck Then Exit Sub
   
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   Screen.MousePointer = 11
   If Not SaveData Then
      MsgBox "存檔失敗,請洽系統管理者", vbCritical
   Else
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
   End If
   Screen.MousePointer = 0
   '2011/5/26 ADD BY SONIA
   If UpForm.Name = "frm010014" Then
      UpForm.cmdSearch_Click
   Else
      UpForm.StrMenu
   End If
   '2011/5/26 END
End Sub

Private Sub Form_Load()
Dim i As Integer, n As Integer
  
   MoveFormToCenter Me
   blnIsSave = False
End Sub

Sub GetData(ByVal Init As Integer)
Dim i As Integer
Dim strName As String
 
   m_CP09 = Mid(frm010015.Tag, 1, 9)
   lbeNum = Mid(frm010015.Tag, 1, 9)
   lbeAccept = ChangeTStringToTDateString(Mid(frm010015.Tag, 10))
   
   'Modify By Sindy 2011/7/7 +CP113
   'Modified by Lydia 2022/04/11 +案源單號CP162
   'Modified by Lydia 2022/05/09 拿掉CP162: 改為該案號案件性質非7401的A類最後收文進度，若有CP162則發CP162
   strExc(1) = "select cp01,cp02,cp03,cp04,cp09,cp10," + _
       " cp14,hc05,hc07,CP113 from caseprogress,hirecase where cp09='" + m_CP09 + _
       "' and CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) "
   intI = 1
   'edit by nickc 2007/02/05 不用 dll 了
   'Set rs = objLawDll.ReadRstMsg(intI, strExc(1))
   Set rs = ClsLawReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      m_CP01 = rs.Fields!cp01
      m_CP02 = rs.Fields!cp02
      m_CP03 = rs.Fields!cp03
      m_CP04 = rs.Fields!cp04
      lbeCaseNum = GiveSymbol(m_CP01, m_CP02, m_CP03, m_CP04)
      If Not IsNull(rs.Fields!hc07) Then
         lbeHC07 = rs.Fields!hc07
      Else
         lbeHC07 = ""
      End If
      lbeCustomer = rs.Fields!hc05
      lbeCusName = GetCustomerName(rs.Fields!hc05, 0)
      
      If Not IsNull(rs.Fields!CP10) Then
         strName = ""
         lbeProperty = rs.Fields!CP10
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(m_CP01, lbeProperty, strName, False) Then
         If ClsPDGetCaseProperty(m_CP01, lbeProperty, strName, False) Then
            lbePropertyName = strName
         End If
      End If
      
      strName = ""
      '承辦人
      If Not IsNull(rs.Fields("CP14")) Then
         Text(0).Text = rs.Fields("CP14")
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(Text(0).Text, strName) Then
         If ClsPDGetStaff(Text(0).Text, strName) Then
            lbe(0).Caption = strName
         End If
      End If
      
      'Add By Sindy 2011/7/7
      If IsNull(rs.Fields!CP113) Then txtCP113 = "" Else txtCP113 = rs.Fields!CP113
   End If
   
   strName = ""
   '協辦人員
   Text(1).Text = strUserNum
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetStaff(Text(1).Text, strName) Then
   If ClsPDGetStaff(Text(1).Text, strName) Then
      lbe(1).Caption = strName
   End If
   
   '目前智權人員
   strName = ""
   'Added by Lydia 2022/04/11 有案源改通知介紹人員
   lbeSalesNo = ""
   'Modified by Lydia 2022/05/09 改為該案號案件性質非7401的A類最後收文進度，若有CP162則發CP162
   'If "" & rs.Fields("CP162") <> "" Then
   '    strExc(1) = "select los04 from lawofficesource where los15='" & rs.Fields("CP162") & "' "
   strExc(0) = ""
   strExc(1) = "select cp05,cp09,cp162 from caseprogress where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "'  and cp04='" & m_CP04 & "' and cp10<>'7401' and substr(cp09,1,1)='A' order by cp05 desc,cp09 desc "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
   If intI = 1 Then
     RsTemp.MoveFirst
     strExc(0) = "" & RsTemp.Fields("cp162")
   End If
   If strExc(0) <> "" Then
       strExc(1) = "select los04 from lawofficesource where los15='" & strExc(0) & "' "
   'end 2022/05/09
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
       If intI = 1 Then
           If "" & RsTemp.Fields("los04") <> "" Then
               strExc(1) = PUB_GetNowStaff("" & RsTemp.Fields("los04"), strExc(2))
               If strExc(2) <> "" Then
                   lbeSalesNo = strExc(2)
                   Label8 = "介紹人員："
               End If
           End If
       End If
   End If
   If lbeSalesNo = "" Then
   'end 2022/04/11
       lbeSalesNo = PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)
   End If 'Added by Lydia 2022/04/11
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetStaff(LbeSalesNo, strName) Then
   If ClsPDGetStaff(lbeSalesNo, strName) Then
      lbeSalesName = strName
   End If
End Sub

Private Function SaveData() As Boolean
Dim strNewNum As String, strNum As String, strTemp As String
Dim i As Integer
Dim strNP22 As String
   
On Error GoTo ErrorHandler
   SaveData = True
   cnnConnection.BeginTrans
   
   i = 1
   
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetAutoNumber("A", strNewNum, 1, 1) Then
   If ClsPDGetAutoNumber("A", strNewNum, 1, 1) Then
      'edit by nickc 2006/03/17
      'strNum = "A" + CStr(Year(Date) - 1911) + strNewNum
      'Modify By Sindy 2010/8/18 比對自動編號年度
      'strNum = "A" + CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911) + strNewNum
      strNum = "A" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) + strNewNum
   End If
   
   'Modify By Sindy 2011/7/7 +CP113
   'Modified by Lydia 2020/04/29 畫面的「目前智權人員」改為「客戶智權人員」，新增收文的業務區+業務員改存成操作者收文部門+操作者。
   'strExc(1) = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp05,CP11,cp12,cp13,cp32,cp43,cp20,cp26,CP27," + _
      " cp10,CP14,CP29,CP64,cp113) values (" + CNULL(strNum) + "," + CNULL(m_CP01) + "," + CNULL(m_CP02) + "," + _
      CNULL(m_CP03) + "," + CNULL(m_CP04) + "," + CNULL(ChangeTStringToWString(Replace(lbeAccept, "/", ""))) + ",'08'," + _
      CNULL(GetST15(LbeSalesNo)) + "," + CNULL(LbeSalesNo) + ",'N'," + CNULL(lbeNum) + ",'N','N'," + strSrvDate(1) + ",'7401'," + _
      CNULL(Text(0)) + "," + CNULL(Text(1)) + "," + CNULL(ChgSQL(Text(2))) + "," & CNULL(txtCP113) & ")"
   strExc(1) = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp05,CP11,cp12,cp13,cp32,cp43,cp20,cp26,CP27," + _
      " cp10,CP14,CP29,CP64,cp113) values (" + CNULL(strNum) + "," + CNULL(m_CP01) + "," + CNULL(m_CP02) + "," + _
      CNULL(m_CP03) + "," + CNULL(m_CP04) + "," + CNULL(ChangeTStringToWString(Replace(lbeAccept, "/", ""))) + ",'08'," + _
      CNULL(GetST15(strUserNum)) + "," + CNULL(strUserNum) + ",'N'," + CNULL(lbeNum) + ",'N','N'," + strSrvDate(1) + ",'7401'," + _
      CNULL(Text(0)) + "," + CNULL(Text(1)) + "," + CNULL(ChgSQL(Text(2))) + "," & CNULL(txtCP113) & ")"
   cnnConnection.Execute strExc(1)
   
   If SaveData Then blnIsSave = True Else blnIsSave = False

   cnnConnection.CommitTrans
   
   Call PUB_SendMail(strUserNum, lbeSalesNo, strNum, "顧問案件" & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & "電話諮詢", "電話諮詢內容：" & CNULL(ChgSQL(Text(2))), "" & vbCrLf & "此案當事人：" & lbeCusName)

   Exit Function

ErrorHandler:
   cnnConnection.RollbackTrans
   SaveData = False
End Function

Private Function ChgType(i As Integer, strText As String) As String
Dim strTemp As String
   
   Select Case i
      Case 0, 1
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(StrText, strTemp) Then ChgType = strTemp
         If ClsPDGetStaff(strText, strTemp) Then ChgType = strTemp
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm010015 = Nothing
End Sub

Private Sub Text_Change(Index As Integer)
   Select Case Index
      Case 0, 1
         If Text(Index) = "" Then lbe(Index) = ""
   End Select
End Sub

Private Sub Text_GotFocus(Index As Integer)
   
   Select Case Index
   Case Index
      TextInverse Text(Index)
   End Select
   
   Select Case Index
   Case 2
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Text(Index).IMEMode = 1
      OpenIme
   Case Else
      'edit by nickc 2007/06/06 切換輸入法改用API
      'Text(Index).IMEMode = 2
      CloseIme
   End Select
End Sub

'Modified by Lydia 2021/09/15 改成Form 2.0
'Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
  KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text_LostFocus(Index As Integer)
   Select Case Index
      Case 2
         'edit by nickc 2007/06/06 切換輸入法改用API
         'Text(Index).IMEMode = 2
         CloseIme
   End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp1 As String, strTemp2 As String
   Select Case Index
      Case 0, 1
         If Text(Index) <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetStaff(Text(Index), strTemp1) Then Lbe(Index) = strTemp1 Else Cancel = True
            If ClsPDGetStaff(Text(Index), strTemp1) Then lbe(Index) = strTemp1 Else Cancel = True
         End If
         If Text(0) = "" And Text(1) = "" Then
         'Modified by Lydia 2015/10/05
            'MsgBox "承辦律師及承辦法務不可同時空白!"
            MsgBox "承辦人及協辦人員不可同時空白!"
            Cancel = True
         End If
      Case 2
         If Text(Index) <> "" Then
            If CheckLengthIsOK(Text(Index), 2000) = False Then
                Cancel = True
            End If
         Else
            MsgBox "進度備註不可空白!"
            Cancel = True
         End If
   End Select
   If Cancel Then TextInverse Text(Index)
End Sub

Private Function AllTextBeforeSaveCheck() As Boolean
Dim strTemp  As String
  
  AllTextBeforeSaveCheck = True
    
  strTemp = ""
  
  If Text(0) <> "" Then
      strTemp = ""
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetStaff(Text(0), strTemp) Then
      If ClsPDGetStaff(Text(0), strTemp) Then
         lbe(0) = strTemp
      Else
         AllTextBeforeSaveCheck = True
         TextInverse Text(0)
         Exit Function
      End If
   End If
  
  If Text(1) <> "" Then
      strTemp = ""
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetStaff(Text(1), strTemp) Then
      If ClsPDGetStaff(Text(1), strTemp) Then
         lbe(1) = strTemp
      Else
         AllTextBeforeSaveCheck = True
         TextInverse Text(1)
         Exit Function
      End If
   End If
  
  If Text(2) <> "" Then
     If CheckLengthIsOK(Text(2), 2000) = False Then
         AllTextBeforeSaveCheck = True
         TextInverse Text(2)
         Exit Function
     End If
  End If
  
  AllTextBeforeSaveCheck = False
End Function

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   Cancel = False
   TxtValidate = False
   
   For Each objTxt In Me.Text
      If objTxt.Enabled = True Then
         Text_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   'Add By Sindy 2011/7/7
   txtCP113_Validate Cancel
   If Cancel = True Then
      txtCP113.SetFocus
      Exit Function
   End If
   '2011/7/7 END
   
   'Added by Lydia 2021/09/15 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   TxtValidate = True
End Function

'Add By Sindy 2011/7/7
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   
End Sub
'2011/7/7 End
