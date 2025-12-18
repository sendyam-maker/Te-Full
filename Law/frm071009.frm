VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071009 
   BorderStyle     =   1  '單線固定
   Caption         =   "法院文書"
   ClientHeight    =   5376
   ClientLeft      =   180
   ClientTop       =   840
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5376
   ScaleWidth      =   9336
   Begin VB.CommandButton cmdSure 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6348
      TabIndex        =   39
      Top             =   70
      Width           =   752
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7128
      TabIndex        =   38
      Top             =   70
      Width           =   1100
   End
   Begin VB.CommandButton CmdEnd 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8256
      TabIndex        =   37
      Top             =   70
      Width           =   800
   End
   Begin MSForms.ComboBox cboGov 
      Height          =   324
      Left            =   1230
      TabIndex        =   5
      Top             =   2425
      Width           =   2844
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "5016;572"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   7
      Left            =   1230
      TabIndex        =   46
      Top             =   2791
      Width           =   4935
      VariousPropertyBits=   671105051
      MaxLength       =   32
      Size            =   "8705;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   588
      Index           =   13
      Left            =   1230
      TabIndex        =   45
      Top             =   4728
      Width           =   8016
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14139;1037"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   9
      Left            =   1230
      TabIndex        =   44
      Top             =   3445
      Width           =   840
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1482;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   585
      Index           =   12
      Left            =   1230
      TabIndex        =   43
      Top             =   4099
      Width           =   7995
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14102;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   8
      Left            =   1230
      TabIndex        =   42
      Top             =   3118
      Width           =   4956
      VariousPropertyBits=   671105051
      MaxLength       =   40
      Size            =   "8742;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   11
      Left            =   1230
      TabIndex        =   41
      Top             =   3772
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   10
      Left            =   5280
      TabIndex        =   7
      Top             =   3772
      Width           =   1215
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "2143;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   1230
      TabIndex        =   1
      Top             =   1784
      Width           =   720
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1270;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   0
      Left            =   5280
      TabIndex        =   0
      Top             =   1470
      Width           =   1215
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "2143;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   4
      Left            =   5280
      TabIndex        =   4
      Top             =   2098
      Width           =   1215
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "2143;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   6
      Left            =   5280
      TabIndex        =   6
      Top             =   2445
      Width           =   2415
      VariousPropertyBits=   671105051
      MaxLength       =   20
      Size            =   "4260;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   2
      Top             =   1784
      Width           =   612
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1080;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   285
      Index           =   3
      Left            =   1230
      TabIndex        =   3
      Top             =   2098
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbeCusName 
      Height          =   285
      Left            =   2340
      TabIndex        =   40
      Top             =   1156
      Width           =   6435
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Size            =   "11351;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label26 
      Caption         =   "協辦人員："
      Height          =   252
      Left            =   240
      TabIndex        =   36
      Top             =   3772
      Width           =   1092
   End
   Begin MSForms.Label lbe 
      Height          =   285
      Index           =   11
      Left            =   2190
      TabIndex        =   35
      Top             =   3772
      Width           =   1815
      VariousPropertyBits=   27
      Size            =   "3201;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbeAccept 
      Height          =   285
      Left            =   1230
      TabIndex        =   34
      Top             =   1470
      Width           =   1452
   End
   Begin VB.Label Label20 
      Caption         =   "機關文號："
      Height          =   252
      Left            =   240
      TabIndex        =   33
      Top             =   3118
      Width           =   972
   End
   Begin MSForms.Label lbe 
      Height          =   285
      Index           =   2
      Left            =   6000
      TabIndex        =   32
      Top             =   1784
      Width           =   3096
      VariousPropertyBits=   27
      Size            =   "5461;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbeCustomer 
      Height          =   285
      Left            =   1230
      TabIndex        =   31
      Top             =   1156
      Width           =   1092
   End
   Begin VB.Label Label16 
      Caption         =   "當  事  人："
      Height          =   252
      Left            =   240
      TabIndex        =   30
      Top             =   1172
      Width           =   972
   End
   Begin VB.Label lbePropertyName 
      Height          =   285
      Left            =   6000
      TabIndex        =   29
      Top             =   528
      Width           =   3108
   End
   Begin VB.Label lbeProperty 
      Height          =   285
      Left            =   5280
      TabIndex        =   28
      Top             =   528
      Width           =   612
   End
   Begin VB.Label Label13 
      Caption         =   "案件性質："
      Height          =   252
      Left            =   4200
      TabIndex        =   27
      Top             =   544
      Width           =   972
   End
   Begin MSForms.Label lbe 
      Height          =   285
      Index           =   1
      Left            =   2190
      TabIndex        =   26
      Top             =   1784
      Width           =   1965
      VariousPropertyBits=   27
      Size            =   "3466;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbe 
      Height          =   285
      Index           =   9
      Left            =   2190
      TabIndex        =   25
      Top             =   3445
      Width           =   1812
      VariousPropertyBits=   27
      Size            =   "3196;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lbeCaseNum 
      Height          =   285
      Left            =   1230
      TabIndex        =   24
      Top             =   842
      Width           =   2052
   End
   Begin VB.Label lbeNum 
      Height          =   285
      Left            =   1230
      TabIndex        =   23
      Top             =   528
      Width           =   1572
   End
   Begin VB.Label Label10 
      Caption         =   "進度備註："
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4099
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "承辦期限："
      Height          =   252
      Left            =   4200
      TabIndex        =   21
      Top             =   3772
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "承  辦  人："
      Height          =   252
      Left            =   240
      TabIndex        =   20
      Top             =   3445
      Width           =   972
   End
   Begin VB.Label Label5 
      Caption         =   "來文性質："
      Height          =   252
      Left            =   240
      TabIndex        =   19
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "判決日期："
      Height          =   252
      Left            =   4200
      TabIndex        =   18
      Top             =   1486
      Width           =   972
   End
   Begin VB.Label Label3 
      Caption         =   "本所期限："
      Height          =   252
      Left            =   240
      TabIndex        =   17
      Top             =   2114
      Width           =   972
   End
   Begin VB.Label Label18 
      Caption         =   "案件備註："
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4728
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "下一程序："
      Height          =   252
      Left            =   4224
      TabIndex        =   15
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label8 
      Caption         =   "法院案號："
      Height          =   252
      Left            =   240
      TabIndex        =   14
      Top             =   2791
      Width           =   972
   End
   Begin VB.Label Label7 
      Caption         =   "法定期限："
      Height          =   252
      Left            =   4200
      TabIndex        =   13
      Top             =   2114
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號：  "
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   858
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號：    "
      Height          =   252
      Left            =   240
      TabIndex        =   11
      Top             =   544
      Width           =   972
   End
   Begin VB.Label Label21 
      Caption         =   "收  受  日："
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   1486
      Width           =   972
   End
   Begin VB.Label Label24 
      Caption         =   "股        別："
      Height          =   252
      Left            =   4200
      TabIndex        =   9
      Top             =   2461
      Width           =   972
   End
   Begin VB.Label Label25 
      Caption         =   "機關代號："
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   2461
      Width           =   972
   End
End
Attribute VB_Name = "frm071009"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/14 改成Form2.0 ;lbeCusName、Text(Index)、lbe(Index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim Rs As New ADODB.Recordset, strCP09() As String, t As Integer
Dim cp01 As String, cp02 As String, cp03 As String, cp04 As String, Strsale As String
Dim blnIsSave As Boolean
Dim m_CP09 As String
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_Nation As String
Dim m_GovListNew As String, m_GovListDef As String  'Added by Lydia 2025/11/18

Private Sub cmdBack_Click()
 Dim yn As Integer
   If blnIsSave = False Then
      yn = MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2)
      If yn = 7 Then
         Exit Sub
      End If
   End If
   frm071008.Show
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
   Unload frm071008
   Unload Me
End Sub

Private Sub cmdSure_Click()
Dim strDay1 As String, strDay2 As String, LcTmp As String
Dim strDate As String
Dim m_StrTo As String, m_contents As String

If AllTextBeforeSaveCheck Then Exit Sub

If Text(10).Text <> "" And Text(3).Text <> "" Then
   If Val(Text(10).Text) > Val(Text(3).Text) Then
      MsgBox "承辦期限不可大於本所期限!", vbExclamation, "法院文書"
      Text(10).SetFocus
      Exit Sub
   End If
End If

LcTmp = cp01 + cp02 + cp03 + cp04

'If Not objLawDll.ChkMRec(ChangeTStringToWString(Replace(lbeAccept, "/", "")), LcTmp, strDay1, strDay2) Then
'   MsgBox "本所案號'" + lbeCaseNum + "'與收受日'" + lbeAccept + "'不存在於來函記錄檔中", vbCritical
'   Exit Sub
'ElseIf strDay1 <> "" Then
'   If Text(3) <> ChangeWStringToTString(strDay1) Or Text(4) <> ChangeWStringToTString(strDay2) Then
'    If MsgBox("來函記錄檔之本所期限 '" + ChangeWStringToTString(strDay1) + "' 與法定期限 '" + ChangeWStringToTString(strDay2) + "' 與畫面輸入不同，是否儲存 ? ", vbCritical + vbYesNo) = vbNo Then
'      Text(3).SetFocus
'      Exit Sub
'    End If
'   End If
'End If
'
'2011/5/17 cancel by sonia
'   If m_CP01 = "LA" Or m_Nation = "000" Then
'      strDate = GetMailRecField(m_CP01, m_CP02, m_CP03, m_CP04, DBDATE(lbeAccept.Caption), "MR16")
'      If (TAIWANDATE(Text(3).Text) <> TAIWANDATE(strDate)) Or strDate = "" Then
'            If MsgBox("輸入的本所期限與來函記錄中的本所期限日期不同", vbYesNo, "資料檢核") = vbNo Then
'               Text(3).SetFocus
'               Exit Sub
'            End If
'      End If
'      strDate = ""
'      strDate = GetMailRecField(m_CP01, m_CP02, m_CP03, m_CP04, DBDATE(lbeAccept.Caption), "MR17")
'      If (TAIWANDATE(Text(4).Text) <> TAIWANDATE(strDate)) Or strDate = "" Then
'            If MsgBox("輸入的法定期限與來函記錄中的法定期限日期不同", vbYesNo, "資料檢核") = vbNo Then
'               Text(4).SetFocus
'               Exit Sub
'            End If
'      End If
'   End If
'2011/5/17 end

   'Add By Cheng 2002/05/24
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   Screen.MousePointer = 11
If Not SaveData Then
   DataErrorMessage (3)
Else
   'Add By Sindy 2011/10/20 發E-Mail給智權人員
   If cp01 = "FCL" Or cp01 = "LIN" Then
      m_StrTo = PUB_GetFCLSalesNo(cp01, cp02, cp03, cp04)
   Else
      m_StrTo = PUB_GetAKindSalesNo(cp01, cp02, cp03, cp04)
   End If
   'Modify By Sindy 2025/10/15 原只發EMAIL給案件最新智權人員，請修改為若有案源則發給案源介紹人(可能多個)，無案源才發給案件最新智權人員。
   strExc(0) = "select * from LawOfficeSource where LOS06='" & lbeNum & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_StrTo = Replace(RsTemp.Fields("LOS04"), ",", ";")
   End If
   '2025/10/15 END
   If m_StrTo > "" Then
      'Modified by Lydia 2015/10/05 '法務人員'改為'協辦人員'
      'Modified by Lydia 2025/11/18 Text(5) & lbe(5) 改為 Trim(cboGov.Text)
      m_contents = "本所案號：" & cp01 & "-" & cp02 & "-" & cp03 & "-" & cp04 & vbCrLf & _
                   "智權人員：" & GetPrjSalesNM(m_StrTo) & vbCrLf & _
                   "當 事 人：" & lbeCusName & vbCrLf & _
                   "收 受 日：" & lbeAccept & vbCrLf & _
                   "判決日期：" & ChangeTStringToTDateString(Text(0)) & vbCrLf & _
                   "來文性質：" & Text(1) & lbe(1) & vbCrLf & _
                   "下一程序：" & Text(2) & lbe(2) & vbCrLf & _
                   "本所期限：" & ChangeTStringToTDateString(Text(3)) & vbCrLf & _
                   "法定期限：" & ChangeTStringToTDateString(Text(4)) & vbCrLf & _
                   "機關代號：" & Trim(cboGov.Text) & vbCrLf & _
                   "股    別：" & Text(6) & vbCrLf & _
                   "法院案號：" & Text(7) & vbCrLf & _
                   "機關文號：" & Text(8) & vbCrLf & _
                   "承 辦 人：" & Text(9) & lbe(9) & vbCrLf & _
                   "協辦人員：" & Text(11) & lbe(11) & vbCrLf & _
                   "承辦期限：" & ChangeTStringToTDateString(Text(10)) & vbCrLf & _
                   "進度備註：" & Text(12) & vbCrLf & _
                   "案件備註：" & Text(13) & vbCrLf
                  PUB_SendMail strUserNum, m_StrTo, "", cp01 & "-" & cp02 & "-" & cp03 & "-" & cp04 & "法院文書", m_contents
   End If
   '2011/10/20 End
   
   Unload frm071009
   Set frm071009 = Nothing
   Unload frm071008
   frm071008.Show
End If
Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
  Dim i As Integer, n As Integer
  
   m_CP01 = frm071008.txtcp01
   m_CP02 = frm071008.txtcp02
   If frm071008.txtcp03 <> "" Then
      m_CP03 = frm071008.txtcp03
   Else
      m_CP03 = "0"
   End If
   
  If frm071008.txtcp04.Text <> "" Then
      m_CP04 = frm071008.txtcp04.Text
   Else
      m_CP04 = "00"
   End If

   MoveFormToCenter Me
   blnIsSave = False
   With frm071008.MSHFlexGrid1
   'n = 0
   For i = 1 To .Rows - 1
      .row = i
      .col = 0
       If .Text = "v" Then
          .col = 2
         'ReDim Preserve strCP09(n)
          m_CP09 = .Text
          lbeNum = .Text
    '     n = n + 1
         .col = 4
         Strsale = .Text
       End If
   Next
   End With
   cp01 = frm071008.txtcp01
   cp02 = frm071008.txtcp02
   cp03 = frm071008.txtcp03
   cp04 = frm071008.txtcp04
   lbeCaseNum = GiveSymbol(cp01, cp02, cp03, cp04)
   lbeAccept = ChangeTStringToTDateString(frm071008.txtAccept)
   lbeCustomer = frm071008.lbeCusNum
   lbeCusName = frm071008.lbeCusName
   GetData (0)
End Sub

Private Sub GetData(ByVal Init As Integer)
 Dim i As Integer
 Dim strName As String
 
   If cp01 <> "LA" Then
      strExc(1) = "select cp09,cp10,cp46,cp25,cp06,cp07,cp71,cp30,cp08,cp13," + _
          " cp14,cp48,cp29,cp64 ,lc15,lc27,lc08 from caseprogress,lawcase where cp09='" + m_CP09 + _
          "' and CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04"
   Else
      strExc(1) = "select cp09,cp10,cp46,cp25,cp06,cp07,cp71,cp30,cp08,cp13," + _
          " cp14,cp48,cp29,cp64 ,hc12,hc09 from caseprogress,hirecase where cp09='" + m_CP09 + _
          "' and CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04"
   End If
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set rs = objLawDll.ReadRstMsg(intI, strExc(1))
   Set Rs = ClsLawReadRstMsg(intI, strExc(1))

 If intI = 1 Then
 'lbeNum = strCP09(init)
      If Not IsNull(Rs.Fields!cp13) Then Strsale = Rs.Fields!cp13
      If Not IsNull(Rs.Fields!CP10) Then lbeProperty = Rs.Fields!CP10: lbePropertyName = ChgType(1, Rs.Fields!CP10)
      If m_CP01 <> "LA" Then
         If Not IsNull(Rs.Fields!lc15) Then
            m_Nation = Rs.Fields("LC15")
         End If
      ElseIf m_CP01 = "LA" Then
         m_Nation = "000"
      End If
      strName = ""
      '承辦人
      If Not IsNull(Rs.Fields("CP14")) Then
         Text(9).Text = Rs.Fields("CP14")
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(Text(9).Text, strName) Then
         If ClsPDGetStaff(Text(9).Text, strName) Then
            lbe(9).Caption = strName
         End If
      End If
      strName = ""
      '協辦人員
      If Not IsNull(Rs.Fields("CP29")) Then
         Text(11).Text = Rs.Fields("CP29")
         'edit by nickc 2007/02/07 不用 dll 了
         'If objPublicData.GetStaff(Text(11).Text, strName) Then
         If ClsPDGetStaff(Text(11).Text, strName) Then
            lbe(11).Caption = strName
         End If
      End If
      'Added by Lydia 2025/11/18 改成下拉選單
      Call PUB_SetGovCmb(Me.cboGov, m_GovListNew)
      m_GovListDef = m_GovListNew
      'end 2025/11/18
 End If
End Sub
Private Function SaveData() As Boolean
 Dim strNewNum As String, strNum As String, strTemp As String
 Dim i As Integer
 Dim strNP22 As String
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
SaveData = True
cnnConnection.BeginTrans
   
   i = 1
   '911028 nick 移到下面
   ' 序號
   'strNP22 = GetNextProgressNo()
   'edit by nickc 2007/02/07 不用 dll 了
   'If objPublicData.GetAutoNumber("C", strNewNum, 1, 1) Then
   If ClsPDGetAutoNumber("C", strNewNum, 1, 1) Then
      'Modify By Sindy 2010/8/18 比對自動編號年度
      'strNum = "C" + CStr(Year(Date) - 1911) + strNewNum
      strNum = "C" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) + strNewNum
   End If
   
   ' 91.04.04 modify by louis (修改單引號)
   'Modify By Cheng 2003/04/07
   '智權人員存最近收文A類接洽記錄單的智權人員
   '2009/9/9 MODIFY BY SONIA 改FCL,LIN之業務區,智權人員
   'Modify By Sindy 2011/10/20 存檔時同時將此C類來文上發文日
   'Modified by Lydia 2025/11/18 CNULL(Text(5)) 改為 CNULL(Trim(Left(cboGov.Text, 3)))
   strExc(1) = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp12,cp13,cp32,cp43,cp20,cp26," + _
      " cp10,cp08,cp35,CP14,CP48,CP29,CP64,CP71,CP30,cp27) values (" + CNULL(strNum) + "," + CNULL(m_CP01) + "," + _
      CNULL(m_CP02) + "," + CNULL(m_CP03) + "," + CNULL(m_CP04) + "," + CNULL(ChangeTStringToWString(Replace(lbeAccept, "/", ""))) + "," + _
      CNULL(ChangeTStringToWString(Text(3))) + "," + CNULL(ChangeTStringToWString(Text(4))) + "," + CNULL(GetSalesArea(IIf(cp01 = "FCL", PUB_GetFCLSalesNo(m_CP01, m_CP02, m_CP03, m_CP04), IIf(cp01 = "LIN", PUB_GetFCLSalesNo(m_CP01, m_CP02, m_CP03, m_CP04), PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04))))) + _
      "," + IIf(cp01 = "FCL", CNULL(PUB_GetFCLSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)), IIf(cp01 = "LIN", CNULL(PUB_GetFCLSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)), CNULL(PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)))) + ",'N'," + CNULL(lbeNum) + ",'N','N'," + CNULL(Text(1)) + "," + CNULL(Text(8)) + "," + _
      CNULL(Text(7)) + "," + CNULL(Text(9)) + "," + CNULL(ChangeTStringToWString(Text(10))) + "," + CNULL(Text(11)) + "," + CNULL(ChgSQL(Text(12))) + "," + CNULL(Trim(Left(cboGov.Text, 3))) + "," + CNULL(Text(6)) + "," + strSrvDate(1) + ")"
   'Add By Cheng 2002/11/07
   cnnConnection.Execute strExc(1)
    
   'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
   Pub_UpdateFromMaxCP27 m_CP01, m_CP02, m_CP03, m_CP04
    
    If Text(2) <> "" Then
      ' 91.04.04 modify by louis (修改單引號)
      '911028 nick 移下來
        ' 序號
        strNP22 = GetNextProgressNo()
        'Modify By Cheng 2003/04/07
        '智權人員存最近收文A類接洽記錄單的智權人員
      strExc(2) = "insert into nextprogress(np01,np02,np03,np04,np05,np07,np10,np13,np15,NP22,NP08,NP09) values (" + CNULL(strNum) + "," + CNULL(m_CP01) + "," + CNULL(m_CP02) + _
         "," + CNULL(m_CP03) + "," + CNULL(m_CP04) + "," + CNULL(Text(2)) + "," + IIf(cp01 = "FCL", CNULL(PUB_GetFCLSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)), IIf(cp01 = "LIN", CNULL(PUB_GetFCLSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)), CNULL(PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)))) + "," + CNULL(Text(8)) + _
         "," + CNULL(ChgSQL(Text(12))) + "," + strNP22 + "," + CNULL(ChangeTStringToWString(Text(3))) + "," + CNULL(ChangeTStringToWString(Text(4))) + ")"
    'Add By Cheng 2002/11/07
    cnnConnection.Execute strExc(2)
      i = 2
    End If
    If Text(0) <> "" Then
      strExc(3) = "update caseprogress set cp25=" + CNULL(ChangeTStringToWString(Text(0))) + " where cp09='" + lbeNum + "' "
    'Add By Cheng 2002/11/07
    cnnConnection.Execute strExc(3)
      i = 3
    End If
'    SaveData = objLawDll.ExecSQL(i, strExc)
   If SaveData Then blnIsSave = True Else blnIsSave = False
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    SaveData = False
End Function

Private Function ChgType(i As Integer, strText As String) As String
 Dim strTemp As String
   
   Select Case i
      Case 3, 4, 10
        ChgType = ChangeWStringToTString(strText)
      Case 1, 2
          'edit by nickc 2007/02/07 不用 dll 了
          'If objPublicData.GetCaseProperty(CP01, StrText, strTemp, False) Then ChgType = strTemp
          If ClsPDGetCaseProperty(cp01, strText, strTemp, False) Then ChgType = strTemp
      Case 5
            'edit by nickc 2007/02/07 不用 dll 了
            'If objLawDll.GetGovName(StrText, strTemp) Then ChgType = strTemp
            If ClsPDGetGovName(strText, strTemp) Then ChgType = strTemp
      Case 9, 11
            'edit by nickc 2007/02/07 不用 dll 了
            'If objPublicData.GetStaff(StrText, strTemp) Then ChgType = strTemp
            If ClsPDGetStaff(strText, strTemp) Then ChgType = strTemp
   End Select
  
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm071009 = Nothing
End Sub

Private Sub Text_Change(Index As Integer)
Select Case Index
Case 1, 2, 5, 9, 11
    If Text(Index) = "" Then lbe(Index) = ""
 End Select
End Sub

Private Sub Text_GotFocus(Index As Integer)
   Select Case Index
   Case Index
      TextInverse Text(Index)
   End Select
   
   Select Case Index
   Case 6, 12, 13
      'edit by nickc 2007/06/11  切換輸入法改用API
      'Text(Index).IMEMode = 1
      OpenIme
   Case Else
      'edit by nickc 2007/06/11  切換輸入法改用API
      'Text(Index).IMEMode = 2
      CloseIme
   End Select
End Sub

'Modified by Lydia 2021/09/14 改成Form 2.0
'Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
  KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text_LostFocus(Index As Integer)
   Select Case Index
   Case 6, 12, 13
      'edit by nickc 2007/06/11  切換輸入法改用API
      'Text(Index).IMEMode = 2
      CloseIme
   Case 1
      If ChkUData(m_Nation, Text(1).Text) Then
      End If
      
   End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim LcTmp As String, strTemp1 As String, strTemp2 As String
Select Case Index
Case 0
 If Text(Index) <> "" Then
    If CheckIsTaiwanDate(Text(Index)) Then
       If Val(GetTaiwanTodayDate) - Val(Text(Index)) < 0 Then
           MsgBox "輸入日期大於系統日", vbCritical
           Cancel = True
        End If
    Else
       Cancel = True
    End If
 End If
Case 1, 2 '來文性質, 下一程序
   If Text(Index) <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetCaseProperty(CP01, Text(Index), strTemp1, False) Then
      If ClsPDGetCaseProperty(cp01, Text(Index), strTemp1, False) Then
         lbe(Index) = strTemp1
      Else
         Cancel = True
      End If
   Else
     If Index = 1 Then Cancel = True: DataErrorMessage 5, "來文性質"
   End If
 '本所期限, 法定期限：可空白，但有輸入下一程序欄時，則此二欄不可空白。檢查日期且本所期限必須≦法定期限。
 Case 3 '本所期限
   '若有輸入本所期限
   If Text(Index) <> "" Then
      If CheckIsTaiwanDate(Text(Index)) Then
         'Add By Cheng 2002/03/11
         '若有輸入本所期限時, 本所期限不可小於系統日
         If (Val(Me.Text(Index).Text) + 19110000) < ServerDate Then
            DataErrorMessage 10, "本所期限"
            Cancel = True
            Exit Sub
         End If
                 
         If Text(4) <> "" Then
            If Val(Text(4)) - Val(Text(Index)) < 0 Then
               DataErrorMessage 13
               Cancel = True
            End If
         End If
      Else
         DataErrorMessage 2, "本所期限"
         Cancel = True
      End If
    '若有輸入下一程序, 但未輸入本所期限
    ElseIf Text(2) <> "" And Text(Index) = "" Then
            MsgBox "本所期限不可空白"
            Cancel = True
    End If

 Case 4 '法定期限
   If Text(Index) <> "" Then
      If CheckIsTaiwanDate(Text(Index)) Then
         If Text(3) <> "" Then
            If Val(Text(Index)) - Val(Text(3)) < 0 Then
               DataErrorMessage 12
               Cancel = True
            End If
         End If
      Else
         Cancel = True
      End If
   '若有輸入下一程序時, 法定期限不可空白
   ElseIf Text(2) <> "" And Text(Index) = "" Then
      MsgBox "法定期限不可空白"
      Cancel = True
   End If
Case 5
   If Text(Index) <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objLawDll.GetGovName(Text(Index), strTemp1) Then Lbe(Index) = strTemp1 Else Cancel = True
      If ClsPDGetGovName(Text(Index), strTemp1) Then lbe(Index) = strTemp1 Else Cancel = True
   End If
   
Case 6
   If CheckLengthIsOK(Text(6), 20) = False Then
      Cancel = True
      Text(6).SetFocus
      TextInverse Text(6)
      Exit Sub
   End If
Case 7
   If CheckLengthIsOK(Text(7), 32) = False Then
      Cancel = True
      Text(7).SetFocus
      TextInverse Text(7)
      Exit Sub
   End If
Case 8
   If CheckLengthIsOK(Text(8), 40) = False Then
      Cancel = True
      Text(8).SetFocus
      TextInverse Text(8)
      Exit Sub
   End If
   
Case 10
 If Text(Index) <> "" Then
'     If CheckIsTaiwanDate(Text(Index)) Then
'       If Val(GetTaiwanTodayDate) - Val(Text(Index)) > 0 Then
'           MsgBox "輸入日期小於系統日", vbCritical
'           Cancel = True
'        End If
'    Else
'       Cancel = True
'    End If
     If CheckIsTaiwanDate(Text(Index)) Then
         If Text(3).Text <> "" Then
             If Val(Text(Index).Text) > Val(Text(3).Text) Then
                MsgBox "承辦期限不可大於本所期限!", vbExclamation, "法院文書"
                Cancel = True
             End If
         End If
      Else
          Cancel = True
      End If
  End If

Case 9, 11
   If Text(Index) <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetStaff(Text(Index), strTemp1) Then Lbe(Index) = strTemp1 Else Cancel = True
      If ClsPDGetStaff(Text(Index), strTemp1) Then lbe(Index) = strTemp1 Else Cancel = True
   End If
Case 12, 13
    If Text(Index) <> "" Then
      If CheckLengthIsOK(Text(Index), 2000) = False Then
          Cancel = True
      End If
    End If
End Select
If Cancel Then TextInverse Text(Index)

End Sub
Private Function ChkUData(strNation As String, strProperty As String) As Boolean
Dim dblTempDays As Double, strDate As Variant
Dim nDays As Integer
  'edit by nickc 2007/02/07 不用 dll 了
  'If objLawDll.GetCaseFee(CP01, strNation, strProperty, dblTempDays) Then
''''edit by nickc 2007/10/12 改抓有時效性的
''''  If ClsLawGetCaseFee(CP01, strNation, strProperty, dblTempDays) Then
''''     nDays = CInt(dblTempDays)
''''     If nDays <> 0 Then
''''        strDate = ChangeWStringToTString(CompWorkDay(nDays, ChangeTStringToWString(ChangeTDateStringToTString(lbeAccept))))
''''     ' strDate = DateAdd("d", dblTempDays, CDate(DateSerial(Val(Left(lbeAccept, 2)) + 1911, Val(Mid(lbeAccept, 4, 2)), Val(Right(lbeAccept, 2)))))
''''       Text(10).Text = strDate
''''      End If
''''  End If
    Text(10).Text = TAIWANDATE(Pub_GetHandleDay(cp01, strNation, strProperty, DBDATE(lbeAccept), Text(3)))
End Function
Private Function AllTextBeforeSaveCheck() As Boolean
  Dim strTemp  As String
  
  AllTextBeforeSaveCheck = True
    
  strTemp = ""
  
   If Text(0) <> "" Then
      If CheckIsTaiwanDate(Text(0)) Then
       If Val(GetTaiwanTodayDate) - Val(Text(0)) < 0 Then
           MsgBox "輸入日期大於系統日", vbCritical
           AllTextBeforeSaveCheck = True
           TextInverse Text(0)
           Exit Function
        End If
     Else
         AllTextBeforeSaveCheck = True
         TextInverse Text(0)
         Exit Function
     End If
   End If
    
   If Text(1) <> "" Then
      'edit by nickc 2007/02/07 不用 dll 了
      'If Not objPublicData.GetCaseProperty(CP01, Text(1), strTemp, False) Then
      If Not ClsPDGetCaseProperty(cp01, Text(1), strTemp, False) Then
         AllTextBeforeSaveCheck = True
         TextInverse Text(1)
         Exit Function
      End If
    Else
       MsgBox "來文性質不可空白", vbCritical
       AllTextBeforeSaveCheck = True
       Text(1).SetFocus
       Exit Function
    End If
   '檢查下一程序
   If Text(2) <> "" Then
      strTemp = ""
      'edit by nickc 2007/02/07 不用 dll 了
      'If Not objPublicData.GetCaseProperty(CP01, Text(2), strTemp, False) Then
      If Not ClsPDGetCaseProperty(cp01, Text(2), strTemp, False) Then
         AllTextBeforeSaveCheck = True
         TextInverse Text(2)
         Exit Function
      End If
   End If
   '本所期限, 法定期限：可空白，但有輸入下一程序欄時，則此二欄不可空白。檢查日期且本所期限必須≦法定期限。
   '檢查本所期限
   '若有輸入本所期限
   If Text(3) <> "" Then
     If CheckIsTaiwanDate(Text(3)) Then
         'Add By Cheng 2002/03/11
         '若有輸入本所期限時, 本所期限不可小於系統日
         If (Val(Me.Text(3).Text) + 19110000) < ServerDate Then
            DataErrorMessage 10, "本所期限"
            AllTextBeforeSaveCheck = True
            TextInverse Text(3)
            Exit Function
         End If
        
        If Text(4) <> "" Then
           If Val(Text(4)) - Val(Text(3)) < 0 Then
               DataErrorMessage 13
               AllTextBeforeSaveCheck = True
               TextInverse Text(3)
               Exit Function
           End If
        End If
      Else
         AllTextBeforeSaveCheck = True
         TextInverse Text(3)
         Exit Function
      End If
   '若有輸入下一程序, 但未輸入本所期限
   ElseIf Text(2) <> "" And Text(3) = "" Then
         MsgBox "本所期限不可空白"
         AllTextBeforeSaveCheck = True
         TextInverse Text(3)
         Exit Function
   End If

   If Text(4) <> "" Then
      If CheckIsTaiwanDate(Text(4)) Then
        If Text(3) <> "" Then
           If Val(Text(4)) - Val(Text(3)) < 0 Then
              DataErrorMessage 12
              AllTextBeforeSaveCheck = True
              TextInverse Text(4)
              Exit Function
           End If
        End If
      Else
         AllTextBeforeSaveCheck = True
         TextInverse Text(4)
         Exit Function
      End If
   ElseIf Text(2) <> "" And Text(4) = "" Then
            MsgBox "法定期限不可空白"
            AllTextBeforeSaveCheck = True
            TextInverse Text(4)
            Exit Function
   End If
   
   'Modified by Lydia 2025/11/18
   'If Text(5) <> "" Then
   '    strTemp = ""
   '   'edit by nickc 2007/02/07 不用 dll 了
   '   'If objLawDll.GetGovName(Text(5), strTemp) Then
   '   If ClsPDGetGovName(Text(5), strTemp) Then
   '      lbe(5) = strTemp
   '   Else
   '       AllTextBeforeSaveCheck = True
   '       TextInverse Text(5)
   '       Exit Function
   '   End If
   'End If
   If Trim(cboGov.Text) <> "" Then
      AllTextBeforeSaveCheck = False
      Call cboGov_Validate(AllTextBeforeSaveCheck)
      If AllTextBeforeSaveCheck = True Then
         Exit Function
      End If
   End If
   'end 2025/11/18
   
   If CheckLengthIsOK(Text(6), 20) = False Then
      AllTextBeforeSaveCheck = True
      Text(6).SetFocus
      TextInverse Text(6)
      Exit Function
   End If

   If CheckLengthIsOK(Text(7), 32) = False Then
      AllTextBeforeSaveCheck = True
      Text(7).SetFocus
      TextInverse Text(7)
      Exit Function
   End If

   If CheckLengthIsOK(Text(8), 40) = False Then
      AllTextBeforeSaveCheck = True
      Text(8).SetFocus
      TextInverse Text(8)
      Exit Function
   End If
   
  If Text(9) <> "" Then
      strTemp = ""
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetStaff(Text(9), strTemp) Then
      If ClsPDGetStaff(Text(9), strTemp) Then
         lbe(9) = strTemp
      Else
         AllTextBeforeSaveCheck = True
         TextInverse Text(9)
         Exit Function
      End If
   End If
  
  If Text(11) <> "" Then
      strTemp = ""
      'edit by nickc 2007/02/07 不用 dll 了
      'If objPublicData.GetStaff(Text(11), strTemp) Then
      If ClsPDGetStaff(Text(11), strTemp) Then
         lbe(11) = strTemp
      Else
         AllTextBeforeSaveCheck = True
         TextInverse Text(11)
         Exit Function
      End If
   End If
  
  If Text(10).Text <> "" Then
     If CheckIsTaiwanDate(Text(10)) Then
         If Text(3).Text <> "" Then
             If Val(Text(10).Text) > Val(Text(3).Text) Then
                MsgBox "承辦期限不可大於本所期限!", vbExclamation, "法院文書"
                 AllTextBeforeSaveCheck = True
                 TextInverse Text(10)
                 Exit Function
             End If
         End If
      Else
            AllTextBeforeSaveCheck = True
            Text(10).SetFocus
            TextInverse Text(10)
            Exit Function
      End If
  End If
  
  If Text(12) <> "" Then
     If CheckLengthIsOK(Text(12), 2000) = False Then
         AllTextBeforeSaveCheck = True
         TextInverse Text(12)
         Exit Function
     End If
  End If
  
  If Text(13) <> "" Then
     If CheckLengthIsOK(Text(13), 2000) = False Then
         AllTextBeforeSaveCheck = True
         TextInverse Text(13)
         Exit Function
     End If
  End If
  
  AllTextBeforeSaveCheck = False

End Function
'Add By Cheng 2002/05/24
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Me.Text
   If objTxt.Enabled = True Then
      Cancel = False
      Text_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Added by Lydia 2021/09/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If


TxtValidate = True
End Function

'Added by Lydia 2025/11/18
Private Sub cboGov_GotFocus()
   TextInverse cboGov
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_Validate(Cancel As Boolean)

   If Trim(cboGov.Text) <> "" And cboGov.Tag <> cboGov.Text Then
      If PUB_ChkGovIsExist(IIf(Val(Trim(Left(cboGov.Text, 3))) > 0, Trim(Left(cboGov.Text, 3)), Trim(cboGov.Text)), strExc(3), strExc(4)) = True Then
         cboGov.Text = strExc(3) & " " & strExc(4)
      Else
         Cancel = True
         cboGov.SetFocus
         cboGov_GotFocus
         Exit Sub
      End If
   End If
   cboGov.Tag = cboGov.Text
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2025/11/18
Private Sub cboGov_DropButtonClick()
   If cboGov.Text <> "" Then
      If Val(Trim(Left(cboGov.Text, 3))) > 0 Then
      Else  '依輸入文字模糊比對
         Call PUB_SetGovCmb(cboGov, m_GovListNew, , Trim(cboGov.Text))
         If m_GovListNew = "" Then
            Call PUB_SetGovCmb(cboGov, m_GovListNew)
         End If
      End If
   Else
      If m_GovListNew <> m_GovListDef Then
         Call PUB_SetGovCmb(cboGov, m_GovListNew)
      End If
   End If
End Sub

