VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050102_7 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文（優先權證明書）"
   ClientHeight    =   4860
   ClientLeft      =   336
   ClientTop       =   996
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8580
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   5280
      MaxLength       =   4
      TabIndex        =   9
      Top             =   3150
      Width           =   540
   End
   Begin VB.TextBox txtChkRltDate 
      Height          =   270
      Left            =   5130
      MaxLength       =   8
      TabIndex        =   7
      Top             =   2820
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   5070
      TabIndex        =   1
      Top             =   1830
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7632
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5580
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6408
      TabIndex        =   12
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   19
      Left            =   1575
      TabIndex        =   6
      Top             =   2820
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   6
      Left            =   6150
      TabIndex        =   5
      Top             =   2490
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   5
      Left            =   6150
      TabIndex        =   3
      Top             =   2160
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   2
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   2190
      Width           =   495
      VariousPropertyBits=   671107099
      MaxLength       =   2
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1860
      Width           =   1095
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   2520
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   1335
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   3450
      Width           =   8355
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14737;2355"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數："
      Height          =   180
      Index           =   18
      Left            =   4320
      TabIndex        =   39
      Top             =   3195
      Width           =   900
   End
   Begin VB.Label lblCaseFee 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   6120
      TabIndex        =   8
      Tag             =   "Y"
      Top             =   2790
      Width           =   255
   End
   Begin VB.Label lblChkRltDate 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   4320
      TabIndex        =   37
      Top             =   2835
      Width           =   765
   End
   Begin VB.Label Label4 
      Caption         =   "是否印傳真封面：        （N:不印）"
      Height          =   255
      Index           =   4
      Left            =   135
      TabIndex        =   36
      Top             =   2835
      Width           =   2685
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容：          （Y：Word）"
      Height          =   180
      Left            =   4320
      TabIndex        =   35
      Top             =   2520
      Width           =   3315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否修改指示信內容：             (Y：Word)"
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   34
      Top             =   2190
      Width           =   3210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請優先權證明書  共                份"
      Height          =   180
      Left            =   120
      TabIndex        =   33
      Top             =   2220
      Width           =   2910
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   6330
      TabIndex        =   32
      Top             =   1860
      Width           =   2175
      VariousPropertyBits=   27
      Size            =   "3836;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label23 
      Caption         =   "代理人："
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "發文日："
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "是否列印通知函：        （N：不印）"
      Height          =   180
      Left            =   120
      TabIndex        =   29
      Top             =   2550
      Width           =   2820
   End
   Begin VB.Label Label2 
      Caption         =   "進度備註："
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   28
      Top             =   3150
      Width           =   975
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   720
      Width           =   2535
   End
   Begin MSForms.Label lblSalesName 
      Height          =   255
      Left            =   6000
      TabIndex        =   20
      Top             =   1080
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   19
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   18
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   17
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   16
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   15
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   14
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label8 
      Caption         =   "智權人員："
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   27
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   "本所期限："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "專利種類："
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   25
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收文號："
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "法定期限："
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   22
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblCaseFees 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   6165
      TabIndex        =   38
      Top             =   2850
      Width           =   255
   End
End
Attribute VB_Name = "frm050102_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (txtCaseField,lblSalesName...)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
'Add By Cheng 2003/09/16
Dim strCountry As String '存放EPC指定國家
Dim m_bActived As Boolean
'Add By Sindy 2018/1/8
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2018/1/8 END
Dim m_strAF01 As String, m_strLD18 As String 'Added by Morgan 2018/8/22

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(1 To 5) As String, intStep As Integer, i As Integer
   EndLetter ET01, cp(9), ET03, strUserNum
   
   Select Case txtCaseField(2)
      Case "1"
         strExc(0) = "one certified copy"
      Case "2"
         strExc(0) = "two certified copies"
      Case "3"
         strExc(0) = "three certified copies"
      Case "4"
         strExc(0) = "four certified copies"
      Case "5"
         strExc(0) = "five certified copies"
      Case "6"
         strExc(0) = "six certified copies"
      Case "7"
         strExc(0) = "seven certified copies"
      Case "8"
         strExc(0) = "eight certified copies"
      Case "9"
         strExc(0) = "nine certified copies"
      Case "10"
         strExc(0) = "ten certified copies"
   End Select
   
   strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
      "','申請優先權證明書幾份','" & strExc(0) & "')"
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(1, strTxt) Then
   If Not ClsLawExecSQL(1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
 Dim i As Integer
 Dim stLetter As String 'Add by Morgan 2006/6/20

   Select Case Index
      Case 0
      
         'Added by Morgan 2015/8/7
         If PUB_ChkFileNP(cp(9)) Then
            MsgBox "下一程序已有提申或收達期限，不可發文！"
            Exit Sub
         End If
         'end 2015/8/7
   
         If txtCaseField(3) = "" And txtCaseField(2) = "" Then
            MsgBox "請輸入優先權證明書份數 !", vbCritical
            txtCaseField(2).SetFocus
            Exit Sub
         End If
         Screen.MousePointer = vbHourglass
         'Modify By Cheng 2002/07/30
'         For i = 0 To 4
         For i = 0 To 6
            If i <> 1 Then
                If txtCaseField(i).Enabled Then
                     If CheckKeyIn(i) <> 1 Then
                        txtCaseField(i).SetFocus
                        txtCaseField_GotFocus (i)
                        Exit For
                     End If
                End If
            'Add By Cheng 2002/08/19
            Else
               If CheckKeyIn(i) <> 1 Then
                  Me.Combo1.SetFocus
                  Exit For
               End If
            End If
         Next i
         'Modify By Cheng 2002/07/30
'         If i = 5 Then
         If i = 7 Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
            
            If SaveDatabase Then
            
               'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
               PUB_CheckEMail cp(44), cp(116)
               PUB_CheckEMail field(75), field(144)
               If field(145) <> "" Then
                  PUB_CheckEMail field(75), field(145)
               End If
               'end 2008/2/20
               
               '指示信
               StartLetter "01", "30"
               'Modify by Morgan 2006/6/20 加傳真封面--禧佩
               'NowPrint cp(9), "01", "30", IIf(Me.txtCaseField(5).Text = "Y", True, False), strUserNum, 0
               '加是否印傳真封面選項
               If txtCaseField(19) <> "N" Then
                  If Me.txtCaseField(5).Text = "Y" Then
                     NowPrint cp(9), "01", "99", False, strUserNum, , , True, stLetter, , , , , , , , , m_strAF01
                  Else
                     NowPrint cp(9), "01", "99", False, strUserNum, , , , , , , , , , , , , m_strAF01
                     stLetter = ""
                  End If
                  If m_strAF01 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/20
               End If
               NowPrint cp(9), "01", "30", IIf(Me.txtCaseField(5).Text = "Y", True, False), strUserNum, 0, stLetter, , , , , , , , , , , m_strAF01
               'end 2006/6/20
               
               'Added by Morgan 2018/8/22 CFP電子化
               If txtCaseField(5).Text = "Y" And m_strAF01 <> "" Then
                  frm1105_1.m_RecNo = m_strAF01
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
                  frm1105_1.Show
                  If txtCaseField(6).Text = "Y" Then
                     MsgBox "指示信編輯中，客戶函請至定稿維護修改！", vbExclamation
                     txtCaseField(6).Text = ""
                  End If
               End If
               'end 2018/8/22
               
               '通知函
               If txtCaseField(3) <> "N" Then
                  StartLetter "01", "00"
                  'Modify By Cheng 2002/07/30
'                  NowPrint cp(9), "01", "00", False, strUserNum, 0
                  NowPrint cp(9), "01", "00", IIf(Me.txtCaseField(6).Text = "Y", True, False), strUserNum, 0, , , , , , , , , , , , m_strLD18
                  
                  'Added by Morgan 2018/8/22 CFP電子化
                  If txtCaseField(6).Text = "Y" And m_strLD18 <> "" Then
                     frm1105_1.m_RecNo = m_strLD18
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".CUS.PDF"
                     frm1105_1.Show
                  End If
                  'end 2018/8/22
               End If
               
               bolLeave = True
               intLeaveKind = 1
               'Add By Cheng 2002/04/30
               '若有未發文資料顯示警告
               PUB_GetCPunIssueDatas "" & Me.lblCaseField(1).Caption
               
               Unload Me
            '911202 nick
            Else
                MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
            End If
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         'Add By Sindy 2013/5/28
         If frm050102_1.bolIsEMPFlow = True Then
            intLeaveKind = 0
            'Unload frm050102_1
            frm090202_4.Show
            frm090202_4.QueryData
         '2013/5/28 End
         'Add By Sindy 2018/1/8
         ElseIf Me.m_strIR01 <> "" Then
            intLeaveKind = 0
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         '2018/1/8 END
         Else
            If Index = 2 Then
               intLeaveKind = 0
            Else
               intLeaveKind = 1
            End If
         End If
         Unload Me
   End Select
   
   ' 發文回前畫面時
   Select Case Index
      Case 0:
         ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
         'Add By Sindy 2013/5/28
         If frm050102_1.bolIsEMPFlow = True Then
            intLeaveKind = 0
            'Unload frm050102_1
            frm090202_4.Show
            frm090202_4.QueryData
         '2013/5/28 End
         'Add By Sindy 2018/1/8
         ElseIf Me.m_strIR01 <> "" Then
            intLeaveKind = 0
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         '2018/1/8 END
         Else
            frm050102_1.Clear
         End If
   End Select
End Sub
Private Function SaveDatabase() As Boolean
Dim strTxt(1 To 10) As String, iStep As Integer
Dim strLetterJudge As String, strSubject As String '指示信判發人/主旨 Added by Morgan 2018/8/22

'911106 nick transation
SaveDatabase = True
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   cp(27) = txtCaseField(0)
   
   'Modify by Morgan 2008/2/21
   'cp(44) = Combo1
   intI = InStr(Combo1, "-")
   If intI > 0 Then
      cp(44) = Left(Combo1, intI - 1)
      cp(116) = Mid(Combo1, intI + 1)
   Else
      cp(44) = Combo1
      cp(116) = ""
   End If
   'end 2008/2/21
   cp(44) = ChangeCustomerL(cp(44))
   
   cp(64) = txtCaseField(4)
   
   cp(113) = txtCP113 'Added by Lydia 2021/05/25 工作時數
   
   'Modified by Morgan 2012/2/15 改呼叫共用函數
   'strExc(0) = "select cp45 from caseprogress where cp01=" + CNULL(cp(1)) + _
   '   " and cp02=" + CNULL(cp(2)) + " and cp03=" + CNULL(cp(3)) + _
   '   " and cp04=" + CNULL(cp(4)) + " and cp44=" + CNULL(cp(44)) + " order by cp27 desc"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   'cp(45) = ""
   'If intI = 1 And Not IsNull(RsTemp.Fields("CP45")) Then cp(45) = RsTemp.Fields("CP45")
   If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   'end 2012/2/15
   
   strTxt(1) = GetCPSQL(cp())
   
   '911106 nick transation
   cnnConnection.Execute strTxt(1)
   'SaveDatabase = objLawDll.ExecSQL(1, strTxt)
    'Add By Cheng 2003/09/16
    '若有ECP指定國家, 則新增案件進度檔資料
    If field(9) = EPC指定國家 And strCountry <> "" Then
        'Modify by Morgan 2006/12/25
        'If Not objPublicData.SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry) Then
        If Not PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry) Then
            GoTo CheckingErr
        End If
    End If
   ' 90.12.05 modify by louis 若案件國家收費表存在代理人收達天數則新增一筆收達的下一程序檔
   If SaveDatabase = True Then
      'Modify by Morgan 2015/8/7 發文收達期限管控改呼叫公用函式
      PUB_SetArriveDate cp(9)
      'end 2015/8/7
   End If
   
   'Added by Morgan 2015/8/7
   '提申管制
   PUB_SetApplyDate cp(1), cp(2), cp(3), cp(4), cp(7), cp(9), cp(10), txtCaseField(0), field(9)
   'end 2015/8/7
   
   'Add by Morgan 2009/8/18
   If txtChkRltDate <> "" Then
      PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), cp(10), cp(43)
   End If
   
   'Add By Sindy 2015/8/3 發文時,若工程師各項日期未輸入者,自動更新為發文日
   Call PUB_UpdEmpDate(cp(9), cp(1), cp(10), DBDATE(cp(27)))
   
   'Add by Sindy 2018/1/8
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050102_1"
   End If
   '2018/1/8 END
   
   'Added by Morgan 2018/8/22 CFP電子化
   If strSrvDate(1) >= CFP指示信電子化啟用日 Then
      strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), cp(10), field(9))
      strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), cp(4), cp(10), field(11), cp(45), field(9))
      PUB_AddAppForm cp(9), True, strLetterJudge, strSubject
      m_strAF01 = cp(9)
   End If
   If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
      If txtCaseField(3) <> "N" Then
         strLetterJudge = PUB_GetLetterJudgeNew("1", field(1), cp(10), field(9))
         PUB_AddLetterProgress cp(9), 0, True, strLetterJudge, False, field(26), cp(10), field(75)
         m_strLD18 = cp(9)
      End If
   End If
   'end 2018/8/22
   
   cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    SaveDatabase = False
     cnnConnection.RollbackTrans
   
End Function
Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String, strTemp1 As String, j As Integer
Dim adoRecord As Object, strSameName As String

On Error GoTo ErrHnd
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.ReadAllData(frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.Row, 5), cp(), field(), intCaseKind, intPWhere) Then
ReDim cp(TF_CP) As String
cp(9) = frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.row, 5)

If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
   lblCaseField(0) = cp(9)
   lblCaseField(1) = cp(1) + " - " + cp(2) + _
      IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
      IIf(cp(4) = "00", "", " - " + cp(4))
   lblCaseField(2) = TransDate(cp(6), 1)
   lblCaseField(4) = cp(13)
   lblCaseField(5) = TransDate(cp(7), 1)
   lblCaseField(3) = field(8)
   txtCaseField(4) = cp(64)
   'Modify By Cheng 2002/08/19
'   If objPublicData.GetCasePreAgent(cp(), strTemp) Then
'      txtCaseField(1) = strTemp
'      CheckKeyIn 1
'   End If
   Set adoRecord = CreateObject("ADODB.Recordset")
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.SelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
   '2007/4/23 MODIFY BY SONIA 加發文日降冪排序
   'If ClsPDSelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
   'Modify by Morgan 2008/2/21 加聯絡人
   'Added by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設
   If cp(31) = "Y" Then
      AddAgent Combo1, cp, , , , cp(9), field(9), field(26)
      If Combo1 <> "" Then CheckKeyIn 1
      
   Else '非新案照原本
        If ClsPDSelectTable("select cp44||decode(cp116,null,null,'-'||cp116) from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "' and cp09<'C' and cp44 is not null order by cp27 desc", adoRecord) Then
        '2007/4/23 END
           Do While adoRecord.EOF = False
              If IsNull(adoRecord.Fields(0).Value) = False Then
                 If strSameName <> adoRecord.Fields(0).Value Then
                    Combo1.AddItem adoRecord.Fields(0).Value
                    strSameName = adoRecord.Fields(0).Value
                 End If
              End If
              adoRecord.MoveNext
           Loop
           Combo1 = Combo1.List(0)
        End If
        
      'Added by Morgan 2023/10/30 已有設定時不必再重新設定(IDS分案會先設,且抓預設代理人時也會剔除)
      If cp(44) <> "" Then
         Combo1 = cp(44) & IIf(cp(116) <> "", "-" & cp(116), "")
         CheckKeyIn 1
      Else
      'end 2023/10/30
        
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCasePreAgent(cp(), strTemp) Then
        If ClsPDGetCasePreAgent(cp(), strTemp) Then
           Combo1 = strTemp
           CheckKeyIn 1
        End If
        
      End If 'Added by Morgan 2023/10/30
        
   End If
   'end 2016/10/27
   
   txtCaseField(5) = "Y"
    'Add By Cheng 2003/09/16
    '讀取ECP指定國家
    'edit by nickc 2007/02/02 不用 dll 了
    'If field(9) = EPC指定國家 Then objPublicData.ReadCountry intCaseKind, cp(), strCountry, True, False
    If field(9) = EPC指定國家 Then ClsPDReadCountry intCaseKind, cp(), strCountry, True, False

   'Add by Morgan 2009/8/18
   If txtCaseField(0).Tag <> txtCaseField(0) Then
      PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8)
      txtCaseField(0).Tag = txtCaseField(0)
   End If
   
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
      
Else
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If

ErrHnd:
ErrorMsg
Screen.MousePointer = varSaveCursor

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo1.Text <> "" Then
      If CheckKeyIn(1) = -1 Then
         Cancel = True
      End If
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If Cancel = False Then
         strNo = Combo1.Text
         'Add by Morgan 2008/2/21 加聯絡人判斷
         iPos = InStr(Combo1.Text, "-")
         If iPos > 0 Then
            strNo = Left(Combo1.Text, iPos - 1)
         End If
         'end 2008/2/21
         
         If PUB_CheckStatus(strNo) = False Then
            Cancel = True
         'Added by Morgan 2012/3/7 發文都要顯示代理人備註--甄妮
         Else
            strExc(0) = "select FA29 from Fagent where " & ChgFagent(strNo) & " and FA29 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "" & RsTemp(0), vbExclamation, "代理人備註"
            End If
         'end 2012/3/7
         End If
      End If
      
      If Cancel Then Combo1.SetFocus
   End If
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String

Select Case Index
   Case 3
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
      If ClsPDGetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
         lblTrademarkKind = strTemp
      End If
   Case 4
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
      If ClsPDGetStaff(lblCaseField(Index), strTemp) Then
         lblSalesName = strTemp
      Else
         lblSalesName = ""
      End If
End Select
End Sub
Private Sub Form_Activate()
   If m_bActived = False Then
      m_bActived = True
      txtCaseField(0) = strSrvDate(2)
      ReadAllData
      txtCaseField(0).SetFocus
      If PUB_ChkFileNP(cp(9)) Then MsgBox "下一程序已有提申或收達期限，若為重新發文時需要先刪除後才可作業！" 'Added by Morgan 2015/8/7
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   
   'Add By Sindy 2018/1/8
   m_strIR01 = frm050102_1.m_strIR01
   m_strIR02 = frm050102_1.m_strIR02
   m_strIR03 = frm050102_1.m_strIR03
   m_strIR04 = frm050102_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/1/8 END
   txtCaseField(19) = "N" 'Added by Morgan 2018/10/22 預設不印傳真封面--慧汶
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2009/8/18
   If intLeaveKind = 1 Then
      frm050102_1.Show
   ElseIf intLeaveKind = 0 Then
     Unload frm050102_1
   End If
   ShowEditForm 'Added by Morgan 2018/8/22
    
   'Set frm050102_7 = Nothing 'Removed by Morgan 2021/12/10 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub
Private Sub txtCaseField_Change(Index As Integer)
Select Case Index
             Case 1
                         lblAgent = ""
End Select
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
      Select Case Index
            Case 3
               If KeyAscii <> 8 And KeyAscii <> 78 Then
                  KeyAscii = 0
               End If
            Case 5, 6
               If KeyAscii <> 8 And KeyAscii <> 89 Then
                  KeyAscii = 0
               End If
   End Select
End Sub
Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = -1 Then
   Cancel = True
End If
If Cancel Then txtCaseField_GotFocus (Index)
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strCusTemp As String

CheckKeyIn = -1
Select Case intIndex
             Case 0
                         If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           CheckKeyIn = 1
                           'Add by Morgan 2009/8/18
                           If txtCaseField(0).Tag <> txtCaseField(0) Then
                              PUB_SetChkResultDate field(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8)
                              txtCaseField(0).Tag = txtCaseField(0)
                           End If
                        End If
             Case 1 '代理人
                        lblAgent = ""
                        If Combo1.Text = "" Then
                           MsgBox "代理人欄不可空白!!!", vbExclamation
                        Else
                           strCusTemp = Combo1
                           'Add by Morgan 2008/2/21 加判斷是否為聯絡人
                           If InStr(strCusTemp, "-") > 0 Then
                              If ClsPDGetContact(strCusTemp, strTemp) Then
                                 Combo1 = strCusTemp
                                 lblAgent.Caption = strTemp
                                 CheckKeyIn = 1
                              End If
                           
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetAgent(strCusTemp, strTemp) Then
                           ElseIf ClsPDGetAgent(strCusTemp, strTemp) Then
                              Combo1 = strCusTemp
                              lblAgent.Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        End If
             Case 2
                        If IsNumeric(txtCaseField(intIndex)) Or txtCaseField(intIndex) = "" Then
                           CheckKeyIn = 1
                        Else
                           MsgBox "請輸入優先權證明書份數 !", vbCritical
                        End If
             Case 3
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
                        
             Case Else
                        CheckKeyIn = 1
End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'add by nickc 2008/05/01
   If IsDebt(field(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
For Each objTxt In Me.txtCaseField
   If objTxt.Enabled = True Then
      Cancel = False
      txtCaseField_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Add by Morgan 2004/9/14
If Combo1.Enabled = True Then
   Cancel = False
   Combo1_Validate Cancel
   If Cancel = True Then
      Combo1.SetFocus
      Exit Function
   End If
End If

'Added by Morgan 2018/9/12 CFP電子化-接洽單檢查
If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
   If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
      If PUB_CheckPDF3(cp(1), cp(2), cp(3), cp(4)) = False Then
         Exit Function
      End If
   End If
End If
'end 2018/9/12

'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
If Pub_ChkACS112isNull(field(1), field(2), field(3), field(4), txtCP113) = True Then
    txtCP113.SetFocus
    txtCP113_GotFocus
    Exit Function
End If
'end 2021/05/25

TxtValidate = True
End Function
'Add by Morgan 2009/8/18
Private Sub lblCaseFee_Click()
   frm12040102_2.txtCF(1) = cp(1)
   frm12040102_2.txtCF(2) = field(9)
   frm12040102_2.txtCF(3) = cp(10)
   frm12040102_2.Show vbModal
   If Val(txtCaseField(0)) > 0 Then
      PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8)
   End If
End Sub

Private Sub lblCaseFee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblCaseFee, lblCaseFees
End Sub

Private Sub lblCaseFee_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblCaseFee, lblCaseFees
End Sub

Private Sub txtChkRltDate_Validate(Cancel As Boolean)
   If txtChkRltDate <> "" Then
      If ChkDate(txtChkRltDate) = False Then
         Cancel = True
      End If
   End If
End Sub

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
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
