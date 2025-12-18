VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010602_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人通知修正"
   ClientHeight    =   5388
   ClientLeft      =   -1176
   ClientTop       =   1608
   ClientWidth     =   8532
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5388
   ScaleWidth      =   8532
   Begin VB.TextBox txtFiles 
      Height          =   300
      Left            =   5535
      MaxLength       =   2
      TabIndex        =   7
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7524
      TabIndex        =   17
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5472
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6300
      TabIndex        =   15
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   8
      Left            =   1020
      TabIndex        =   8
      Top             =   3540
      Width           =   975
      VariousPropertyBits=   671107099
      Size            =   "1720;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   9
      Left            =   5340
      TabIndex        =   9
      Top             =   3540
      Width           =   855
      VariousPropertyBits=   671107099
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   615
      Index           =   10
      Left            =   1020
      TabIndex        =   10
      Top             =   3840
      Width           =   7395
      VariousPropertyBits=   -1467987941
      MaxLength       =   170
      ScrollBars      =   2
      Size            =   "13044;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      Height          =   300
      Left            =   1020
      TabIndex        =   12
      Top             =   1140
      Width           =   7335
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12938;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   6
      Left            =   1380
      TabIndex        =   6
      Top             =   3240
      Width           =   615
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "1085;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   5
      Left            =   5340
      TabIndex        =   5
      Top             =   2940
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   4
      Left            =   1020
      TabIndex        =   4
      Top             =   2940
      Width           =   975
      VariousPropertyBits=   671107099
      MaxLength       =   6
      Size            =   "1720;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   3
      Left            =   5340
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   1
      Left            =   5340
      TabIndex        =   1
      Top             =   2340
      Width           =   735
      VariousPropertyBits=   671107099
      MaxLength       =   4
      Size            =   "1296;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   2
      Left            =   1020
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Top             =   2340
      Width           =   1215
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   615
      Index           =   7
      Left            =   1020
      TabIndex        =   11
      Top             =   4560
      Width           =   7395
      VariousPropertyBits=   -1467987941
      ScrollBars      =   2
      Size            =   "13044;1085"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label112 
      Caption         =   "費　　用："
      Height          =   255
      Left            =   60
      TabIndex        =   55
      Top             =   3555
      Width           =   975
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "點　　數："
      Height          =   180
      Index           =   1
      Left            =   4260
      TabIndex        =   54
      Top             =   3555
      Width           =   900
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "報價備註："
      Height          =   180
      Left            =   60
      TabIndex        =   53
      Top             =   3855
      Width           =   900
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "附件檔案數量："
      Height          =   180
      Left            =   4260
      TabIndex        =   52
      Top             =   3255
      Width           =   1260
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   8
      Left            =   5220
      TabIndex        =   27
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   1020
      TabIndex        =   32
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   900
      TabIndex        =   30
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   6
      Left            =   900
      TabIndex        =   29
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   7
      Left            =   5220
      TabIndex        =   28
      Top             =   1740
      Width           =   2175
   End
   Begin MSForms.Label lblNation 
      Height          =   255
      Left            =   5700
      TabIndex        =   46
      Top             =   2040
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   1620
      TabIndex        =   44
      Top             =   2040
      Width           =   2175
      VariousPropertyBits=   27
      Size            =   "3836;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblCaseProperty 
      Height          =   255
      Left            =   1500
      TabIndex        =   43
      Top             =   1440
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblSales 
      Height          =   255
      Left            =   5940
      TabIndex        =   42
      Top             =   1440
      Width           =   2055
      VariousPropertyBits=   27
      Size            =   "3625;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1020
      TabIndex        =   35
      Top             =   540
      Width           =   2535
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5100
      TabIndex        =   34
      Top             =   540
      Width           =   2535
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1020
      TabIndex        =   33
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   5220
      TabIndex        =   31
      Top             =   1440
      Width           =   615
   End
   Begin MSForms.Label lblPromoter 
      Height          =   255
      Left            =   1860
      TabIndex        =   26
      Top             =   1740
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCode 
      Height          =   255
      Index           =   2
      Left            =   2220
      TabIndex        =   49
      Top             =   900
      Width           =   255
   End
   Begin VB.Label lblCode 
      Height          =   255
      Index           =   3
      Left            =   2460
      TabIndex        =   48
      Top             =   900
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "是否算案件數：              （N：不算）"
      Height          =   255
      Index           =   13
      Left            =   60
      TabIndex        =   25
      Top             =   3255
      Width           =   3015
   End
   Begin VB.Label lblPromoterName 
      Height          =   255
      Left            =   2205
      TabIndex        =   24
      Top             =   2940
      Width           =   1500
   End
   Begin VB.Label Label15 
      Caption         =   "承辦期限："
      Height          =   255
      Index           =   12
      Left            =   4260
      TabIndex        =   23
      Top             =   2955
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "承辦人："
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   22
      Top             =   2955
      Width           =   975
   End
   Begin VB.Label lblNextCaseProperty 
      Height          =   255
      Left            =   6180
      TabIndex        =   21
      Top             =   2340
      Width           =   2175
   End
   Begin VB.Label Label15 
      Caption         =   "法定期限："
      Height          =   255
      Index           =   11
      Left            =   4260
      TabIndex        =   20
      Top             =   2655
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "下一程序："
      Height          =   255
      Index           =   10
      Left            =   4260
      TabIndex        =   19
      Top             =   2355
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "進度備註："
      Height          =   255
      Index           =   9
      Left            =   60
      TabIndex        =   18
      Top             =   4620
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "本所期限："
      Height          =   255
      Index           =   8
      Left            =   60
      TabIndex        =   16
      Top             =   2655
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "櫃台收文日："
      Height          =   255
      Index           =   7
      Left            =   60
      TabIndex        =   14
      Top             =   2355
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   255
      Left            =   60
      TabIndex        =   51
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "證書號數："
      Height          =   255
      Index           =   0
      Left            =   4260
      TabIndex        =   50
      Top             =   540
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "申請案號："
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   47
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label29 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   4260
      TabIndex        =   45
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "智權人員："
      Height          =   255
      Left            =   4260
      TabIndex        =   41
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblIssue 
      Caption         =   "發文日："
      Height          =   255
      Left            =   4260
      TabIndex        =   40
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   39
      Top             =   1740
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "代理人："
      Height          =   255
      Left            =   60
      TabIndex        =   38
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "案件性質："
      Height          =   255
      Left            =   60
      TabIndex        =   37
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   36
      Top             =   1140
      Width           =   975
   End
End
Attribute VB_Name = "frm02010602_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/9 改成Form2.0 (cboCaseName,txtCaseField,lblSales...)
'Memo By Morgan 2012/12/17 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
'Add by Morgan 2004/1/12
Dim m_blnCustReturnSheet As Boolean '判斷是否列印案件回覆單
'Add by Morgan 2004/2/18
'若承辦人是王協理且未發文則要發EMail通知
Dim stCP09 As String, stCP14 As String, stCP27 As String
Dim m_CP14ST06 As String '2010/1/20 add by sonia 承辦人所別
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim m_PrevForm As Form 'Add By Sindy 2016/10/11
Dim m_bolAddLP As Boolean, m_strCP10 As String, m_strLD18 As String 'Added by Morgan 2018/7/18 CFP電子化
Dim m_bolActive As Boolean 'Added by Morgan 2023/7/3

'Add By Sindy 2016/10/11
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim varSaveCursor  As Variant, i As Integer

Select Case Index
             Case 0
                        varSaveCursor = Screen.MousePointer
                        Screen.MousePointer = vbHourglass
                        For i = 0 To 7
                           If i = 1 Then
                           Else
                               If txtCaseField(i).Enabled Then
                                  If CheckKeyIn(i) <> 1 Then
                                     txtCaseField(i).SetFocus
                                     txtCaseField_GotFocus (i)
                                     Exit For
                                  End If
                               End If
                           End If
                        Next
                        If i = 8 Then
                           'Add By Cheng 2002/05/22
                           '重新檢查欄位有效性
                           If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
                           
'                           'Add by Sindy 2017/10/23 已收達信件要歸卷
'                           If m_strIR01 <> "" And frm02010601_1.txtChoose.Text = "1" Then
'                              '下載信件檔,上傳卷宗區
'                              If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, cp(9), "ACK") = False Then
'                                 Screen.MousePointer = vbDefault
'                                 Exit Sub
'                              End If
'                           End If
'                           '2017/10/23 END
                           
                           'Add By Sindy 2022/7/1
                           If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
                              If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
                                 Screen.MousePointer = vbDefault
                                 Exit Sub
                              End If
                           End If
                           '2022/7/1 END
                           
                           If SaveDatabase Then
                              'Add by Morgan 2004/2/18
                              '若承辦人是王協理且未發文則要發EMail通知
                              'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
                              If stCP14 = "99050" And stCP27 = "" Then
                                  Call PUB_SendMail(strUserNum, "99050", stCP09, "分案通知")
                              End If
                              
                              'Add By Sindy 2016/10/7
                              If Me.m_strIR01 <> "" Then
                                 bolLeave = True
                                 intLeaveKind = 0
                                 Unload frm02010602_1
                                 'Add By Sindy 2016/10/11
                                 If Not m_PrevForm Is Nothing Then
                                    Call m_PrevForm.GoNext
                                 End If
                                 '2016/10/11 END
                                 Unload Me
                              Else
                              '2016/10/7 END
                                 bolLeave = True
                                 intLeaveKind = 1
                                 frm02010602_1.Clear
                                 Unload Me
                              End If
                           'Add By Cheng 2002/11/06
                           Else
                                MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
                           End If
                        End If
                        Screen.MousePointer = Default
             Case 1, 2
                        If Index = 2 Then
                           intLeaveKind = 0
                        Else
                           intLeaveKind = 1
                        End If
                        Unload Me
End Select
End Sub

Private Function SaveDatabase() As Boolean
 Dim strTxt(1 To 30) As String, iStep As Integer
 Dim lMax As Long
 Dim bolNP22 As Boolean, NP22(1 To 3) As String, iNP22 As Integer
 Dim stCP12 As String, stCP13 As String 'Add by Morgan 2004/2/6
 Dim bolSavPdf As Boolean 'Added by Morgan 2018/10/2

'Add By Cheng 2002/11/05
On Error GoTo ErrorHandler
SaveDatabase = True
cnnConnection.BeginTrans

   iStep = 1
   If cp(46) = "" And txtCaseField(0) <> "" Then
      strTxt(iStep) = "UPDATE CASEPROGRESS SET CP46=" & CNULL(TransDate(txtCaseField(0), 2)) & " WHERE CP09='" & cp(9) & "'"
        'Add By Cheng 2002/11/05
      cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
   End If
   
   '2011/5/12 add by sonia 更新下一程序收達期限
   strTxt(iStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & cp(9) & "' AND NP06 IS NULL AND NP07='997' "
   cnnConnection.Execute strTxt(iStep)
   iStep = iStep + 1
   '2011/5/12 END

   'edit by nickc 2007/02/02
   'Dim strDataTemp(1 To T_CP) As String
   Dim strDataTemp() As String
   ReDim strDataTemp(1 To TF_CP) As String
   
   strDataTemp(1) = cp(1)
   strDataTemp(2) = cp(2)
   strDataTemp(3) = cp(3)
   strDataTemp(4) = cp(4)
   strDataTemp(5) = strSrvDate(1)
   strDataTemp(6) = txtCaseField(2)
   strDataTemp(7) = txtCaseField(3)
   strDataTemp(9) = 主管機關來函
   '2010/1/19 MODIFY BY SONIA
   'strDataTemp(10) = 通知修正
   strDataTemp(10) = "1224"
   '2010/1/19 END
   '2009/12/30 MODIFY BY SONIA
   'strDataTemp(13) = cp(13)
   'strDataTemp(12) = cp(12)
   strDataTemp(13) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   strDataTemp(12) = GetSalesArea(strDataTemp(13))
   '2009/12/30 END
   strDataTemp(14) = txtCaseField(4)
   strDataTemp(26) = txtCaseField(6)
   strDataTemp(48) = txtCaseField(5)
   strDataTemp(43) = cp(9)
   strDataTemp(20) = "N"
   strDataTemp(32) = "N"
   '2018/8/22 add by sonia 加報價備註欄
   If Val(txtCaseField(8)) > 0 Then strDataTemp(144) = "費用" & Format(txtCaseField(8), "#,###") & "(" & txtCaseField(9) & "P);" & txtCaseField(10)

   '2008/8/26 modify by sonia 櫃台收文日改存 cp119
   '2008/10/24 MODIFY BY SONIA CP64仍存
   If txtCaseField(7) = "" Then
      strDataTemp(64) = "櫃台收文日：" & txtCaseField(0)
   Else
      strDataTemp(64) = "櫃台收文日：" & txtCaseField(0) & "，" & txtCaseField(7)
   End If
   strDataTemp(119) = ChangeTStringToWString(Me.txtCaseField(0).Text)
   '2008/8/26 end
   
   strTxt(iStep) = GetCPSQL(strDataTemp(), False)
   'Add By Cheng 2002/11/05
   cnnConnection.Execute strTxt(iStep)
   '2010/1/20 add by sonia 承辦人為分所人員以系統日的下一個工作天上齊備日
   If m_CP14ST06 <> "1" And strDataTemp(27) = "" Then
      strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & CompWorkDay(2, strSrvDate(1), 0) & " WHERE EP02='" & strDataTemp(9) & "'"
      cnnConnection.Execute strSql
   End If
   '2010/1/20 end
    
    'Add by Morgan 2004/2/18
    '若承辦人是王協理且未發文則要發EMail通知
    stCP09 = strDataTemp(9): stCP14 = strDataTemp(14): stCP27 = ""
    
'2009/12/30 CANCEL BY SONIA
'   iStep = iStep + 1
'    'Add By Cheng 2003/04/03
'    '業務員存最近收文A類接洽記錄單的業務員
'    'Modify by Morgan 2004/2/6
'    'strTxt(iStep) = "Update Caseprogress Set CP13='" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "' Where CP09='" & strDataTemp(9) & "' "
'    stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
'    stCP12 = GetSalesArea(stCP13)
'    strTxt(iStep) = "Update Caseprogress Set CP12='" & stCP12 & "',CP13='" & stCP13 & "' Where CP09='" & strDataTemp(9) & "' "
'    'Modify end 2004/2/6
'   cnnConnection.Execute strTxt(iStep)
'2009/12/30 END
    
   iStep = iStep + 1
   
   lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
   bolNP22 = False
   iNP22 = 1
   
   If txtCaseField(1) <> "" Then
      If txtCaseField(1) = 催審 Or txtCaseField(1) = 提申 Or txtCaseField(1) = 收達 Or txtCaseField(1) = 補文件 Then
        'Modify By Cheng 2003/04/03
        '業務員存最近收文A類接洽記錄單的業務員
         strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np15,np22) values (" + _
            CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
            "," + CNULL(strDataTemp(4)) + "," + CNULL(txtCaseField(1)) + "," & CNULL(TransDate(txtCaseField(2), 2)) & "," & _
            CNULL(TransDate(txtCaseField(3), 2)) & _
            "," + CNULL(PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))) + "," + CNULL(txtCaseField(7)) + "," & lMax & ")"
      Else
        'Modify By Cheng 2003/04/03
        '業務員存最近收文A類接洽記錄單的業務員
         strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np15,np22) values (" + _
            CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
            "," + CNULL(strDataTemp(4)) + "," + CNULL(txtCaseField(1)) + "," & CNULL(TransDate(txtCaseField(2), 2)) & "," & _
            CNULL(TransDate(txtCaseField(3), 2)) & _
            "," + CNULL(PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))) + "," + CNULL(txtCaseField(7)) + "," & lMax & ")"
      End If
      bolNP22 = True
      NP22(iNP22) = lMax
      iNP22 = iNP22 + 1
        'Add By Cheng 2002/11/05
        cnnConnection.Execute strTxt(iStep)
'      lMax = lMax + 1
        lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax

      iStep = iStep + 1
   End If
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", strDataTemp(9), "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010602_1", IIf(Pub_StrUserSt03 = "F22", strDataTemp(9), "")
   End If
   '2016/10/7 END
   
   'Added by Morgan 2018/7/18 CFP電子化
   If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
      'Added by Morgan 2018/10/2 新增B類回覆代理人902 --郭
      'Modified by Morgan 2018/10/12 改收文回覆委任代理人936 --慧汶
      strExc(1) = AutoNo("B", 6)
      'Modified by Morgan 2024/3/8 +CP44,CP45 先設定相關收文號的CF代理人及彼號，IDS的回代發文時才能正確預設(非IDS也能適用)
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
         "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP44,CP45,cp48) VALUES " & _
         "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & "," & CNULL(DBDATE(txtCaseField(2)), True) & "," & CNULL(DBDATE(txtCaseField(3)), True) & _
         ",'" & strExc(1) & "','936','90','" & strDataTemp(12) & "','" & strDataTemp(13) & "'" & _
         ",'" & txtCaseField(4) & "','N','N','N','" & strDataTemp(9) & "','" & ChangeCustomerL(cp(44)) & "','" & cp(45) & "'," & CNULL(DBDATE(txtCaseField(5)), True) & ") "
      cnnConnection.Execute strSql, intI
      
      strSql = "update nextprogress set np06='Y' where np01='" & strDataTemp(9) & "' and np07='936' and np06 is null"
      cnnConnection.Execute strSql, intI
      'end 2018/10/2
   
      m_strLD18 = strDataTemp(9)
      m_strCP10 = strDataTemp(10)
      strExc(1) = PUB_GetLetterJudgeNew("1", field(1), m_strCP10, field(9), cp(10))
      PUB_AddLetterProgress m_strLD18, 1 + Val(txtFiles), False, strExc(1), True, field(26), m_strCP10, field(75)
      m_bolAddLP = True
      
      'Added by Morgan 2018/10/2
      If txtCaseField(4) <> "" And Left(cp(12), 1) <> "F" Then
         Pub_COrderInform strDataTemp(9)
         bolSavPdf = True
      End If
      'end 2018/10/2
   End If
   'end 2018/7/18
   
    'Modify By Cheng 2002/11/05
'   SaveDatabase = objLawDll.ExecSQL(iStep - 1, strTxt())
    cnnConnection.CommitTrans
    
   Dim i As Integer
   'Add by Morgan 2004/1/12
   '先預設不列印案件回覆單
   m_blnCustReturnSheet = False
   If SaveDatabase And bolNP22 Then
      'Removed by Morgan 2018/10/2 99/5/1 起就已不列印
      'For i = 1 To iNP22 - 1
      '   g_PrtForm001.PrintForm NP22(i), cp(1), cp(2), cp(3), cp(4)
      'Next
      'end 2018/10/2
      
      'Add by Morgan 2004/1/12
      '設定要列印案件回覆單
      m_blnCustReturnSheet = True
   End If
   
   'Add by Morgan 2004/1/12
   '列印案件回覆單
   If m_blnCustReturnSheet = True Then
      'Modified by Morgan 2018/7/18 CFP電子化, 配合轉pdf到卷宗區(LP41='Y')改共用一般來函的定稿CFP-03-000-12(內容相同)
      'StartLetter "07", stCP09, "00"
      'NowPrint stCP09, "07", "00", False, strUserNum
      StartLetter "03", stCP09, "12"
      NowPrint stCP09, "03", "12", False, strUserNum, , , , , , , , , , , , , m_strLD18
      'end 2018/7/18
   End If
   
   '列印C類接洽記錄單 92.1.28 ADD BY SONIA
   g_PrtForm001.PrintCFForm strDataTemp(9), , bolSavPdf

'Add By Cheng 2002/11/05
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    SaveDatabase = False
End Function

Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String, strTemp1 As String, j As Integer

On Error GoTo ErrHnd
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.ReadAllData(frm02010602_1.grdDataList.TextMatrix(frm02010602_1.grdDataList.Row, 0), cp(), field(), intCaseKind, intPWhere) Then
   ReDim cp(TF_CP) As String
   cp(9) = frm02010602_1.grdDataList.TextMatrix(frm02010602_1.grdDataList.row, 0)
   If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
   
   If cp(1) = 馬德里案 Then
      lblCaseField(0) = cp(1) + " - " + Left(cp(2), 5) + _
         IIf(Right(cp(2), 1) = "0", "", " - " + Right(cp(2), 1)) + _
         IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
         IIf(cp(4) = "00", "", " - " + cp(4))
   Else
      lblCaseField(0) = cp(1) + " - " + cp(2) + _
         IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
         IIf(cp(4) = "00", "", " - " + cp(4))
   End If
   Select Case intCaseKind
                Case 專利
                           lblCaseField(1) = field(22)
                           lblCaseField(2) = field(11)
                           lblCaseField(8) = field(9)
                Case 商標
                           lblCaseField(1) = field(15)
                           lblCaseField(2) = field(12)
                           lblCaseField(8) = field(10)
                Case Else
                           lblCaseField(1) = field(14)
                           lblCaseField(2) = field(11)
                           lblCaseField(8) = field(9)
   End Select
   lblCaseField(3) = cp(10)
   lblCaseField(4) = cp(13)
   lblCaseField(5) = cp(14)
   lblCaseField(6) = cp(44)
'   If intPWhere <> 國外_CF Then
'      lblCaseField(7) = ChangeTStringToTDateString(cp(27))
'   Else
'      lblCaseField(7) = ChangeWStringToWDateString(cp(27))
'   End If
   lblCaseField(7) = ChangeTStringToTDateString(TransDate(cp(27), 1))
   SetNameToCombo cboCaseName, field(5), field(6), field(7)
   txtCaseField(4) = cp(14)
   'add by sonia 2024/7/15 A7010柯昱安調離改為李柏翰經理99050
   If GetStaffDepartment(txtCaseField(4)) >= "P10" And GetStaffDepartment(txtCaseField(4)) <= "P11" Then
   Else
      txtCaseField(4) = "99050"
   End If
   'end 2024/7/15
   'Modified by Morgan 2023/7/4 承辦人離職彈訊息後會當,先改不彈 CFP-31852
   'CheckKeyIn (4)
   lblPromoterName = GetStaffName(txtCaseField(4), True)
   m_CP14ST06 = PUB_GetST06(txtCaseField(4))
   'end 2023/7/4
   '2010/1/20 modify by sonia 承辦人為北所人員以系統日計算承辦期限,分所人員以系統日的下一個工作天計算
   'txtCaseField(5) = TransDate(Pub_GetHandleDay(cp(1), lblCaseField(8), 通知修正), 1)
   If m_CP14ST06 <> "1" Then
      txtCaseField(5) = TransDate(Pub_GetHandleDay(cp(1), lblCaseField(8), "1224", CompWorkDay(2, strSrvDate(1), 0), IIf(txtCaseField(2) = "", "", TransDate(txtCaseField(2), 2))), 1)
   Else
      txtCaseField(5) = TransDate(Pub_GetHandleDay(cp(1), lblCaseField(8), "1224", , IIf(txtCaseField(2) = "", "", TransDate(txtCaseField(2), 2))), 1)
   End If
   
   '2010/1/20 end
   'Add by Amy 2014/09/18 承辦人期限隱藏
    Label15(12).Visible = False
    txtCaseField(5).Visible = False
    'end 2014/09/18
   
   'Added by Morgan 2018/10/8
   'Modified by Morgan 2018/10/12 改收文回覆委任代理人936 --慧汶
   'txtCaseField(1) = "902"
   txtCaseField(1) = "936"
   txtCaseField(1).Enabled = False
   'end 2018/10/8
Else
err1:
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If
Screen.MousePointer = vbDefault
Exit Sub
ErrHnd:
Screen.MousePointer = vbDefault
ErrorMsg
End Sub
Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String, strCusTemp As String, bolIsChina As Boolean

Select Case Index
             Case 3
                       If lblCaseField(8) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                       'edit by nickc 2007/02/02 不用 dll 了
                       'If objPublicData.GetCaseProperty(cp(1), lblCaseField(Index), strTemp, bolIsChina) Then
                       If ClsPDGetCaseProperty(cp(1), lblCaseField(Index), strTemp, bolIsChina) Then
                           lblCaseProperty.Caption = strTemp
                       End If
             Case 4
                        ' 90.08.09 modify by louis
                        'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
                        '   lblSales = strTemp
                        'Else
                        '   lblSales = ""
                        'End If
                        lblSales = GetStaffName(lblCaseField(Index), True)
             Case 5
                        ' 90.08.09 modify by louis
                        'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
                        '   lblPromoter = strTemp
                        'Else
                        '   lblPromoter = ""
                        'End If
                        lblPromoter = GetStaffName(lblCaseField(Index), True)
             Case 6
                        strCusTemp = lblCaseField(Index)
                        If lblCaseField(Index) <> "" Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetAgent(strCusTemp, strTemp) Then
                           If ClsPDGetAgent(strCusTemp, strTemp) Then
                              lblCaseField(Index) = strCusTemp
                              lblAgent.Caption = strTemp
                           End If
                        End If
             Case 8
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetNation(lblCaseField(Index), strTemp) Then
                        If ClsPDGetNation(lblCaseField(Index), strTemp) Then
                           lblNation.Caption = strTemp
                        End If
End Select
End Sub
Private Sub Form_Activate()
   'Added by Morgan 2023/7/3
   If m_bolActive = True Then Exit Sub
   m_bolActive = True
   'end 2023/7/3
   ReadAllData
End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   txtCaseField_Change 1
   'Add By Sindy 2017/12/28
   m_strIR01 = frm02010602_1.m_strIR01
   m_strIR02 = frm02010602_1.m_strIR02
   m_strIR03 = frm02010602_1.m_strIR03
   m_strIR04 = frm02010602_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/28 END
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2018/10/2
   If m_strIR01 <> "" Then intLeaveKind = 0 'Add By Sindy 2017/10/17
'   'Add By Sindy 2016/10/13
'   If Me.m_strIR01 <> "" Then intLeaveKind = 0
'   '2016/10/13 END
      If intLeaveKind = 1 Then
         frm02010602_1.Show
      ElseIf intLeaveKind = 0 Then
        Unload frm02010602_1
      End If
'   End If
   
   'Add By Sindy 2016/10/11
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   'Add By Cheng 2002/07/18
   Set frm02010602_2 = Nothing
End Sub
Private Sub txtCaseField_Change(Index As Integer)
Select Case Index
             Case 1
                        lblNextCaseProperty = ""
                        If txtCaseField(Index) = "" Then
                           txtCaseField(2).Enabled = False
                           txtCaseField(3).Enabled = False
                        Else
                           txtCaseField(2).Enabled = True
                           txtCaseField(3).Enabled = True
                        End If
             Case 4
                        lblPromoterName = ""
End Select
End Sub
Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
             Case 4, 6
                       KeyAscii = UpperCase(KeyAscii)
End Select
End Sub
Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If CheckKeyIn(Index) = -1 Then
   Cancel = True
End If
If Cancel Then txtCaseField_GotFocus (Index)
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strCusTemp As String, bolIsChina As Boolean, strTemp1 As String

CheckKeyIn = -1
Select Case intIndex
             Case 0
                     If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                        If Val(txtCaseField(intIndex)) > Val(strSrvDate(2)) Then
                           MsgBox "不可大於系統日 !", vbCritical
                        Else
                           CheckKeyIn = 1
                        End If
                     End If
             Case 1 '下一程序
                       If txtCaseField(intIndex) = "" Then
                          txtCaseField(2) = ""
                          txtCaseField(3) = ""
                          CheckKeyIn = 1
                       Else
                        
                        'Add By Cheng 2002/01/04
                        If Len(Me.txtCaseField(intIndex)) <> 3 Then
                           MsgBox "下一程序欄位值必須為三碼 !", vbExclamation
                           Exit Function
                        End If
                       
                          If lblCaseField(8) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                          'edit by nickc 2007/02/02 不用 dll 了
                          'If objPublicData.GetCaseProperty(cp(1), txtCaseField(intIndex), strTemp, bolIsChina) Then
                          If ClsPDGetCaseProperty(cp(1), txtCaseField(intIndex), strTemp, bolIsChina) Then
                              lblNextCaseProperty = strTemp
                              CheckKeyIn = 1
                           End If
                        End If
             Case 2 '本所期限
                     If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                        If Val(txtCaseField(intIndex)) < Val(strSrvDate(2)) Then
                           MsgBox "本所期限不可小於系統日期!!!", vbCritical
                        Else
                            'Add By Cheng 2003/12/08
                            '若本所期限非工作天則直接調整至最近的工作天
                            Me.txtCaseField(intIndex).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(intIndex).Text, True), 1)
                           CheckKeyIn = 1
                        End If
                     End If
             Case 3 '法定期限
                        'Added by Morgan 2018/10/22 改可以不輸入法定期限--慧汶
                        If txtCaseField(intIndex).Text = "" Then
                           CheckKeyIn = 1
                        Else
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              'Modify by Morgan 2010/8/18 百年蟲
                              'If txtCaseField(2) <= txtCaseField(3) Then
                              If Val(txtCaseField(2)) <= Val(txtCaseField(3)) Then
                                 CheckKeyIn = 1
                              Else
                                 ShowMsg MsgText(1033)
                              End If
                           End If
                        End If
             Case 4
                        m_CP14ST06 = "1" '2010/1/20 add by sonia
                        If txtCaseField(intIndex) <> "" Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetStaff(txtCaseField(intIndex).Text, strTemp) Then
                           If ClsPDGetStaff(txtCaseField(intIndex).Text, strTemp) Then
                              lblPromoterName = strTemp
                              CheckKeyIn = 1
                              m_CP14ST06 = PUB_GetST06(txtCaseField(intIndex))  '2010/1/20 add by sonia
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
                        '2010/1/20 add by sonia 重新依承辦人所別以系統日或下一個工作天計算承辦期限
                        If m_CP14ST06 <> "1" Then
                           txtCaseField(5) = TransDate(Pub_GetHandleDay(cp(1), lblCaseField(8), "1224", CompWorkDay(2, strSrvDate(1), 0), IIf(txtCaseField(2) = "", "", TransDate(txtCaseField(2), 2))), 1)
                        Else
                           txtCaseField(5) = TransDate(Pub_GetHandleDay(cp(1), lblCaseField(8), "1224", , IIf(txtCaseField(2) = "", "", TransDate(txtCaseField(2), 2))), 1)
                        End If
                       '2010/1/20 end
             Case 5 '承辦期限
                        '2010/1/20 modify by sonia
                        'If txtCaseField(intIndex).Text <> "" Then
                        If Len(Me.txtCaseField(2).Text) > 0 And Len(Me.txtCaseField(5).Text) > 0 Then
                           If Val(txtCaseField(intIndex).Text) > Val(txtCaseField(2).Text) Then
                              MsgBox "輸入日期大於本所期限 !", vbCritical
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        'Add By Cheng 2002/05/06
                        '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
                        If Len(Me.txtCaseField(2).Text) > 0 And Len(Me.txtCaseField(5).Text) > 0 Then
                           If Val(Me.txtCaseField(2).Text) < Val(Me.txtCaseField(5).Text) Then
                              MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
                              CheckKeyIn = -1
                              Exit Function
                           End If
                        End If
                        
                        If txtCaseField(intIndex) = "" Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetCaseWorkDays(cp(1), lblCaseField(8), txtCaseField(1), strTemp) Then
                           'Modify by Morgan 2007/10/16 工作天函數統一
                           'If ClsPDGetCaseWorkDays(cp(1), lblCaseField(8), txtCaseField(1), strTemp) Then
                           '2010/1/20 MODIFY BY SONIA 應以來函性質判斷有無工作天
                           'strTemp = GetWorkDays(cp(1), lblCaseField(8), txtCaseField(1))
                           strTemp = GetWorkDays(cp(1), lblCaseField(8), "1224")
                           '2010/1/20 END
                           'end 2007/10/16
                              If strTemp <> "" And txtCaseField(1) <> "" Then
                                 ShowMsg MsgText(1049)
                              Else
                                 CheckKeyIn = 1
                              End If
                           'End If
                        Else
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                               CheckKeyIn = 1
                           End If
                        End If
             Case 6
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
   TextInverse txtCaseField(Index)
   txtCaseField(Index).SetFocus 'Added by Morgan 2021/12/9
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False


   'Added by Morgan 2021/12/9 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/9
   
For Each objTxt In Me.txtCaseField
   If objTxt.Enabled = True Then
      Cancel = False
      txtCaseField_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Added by Morgan 2018/7/18 CFP電子化
If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
   If txtFiles = "" Then
      MsgBox "請輸入附件案檔案數量!!", vbExclamation
      txtFiles.SetFocus
      Exit Function
   End If
End If
'end 2018/7/18

TxtValidate = True
End Function

'Add by Morgan 2004/1/12
'儲存定稿例外欄位
Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
   
   Dim strTxt(1 To 2) As String, Jjj As Integer
   
   EndLetter ET01, ET02, ET03, strUserNum
   
   Jjj = 0
   If txtCaseField(1).Text <> "" Then
      Jjj = Jjj + 1
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
          "','下一程序','" & txtCaseField(1).Text & "')"
   End If
   If lblNextCaseProperty.Caption <> "" Then
      Jjj = Jjj + 1
      strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
          "','下一程序名稱','" & lblNextCaseProperty.Caption & "')"
   End If
   
   If Jjj > 0 Then
      'edit by nickc 2007/02/05 不用 dll 了
      'If Not objLawDll.ExecSQL(Jjj, strTxt) Then
      If Not ClsLawExecSQL(Jjj, strTxt) Then
         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
      End If
   End If
End Sub

