VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050102_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文（變更）"
   ClientHeight    =   5316
   ClientLeft      =   492
   ClientTop       =   1176
   ClientWidth     =   8568
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5316
   ScaleWidth      =   8568
   Begin VB.TextBox txtCP113 
      Height          =   300
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   6
      Top             =   3075
      Width           =   540
   End
   Begin VB.TextBox txtChkRltDate 
      Height          =   300
      Left            =   5265
      MaxLength       =   8
      TabIndex        =   7
      Top             =   3075
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   5100
      TabIndex        =   1
      Top             =   1890
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "輸入"
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   405
      Index           =   4
      Left            =   4392
      TabIndex        =   10
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6444
      TabIndex        =   12
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5616
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7668
      TabIndex        =   13
      Top             =   70
      Width           =   800
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   5
      Left            =   6210
      TabIndex        =   3
      Top             =   2280
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
      Height          =   1605
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   3630
      Width           =   8295
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14631;2831"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
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
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1920
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
      Left            =   5880
      TabIndex        =   5
      Top             =   2640
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
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
      Left            =   120
      TabIndex        =   38
      Top             =   3120
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
      Left            =   6255
      TabIndex        =   8
      Tag             =   "Y"
      Top             =   3075
      Width           =   255
   End
   Begin VB.Label lblChkRltDate 
      Caption         =   "催審期限："
      Height          =   255
      Left            =   4320
      TabIndex        =   36
      Top             =   3090
      Width           =   990
   End
   Begin VB.Label Label4 
      Caption         =   "是否修改通知函內容：            （Y:WORD）"
      Height          =   255
      Left            =   4320
      TabIndex        =   35
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "變更事項："
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   2640
      Width           =   975
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   6360
      TabIndex        =   33
      Top             =   1920
      Width           =   2055
      VariousPropertyBits=   27
      Size            =   "3625;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "進度備註："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "代理人："
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "是否列印通知函：          （N：不印）"
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   2280
      Width           =   4065
   End
   Begin VB.Label Label11 
      Caption         =   "發文日："
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label21 
      Caption         =   "是否修改指示信：            （Y:WORD）"
      Height          =   255
      Left            =   4320
      TabIndex        =   28
      Top             =   2640
      Width           =   4095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   21
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   20
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   19
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   18
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   5280
      TabIndex        =   17
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   5280
      TabIndex        =   16
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label lblSalesName 
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "法定期限："
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   27
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收文號："
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   25
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "專利種類："
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   24
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "本所期限："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "智權人員："
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   22
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label lblCaseFees 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   6300
      TabIndex        =   37
      Top             =   3090
      Width           =   255
   End
End
Attribute VB_Name = "frm050102_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/5 改成Form2.0 (txtCaseField,lblAgent)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
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
'strChange存變更事項   bolIsChange是否有輸入變更事項   intFunction 0:其他  1:內商著作權   2:專利
Dim strChange As String, bolIsChange As Boolean, intFunction As Integer
'Add By Cheng 2003/09/16
Dim strCountry As String '存放EPC指定國家
'Add By Sindy 2018/1/8
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2018/1/8 END
Dim m_strAF01 As String, m_strLD18 As String 'Added by Morgan 2018/8/22

Private Sub cmdChange_Click()
If cp(1) = 內商著作權 Then
   intFunction = 1
ElseIf intPCaseKind = 專利 Then
   intFunction = 2
Else
   intFunction = 0
End If
'Modify by Morgan 2009/7/24 +傳cp09
ModifyChange strChange, bolIsChange, intFunction, cp(9)
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String, Optional ET06 As String)
   Dim strTxt(1 To 5) As String, intStep As Integer, i As Integer
   EndLetter ET01, cp(9), ET03, strUserNum
   'Add by Morgan 2008/1/14
   If Trim(ET06) = Empty Then
      strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "SELECT '" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','變更項目',CE61 FROM CHANGEEVENT WHERE CE01='" & cp(9) & "'"
   Else
   'end 2008/1/14
      strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
         "','變更項目'," & CNULL(ET06) & ")"
   End If
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(1, strTxt) Then
   If Not ClsLawExecSQL(1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
 Dim i As Integer, strTmp As String, bolChk As Boolean
   Select Case Index
      Case 0, 4 '確定, 同時發文
      
         'Added by Morgan 2015/8/7
         If PUB_ChkFileNP(cp(9)) Then
            MsgBox "下一程序已有提申或收達期限，不可發文！"
            Exit Sub
         End If
         'end 2015/8/7
   
         'Modify By Cheng 2002/07/30
'         For i = 0 To 4
         For i = 0 To 5
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
         If i = 6 Then
            If bolIsChange Then
               If strChange = "" Then
                  MsgBox "請先輸入變更事項的資料", vbOKOnly + vbCritical, "檢核資料"
                  Exit Sub
               End If
               'Add By Cheng 2002/05/22
               '重新檢查欄位有效性
               If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
               Screen.MousePointer = vbHourglass
               If SaveDatabase Then
                  'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
                  PUB_CheckEMail cp(44), cp(116)
                  PUB_CheckEMail field(75), field(144)
                  If field(145) <> "" Then
                     PUB_CheckEMail field(75), field(145)
                  End If
                  'end 2008/2/20
                  strTmp = ""
                  '指示信
                  If txtCaseField(3) = "Y" Then
                     bolChk = True
                  Else
                     bolChk = False
                  End If
                  '變更地址 31
                  strExc(0) = "SELECT DECODE(CE23||CE24||CE25||CE26||CE27||CE28||CE29||CE30||CE31||CE32||CE33||CE34||CE35||CE36||CE37,NULL,0,1) FROM CHANGEEVENT WHERE CE01='" & cp(9) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp.Fields(0) = 1 Then
                        strTmp = "地址"
                        StartLetter "01", "31", "address"
                        NowPrint cp(9), "01", "31", bolChk, strUserNum, 0, , , , , , , , , , , , m_strAF01
                     Else
                        '變更代表人 32
                        strExc(0) = "SELECT DECODE(CE10||CE11||CE12||CE13||CE14||CE15,NULL,0,1) FROM CHANGEEVENT WHERE CE01='" & cp(9) & "'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If RsTemp.Fields(0) = 1 Then
                              strTmp = "代表人"
                              StartLetter "01", "32", "legal representative"
                              NowPrint cp(9), "01", "32", bolChk, strUserNum, 0, , , , , , , , , , , , m_strAF01
                           Else
                              '變更申請人 33
                              'Modify by Morgan 2005/11/16 加變更代理人
                              strExc(0) = "SELECT DECODE(CE04||CE05||CE06||CE07||CE08,NULL,0,1), NVL(CE55,0) FROM CHANGEEVENT WHERE CE01='" & cp(9) & "'"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 If RsTemp.Fields(0) = 1 Then
                                    strTmp = "名稱"
                                    StartLetter "01", "33", "name"
                                    NowPrint cp(9), "01", "33", bolChk, strUserNum, 0, , , , , , , , , , , , m_strAF01
                                 'Add by Morgan 2005/11/16
                                 ElseIf RsTemp.Fields(1) = 1 Then
                                    strTmp = "代理人"
                                    NowPrint cp(9), "01", "34", bolChk, strUserNum, 0, , , , , , , , , , , , m_strAF01
                                 Else
                                    strTmp = " "
                                    StartLetter "01", "30", " "
                                    NowPrint cp(9), "01", "30", bolChk, strUserNum, 0, , , , , , , , , , , , m_strAF01
                                 End If
                              End If
                           End If
                        End If
                     End If
                  Else
                     strTmp = " "
                     StartLetter "01", "30", " "
                     NowPrint cp(9), "01", "30", bolChk, strUserNum, 0, , , , , , , , , , , , m_strAF01
                  End If
                  
                  'Added by Morgan 2018/8/22 CFP電子化
                  If bolChk = True And m_strAF01 <> "" Then
                     frm1105_1.m_RecNo = m_strAF01
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
                     frm1105_1.Show
                     If txtCaseField(5).Text = "Y" Then
                        MsgBox "指示信編輯中，客戶函請至定稿維護修改！", vbExclamation
                        txtCaseField(5).Text = ""
                     End If
                  End If
                  'end 2018/8/22
                  
                  '通知函
                  If txtCaseField(2) <> "N" Then
                     StartLetter "01", "00", strTmp
                     NowPrint cp(9), "01", "00", IIf(Me.txtCaseField(5).Text = "Y", True, False), strUserNum, 0, , , , , , , , , , , , m_strLD18
                     
                     'Added by Morgan 2018/8/22 CFP電子化
                     If txtCaseField(5).Text = "Y" And m_strLD18 <> "" Then
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
                    'Add By Cheng 2003/11/27
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
                           frm050102_1.Show
                           frm050102_1.Clear
                        End If
                    Case 4:
                        '若尚有未發文資料
                        If PUB_ChkUnissueDatas(Me.lblCaseField(1).Caption) = True Then
                            ' 90.07.12 modify by louis (回發文主畫面並重新查詢)
                           'Add By Sindy 2013/5/28
                           If frm050102_1.bolIsEMPFlow = True Then
                              frm090202_4.QueryData
                           'End If
                           '2013/5/28 End
                           'Add By Sindy 2018/1/8
                           ElseIf Me.m_strIR01 <> "" Then
                              'intLeaveKind = 0
                              'Modify By Sindy 2022/5/20
                              'frm04010519.GoNext
                              Forms(0).Tmpfrm04010519.GoNext
                              Set Forms(0).Tmpfrm04010519 = Nothing
                              '2022/5/20 END
                           '2018/1/8 END
                           End If
                           frm050102_1.Show
                           frm050102_1.ReQuery
                        '若無未發文資料
                        Else
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
                              frm050102_1.Show
                              frm050102_1.Clear
                           End If
                        End If
                    End Select
                    'End
                  Unload Me

               '911202 nick
               Else
                  MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
               End If
            Else
               ShowMsg MsgText(9187)
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
    'Modify By Cheng 2003/11/27
    '本段程式往上移
'   ' 發文回前畫面時
'   Select Case Index
'      Case 0:
'         ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
'         frm050102_1.Clear
'      Case 4:
'         ' 90.07.12 modify by louis (回發文主畫面並重新查詢)
'         frm050102_1.ReQuery
'   End Select
    'End
End Sub

Private Function SaveDatabase() As Boolean
Dim strTmp As String
Dim strTxt(1 To 10) As String, iStep As Integer
'Add By Cheng 2003/03/04
Dim arrChgEvent '變更事項
Dim StrSQLa As String
'Add By Cheng 2003/03/18
Dim strReceiveNo As String
Dim i As Integer, jj As Integer
Dim strCe(99) As String, bolChk As Boolean
Dim strTmpA(1 To 5) As String
Dim intStep As Integer
Dim strLetterJudge As String, strSubject As String '指示信判發人/主旨 Added by Morgan 2018/8/22
 
'911106 nick transation
SaveDatabase = True
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   cp(27) = txtCaseField(0)
   'Modify By Cheng 2002/08/19
'   cp(44) = txtCaseField(1)
   
   'Modify by Morgan 2008/2/14
   'cp(44) = Combo1
   intI = InStr(Combo1, "-")
   If intI > 0 Then
      cp(44) = Left(Combo1, intI - 1)
      cp(116) = Mid(Combo1, intI + 1)
   Else
      cp(44) = Combo1
      cp(116) = ""
   End If
   'end 2008/2/14
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
   
'    'Add By Cheng 2003/03/04
'    '判斷是否有變更申請人
'    If strChange <> "" Then
'        arrChgEvent = Split(strChange, ",")
'        '更新進度檔進度備註
'        If field(26) <> "" And arrChgEvent(3) <> "" Then
'            cp(64) = cp(64) & field(26) & ","
'        End If
'        If field(27) <> "" And arrChgEvent(4) <> "" Then
'            cp(64) = cp(64) & field(27) & ","
'        End If
'        If field(28) <> "" And arrChgEvent(5) <> "" Then
'            cp(64) = cp(64) & field(28) & ","
'        End If
'        If field(29) <> "" And arrChgEvent(6) <> "" Then
'            cp(64) = cp(64) & field(29) & ","
'        End If
'        If field(30) <> "" And arrChgEvent(7) <> "" Then
'            cp(64) = cp(64) & field(30) & ","
'        End If
'        '更新基本檔申請人及相關資料
'        '申請人1
'        If arrChgEvent(3) <> "" Then
'            strSQLA = "Update Patent Set PA26='" & ChangeCustomerL("" & arrChgEvent(3)) & "' " & _
'                            " ,PA31='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(3)), "1")) & "' " & _
'                            " ,PA36='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(3)), "2")) & "' " & _
'                            " ,PA41='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(3)), "3")) & "' " & _
'                            " Where " & ChgPatent(field(1) & field(2) & field(3) & field(4))
'            cnnConnection.ext strSQLA
'        End If
'        '申請人2
'        If arrChgEvent(4) <> "" Then
'            strSQLA = "Update Patent Set PA27='" & ChangeCustomerL("" & arrChgEvent(4)) & "' " & _
'                            " ,PA32='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(4)), "1")) & "' " & _
'                            " ,PA37='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(4)), "2")) & "' " & _
'                            " ,PA42='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(4)), "3")) & "' " & _
'                            " Where " & ChgPatent(field(1) & field(2) & field(3) & field(4))
'            cnnConnection.ext strSQLA
'        End If
'        '申請人3
'        If arrChgEvent(5) <> "" Then
'            strSQLA = "Update Patent Set PA28='" & ChangeCustomerL("" & arrChgEvent(5)) & "' " & _
'                            " ,PA33='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(5)), "1")) & "' " & _
'                            " ,PA38='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(5)), "2")) & "' " & _
'                            " ,PA43='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(5)), "3")) & "' " & _
'                            " Where " & ChgPatent(field(1) & field(2) & field(3) & field(4))
'            cnnConnection.ext strSQLA
'        End If
'        '申請人4
'        If arrChgEvent(6) <> "" Then
'            strSQLA = "Update Patent Set PA29='" & ChangeCustomerL("" & arrChgEvent(6)) & "' " & _
'                            " ,PA34='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(6)), "1")) & "' " & _
'                            " ,PA39='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(6)), "2")) & "' " & _
'                            " ,PA44='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(6)), "3")) & "' " & _
'                            " Where " & ChgPatent(field(1) & field(2) & field(3) & field(4))
'            cnnConnection.ext strSQLA
'        End If
'        '申請人5
'        If arrChgEvent(7) <> "" Then
'            strSQLA = "Update Patent Set PA30='" & ChangeCustomerL("" & arrChgEvent(7)) & "' " & _
'                            " ,PA35='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(7)), "1")) & "' " & _
'                            " ,PA40='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(7)), "2")) & "' " & _
'                            " ,PA45='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL("" & arrChgEvent(7)), "3")) & "' " & _
'                            " Where " & ChgPatent(field(1) & field(2) & field(3) & field(4))
'            cnnConnection.ext strSQLA
'        End If
'    End If
   strTxt(1) = GetCPSQL(cp())
   '911106 nick transation
   'SaveDatabase = objLawDll.ExecSQL(1, strTxt)
   cnnConnection.Execute strTxt(1)
   
   'Add by Morgan 2005/11/24 清除舊資料
   strSql = "DELETE FROM CHANGEEVENT WHERE CE01='" & cp(9) & "'"
   cnnConnection.Execute strSql
   '2005/11/24 end
   
   strTmp = cp(9) + strChange
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.SaveChangeEventDatabase(strTmp, 國外_CF) Then
   If ClsPDSaveChangeEventDatabase(strTmp, 國外_CF) Then
      
   End If
    'Add By Cheng 2003/03/18
   strReceiveNo = cp(9)
   If cp(10) = 變更 Then
      strExc(0) = "SELECT * FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
            For i = 1 To 99
               If IsNull(.Fields(i - 1)) Then
                  strCe(i) = ""
               Else
                  strCe(i) = .Fields(i - 1)
               End If
            Next
         End With
         strExc(1) = ""
         strExc(2) = ""
         strExc(3) = ""
   
         '申請日 10
         If strCe(2) <> "" Then
             strExc(1) = strExc(1) & "申請日 : " & strCe(2) & ","
            If intCaseKind = 專利 Then
               strExc(2) = strExc(2) & "PA10=" & strCe(2) & ","
            Else
               strExc(2) = strExc(2) & "SP10=" & strCe(2) & ","
            End If
            strExc(3) = strExc(3) & "CE03='1',"
         End If
         
         '申請人 26-30
         bolChk = False
         For i = 4 To 8
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If bolChk = True Then
            strExc(1) = strExc(1) & "申請人 : "
            For i = 4 To 8
               If strCe(i) <> "" Then
                  strExc(1) = strExc(1) & strCe(i) & ","
                  'edit by nickc 2007/02/02 不用 dll 了
                  'If objPublicData.GetCustomerNameAndAddress(strCe(i), strTmpA(5), strTmpA(1), strTmpA(2), strTmpA(3)) Then
                  If ClsPDGetCustomerNameAndAddress(strCe(i), strTmpA(5), strTmpA(1), strTmpA(2), strTmpA(3)) Then
                     If intCaseKind = 專利 Then
                        strExc(2) = strExc(2) & "PA" & i + 27 & "=" & CNULL(ChgSQL(strTmpA(1))) & ",PA" & i + 32 & "=" & CNULL(ChgSQL(strTmpA(2))) & ",PA" & i + 37 & "=" & CNULL(ChgSQL(strTmpA(3))) & ","
                     End If
                  End If
               End If
               If intCaseKind = 專利 Then
                    'Modify By Cheng 2003/05/13
'                  strExc(2) = strExc(2) & "PA" & i + 22 & "=" & CNULL(strCe(i)) & ","
                  strExc(2) = strExc(2) & "PA" & i + 22 & "=" & CNULL(ChangeCustomerL(strCe(i))) & ","
               End If
            Next
            If intCaseKind <> 專利 Then
               strExc(2) = strExc(2) & "SP08=" & CNULL(strCe(4)) & "," & "SP58=" & CNULL(strCe(5)) & "," & "SP59=" & CNULL(strCe(6)) & ","
            End If
            strExc(3) = strExc(3) & "CE09='1',"
         Else
            '申請地址 31-45
            bolChk = False
            For i = 23 To 37
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
            If bolChk = True Then
               strExc(1) = strExc(1) & "申請地址 : "
               'Modify by Morgan 2011/6/10
               'For i = 23 To 37
               '   If strCe(i) <> "" Then
               '      strExc(1) = strExc(1) & strCe(i) & ","
               '   End If
               '   strExc(2) = strExc(2) & "PA" & i + 8 & "=" & CNULL(strCe(i)) & ","
               'Next
               For i = 23 + jj To 37 Step 3
                  For jj = 0 To 2
                     If strCe(i) <> "" Then
                        strExc(1) = strExc(1) & strCe(i + jj) & ","
                     End If
                     strExc(2) = strExc(2) & "PA" & (i + 8 + 5 * jj) & "=" & CNULL(strCe(i + jj)) & ","
                  Next
               Next
               'end 2011/6/10
               
               strExc(3) = strExc(3) & "CE38='1',"
            End If
         End If
         
         '專利商標種類代號 08
         If strCe(39) <> "" Then
            strExc(1) = strExc(1) & "專利商標種類代號 : " & strCe(39) & ","
            If intCaseKind = 專利 Then
               strExc(2) = strExc(2) & "PA08='" & strCe(39) & "',"
            End If
            strExc(3) = strExc(3) & "CE40='1',"
         End If
   
         '案件名稱 05-07
         bolChk = False
         For i = 41 To 43
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If bolChk = True Then
            strExc(1) = strExc(1) & "案件名稱 : "
            For i = 41 To 43
               If strCe(i) <> "" Then
                  strExc(1) = strExc(1) & strCe(i) & ","
               End If
               If intCaseKind = 專利 Then
                  strExc(2) = strExc(2) & "PA" & Format(i - 36, "00") & "=" & CNULL(strCe(i)) & ","
               Else
                  strExc(2) = strExc(2) & "SP" & Format(i - 36, "00") & "=" & CNULL(strCe(i)) & ","
               End If
            Next
            strExc(3) = strExc(3) & "CE44='1',"
         End If
   
         '代表人 79-84
         bolChk = False
         For i = 10 To 15
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If Not bolChk Then
            For i = 68 To 91
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
         End If
   
         If bolChk Then
            strExc(1) = strExc(1) & "代表人 : "
            For i = 10 To 15
               If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
               If intCaseKind = 專利 Then
                  strExc(2) = strExc(2) & "PA" & i + 69 & "=" & CNULL(strCe(i)) & ","
               End If
            Next
            For i = 68 To 91
               If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
               If intCaseKind = 專利 Then
                  strExc(2) = strExc(2) & "PA" & i + 41 & "=" & CNULL(strCe(i)) & ","
               End If
            Next
            If intCaseKind <> 專利 Then
               strExc(2) = strExc(2) & "SP42=" & CNULL(strCe(10)) & ","
            End If
            strExc(3) = strExc(3) & "CE16='1',"
         End If
         
         '代表人中譯文
         If Not bolChk Then
            bolChk = False
            For i = 63 To 64
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
            If Not bolChk Then
               For i = 92 To 99
                  If strCe(i) <> "" Then
                     bolChk = True
                     Exit For
                  End If
               Next
            End If
            If bolChk Then
               strExc(1) = strExc(1) & "代表人中譯文 : "
               If intCaseKind = 專利 Then
                  strExc(2) = strExc(2) & "PA79=" & CNULL(strCe(63)) & ",PA82=" & CNULL(strCe(64)) & "," & _
                     "PA109=" & CNULL(strCe(92)) & ",PA112=" & CNULL(strCe(93)) & ",PA115=" & CNULL(strCe(94)) & "," & _
                     "PA118=" & CNULL(strCe(95)) & ",PA121=" & CNULL(strCe(96)) & ",PA124=" & CNULL(strCe(97)) & "," & _
                     "PA127=" & CNULL(strCe(98)) & ",PA130=" & CNULL(strCe(99)) & ","
               End If
               For i = 63 To 64
                  If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
               Next
               For i = 92 To 99
                  If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
               Next
               strExc(3) = strExc(3) & "CE65='1',"
            End If
         End If
   
         If strExc(1) <> "" Then
            For i = 2 To 3
               If Right(strExc(i), 1) = "," Then strExc(i) = Left(strExc(i), Len(strExc(i)) - 1)
            Next
            intStep = intStep + 1
            strTxt(intStep) = "UPDATE CASEPROGRESS SET CP64=CP64||'" & strExc(1) & "' WHERE CP09='" & strReceiveNo & "'"
            'Add By Cheng 2002/11/05
            '91.12.26 MODIFY BY SONIA
            'cnnConnection.Execute strTxt(iStep)
            cnnConnection.Execute strTxt(intStep)
            '91.12.26 END
            intStep = intStep + 1
            If intCaseKind = 專利 Then
               strTxt(intStep) = "UPDATE PATENT SET " & strExc(2) & " WHERE " & ChgPatent(field(1) & field(2) & field(3) & field(4))
            Else
               strTxt(intStep) = "UPDATE SERVICEPRACTICE SET " & strExc(2) & " WHERE " & ChgPatent(field(1) & field(2) & field(3) & field(4))
            End If
            'Add By Cheng 2002/11/05
            '91.12.26 MODIFY BY SONIA
            'cnnConnection.Execute strTxt(iStep)
            cnnConnection.Execute strTxt(intStep)
            '91.12.26 END
            intStep = intStep + 1
'            strTxt(intStep) = "UPDATE CHANGEEVENT SET " & strExc(3) & " WHERE CE01='" & strReceiveNo & "'"
'            'Add By Cheng 2002/11/05
'            '91.12.26 MODIFY BY SONIA
'            'cnnConnection.Execute strTxt(iStep)
'            cnnConnection.Execute strTxt(intStep)
'            '91.12.26 END
'            intStep = intStep + 1
         End If
      End If
   End If
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
      'Modified by Morgan 2015/8/7 改呼叫共用
      PUB_SetArriveDate cp(9)
      'end 2015/8/7
      
      'Added by Morgan 2015/8/7
      '提申管制
      PUB_SetApplyDate cp(1), cp(2), cp(3), cp(4), cp(7), cp(9), cp(10), txtCaseField(0), field(9)
      'end 2015/8/7
   End If
   
   'Add by Sindy 2018/1/8
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050102_1"
   End If
   '2018/1/8 END
   
   'Added by Morgan 2018/8/13 變更也要掛催審期限--郭雅娟
   If txtChkRltDate <> "" Then
      PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), cp(10), cp(43), field(9)
   End If
   'end 2018/8/13
   
   'Added by Morgan 2018/8/22 CFP電子化
   If strSrvDate(1) >= CFP指示信電子化啟用日 Then
      strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), cp(10), field(9))
      strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), cp(4), cp(10), field(11), cp(45), field(9))
      PUB_AddAppForm cp(9), True, strLetterJudge, strSubject
      m_strAF01 = cp(9)
   End If
   If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
      If txtCaseField(2) <> "N" Then
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
   'Modify by Morgan 2008/2/18 加聯絡人
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
           'Modified by Lydia 2016/10/27
           'CheckKeyIn 13
           CheckKeyIn 1
        End If
        
      End If 'Added by Morgan 2023/10/30
   End If
   'end 2016/10/27
   
    'Add By Cheng 2003/09/16
    '讀取ECP指定國家
    'edit by nickc 2007/02/02 不用 dll 了
    'If field(9) = EPC指定國家 Then objPublicData.ReadCountry intCaseKind, cp(), strCountry, True, False
    If field(9) = EPC指定國家 Then ClsPDReadCountry intCaseKind, cp(), strCountry, True, False
Else
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If
txtCaseField(3) = "Y"

   'Added by Morgan 2018/8/13
   If txtCaseField(0).Tag <> txtCaseField(0) Then
      PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8), field(16)
      txtCaseField(0).Tag = txtCaseField(0)
   End If
   'end 2018/8/13
   
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
    
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
         'Add by Morgan 2008/2/18 加聯絡人判斷
         iPos = InStr(Combo1.Text, "-")
         If iPos > 0 Then
            strNo = Left(Combo1.Text, iPos - 1)
         End If
         'end 2008/2/18
         
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

Private Sub lblCaseFee_Click()
   frm12040102_2.txtCF(1) = cp(1)
   frm12040102_2.txtCF(2) = field(9)
   frm12040102_2.txtCF(3) = cp(10)
   frm12040102_2.Show vbModal
   If Val(txtCaseField(0)) > 0 Then
      PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8), field(16)
   End If
End Sub

Private Sub lblCaseFee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblCaseFee, lblCaseFees
End Sub

Private Sub lblCaseFee_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblCaseFee, lblCaseFees
End Sub

Private Sub txtChkRltDate_GotFocus()
   TextInverse txtChkRltDate
   CloseIme
End Sub

Private Sub txtChkRltDate_Validate(Cancel As Boolean)
   If txtChkRltDate <> "" Then
      If ChkDate(txtChkRltDate) = False Then
         Cancel = True
      End If
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
Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   strChange = ""
   txtCaseField(0) = strSrvDate(2)
   
   'Add By Sindy 2018/1/8
   m_strIR01 = frm050102_1.m_strIR01
   m_strIR02 = frm050102_1.m_strIR02
   m_strIR03 = frm050102_1.m_strIR03
   m_strIR04 = frm050102_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/1/8 END
   
   ReadAllData
   If PUB_ChkFileNP(cp(9)) Then MsgBox "下一程序已有提申或收達期限，若為重新發文時需要先刪除後才可作業！" 'Added by Morgan 2015/8/7
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
   
   'Set frm050102_4 = Nothing'Removed by Morgan 2021/12/10 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub
Private Sub txtCaseField_Change(Index As Integer)
Select Case Index
             Case 1
                         lblAgent = ""
End Select
End Sub
Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
             Case 1, 2, 3
                       KeyAscii = UpperCase(KeyAscii)
            'Add By Cheng 2002/07/30
            Case 5 '是否修改通知函內容
               KeyAscii = UpperCase(KeyAscii)
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
                        End If
                        'Added by Morgan 2018/8/13
                        If txtCaseField(0).Tag <> txtCaseField(0) Then
                           PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8), field(16)
                           txtCaseField(0).Tag = txtCaseField(0)
                        End If
                        'end 2018/8/13
                        
             Case 1 '代理人
                        lblAgent = ""
                        If Combo1.Text = "" Then
                           MsgBox "代理人欄不可空白!!!", vbExclamation
                        Else
                           strCusTemp = Combo1
                           'Add by Morgan 2008/2/14 加判斷是否為聯絡人
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
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 3
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "Y" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
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
