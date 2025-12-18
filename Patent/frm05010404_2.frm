VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm05010404_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "消滅函輸入"
   ClientHeight    =   4520
   ClientLeft      =   110
   ClientTop       =   820
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4520
   ScaleWidth      =   9000
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   26
      Top             =   463
      Width           =   855
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2475
      MaxLength       =   1
      TabIndex        =   25
      Top             =   463
      Width           =   330
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2790
      MaxLength       =   2
      TabIndex        =   24
      Top             =   463
      Width           =   435
   End
   Begin VB.TextBox txtSystem 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   23
      Text            =   "CFP"
      Top             =   463
      Width           =   465
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   7335
      TabIndex        =   8
      Top             =   30
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   6705
      TabIndex        =   7
      Top             =   30
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   8250
      TabIndex        =   9
      Top             =   30
      Width           =   600
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   6
      Left            =   6585
      TabIndex        =   4
      Top             =   2640
      Width           =   1065
      VariousPropertyBits=   671107097
      MaxLength       =   8
      Size            =   "1879;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   990
      Index           =   4
      Left            =   1320
      TabIndex        =   5
      Top             =   3015
      Width           =   7140
      VariousPropertyBits=   -1467987941
      ScrollBars      =   2
      Size            =   "12594;1746"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Width           =   300
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "529;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   5235
      TabIndex        =   1
      Top             =   1890
      Width           =   300
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "529;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1335
      TabIndex        =   2
      Top             =   2250
      Width           =   300
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "529;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   1305
      TabIndex        =   3
      Top             =   2640
      Width           =   1065
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1879;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   5
      Left            =   1320
      TabIndex        =   6
      Top             =   4110
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1140
      TabIndex        =   10
      Top             =   840
      Width           =   7395
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "13044;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      Caption         =   "專利權終止通知書發出日："
      Height          =   255
      Left            =   4380
      TabIndex        =   38
      Top             =   2640
      Width           =   2160
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   5445
      TabIndex        =   36
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "(Y:閉卷)"
      Height          =   255
      Left            =   5895
      TabIndex        =   35
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label39 
      Caption         =   "是否閉卷："
      Height          =   255
      Index           =   1
      Left            =   4380
      TabIndex        =   34
      Top             =   1560
      Width           =   990
   End
   Begin VB.Label Label36 
      Caption         =   "是否列印客戶通知函：          (N:不印)"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   33
      Top             =   1920
      Width           =   3030
   End
   Begin VB.Label Label36 
      Caption         =   "是否修改通知函內容：        (Y:是)"
      Height          =   255
      Index           =   2
      Left            =   3420
      TabIndex        =   32
      Top             =   1920
      Width           =   2625
   End
   Begin VB.Label Label43 
      Caption         =   "進度備註:"
      Height          =   255
      Left            =   180
      TabIndex        =   31
      Top             =   3030
      Width           =   855
   End
   Begin VB.Label Label38 
      Caption         =   "消滅日："
      Height          =   255
      Left            =   180
      TabIndex        =   29
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Label Label39 
      Caption         =   "是否閉卷:"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   28
      Top             =   4110
      Width           =   855
   End
   Begin VB.Label Label40 
      Caption         =   "(Y:閉卷)"
      Height          =   255
      Left            =   1860
      TabIndex        =   27
      Top             =   4110
      Width           =   855
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5430
      TabIndex        =   22
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label8 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   4380
      TabIndex        =   21
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "櫃台收文日："
      Height          =   255
      Left            =   180
      TabIndex        =   20
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   19
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   1380
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1230
      TabIndex        =   17
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label22 
      Caption         =   "專利種類："
      Height          =   255
      Left            =   180
      TabIndex        =   16
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   1710
      TabIndex        =   14
      Top             =   1200
      Width           =   2565
   End
   Begin MSForms.Label lblNation 
      Height          =   255
      Left            =   6135
      TabIndex        =   13
      Top             =   1200
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   4380
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5430
      TabIndex        =   11
      Top             =   1200
      Width           =   645
   End
   Begin VB.Label lblCaseField 
      Caption         =   "年費"
      Height          =   255
      Index           =   6
      Left            =   4005
      TabIndex        =   37
      Top             =   2280
      Width           =   645
   End
   Begin VB.Label Label36 
      Caption         =   "消滅函類別：         (1:逾期未辦 2:未領證 3:未續繳年費       4:屆滿 5.未繳實審)"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   30
      Top             =   2280
      Width           =   7485
   End
End
Attribute VB_Name = "frm05010404_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/8 改成Form2.0 (cboCaseName,Text1,lblNation)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Morgan 2008/12/10
Option Explicit

Public frmParent As Form
Dim bolActived As Boolean
Dim pa() As String
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'intLeaveKind判斷離開時，是2:結束  1:回上一畫面
Dim intLeaveKind As Integer
Dim m_NewCP09 As String, m_NP07 As String, m_NP07Desc As String
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END

'Added by Morgan 2018/4/18
Dim m_HKNewCP09 As String, m_HKCloser As String
Dim m_strHK(1 To 4) As String '香港案
'end 2018/4/18
Dim m_bolAddLP As Boolean, m_strCP10 As String, m_strLD18 As String 'Added by Morgan 2018/7/18 CFP電子化

Private Sub cmdok_Click(Index As Integer)
   Dim bEdit As Boolean
   Dim strTmp As String
   
   intLeaveKind = Index
   bolLeave = False
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         If TxtValidate = True Then
            If fnChoiceCheck = True Then
               'If Text1(2) = "4" Then     'cancel by sonia 2025/4/23 郭因CFP-032291閉卷後又收文分析，但沒有後續，且消滅函並非屆滿故改為不分類別都詢問
                  If CheckCP109 = True Then
                     '詢問是否可結餘
                     '2011/11/8 modify by sonia TF子案不可結餘故加傳本所案號
                     'Pub_EndModCashMsg pa(9)
                     Pub_EndModCashMsg pa(9), pa(1), pa(2), pa(3), pa(4)
                  End If
               'End If                     'cancel by sonia 2025/4/23 郭因CFP-032291閉卷後又收文分析，但沒有後續，且消滅函並非屆滿故改為不分類別都詢問
               If SaveData = False Then
                  MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
               Else
                  If Text1(0).Text <> "N" Then '通知函
                     If Text1(1).Text = "Y" Then
                        bEdit = True
                     Else
                        bEdit = False
                     End If
                     
                     Select Case Text1(2)
                        Case "1"               '逾期未辦 01
                           strTmp = "01"
                           '2011/11/14 modify by sonia改用列印備註
                           ''2011/9/21 ADD BY SONIA 美國先改,其他待慧汶整理
                           'If pa(9) = "101" Then strTmp = "06"
                           StartLetter "03", strTmp
                        Case "2"               '未領證 02
                           strTmp = "02"
                        Case "3"               '未續繳年(延展,維持)費 03
                           strTmp = "03"
                           'Added by Morgan 2021/10/15 寶齡富錦 Y55435 案件
                            If ChangeCustomerS(pa(75)) = "Y55435" Then
                               strTmp = "99"
                            End If
                            'end 2021/10/15
                        Case "4"               '屆滿04
                           strTmp = "04"
                        Case "5"               '未繳實審 05
                           strTmp = "05"
                     End Select
                     If strTmp <> "" Then
                        StartLetter "03", strTmp
                        NowPrint m_NewCP09, "03", strTmp, bEdit, strUserNum, , , , , , , , , , , , , m_strLD18
                        'Added by Morgan 2018/7/18 CFP電子化
                        If m_bolAddLP And bEdit Then
                           frm1105_1.m_RecNo = m_strLD18
                           frm1105_1.m_PdfName = PUB_CaseNo2FileName(pa(1), pa(2), pa(3), pa(4)) & "." & m_strCP10 & ".CUS.PDF"
                           frm1105_1.Show
                        End If
                        'end 2018/7/18
                     End If
                  End If
                  
                  'Added by Morgan 2018/4/18
                  If m_strHK(1) <> "" And m_HKNewCP09 <> "" Then
                     '抓最近發文收文號(代理人)
                     strExc(0) = "select cp09 from caseprogress where cp01='" & m_strHK(1) & "' and cp02='" & m_strHK(2) & "'" & _
                        " and cp03='" & m_strHK(3) & "' and cp04='" & m_strHK(4) & "' and cp09<'C' and cp44 is not null and cp10<>'421'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
                        'strExc(2) = Pub_GetSpecMan("PS4")
                        strExc(2) = PUB_GetLetterJudgeNew("2", m_strHK(1), "913", "013", , m_HKCloser)
                        PUB_AddAppForm m_HKNewCP09, True, strExc(2)
                        NowPrint RsTemp("cp09"), "14", "42", False, m_HKCloser, , , , , , , , , , , , , m_HKNewCP09
                     End If
                  End If
                  'end 2018/4/18
         
                  'Add By Sindy 2016/10/7
                  If Me.m_strIR01 <> "" Then
                     bolLeave = True
                     Unload frm05010404_1
                     Unload Me
                     'Modify By Sindy 2022/5/20
                     'frm04010519.GoNext
                     Forms(0).Tmpfrm04010519.GoNext
                     Set Forms(0).Tmpfrm04010519 = Nothing
                     '2022/5/20 END
                  Else
                  '2016/10/7 END
                     bolLeave = True
                     Unload Me
                  End If
               End If
            End If
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         Unload Me
   End Select
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt() As String
   Dim ii As Integer
   Dim stNP07 As String, stDesc As String
   
   EndLetter ET01, m_NewCP09, ET03, strUserNum
   
   ii = 0
   
   '2011/11/14 add by sonia
   If ET03 = "01" Then
      ii = ii + 1
      ReDim Preserve strTxt(ii) As String
      Select Case pa(9)
         Case "101"
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
               "','列印備註','視為放棄')"
         'Modified by Morgan 2013/9/23 +法國 203
         Case "221", "203"
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
               "','列印備註','視為撤回')"
         Case Else
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
               "','列印備註','視為放棄（視為撤回）')"
      End Select
   End If
   '2011/11/14 end
   
   If Text1(3) <> "" Then
      ii = ii + 1
      ReDim Preserve strTxt(ii) As String
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','專利權消滅日','" & DBDATE(Text1(3)) & "')"
   End If
   
   If Text1(2) = "3" Then
      If m_NP07 = "605" Then
         strExc(2) = PUB_GetNextYear(pa, strExc(1))
         If strExc(1) <> "" Then strExc(2) = strExc(1)
         '2009/10/23 modify by sonia CFP-019149
         'strExc(0) = "第" & strExc(2) & "年" & m_NP07Desc
         strExc(0) = strExc(2) & m_NP07Desc
         '2009/10/23 end
      Else
         strExc(2) = PUB_GetNextTime(pa, strExc(1))
         If strExc(1) <> "" Then strExc(2) = strExc(1)
         '2009/10/23 modify by sonia CFP-019149
         'strExc(0) = "第" & strExc(2) & m_NP07Desc
         strExc(0) = strExc(2) & m_NP07Desc
         '2009/10/23 end
      End If
   
      ii = ii + 1
      ReDim Preserve strTxt(ii) As String
            
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','費用說明','" & strExc(0) & "')"
   End If
   
   'Add by Morgan 2011/4/20
   If Text1(6) <> "" Then
      strExc(1) = CompDate(1, 2, Text1(6))
      ii = ii + 1
      ReDim Preserve strTxt(ii) As String
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','復權期限','" & strExc(1) & "')"
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean

   'Added by Morgan 2021/12/8 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If

   If Text1(2) = "" Then
      MsgBox "消滅函類別不可空白!!", vbExclamation
      Text1(2).SetFocus
      Exit Function
   End If
   If Text1(2) = "3" Or Text1(2) = "4" Then
      If Text1(3) = "" Then
         MsgBox "該消滅函類別，消滅日不可空白!!", vbExclamation
         Text1(3).SetFocus
         Exit Function
      Else
         Text1_Validate 3, bCancel
         If bCancel Then
            Exit Function
         End If
      End If
   End If
   
   'Add by Morgan 2011/4/20
   'EU設計要輸入專利權終止日
   '2012/5/3 MODIFY BY SONAI 慧汶說EPC及EU的逾期未辦,未續繳年費,未繳實審都要輸入
   'If Text1(0) <> "N" And Text1(2) = "3" And Text1(6).Enabled = True And Text1(6) = "" Then
   'Modify by Amy 2013/05/22
   'If Text1(0) <> "N" And (Text1(2) = "1" Or Text1(2) = "3" Or Text1(2) = "5") And Text1(6).Enabled = True And Text1(6) = "" Then
   If Text1(0) <> "N" And (Text1(2) = "1" Or Text1(2) = "2" Or Text1(2) = "3" Or Text1(2) = "5") And Text1(6).Enabled = True And Text1(6) = "" Then
      If pa(9) = "239" And pa(8) = "3" Then
         MsgBox "EU設計要輸入專利權終止日以便計算定稿內相關日期！"
         Text1(6).SetFocus
         Exit Function
      'Add by Morgan 2011/8/15
      ElseIf pa(9) = "221" And pa(8) = "1" Then
         MsgBox "EPC發明案要輸入專利權終止日以便計算定稿內相關日期！"
         Text1(6).SetFocus
         Exit Function
      'Added by Morgan 2103/9/23  2013/12/25加新型
      ElseIf pa(9) = "203" And (pa(8) = "1" Or pa(8) = "2") Then
         MsgBox "法國發明案要輸入專利權終止日以便計算定稿內相關日期！"
         Text1(6).SetFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

Private Function SaveData() As Boolean

   Dim stCP13 As String, stCP12 As String
   Dim cp() As String
   ReDim cp(1 To TF_CP) As String
   Dim stUpdatePA As String
   Dim str111np01 As String, str111np22 As String 'Added by Morgan 2018/4/18

On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
On Error GoTo ErrorHandler1
   
   
   '屆滿 更新可結餘日
   'If Text1(2) = "4" Then   'cancel by sonia 2025/4/23 郭因CFP-032291閉卷後又收文分析，但沒有後續，且消滅函並非屆滿故改為不分類別都詢問
      '函數有控制存檔前詢問確認要的才會更新
      Pub_UpdateEndModCash pa(1), pa(2), pa(3), pa(4)
   'End If                   'cancel by sonia 2025/4/23 郭因CFP-032291閉卷後又收文分析，但沒有後續，且消滅函並非屆滿故改為不分類別都詢問
   
   stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
   stCP12 = GetSalesArea(stCP13)
   cp(1) = pa(1)
   cp(2) = pa(2)
   cp(3) = pa(3)
   cp(4) = pa(4)
   cp(5) = strSrvDate(1)
   cp(9) = 主管機關來函
   cp(10) = 專利權消滅
   cp(12) = stCP12
   cp(13) = stCP13
   cp(14) = strUserNum
   cp(27) = strSrvDate(1)
   cp(20) = "N"
   cp(26) = "N"
   cp(32) = "N"
   cp(64) = Text1(4)
   cp(119) = DBDATE(lblCaseField(4))
   
   'Add by Morgan 2011/4/20
   If Text1(6) <> "" Then
      cp(64) = "專利權終止通知書發出日：" & Text1(6) & ";" & cp(64)
   End If
   
   strSql = GetCPSQL(cp(), False)
   cnnConnection.Execute strSql, intI
   
   m_NewCP09 = cp(9)
   
   '抓最新的AB類發文代理人更新
   Pub_UpdateFromMaxCP27 pa(1), pa(2), pa(3), pa(4)
   
   stUpdatePA = "PA17='N'"
   
   '未閉卷者才更新
   If lblCaseField(5) = "" And Text1(5) = "Y" Then
      stUpdatePA = stUpdatePA & ",PA57='Y',PA58=" & strSrvDate(1) & ", PA59='89'"
   End If
   strSql = "UPDATE PATENT SET " & stUpdatePA & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   cnnConnection.Execute strSql, intI
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm05010404_1"
   End If
   '2016/10/7 END
   
   'Added by Morgan 2018/4/18
   '若有香港案 1.有已收文未發文則email通知智權人員銷案 2.若無但有智權人員期限者通知結案 3.其他則通知智權人員系統自動結案(閉卷)
   m_strHK(1) = ""
   str111np01 = "": str111np22 = ""
   strExc(0) = "select cm01,cm02,cm03,cm04 from casemap,patent" & _
      " where cm10='4' and cm05='" & pa(1) & "' and cm06='" & pa(2) & "'" & _
      " and cm07='" & pa(3) & "' and cm08='" & pa(4) & "'" & _
      " and pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 and pa09='013' and pa57 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_strHK(1) = RsTemp("cm01")
      m_strHK(2) = RsTemp("cm02")
      m_strHK(3) = RsTemp("cm03")
      m_strHK(4) = RsTemp("cm04")
      
      strExc(1) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
      strExc(2) = m_strHK(1) & "-" & m_strHK(2) & IIf(m_strHK(3) & m_strHK(4) = "000", "", "-" & m_strHK(3) & "-" & m_strHK(4))
      strExc(3) = ""
      strExc(0) = "select cpm04,cp10 from caseprogress,casepropertymap" & _
         " where cp01='" & m_strHK(1) & "' and cp02='" & m_strHK(2) & "' and cp03='" & m_strHK(3) & "' and cp04='" & m_strHK(4) & "'" & _
         " and cp09<'B' and cp27||cp57 is null and cpm01(+)=cp01 and cpm02(+)=cp10"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(3) = lblNation & "案" & strExc(1) & "已消滅,香港案" & strExc(2) & "(" & RsTemp("cpm04") & ")將無法辦理，請銷案。"
      Else
         '判斷香港第二階段111未收文(EX.香港案P98784,大陸案P95625)
         strExc(0) = "select cp10 from caseprogress where cp01='" & m_strHK(1) & "' and cp02='" & m_strHK(2) & "' and cp03='" & m_strHK(3) & "' and cp04='" & m_strHK(4) & "'" & _
            " and cp09<'B' and cp10 = '111' and cp159=0 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            strExc(3) = "相關" & lblNation & "案" & strExc(1) & "已消滅,香港案" & strExc(2) & "已自動結案。"
            strSql = "update patent set PA57='Y',PA58=" & DBDATE(lblCaseField(4)) & ", PA59='99',pa91=pa91||';相關" & lblNation & "案" & strExc(1) & "已消滅本案自動上閉卷(" & DBDATE(lblCaseField(4)) & ")。' where pa01='" & m_strHK(1) & "' and pa02='" & m_strHK(2) & "' and pa03='" & m_strHK(3) & "' and pa04='" & m_strHK(4) & "'"
            cnnConnection.Execute strSql, intI
            '抓香港案下一程序之標準專利批准記錄請求未收文期限的總收文號
            strExc(0) = "SELECT NP01,NP22 FROM NEXTPROGRESS WHERE NP02='" & m_strHK(1) & "' AND NP03='" & m_strHK(2) & "' AND NP04='" & m_strHK(3) & "' AND NP05='" & m_strHK(4) & "' AND NP07='111' AND NP06 IS NULL"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               str111np01 = "" & RsTemp.Fields(0)
               str111np22 = "" & RsTemp.Fields(1)
               '更新下一程序之標準專利批准記錄請求未收文期限的續辦欄及解除期限日期,原因
               strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & strSrvDate(1) & ",NP12='99' WHERE NP01='" & str111np01 & "' AND NP22='" & str111np22 & "' AND NP02='" & m_strHK(1) & "' AND NP03='" & m_strHK(2) & "' AND NP04='" & m_strHK(3) & "' AND NP05='" & m_strHK(4) & "'"
               cnnConnection.Execute strSql, intI
            End If
            
            '產生閉卷進度
            m_HKNewCP09 = AutoNo("B", 6)
            'Added by Morgan 2025/1/24
            If strSrvDate(1) >= P業務區劃分啟用日 Then
               m_HKCloser = PUB_GetPHandler(m_strHK(1) & m_strHK(2) & m_strHK(3) & m_strHK(4))
            Else
            'end 2025/1/24
               m_HKCloser = Pub_GetSpecMan("專利處轉信非台灣程序2")
            End If 'Added by Morgan 2025/1/24
            strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
               "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP30,CP57,CP58,CP64) VALUES " & _
               "('" & m_strHK(1) & "','" & m_strHK(2) & "','" & m_strHK(3) & "','" & m_strHK(4) & "'," & strSrvDate(1) & _
               ",'" & m_HKNewCP09 & "','913','90'," & CNULL(GetSalesArea(PUB_GetAKindSalesNo(m_strHK(1), m_strHK(2), m_strHK(3), m_strHK(4)))) & "," & CNULL(PUB_GetAKindSalesNo(m_strHK(1), m_strHK(2), m_strHK(3), m_strHK(4))) & _
               ",'" & m_HKCloser & "','N','N'," & strSrvDate(1) & ",'N','" & str111np01 & "','" & str111np22 & "'," & strSrvDate(1) & ",'99','相關" & lblNation & "案" & strExc(1) & "已消滅本案自動上閉卷(" & DBDATE(lblCaseField(4)) & ")。') "
            cnnConnection.Execute strSql, intI
         
            'EMail通知閉卷承辦人(同操作人時也要寄)
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " values('" & strUserNum & "','" & m_HKCloser & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'香港案" & strExc(2) & "已閉卷請至待處理區上傳指示信','如旨')"
            cnnConnection.Execute strSql, intI
         End If
      End If
      '通知香港案智權人員
      If strExc(3) <> "" Then
         strExc(4) = PUB_GetAKindSalesNo(m_strHK(1), m_strHK(2), m_strHK(3), m_strHK(4))
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " values('" & strUserNum & "','" & strExc(4) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & strExc(3) & "','如旨')"
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2018/4/18
   
   'Added by Morgan 2018/7/18 CFP電子化
   If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
      m_strLD18 = cp(9)
      m_strCP10 = cp(10)
      strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), m_strCP10, pa(9))
      '逾期未辦(1) & 未領證(2)掛號
      'Modified by Morgan 2024/3/11
      'PUB_AddLetterProgress m_strLD18, 2, True, strExc(1), IIf(Text1(2) = "1" Or Text1(2) = "2", True, False), pa(26), m_strCP10, pa(75)
      PUB_AddLetterProgress m_strLD18, 2, IIf(Text1(0).Text <> "N", True, False), strExc(1), IIf(Text1(2) = "1" Or Text1(2) = "2", True, False), pa(26), m_strCP10, pa(75)
      'end 2024/3/11
      m_bolAddLP = True
   End If
   'end 2018/7/18
   
   cnnConnection.CommitTrans
   SaveData = True
   Exit Function
   
ErrorHandler1:
   cnnConnection.RollbackTrans
   
ErrorHandler:
   'MsgBox Err.Description, vbCritical

End Function

Private Sub Form_Activate()
   If Not bolActived Then
      bolActived = True
      ReadPatent
      'Add by Morgan 2011/5/12
      'Modify by Morgan 2011/8/15 +EPC 發明
      'Modified by Morgan 2103/9/23 +法國 發明  2013/12/25 加新型
      If (pa(8) = "3" And pa(9) = "239") Or (pa(8) = "1" And pa(9) = "221") Or ((pa(8) = "1" Or pa(8) = "2") And pa(9) = "203") Then
         Text1(6).Enabled = True
      End If
   End If
End Sub

Private Sub Form_Initialize()
   ReDim pa(TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Add By Sindy 2017/12/28
   m_strIR01 = frm05010404_1.m_strIR01
   m_strIR02 = frm05010404_1.m_strIR02
   m_strIR03 = frm05010404_1.m_strIR03
   m_strIR04 = frm05010404_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/28 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2018/4/18
   'Add By Sindy 2016/10/13
   If Me.m_strIR01 = "" Then
   '2016/10/13 END
      If intLeaveKind = 2 Then
         Unload frmParent
      Else
         frmParent.Show
         If intLeaveKind = 0 Then
            frmParent.Clear
         End If
      End If
   End If
   Set frm05010404_2 = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bolLeave = False Then
      If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Cancel = 1
      End If
   End If
End Sub

Private Sub ReadPatent()
   Dim strTemp As String
   pa(1) = txtSystem
   pa(2) = txtCode(0)
   pa(3) = txtCode(1)
   pa(4) = txtCode(2)
   If ClsPDReadPatentDatabase(pa(), intPWhere) Then
      '申請案號
      lblCaseField(1) = pa(11)
      '案件名稱
      SetNameToCombo cboCaseName, pa(5), pa(6), pa(7)
      '專利種類
      lblCaseField(2) = pa(8)
      If ClsPDGetPatentTrademarkKind(Trim(intPCaseKind), lblCaseField(2), strTemp, , pa(9)) Then
         lblTrademarkKind = strTemp
      End If
      '申請國家
      lblCaseField(3) = pa(9)
      If ClsPDGetNation(lblCaseField(3), strTemp) Then
         lblNation.Caption = strTemp
      End If
      Text1(5) = pa(57)
      lblCaseField(5) = Text1(5)
      If GetNP07(pa(9), pa(8), m_NP07) = True Then
         ClsPDGetCaseProperty pa(1), m_NP07, m_NP07Desc
         lblCaseField(6) = m_NP07Desc
      End If
   End If
   If Text1(5) = "" Then Text1(5) = "Y"   'add by sonia 2018/3/15
End Sub


Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   Select Case Index
      Case 4
         OpenIme
      Case Else
         CloseIme
   End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 0
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      Case 1, 5
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
      Case 2
         If KeyAscii <> 8 And Not (KeyAscii >= Asc("1") And KeyAscii <= Asc("5")) Then
            KeyAscii = 0
            Beep
         Else
            Text1(4) = Replace(Text1(4), ",消滅-逾期未辦", "")
            Text1(4) = Replace(Text1(4), ",消滅-未領證", "")
            Text1(4) = Replace(Text1(4), ",消滅-未續繳" & m_NP07Desc, "")
            Text1(4) = Replace(Text1(4), ",消滅-屆滿", "")
            Text1(4) = Replace(Text1(4), ",消滅-未繳實審", "")
            Select Case Chr(KeyAscii)
               Case "1"
                  Text1(4) = Text1(4) & ",消滅-逾期未辦"
               Case "2"
                  Text1(4) = Text1(4) & ",消滅-未領證"
               Case "3"
                  Text1(4) = Text1(4) & ",消滅-未續繳" & m_NP07Desc
               Case "4"
                  Text1(4) = Text1(4) & ",消滅-屆滿"
               Case "5"
                  Text1(4) = Text1(4) & ",消滅-未繳實審"
            End Select
         End If
      
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 3
         If Text1(Index) <> "" Then
            If CheckIsDate(Text1(Index)) = False Then
               Cancel = True
               TextInverse Text1(Index)
            ElseIf Val(Text1(Index)) > Val(strSrvDate(1)) Then
               ShowMsg MsgText(1050)
               Cancel = True
               TextInverse Text1(Index)
            End If
         End If
      Case 5
         If lblCaseField(5) = "" And Text1(5) = "Y" Then
            If MsgBox("是否確定閉卷 ?", vbYesNo + vbQuestion) = vbNo Then
               Cancel = True
               TextInverse Text1(Index)
            End If
         End If
      Case 6 'Add by Morgan 2011/4/20
         If Text1(Index) <> "" Then
            If CheckIsDate(Text1(Index)) = False Then
               Cancel = True
               TextInverse Text1(Index)
            ElseIf Val(Text1(Index)) > Val(strSrvDate(1)) Then
               MsgBox "專利權終止通知書發出日不可晚於系統日！"
               Cancel = True
               TextInverse Text1(Index)
            End If
         End If
   End Select
End Sub

'消滅函類別選擇檢查
Private Function fnChoiceCheck() As Boolean
   
   Select Case Text1(2).Text
      '選1時
      Case "1"
         '若有專用期間則確認
         If pa(24) <> "" Then
            If MsgBox("此案有專用期間，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
      '選2時
      Case "2"
         '1.若下一程序領證未上不續辦('N')則確認
         strExc(0) = "SELECT 1 FROM NEXTPROGRESS" & _
            " WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "'" & _
            " AND NP07='601' AND (NP06 IS NULL or NP06='Y')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If MsgBox("此案領證未結案，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
         '2.若有專用期間則確認
         If pa(24) <> "" Then
            If MsgBox("此案有專用期間，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
         '3.若非核准案則確認
         If pa(16) <> "1" Then
            If MsgBox("此案非核准案，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
      '選3時
      Case "3"
         '1.若無專用期間則確認
         If pa(24) = "" Then
            If MsgBox("此案無專用期間，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
         strExc(0) = "SELECT MAX(NP08||NVL(NP06,'X')) FROM NEXTPROGRESS" & _
            " WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "'" & _
            " AND NP07='605'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '2.若下一程序最大年費未上'N'則確認
            If Right("" & RsTemp.Fields(0), 1) = "X" Then
               If MsgBox("此案年費未結案，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
            End If
            '3.若下一程序最大年費期限>=系統日則確認
            If Val(Left("" & RsTemp.Fields(0), 8)) >= Val(strSrvDate(1)) Then
               If MsgBox("此案年費未逾期，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
            End If
         End If
      '選4時
      Case "4"
         '1.若無專用期間則確認
         If pa(24) = "" Then
            If MsgBox("此案無專用期間，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
         '2.若專用期止日>=系統日則確認
         If Val(pa(25)) >= Val(strSrvDate(1)) Then
            If MsgBox("此案未屆滿，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
      '選5時
      Case "5"
         '1.若有專用期間則確認
         If pa(24) <> "" Then
            If MsgBox("此案有專用期間，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
         strExc(0) = "SELECT MAX(NP08||NVL(NP06,'X')) FROM NEXTPROGRESS" & _
            " WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "'" & _
            " AND NP07='416'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '2.若下一程序最大實審費未上'N'則確認
            If Right("" & RsTemp.Fields(0), 1) = "X" Then
               If MsgBox("此案實審費未結案，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
            End If
            '3.若下一程序最大實審費期限>=系統日則確認
            If Val(Left("" & RsTemp.Fields(0), 8)) >= Val(strSrvDate(1)) Then
               If MsgBox("此案實審費未逾期，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
            End If
         End If
   End Select
   fnChoiceCheck = True
   
flgErr:
   
End Function
'是否有尚未結餘的程序
Private Function CheckCP109() As Boolean
   strExc(0) = "select * from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp57 is null and cp27>0 and cp59 is null and cp109 is null and rownum<2"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      CheckCP109 = True
   End If
   
End Function
