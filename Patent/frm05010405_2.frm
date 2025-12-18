VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm05010405_2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "年費逾期補繳通知函輸入"
   ClientHeight    =   3765
   ClientLeft      =   105
   ClientTop       =   810
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   9000
   Begin VB.TextBox text1 
      Height          =   270
      Index           =   1
      Left            =   1665
      MaxLength       =   1
      TabIndex        =   0
      Top             =   3375
      Width           =   300
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2115
      MaxLength       =   6
      TabIndex        =   20
      Top             =   463
      Width           =   855
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2970
      MaxLength       =   1
      TabIndex        =   19
      Top             =   463
      Width           =   330
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3285
      MaxLength       =   2
      TabIndex        =   18
      Top             =   463
      Width           =   435
   End
   Begin VB.TextBox txtSystem 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1665
      MaxLength       =   3
      TabIndex        =   17
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
      TabIndex        =   2
      Top             =   30
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   6705
      TabIndex        =   1
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
      TabIndex        =   3
      Top             =   30
      Width           =   600
   End
   Begin MSForms.ComboBox cboCustName 
      Height          =   300
      Left            =   1665
      TabIndex        =   21
      Top             =   1230
      Width           =   6960
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12277;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cboCaseName 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1665
      TabIndex        =   4
      Top             =   840
      Width           =   6945
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12250;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1665
      TabIndex        =   33
      Top             =   2670
      Width           =   1995
   End
   Begin VB.Label lblCaseField 
      Alignment       =   2  '置中對齊
      Caption         =   "此案已閉卷"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   4770
      TabIndex        =   32
      Top             =   1980
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label17 
      Caption         =   "  (Y : Word )"
      Height          =   255
      Index           =   0
      Left            =   2070
      TabIndex        =   31
      Top             =   3375
      Width           =   915
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請人１："
      Height          =   255
      Index           =   3
      Left            =   660
      TabIndex        =   30
      Top             =   1245
      Width           =   900
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "未繳年度(次數)："
      Height          =   255
      Index           =   5
      Left            =   165
      TabIndex        =   29
      Top             =   2685
      Width           =   1395
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "原繳費期限："
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   28
      Top             =   3015
      Width           =   1080
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "最後期限："
      Height          =   255
      Index           =   7
      Left            =   4785
      TabIndex        =   27
      Top             =   3015
      Width           =   990
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "是否修改定稿："
      Height          =   255
      Index           =   8
      Left            =   300
      TabIndex        =   26
      Top             =   3375
      Width           =   1260
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   6
      Left            =   1665
      TabIndex        =   25
      Top             =   2325
      Width           =   6435
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   8
      Left            =   5880
      TabIndex        =   24
      Top             =   3015
      Width           =   1995
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "已繳年度："
      Height          =   255
      Index           =   9
      Left            =   570
      TabIndex        =   23
      Top             =   2325
      Width           =   990
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   7
      Left            =   1665
      TabIndex        =   22
      Top             =   3015
      Width           =   1995
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5835
      TabIndex        =   16
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請案號："
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "櫃台收文日："
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   1980
      Width           =   1080
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   585
      TabIndex        =   13
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   1665
      TabIndex        =   12
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1665
      TabIndex        =   11
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label Label22 
      Alignment       =   1  '靠右對齊
      Caption         =   "專利種類："
      Height          =   255
      Left            =   660
      TabIndex        =   10
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   585
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   2115
      TabIndex        =   8
      Top             =   1620
      Width           =   2430
   End
   Begin MSForms.Label lblNation 
      Height          =   255
      Left            =   6540
      TabIndex        =   7
      Top             =   1620
      Width           =   2055
      VariousPropertyBits=   27
      Size            =   "3625;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請國家："
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5835
      TabIndex        =   5
      Top             =   1620
      Width           =   645
   End
End
Attribute VB_Name = "frm05010405_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/8 改成Form2.0 (cboCaseName,cboCustName,lblNation)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Morgan 2008/12/15
Option Explicit

Public frmParent As Form
Dim bolActived As Boolean
Dim pa() As String
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'intLeaveKind判斷離開時，是2:結束  1:回上一畫面
Dim intLeaveKind As Integer
Dim m_NewCP09 As String, m_NP07 As String, m_NP07Desc As String, m_DelayDesc As String
Dim m_strMaxPA72 As String '目前繳費年度
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim m_bolAddLP As Boolean, m_strCP10 As String, m_strLD18 As String 'Added by Morgan 2018/7/18 CFP電子化

Private Sub cmdOK_Click(Index As Integer)
   Dim bEdit As Boolean
   Dim strTmp As String
   
   intLeaveKind = Index
   bolLeave = False
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         If TxtValidate = True Then
            If SaveData = False Then
               MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
            Else
               If Text1(1).Text = "Y" Then
                  bEdit = True
               Else
                  bEdit = False
               End If
               StartLetter "03", "01"
               NowPrint m_NewCP09, "03", "01", bEdit, strUserNum, , , , , , , , , , , , , m_strLD18
               'Added by Morgan 2018/7/18 CFP電子化
               If m_bolAddLP And bEdit Then
                  frm1105_1.m_RecNo = m_strLD18
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(pa(1), pa(2), pa(3), pa(4)) & "." & m_strCP10 & ".CUS.PDF"
                  frm1105_1.Show
               End If
               'end 2018/7/18
               
               'Add By Sindy 2016/10/7
               If Me.m_strIR01 <> "" Then
                  bolLeave = True
                  Unload frm05010405_1
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
   ii = ii + 1
   ReDim Preserve strTxt(ii) As String
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & "','年費法定期限'," & DBDATE(lblCaseField(7)) & ")"
   
   '最後期限(本所)
   ii = ii + 1
   ReDim Preserve strTxt(ii) As String
   Dim strDate(0 To 3) As String
   strDate(1) = pa(1)     '系統別
   strDate(2) = pa(9)     '申請國家
   strDate(3) = DBDATE(lblCaseField(8))  '下次法定期限
   GetCtrlDT strDate()
   strExc(0) = ChangeWStringToTString(strDate(0))
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & "','下次繳年費日','" & DBDATE(strDate(0)) & "')"

   ii = ii + 1
   ReDim Preserve strTxt(ii) As String
         
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
      "','費用說明','" & lblCaseField(0) & "')"
   
   'Added by Morgan 2012/3/28 要區分年費或維持費--禧佩 Ex.CFP-009805
   ii = ii + 1
   ReDim Preserve strTxt(ii) As String
   
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
      "','案件性質','" & m_NP07Desc & "')"
   'end 2012/3/28
   
   ii = ii + 1
   ReDim Preserve strTxt(ii) As String
         
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
      "','列印備註','" & m_DelayDesc & "')"
      
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   If lblCaseField(0) = "" Then
      MsgBox "無法取得未繳年度(次數)資料!!", vbExclamation
      Exit Function
   End If
   If lblCaseField(7) = "" Then
      MsgBox "無法取得原繳費期限!!", vbExclamation
      Exit Function
   End If
   
   If lblCaseField(8) = "" Then
      MsgBox "無法取得最後期限!!", vbExclamation
      Exit Function
   End If
   
   If Val(DBDATE(lblCaseField(7))) > Val(strSrvDate(1)) Then
      MsgBox "原繳費期限不可大於系統日!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If Val(DBDATE(lblCaseField(8))) < Val(strSrvDate(1)) Then
      MsgBox "最後期限不可小於或等於系統日!!!", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   '重複通知檢查
   If DupeCheck = False Then
      Exit Function
   End If
                  
   TxtValidate = True
End Function

Private Function SaveData() As Boolean

   Dim stCP13 As String, stCP12 As String
   Dim cp() As String
   ReDim cp(1 To TF_CP) As String
   Dim stUpdatePA As String
 
On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
On Error GoTo ErrorHandler1
  
   stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
   stCP12 = GetSalesArea(stCP13)
   cp(1) = pa(1)
   cp(2) = pa(2)
   cp(3) = pa(3)
   cp(4) = pa(4)
   cp(5) = strSrvDate(1)
   cp(9) = 主管機關來函
   cp(10) = "1605"
   cp(12) = stCP12
   cp(13) = stCP13
   cp(14) = strUserNum
   cp(27) = strSrvDate(1)
   cp(20) = "N"
   cp(26) = "N"
   cp(32) = "N"
   cp(64) = "未繳年度(次數):" & lblCaseField(0)
   cp(119) = DBDATE(lblCaseField(4))   '2012/11/5 add by sonia
   
   strSql = GetCPSQL(cp(), False)
   cnnConnection.Execute strSql, intI
   
   m_NewCP09 = cp(9)
   If pa(9) <> "000" Then
      '抓最新的AB類發文代理人更新
      Pub_UpdateFromMaxCP27 cp(1), cp(2), cp(3), cp(4)
   End If
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm05010405_1"
   End If
   '2016/10/7 END
   
   'Added by Morgan 2018/7/18 CFP電子化
   If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
      m_strLD18 = cp(9)
      m_strCP10 = cp(10)
      strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), m_strCP10, pa(9))
      PUB_AddLetterProgress m_strLD18, 2, True, strExc(1), True, pa(26), m_strCP10, pa(75)
      m_bolAddLP = True
   End If
   'end 2018/7/18
   
   cnnConnection.CommitTrans
   SaveData = True
   Exit Function
   
ErrorHandler1:
   cnnConnection.RollbackTrans
   
ErrorHandler:

End Function

Private Sub Form_Activate()
   If Not bolActived Then
      bolActived = True
      ReadPatent
   End If
End Sub

Private Sub Form_Initialize()
   ReDim pa(TF_PA) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'Add By Sindy 2017/12/28
   m_strIR01 = frm05010405_1.m_strIR01
   m_strIR02 = frm05010405_1.m_strIR02
   m_strIR03 = frm05010405_1.m_strIR03
   m_strIR04 = frm05010405_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/28 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
   Set frm05010405_2 = Nothing
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
   Dim arrPA72
   
   pa(1) = txtSystem
   pa(2) = txtCode(0)
   pa(3) = txtCode(1)
   pa(4) = txtCode(2)
   If ClsPDReadPatentDatabase(pa(), intPWhere) Then
      '申請案號
      lblCaseField(1) = pa(11)
      '案件名稱
      SetNameToCombo cboCaseName, pa(5), pa(6), pa(7)
      '申請人1(中-->英-->日)
      If pa(26) <> "" Then
         strExc(0) = "select cu04,cu05,cu06,cu88,cu89,cu90 from customer where cu01='" & Left(pa(26) & "000", 8) & "' and cu02='" & Mid(pa(26) & "000", 9, 1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            With RsTemp
            SetNameToCombo cboCustName, "" & .Fields("CU04"), Trim(.Fields("CU05") & " " & .Fields("CU88") & " " & .Fields("CU89") & " " & .Fields("CU90")), "" & .Fields("CU06")
            End With
         End If
      End If
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
      If pa(57) = "Y" Then
         lblCaseField(5).Visible = True
      Else
         lblCaseField(5).Visible = False
      End If
      
      '已繳年度
      lblCaseField(6) = pa(72)
      If pa(72) <> "" Then
         arrPA72 = Split(pa(72), ",")
         m_strMaxPA72 = arrPA72(UBound(arrPA72))
      End If
      '原期限,補繳期限
      If GetNP07(pa(9), pa(8), m_NP07) = True Then
         ClsPDGetCaseProperty pa(1), m_NP07, m_NP07Desc
         If m_NP07 = "605" Then
            strExc(2) = PUB_GetNextYear(pa, strExc(1))
            strExc(3) = strExc(2)
            If strExc(1) <> "" Then strExc(2) = strExc(1)
            '2009/9/23 modify by sonia CFP-007508-0-01
            'lblCaseField(0) = "第" & strExc(2) & "年" & m_NP07Desc
            lblCaseField(0) = strExc(2) & m_NP07Desc
            '2009/9/23 end
         Else
            strExc(2) = PUB_GetNextTime(pa, strExc(1))
            strExc(3) = strExc(2)
            If strExc(1) <> "" Then strExc(2) = strExc(1)
            '2009/9/23 modify by sonia
            'lblCaseField(0) = "第" & strExc(2) & m_NP07Desc
            lblCaseField(0) = strExc(2) & m_NP07Desc
            '2009/9/23 end
         End If
         
         'Add by Morgan 2009/7/2 EPC子案抓各國已繳費年度計算 Ex.CFP-015807-0-06(波蘭)
         If pa(4) <> "00" Then
            '子案期限為申請日+最大繳費年(起算日照EPC,繳費年依各國)
            'Modify by Morgan 2009/12/4 改以下次繳費年-1計算 Ex.CFP-017891-0-05(英國)
            strExc(1) = CompDate(0, Val(strExc(3) - 1), pa(10))
            lblCaseField(7).Caption = ChangeWStringToTString(strExc(1))
         Else
         'end 2009/7/2
            strExc(0) = "Select MAX(NP09) From NextProgress Where NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP07='" & m_NP07 & "' and (np06 is null or np06='N')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With RsTemp
               'Add by Morgan 2009/5/25
               '若無下一程序時要再抓進度檔,因為有可能收文後取消(Ex. CFP-19716)
               If IsNull(.Fields(0)) Then
                  strExc(0) = "Select MAX(cp07) From caseProgress Where cP01='" & pa(1) & "' AND cP02='" & pa(2) & "' AND cP03='" & pa(3) & "' AND cP04='" & pa(4) & "' AND cP10='" & m_NP07 & "' and cp27 is null"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     lblCaseField(7).Caption = ChangeWStringToTString("" & RsTemp.Fields(0).Value)
                  End If
               'end 2009/5/25
               Else
                  lblCaseField(7).Caption = ChangeWStringToTString("" & .Fields(0).Value)
               End If
               End With
            End If
         End If
         
         If lblCaseField(7).Caption <> "" Then
            strExc(0) = "Select CF12,CF28 From CASEFEE Where CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & m_NP07 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With RsTemp
                  If Val("" & .Fields("CF12")) > 0 Then
                     strExc(0) = .Fields("CF12")
                     m_DelayDesc = strExc(0) & "天"
                     lblCaseField(8).Caption = ChangeWStringToTString(CompDate(2, Val(strExc(0)), Format(lblCaseField(7).Caption)))
                  ElseIf Val("" & .Fields("CF28")) > 0 Then
                     strExc(0) = .Fields("CF28")
                     m_DelayDesc = strExc(0) & "個月"
                     lblCaseField(8).Caption = ChangeWStringToTString(CompDate(1, Val(strExc(0)), Format(lblCaseField(7).Caption)))
                  End If
               End With
            End If
         End If
      End If
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      Case 1
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Function DupeCheck() As Boolean
   DupeCheck = True
   strExc(0) = "select cp10, cp27 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "'  and cp04='" & pa(4) & "' and cp57 is null order by cp27 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields("cp10") = "1605" Then
         If MsgBox("本案件已於 " & ChangeWStringToTDateString("" & RsTemp.Fields(1)) & " 通知逾期補繳，是否再次通知？", vbYesNo + vbDefaultButton2) = vbNo Then
            DupeCheck = False
         End If
      End If
   End If
End Function
