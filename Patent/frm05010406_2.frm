VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm05010406_2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "實體審查、領證費逾期補繳通知函"
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
      Index           =   4
      Left            =   7230
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2955
      Width           =   1245
   End
   Begin VB.TextBox text1 
      Height          =   270
      Index           =   3
      Left            =   4410
      MaxLength       =   8
      TabIndex        =   2
      Top             =   2955
      Width           =   1245
   End
   Begin VB.TextBox text1 
      Height          =   270
      Index           =   2
      Left            =   1620
      MaxLength       =   8
      TabIndex        =   1
      Top             =   2955
      Width           =   1245
   End
   Begin VB.TextBox text1 
      Height          =   270
      Index           =   0
      Left            =   1620
      MaxLength       =   1
      TabIndex        =   0
      Top             =   2340
      Width           =   300
   End
   Begin VB.TextBox text1 
      Height          =   270
      Index           =   1
      Left            =   1665
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3315
      Width           =   300
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2115
      MaxLength       =   6
      TabIndex        =   24
      Top             =   463
      Width           =   855
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2970
      MaxLength       =   1
      TabIndex        =   23
      Top             =   463
      Width           =   330
   End
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3285
      MaxLength       =   2
      TabIndex        =   22
      Top             =   463
      Width           =   435
   End
   Begin VB.TextBox txtSystem 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1665
      MaxLength       =   3
      TabIndex        =   21
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
      TabIndex        =   6
      Top             =   30
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   6705
      TabIndex        =   5
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
      TabIndex        =   7
      Top             =   30
      Width           =   600
   End
   Begin MSForms.ComboBox cboCustName 
      Height          =   300
      Left            =   1665
      TabIndex        =   25
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
      TabIndex        =   8
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
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "約定期限："
      Height          =   255
      Index           =   4
      Left            =   6165
      TabIndex        =   35
      Top             =   2955
      Width           =   990
   End
   Begin VB.Label Label17 
      Caption         =   "2 : 領證"
      Height          =   255
      Index           =   2
      Left            =   2025
      TabIndex        =   34
      Top             =   2610
      Width           =   1815
   End
   Begin VB.Label Label17 
      Caption         =   "1 : 實體審查 "
      Height          =   255
      Index           =   1
      Left            =   2025
      TabIndex        =   33
      Top             =   2340
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "通知函類別："
      Height          =   255
      Left            =   495
      TabIndex        =   32
      Top             =   2340
      Width           =   1080
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
      TabIndex        =   31
      Top             =   1980
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label17 
      Caption         =   "  (Y : Word )"
      Height          =   255
      Index           =   0
      Left            =   2070
      TabIndex        =   30
      Top             =   3315
      Width           =   915
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請人１："
      Height          =   255
      Index           =   3
      Left            =   660
      TabIndex        =   29
      Top             =   1245
      Width           =   900
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "原繳費期限："
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   28
      Top             =   2955
      Width           =   1080
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "補繳期限："
      Height          =   255
      Index           =   7
      Left            =   3345
      TabIndex        =   27
      Top             =   2955
      Width           =   990
   End
   Begin VB.Label Label17 
      Alignment       =   1  '靠右對齊
      Caption         =   "是否修改定稿："
      Height          =   255
      Index           =   8
      Left            =   300
      TabIndex        =   26
      Top             =   3315
      Width           =   1260
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5835
      TabIndex        =   20
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      Caption         =   "申請案號："
      Height          =   255
      Left            =   4800
      TabIndex        =   19
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "櫃台收文日："
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   1980
      Width           =   1080
   End
   Begin VB.Label Label4 
      Alignment       =   1  '靠右對齊
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   585
      TabIndex        =   17
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   1665
      TabIndex        =   16
      Top             =   1980
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1665
      TabIndex        =   15
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label Label22 
      Alignment       =   1  '靠右對齊
      Caption         =   "專利種類："
      Height          =   255
      Left            =   660
      TabIndex        =   14
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label6 
      Alignment       =   1  '靠右對齊
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   585
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   2115
      TabIndex        =   12
      Top             =   1620
      Width           =   2430
   End
   Begin MSForms.Label lblNation 
      Height          =   255
      Left            =   6540
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5835
      TabIndex        =   9
      Top             =   1620
      Width           =   645
   End
End
Attribute VB_Name = "frm05010406_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/8 改成Form2.0 (cboCaseName,cboCustName,lblNation)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Create by Morgan 2010/6/15
Option Explicit

Public frmParent As Form
Dim bolActived As Boolean
Dim pa() As String
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'intLeaveKind判斷離開時，是2:結束  1:回上一畫面
Dim intLeaveKind As Integer
Dim m_NewCP09 As String, m_DelayDesc As String
Dim m_CP64 As String
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim m_bolAddLP As Boolean, m_CP10 As String, m_LD18 As String 'Added by Morgan 2018/7/18 CFP電子化

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
               If Text1(0) = "1" Then
                  strTmp = "01"
               Else
                  strTmp = "02"
               End If
               StartLetter "03", strTmp
               NowPrint m_NewCP09, "03", strTmp, bEdit, strUserNum, , , , , , , , , , , , , m_LD18
               'Added by Morgan 2018/7/18 CFP電子化
               If m_bolAddLP And bEdit Then
                  frm1105_1.m_RecNo = m_LD18
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(pa(1), pa(2), pa(3), pa(4)) & "." & m_CP10 & ".CUS.PDF"
                  frm1105_1.Show
               End If
               'end 2018/7/18
               'Add By Sindy 2016/10/7
               If Me.m_strIR01 <> "" Then
                  bolLeave = True
                  Unload frm05010406_1
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
      "('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & "','原繳費期限'," & DBDATE(Text1(2)) & ")"
   
   ii = ii + 1
   ReDim Preserve strTxt(ii) As String
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & "','補繳期限','" & DBDATE(Text1(3)) & "')"

   ii = ii + 1
   ReDim Preserve strTxt(ii) As String
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & "','約定期限','" & DBDATE(Text1(4)) & "')"
      
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   
   If Text1(0) = "" Then
      MsgBox "請輸入通知函類別!!", vbExclamation
      Text1(0).SetFocus
      Exit Function
   End If
   
   If Text1(2) = "" Then
      MsgBox "請輸入原繳費期限!!", vbExclamation
      Text1(2).SetFocus
      Exit Function
   Else
      Text1_Validate 2, bCancel
      If bCancel Then Exit Function
   End If
   
   If Text1(3) = "" Then
      MsgBox "請輸入補繳期限!!", vbExclamation
      Text1(3).SetFocus
      Exit Function
   Else
      Text1_Validate 3, bCancel
      If bCancel Then Exit Function
   End If
   
   If Text1(4) = "" Then
      MsgBox "請輸入約定期限!!", vbExclamation
      Text1(3).SetFocus
      Exit Function
   Else
      Text1_Validate 3, bCancel
      If bCancel Then Exit Function
   End If
   
   If Val(DBDATE(Text1(2))) > Val(strSrvDate(1)) Then
      MsgBox "原繳費期限不可大於系統日!!!", vbExclamation + vbOKOnly
      Text1(2).SetFocus
      Exit Function
   End If
   
   If Val(DBDATE(Text1(3))) <= Val(strSrvDate(1)) Then
      MsgBox "補繳期限不可小於或等於系統日!!!", vbExclamation + vbOKOnly
      Text1(3).SetFocus
      Exit Function
   End If
   
   If Val(DBDATE(Text1(3))) < Val(DBDATE(Text1(4))) Then
      MsgBox "約定期限不可大於補繳期限!!!", vbExclamation + vbOKOnly
      Text1(4).SetFocus
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
   'Modified by Morgan 2022/5/31
   'cp(10) = "1902"
   cp(10) = m_CP10
   'end 2022/5/31
   cp(12) = stCP12
   cp(13) = stCP13
   cp(14) = strUserNum
   cp(27) = strSrvDate(1)
   cp(20) = "N"
   cp(26) = "N"
   cp(32) = "N"
   'cp(64) = m_CP64 'Removed by Morgan 2022/5/31 已新增案件性質
   cp(119) = DBDATE(lblCaseField(4))    '2012/11/5 add by sonia
   
   strSql = GetCPSQL(cp(), False)
   cnnConnection.Execute strSql, intI
   
   m_NewCP09 = cp(9)
   If pa(9) <> "000" Then
      '抓最新的AB類發文代理人更新
      Pub_UpdateFromMaxCP27 cp(1), cp(2), cp(3), cp(4)
   End If
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm05010406_1"
   End If
   '2016/10/7 END
   
   'Added by Morgan 2018/7/18 CFP電子化
   If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
      m_LD18 = cp(9)
      m_CP10 = cp(10)
      'Modified by Morgan 2022/5/23 應同1605(通知年費逾期)設為自行判發(因無專用案件性質，故改傳1605抓判發人)--玫音
      'strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), m_strCP10, pa(9))
      'Modified by Morgan 2022/6/8 已有專用的增案件質
      'strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), "1605", pa(9))
      strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10), pa(9))
      'end 2022/6/8
      'end 2022/5/23
      'end 2022/5/23
      PUB_AddLetterProgress m_LD18, 2, True, strExc(1), True, pa(26), m_CP10, pa(75)
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
   m_strIR01 = frm05010406_1.m_strIR01
   m_strIR02 = frm05010406_1.m_strIR02
   m_strIR03 = frm05010406_1.m_strIR03
   m_strIR04 = frm05010406_1.m_strIR04
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
   Set frm05010406_2 = Nothing
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
   End If
End Sub

Private Sub Text1_Change(Index As Integer)
   If Index = 0 Then
      Select Case Text1(0)
         Case "1"
            m_CP10 = "1237" 'Added by Morgan 2022/5/31
            m_CP64 = "實體審查逾期補繳通知函"
            SetDate "416"
            
         Case "2"
            m_CP10 = "1609" 'Added by Morgan 2022/5/31
            m_CP64 = "領證逾期補繳通知函"
            SetDate "601"
            
         Case Else
            m_CP10 = "" 'Added by Morgan 2022/5/31
            m_CP64 = ""
      End Select
   End If
End Sub

Private Sub SetDate(p_NP07 As String)
   '原繳費期限
   strExc(0) = "select np09 from nextprogress where np02='" & pa(1) & "'" & _
      " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'" & _
      " and ( np06='N' or np06 is null) and np07='" & p_NP07 & "' order by np09 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Text1(2) = RsTemp(0)
   Else
      strExc(0) = "select cp07 from caseprogress where cp01='" & pa(1) & "'" & _
         " and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
         " and cp10='" & p_NP07 & "' and cp27 is null order by cp07 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Text1(2) = RsTemp(0)
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
      Case 0
         If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
            KeyAscii = 0
            Beep
         End If
         
      Case 1
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Function DupeCheck() As Boolean
   DupeCheck = True
   'Modified by Morgan 2022/5/31 增加檢查來函性質
   strExc(0) = "select cp10, cp27,cp64" & _
      " from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "'  and cp04='" & pa(4) & "' and cp57 is null" & _
      " and cp10 in ('1902','" & m_CP10 & "') and cp27>19221111 order by cp27 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modified by Morgan 2022/5/31
      'If RsTemp.Fields("cp10") = "1902" And InStr(RsTemp.Fields("cp64"), m_CP64) > 0 Then
      '   If MsgBox("本案件已於 " & ChangeWStringToTDateString("" & RsTemp.Fields(1)) & " 通知逾期補繳，是否再次通知？", vbYesNo + vbDefaultButton2) = vbNo Then
      '      DupeCheck = False
      '   End If
      'End If
      With RsTemp
      Do While Not .EOF
         If (.Fields("cp10") = "1902" And InStr("" & .Fields("cp64"), m_CP64) > 0) Or .Fields("cp10") = m_CP10 Then
            If MsgBox("本案件已於 " & ChangeWStringToTDateString(.Fields(1)) & " 通知逾期補繳，是否再次通知？", vbYesNo + vbDefaultButton2) = vbNo Then
               DupeCheck = False
            End If
            Exit Do
         End If
         .MoveNext
      Loop
      End With
      'end 2022/5/31
   End If
End Function

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Index = 2 Or Index = 3 Or Index = 4 Then
      If Text1(Index) <> "" Then
         If Not ChkDate(Text1(Index)) Then
            Cancel = True
         End If
      End If
   End If
End Sub
