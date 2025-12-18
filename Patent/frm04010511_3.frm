VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010511_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "消滅函／視為撤回輸入"
   ClientHeight    =   6060
   ClientLeft      =   108
   ClientTop       =   816
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9000
   Begin VB.TextBox Text27 
      Height          =   270
      Index           =   0
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4365
      Width           =   375
   End
   Begin VB.TextBox Text26 
      Height          =   270
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   5
      Top             =   2805
      Width           =   975
   End
   Begin VB.TextBox Text15 
      Height          =   270
      Index           =   1
      Left            =   5100
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2010
      Width           =   255
   End
   Begin VB.TextBox Text15 
      Height          =   270
      Index           =   0
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2010
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1608
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1620
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4512
      TabIndex        =   16
      Top             =   48
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   15
      Top             =   60
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   14
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   13
      Top             =   60
      Width           =   255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   12
      Top             =   60
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7464
      TabIndex        =   9
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6840
      TabIndex        =   8
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   2
      Left            =   8388
      TabIndex        =   10
      Top             =   0
      Width           =   600
   End
   Begin VB.TextBox Text15 
      Height          =   270
      Index           =   2
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2430
      Width           =   255
   End
   Begin MSForms.TextBox Text29 
      Height          =   1050
      Left            =   1320
      TabIndex        =   6
      Top             =   3240
      Width           =   7095
      VariousPropertyBits=   -1467987941
      ScrollBars      =   2
      Size            =   "12515;1852"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   11
      Top             =   330
      Width           =   5535
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9763;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label40 
      Caption         =   "(Y:閉卷)"
      Height          =   255
      Left            =   1875
      TabIndex        =   39
      Top             =   4365
      Width           =   855
   End
   Begin VB.Label Label39 
      Caption         =   "是否閉卷:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   4365
      Width           =   855
   End
   Begin VB.Label Label38 
      Caption         =   "消  滅  日:"
      Height          =   255
      Left            =   90
      TabIndex        =   37
      Top             =   2805
      Width           =   1215
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      Caption         =   "大陸消滅函類別:         (1:逾期未辦 2:未領證 3:未續繳年費 4:屆滿 5.未繳實審 6.被主張國內優先權 7.復審結案)"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   36
      Top             =   2445
      Width           =   8445
   End
   Begin VB.Label Label11 
      Caption         =   " 官方發文日期:"
      Height          =   252
      Left            =   120
      TabIndex        =   35
      Top             =   1608
      Width           =   1212
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3240
      TabIndex        =   34
      Top             =   1650
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label43 
      Caption         =   "進度備註:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   3270
      Width           =   855
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容:       (Y:是)"
      Height          =   180
      Index           =   2
      Left            =   3420
      TabIndex        =   32
      Top             =   2025
      Width           =   2445
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      Caption         =   "是否列印客戶通知函:             (N:不印)"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   31
      Top             =   2025
      Width           =   2895
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   8910
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否閉卷"
      Height          =   180
      Index           =   3
      Left            =   6300
      TabIndex        =   30
      Top             =   900
      Width           =   720
   End
   Begin VB.Label lblPA57 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Left            =   7260
      TabIndex        =   29
      Top             =   900
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(Y:閉卷)"
      Height          =   180
      Index           =   4
      Left            =   7980
      TabIndex        =   28
      Top             =   900
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Left            =   3960
      TabIndex        =   27
      Top             =   672
      Width           =   588
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   26
      Top             =   360
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3528
      TabIndex        =   25
      Top             =   96
      Width           =   768
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   60
      Width           =   768
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   648
      Width           =   768
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   1
      Left            =   1152
      TabIndex        =   22
      Top             =   672
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   2
      Left            =   5040
      TabIndex        =   21
      Top             =   648
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "收文號"
      Height          =   180
      Index           =   1
      Left            =   144
      TabIndex        =   20
      Top             =   888
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日"
      Height          =   180
      Index           =   2
      Left            =   3960
      TabIndex        =   19
      Top             =   912
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   5
      Left            =   1152
      TabIndex        =   18
      Top             =   912
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   6
      Left            =   5040
      TabIndex        =   17
      Top             =   912
      Width           =   480
   End
   Begin VB.Line Line1 
      DrawMode        =   16  'Merge Pen
      Index           =   1
      X1              =   120
      X2              =   8910
      Y1              =   1290
      Y2              =   1290
   End
End
Attribute VB_Name = "frm04010511_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Text29,Combo1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strTemp As String
Dim pa() As String, cp() As String
Dim intWhere As Integer, intLastRow As Integer
Dim m_NewCP09 As String
Dim m_HKNewCP09 As String, m_HKCloser As String 'Added by Morgan 2016/6/15
Public MPa9 As String

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim m_bln_FieldValid As Boolean 'False:欄位值無效, True:欄位值有效
Dim m_bolFMP As Boolean 'Add by Morgan 2010/1/4
Dim m_strHK(1 To 4) As String 'Add by Morgan 2010/6/7 香港案
Dim m_bolHKletterAsk As Boolean 'Added by Morgan 2012/5/25 詢問香港案是否出定稿
Dim m_strMA(1 To 4) As String 'add by sonia 2015/9/16 澳門案
Dim m_EndModCash As Boolean   'add by sonia 2018/12/24 是否已上過可結餘日
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/05/17 是否為寰華案
'Added by Morgan 2014/1/14
Public m_DocWord As String 'Added by Morgan 2014/4/17
Public m_DocNo As String
Public m_AppNo As String
'end 2014/1/14
Dim RC_cp10 As String 'Add by Lydia 2014/11/18 傳案件性質
Dim stSQL As String, rsQuery As New ADODB.Recordset
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END
Dim m_CP13 As String, m_CP12 As String 'Added by Morgan 2019/8/29
Dim m_bolNoCP27 As Boolean '不上發文 Added by Morgan 2020/1/17
Dim m_bolRegMail As Boolean '是否掛號 'Added by Morgan 2020/1/17
Dim m_bolFMPNoPrint As Boolean 'Added by Morgan 2023/4/11 FMP案是否列印中文定稿

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(1 To 5) As String
 Dim Jjj As Integer
   EndLetter ET01, m_NewCP09, ET03, strUserNum
   
   Jjj = 1
   If CheckStr(Text26.Text) <> "" Then
      'Modified by Morgan 2022/10/13 改存西元日期
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','專利權消滅日','" & DBDATE(Text26.Text) & "')"
      Jjj = Jjj + 1
   End If
   '2008/3/5 add by sonia
   If Text15(2) = "6" Then
      'Modified by Morgan 2018/6/12 被主張國內優先權須抓相同申請國家 Ex.P-116763 --玲玲
      strExc(0) = "select pa11 from pridate,patent where pd06='" & pa(11) & "'" & _
         " and pa01(+)=pd01 and pa02(+)=pd02 and pa03(+)=pd03 and pa04(+)=pd04 and pa09='" & pa(9) & "' and pa57 is null order by 1 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','主張案申請案號','" & RsTemp(0) & "')"
         Jjj = Jjj + 1
      End If
   End If
   '2008/3/5 end
   
   'Added by Morgan 2022/10/14
   strExc(0) = CompDate(1, 2, Text6) '官方發文日+2個月
   strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
      "','恢復權利法定期限','" & strExc(0) & "')"
   Jjj = Jjj + 1
   'end 2022/10/14
   
   'Added by Morgan 2022/10/17 抓核准函的年費起始年度
   'Modified by Morgan 2024/7/9 改抓最大收文號且為FMP案未領證才要(Ex:P124900,台灣且核准兩次,抓到兩筆而造成錯誤,)
   If m_bolFMP And Text15(2) = "2" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "select '" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','大陸年費年度',substr(max(cp09||cp53),9) cp53 from caseprogress,nextprogress" & _
         " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='1001' and np01(+)=cp09 and np07='601'"
      Jjj = Jjj + 1
   End If
   'end 2022/10/17
   
   'Added by Morgan 2023/6/21
   If Text15(2) = "1" Then
      'Modified by Morgan 2025/1/9 審查意見可能有多筆且為不同相關收文號(申請,復審)，改用本所號抓最後收文
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "select '" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','審查意見官方發文日',cp133 from caseprogress a" & _
         " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='1202'" & _
         " and not exists(select * from caseprogress b where cp01=a.cp01 and cp02=a.cp02 and cp03=a.cp03 and cp04=a.cp04 and cp10=a.cp10 and cp05>a.cp05)"
      Jjj = Jjj + 1
   End If
   'end 2023/6/21
      
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(Jjj - 1, strTxt) Then
   If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Removed by Morgan 2018/4/18
'Private Sub StartLetter1(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
'   Dim strTxt(1 To 5) As String
'   Dim Jjj As Integer
'   EndLetter ET01, ET02, ET03, strUserNum
'
'   Jjj = 1
'
'   strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'      "','大陸案本所案號','" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & "')"
'   Jjj = Jjj + 1
'
'   strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'      "','大陸案申請案號','" & pa(11) & "')"
'   Jjj = Jjj + 1
'
'   If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
'      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
'   End If
'End Sub

'Add by Morgan 2004/3/31
'非國內案消滅函類別選擇檢查
Private Function fnChoiceCheck() As Boolean
   
   '國內案不檢查
   If pa(9) = "000" Then fnChoiceCheck = True: Exit Function
   
On Error GoTo flgErr

   Select Case Text15(2).Text
      '選1時
      Case "1"
         '若有專用期間則確認
         If pa(24) <> "" Then
            If MsgBox("此案有專用期間，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
      '選2時
      Case "2"
         '1.若下一程序領證未上不續辦('N')則確認
         stSQL = "SELECT 1 FROM NEXTPROGRESS" & _
            " WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "'" & _
            " AND NP07='601' AND (NP06 IS NULL or NP06='Y')"
         If rsQuery.State = adStateOpen Then rsQuery.Close
         rsQuery.CursorLocation = adUseClient
         rsQuery.Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
         If rsQuery.RecordCount > 0 Then
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
         stSQL = "SELECT MAX(NP08||NVL(NP06,'X')) FROM NEXTPROGRESS" & _
            " WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "'" & _
            " AND NP07='605'"
         If rsQuery.State = adStateOpen Then rsQuery.Close
         rsQuery.CursorLocation = adUseClient
         rsQuery.Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
         If Not IsNull(rsQuery.Fields(0)) Then
            '2.若下一程序最大年費未上'N'則確認
            If Right("" & rsQuery.Fields(0), 1) = "X" Then
               If MsgBox("此案年費未結案，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
            End If
            '3.若下一程序最大年費期限>=系統日則確認
            If Left("" & rsQuery.Fields(0), 8) >= strSrvDate(1) Then
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
         If Val(pa(25)) >= Val(strSrvDate(2)) Then
            If MsgBox("此案未屆滿，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
      '2008/1/28 add by sonia 未繳實審
      '選5時
      Case "5"
         '1.若有專用期間則確認
         If pa(24) <> "" Then
            If MsgBox("此案有專用期間，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
         End If
         stSQL = "SELECT MAX(NP08||NVL(NP06,'X')) FROM NEXTPROGRESS" & _
            " WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "'" & _
            " AND NP07='416'"
         If rsQuery.State = adStateOpen Then rsQuery.Close
         rsQuery.CursorLocation = adUseClient
         rsQuery.Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
         If Not IsNull(rsQuery.Fields(0)) Then
            '2.若下一程序最大實審費未上'N'則確認
            If Right("" & rsQuery.Fields(0), 1) = "X" Then
               If MsgBox("此案實審費未結案，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
            End If
            '3.若下一程序最大實審費期限>=系統日則確認
            If Left("" & rsQuery.Fields(0), 8) >= strSrvDate(1) Then
               If MsgBox("此案實審費未逾期，是否要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then GoTo flgErr
            End If
         End If
      '2008/1/28 end
   End Select
   fnChoiceCheck = True
   
flgErr:
   Set rsQuery = Nothing
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub cmdok_Click(Index As Integer)
 Dim bolChk As Boolean, strTmp As String, strTmp2 As String
   strTmp = ""
   Select Case Index
      Case 0 '確定
'CANCEL BY SONIA 2014/12/30 智慧局函已無機關文號
'        'Add By Cheng 2003/03/26
'        '檢查機關文號
'        If pa(9) = 台灣國家代號 Then
'            If Me.Text9.Tag = Me.Text9.Text Then
'                MsgBox "請輸入機關文號!!!", vbExclamation + vbOKOnly
'                Me.Text9.SetFocus
'                Text9_GotFocus
'                Exit Sub
'            End If
'        End If
'END 2014/12/30
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         'Add By Sindy 2022/7/1
         'Mark by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail:可取消外專系統收件區，key來函承辦人掛程序人員，則按確定，信件會再打開一次的設定。
         'If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
         '   If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
          '     Exit Sub
          '  End If
         'End If
         ''2022/7/1 END
         'end 2023/05/17
         
         'Add by Morgan 2004/3/31
         If fnChoiceCheck = False Then Exit Sub
         'Add by Morgan 2005/5/20
         '非台灣 屆滿 詢問是否計算結餘
         '2005/9/27 MODIFY BY SONIA 已上可結餘日者不必做
         'If Text15(2) = "4" Then
         'modify by sonia 2018/12/24 不可判斷cp(109),因為新申請案之前都一定算過結餘
         'If Text15(2) = "4" And cp(109) = "" Then
         'Modify by Amy 2022/07/19 P案非台灣案收到專利權消滅函時由系統自動上結餘日期-雅娟
         'If Text15(2) = "4" And m_EndModCash = True Then
         If m_EndModCash = True Then
             'Modified by Lydia 2015/03/03 +pa01,pa02,pa03,pa04
            Pub_EndModCashMsg pa(9), pa(1), pa(2), pa(3), pa(4)
         End If
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         If Text15(0).Text <> "N" Then '通知函
            If Text15(1).Text = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            '91.8.26 modify by sonia
            Select Case pa(9)
               Case 台灣國家代號  '台灣 08
                     'add by toni  20080901 for 大對台定稿
                     If PUB_CheckCuNation(pa(26), Text2, Text3, Text4, Text5) = "1" Then
                           strTmp = "15"     '大-->台定稿
                     Else
                           strTmp = "08"
                     End If
                     
               Case 大陸國家代號  '大陸
                     Select Case Text15(2)
                        Case "1"               '陳述意見未辦撤回通知書 09
                           strTmp = "09"
                           'Added by Morgan 2023/6/21 寶齡富錦 Y55435 案件
                            If ChangeCustomerS(pa(75)) = "Y55435" Then
                              strTmp = "99"
                            End If
                            'end 2023/6/21
                        Case "2"               '放棄專利權(未領證) 10
                           strTmp = "10"
                        Case "3"               '專利權終止(未繳年費) 11
                           strTmp = "11"
                        Case "4"               '專利權終止(屆滿) 12
                           strTmp = "12"
                        Case "5"               '未繳實審撤回通知書 13
                           strTmp = "13"
                        '2008/3/5 add by sonia
                        Case "6"               '被主張國內優先權 14
                           strTmp = "14"
                        'add by sonia 2014/8/14
                        Case "7"               '復審結案 16
                           strTmp = "16"
                     End Select
               Case "013"         '香港 17
                     strTmp = "17"
                     'Added by Morgan 2024/11/8+無專用期定稿--品薇
                     If pa(24) = "" Then
                        strTmp = "18"
                     End If
                     'end 2024/11/8
            End Select
            StartLetter "07", strTmp
            If strTmp <> "" Then
               'Modify by Morgan 2010/1/4 FMP只要印1張
               If m_bolFMP Then
                  'Modified by Morgan 2023/4/11 +m_bolFMPNoPrint
                  NowPrint m_NewCP09, "07", strTmp, bolChk, strUserNum, , , , , 1, , , , , , , , m_NewCP09, , , , , m_bolFMPNoPrint
               Else
                  NowPrint m_NewCP09, "07", strTmp, bolChk, strUserNum, , , , , , , , , , , , , m_NewCP09
               End If
            End If
            
            'Added by Morgan 2022/10/13
            If m_bolFMP Then
               strTmp2 = ""
               Select Case Text15(2)
                  Case "1" '陳述意見未辦
                     If pa(9) = "020" Then
                        strTmp2 = "53" '大
                     End If
                  Case "2" '放棄專利權(未領證)
                     If pa(9) = "020" Then
                        strTmp2 = "55" '大
                     End If
                  Case "3" '專利權終止(未繳年費)
                     If pa(9) = "020" Then
                        strTmp2 = "54" '大
                     End If
                  Case "4" '專利權終止(屆滿)
                     If pa(9) = "020" Then
                        strTmp2 = "51" '大
                     Else
                        strTmp2 = "52" '港/澳
                     End If
               End Select
               
               If strTmp2 <> "" Then
                  strUserNum = strFMPNum
                  StartLetter "07", strTmp2
                  NowPrint m_NewCP09, "07", strTmp2, False, strUserNum
                  strUserNum = strUser1Num
               End If
            End If
            'end 2022/10/13
         End If
         'Add by Morgan 2010/6/7
         'modify by sonia 2014/5/12 下一程序不管有無掛期限均出指示信通知香港代理人無法辦理第二階段
         'If m_strHK(1) <> "" And m_bolHKletterAsk = True Then
         'Modified by Lydia 2017/04/07 有香港案的閉卷收文才執行
         'If m_strHK(1) <> "" Then
         If m_strHK(1) <> "" And m_HKNewCP09 <> "" Then
            'Removed by Morgan 2018/4/17 外專人員操作時不用彈維護畫面,改都不問--玲玲確認
            'strExc(1) = m_strHK(1) & "-" & m_strHK(2) & IIf(m_strHK(3) & m_strHK(4) = "000", "", "-" & m_strHK(3) & "-" & m_strHK(4))
            'strExc(2) = "香港案 " & strExc(1) & " 是否列印結案定稿？"
            'If MsgBox(strExc(2), vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
            'end 2018/4/17
               '抓最近發文收文號(代理人)
               If rsQuery.State = adStateOpen Then rsQuery.Close
               stSQL = "select cp09 from caseprogress where cp01='" & m_strHK(1) & "' and cp02='" & m_strHK(2) & "'"
               stSQL = stSQL & " and cp03='" & m_strHK(3) & "' and cp04='" & m_strHK(4) & "' and cp09<'C' and cp44 is not null and cp10<>'421'"
               rsQuery.CursorLocation = adUseClient
               rsQuery.Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
               If rsQuery.RecordCount > 0 Then
                  
                  'Modified by Morgan 2016/6/15
                  '指示信電子化
                  'Removed by Morgan 2018/4/17 香港結案固定由內專程序承辦
                  'StartLetter1 "14", "" & rsQuery.Fields(0), "41"
                  'If Left(Pub_StrUserSt03, 1) = "F" Then
                  '   NowPrint "" & rsQuery.Fields(0), "14", "41", False, strUserNum, 0
                  '   PUB_PrintLetter "" & rsQuery.Fields(0)
                  'Else
                  'end 2018/4/17
                  'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
                  'strExc(2) = Pub_GetSpecMan("PS4")
                  strExc(2) = PUB_GetLetterJudgeNew("2", m_strHK(1), "913", "013", , m_HKCloser)
                  PUB_AddAppForm m_HKNewCP09, True, strExc(2)
                     
                  'Added by Morgan 2018/4/17 閉卷承辦人非操作人員改從待處理區維護並上傳指示信(為與CFP共用改定稿狀態為42)
                  If m_HKCloser <> strUserNum Then
                     NowPrint "" & rsQuery.Fields(0), "14", "42", False, m_HKCloser, , , , , , , , , , , , , m_HKNewCP09
                  Else
                  'end 2018/4/17
                  
                     NowPrint "" & rsQuery.Fields(0), "14", "42", True, strUserNum, , , , , , , , , , , , , m_HKNewCP09
                     frm1105_1.m_RecNo = m_HKNewCP09
                     'Modified by Morgan 2016/9/8 指示信檔名應該是香港案案號
                     'frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & ".913.DATA.PDF"
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(m_strHK(1), m_strHK(2), m_strHK(3), m_strHK(4)) & ".913.DATA.PDF"
                     'end 2016/9/8
                     frm1105_1.Show
                  End If
                  'end 2016/6/15
               End If
               If rsQuery.State = adStateOpen Then rsQuery.Close
               Set rsQuery = Nothing
           'End If 'Removed by Morgan 2018/4/17
         End If
         'Add by Lydia 2014/11/18 台灣案主管機關來函輸入，若此案有工程師未發文的程序，發E-MAIL通知工程師收到來函的內容
         'Modified by Lydia 2022/08/15 開放P大陸案
         'If pa(9) = "000" And pa(1) = "P" Then
         'Modified by Lydia 2022/10/11 經查此設定並不適用於外專及日專，故請協助排除FMP案
         'If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" Then
         If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" And m_bolFMP = False Then
            'Modified by Lydia 2022/08/16 +申請國家
            'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), RC_cp10, m_NewCP09
            PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), RC_cp10, pa(9), m_NewCP09
         End If
         
         'Add By Sindy 2016/10/5
         If Me.m_strIR01 <> "" Then
            Unload frm04010511_1
            Unload frm04010511_2
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         ElseIf Me.m_DocNo <> "" Then
         'Added by Morgan 2014/1/14
         'If Me.m_DocNo <> "" Then
         '2016/10/5 END
            Unload frm04010511_1
            Unload frm04010511_2
            Unload Me
            frm04010516.GoNext
         Else
         'end 2014/1/14
            frm04010511_1.Show
            frm04010511_1.Clear
            Unload frm04010511_2
            Unload Me
         End If 'Added by Morgan 2014/1/14
         
      Case 1
         frm04010511_2.Show
         Unload Me
      Case 2
         Unload frm04010511_1
         Unload frm04010511_2
         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean
Dim intStep As Integer, strTxt(1 To 10) As String, strTmp As String, i As Integer
'edit by nickc 2007/02/02
'Dim Ncp(1 To T_CP) As String
Dim Ncp() As String
ReDim Ncp(1 To TF_CP) As String
Dim strProgressNo As String '(暫存列印接洽結案單下一程序序號)
Dim BlnCheck As Boolean '判斷是否有勾選本案期限
Dim strDate1 As String '本所期限
Dim StrDate2 As String '法定期限
'add by sonia 2015/3/18
'Dim Newcp09 As String '結案指示信要判發改全域變數 m_HKNewCP09 Morgan 2016/6/15
Dim str111np01 As String
Dim str111np22 As String
'end 2015/3/18

On Error GoTo ErrorHandler
'FormSave = True
FormSave = False
cnnConnection.BeginTrans

   intStep = 1
   
   '1
      Ncp(1) = cp(1)
      Ncp(2) = cp(2)
      Ncp(3) = cp(3)
      Ncp(4) = cp(4)
      Ncp(5) = Label3(6)
      'Ncp(8) = Text9   'CANCEL BY SONIA 2014/12/30 智慧局函已無機關文號
      If Text9.Visible = True Then Ncp(8) = Text9 'Added by Morgan 2016/1/22 目前偶有電子公文
      'Modify by Morgan 2011/2/24 修正百年收文號問題
      'Ncp(9) = "C" & Left(strSrvDate(2), 2)
      Ncp(9) = "C" & CompAutoNumberYear(GetTaiwanThisYear)
      
      '案件性質
      'Added by Lydia 2014/01/17 1.逾期未辦、5未繳實體審查、6被主張國內優先權與7復審結案時，請直接帶視為撤回1610
      If pa(9) <> "000" And InStr("1,5,6,7", Me.Text15(2)) > 0 Then
         Ncp(10) = "1610"
      Else
      'end 2024/01/17
         Ncp(10) = "1604"
      End If
      RC_cp10 = Ncp(10)
      '2009/12/30 MODIFY BY SONIA
      'Ncp(12) = cp(12)
      'Ncp(13) = cp(13)
      'Modified by Morgan 2019/8/29
      'Ncp(13) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
      'Ncp(12) = GetSalesArea(Ncp(13))
      Ncp(13) = m_CP13
      Ncp(12) = m_CP12
      'end 2019/8/29
      '2009/12/30 END
      
      '承辦人
      Ncp(14) = strUserNum
      Ncp(26) = "N"
      If "1604" = 專利權消滅 Then Ncp(25) = TransDate(Text26, 1)
      '發文日
      'Added by Morgan 2020/1/17
      If m_bolNoCP27 = True Then
         Ncp(27) = ""
      Else
      'end 2020/1/17
         Ncp(27) = strSrvDate(2)
      End If
      Ncp(32) = "N"
      Ncp(43) = cp(9)
      Ncp(64) = Text29
      Ncp(133) = DBDATE(Text6) 'Add by Morgan 2010/1/8
      Ncp(119) = DBDATE(Label3(6)) 'Added by Morgan 2012/4/30 +cp119=櫃檯收文日
      
      If m_bolFMP Then
         'Modified by Lydia 2024/01/17 改成變數"1604"=>Ncp(10)
         Ncp(20) = PUB_GetCP20(pa(1), Ncp(10), , pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
      End If
      
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.SaveNewCaseProgressDatabase("C", Ncp, intWhere) Then
      If Not ClsPDSaveNewCaseProgressDatabase("C", Ncp, intWhere) Then
        'Modify By Cheng 2002/11/06
'         Exit Function
        GoTo ErrorHandler
      End If
      '暫存新的收文號
      m_NewCP09 = Ncp(9)
      
      
      PUB_DualCaseInform Ncp(9) 'Added by Morgan 2022/4/7
      
      'Added by Morgan 2014/1/14
      If m_DocNo <> "" Then
         PUB_UpdateEdocRec m_DocNo, Ncp(9), pa(1), pa(2), pa(3), pa(4), Ncp(10)
      End If
      'end 2014/1/14
      
      'Added by Morgan 2014/4/11 電子化-新增信函進度檔
      If pa(9) = "000" Then
         strExc(1) = ""
         If Text15(0) <> "N" Then
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), Ncp(10), , , pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), Ncp(10))
         End If
         'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
         'Modified by Morgan 2020/1/17
         'PUB_AddLetterProgress Ncp(9), 1, IIf(Text15(0) <> "N", True, False), strExc(1), False, pa(26), Ncp(10), pa(75)
         PUB_AddLetterProgress Ncp(9), 1, IIf(Text15(0) <> "N", True, False), strExc(1), m_bolRegMail, pa(26), Ncp(10), pa(75)
      'Added by Morgan 2016/6/8
      ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(pa(1), Ncp(10), , pa(9), pa(1), pa(2), pa(3), pa(4))
         strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), Ncp(10), pa(9), , , m_bolFMP)
         'Modified by Morgan 2016/7/5 除屆滿外都掛號
         'Modified by Morgan 2016/12/20 香港目前只有未續繳年費定稿無期限不可補繳(不直寄)
         'PUB_AddLetterProgress Ncp(9), 2, IIf(Text15(0) <> "N", True, False), strExc(1), IIf(Text15(2) = "4", False, True), pa(26), Ncp(10), pa(75)
         'Modified by Morgan 2020/1/17
         'PUB_AddLetterProgress Ncp(9), 2, IIf(Text15(0) <> "N", True, False), strExc(1), IIf(Text15(2) = "4" Or pa(9) = "013", False, True), pa(26), Ncp(10), pa(75)
         PUB_AddLetterProgress Ncp(9), 2, IIf(Text15(0) <> "N", True, False), strExc(1), m_bolRegMail, pa(26), Ncp(10), pa(75)
      'end 2016/6/8
      End If
      'end 2014/4/11
   
   '2
   If Text27(0) = "Y" And lblPA57 = "" Then
      '2005/9/15 MODIFY BY SONIA 未閉卷者才更新
      'strTxt(intStep) = "UPDATE PATENT SET PA57='Y',PA58=" & TransDate(Label3(6), 2) & ", PA59='89' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      strTxt(intStep) = "UPDATE PATENT SET PA57='Y',PA58=" & TransDate(Label3(6), 2) & ", PA59='89' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4)) & " AND PA57 IS NULL "
      '2005/9/15 END
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   '92.1.8 ADD BY SONIA
   strTxt(intStep) = "UPDATE PATENT SET PA17='N' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   '92.1.8 END

   'Add by Morgan 2005/5/20
   '非台灣 屆滿 更新結餘
   '2005/9/27 MODIFY BY SONIA 已上可結餘日者不必再做
   'If Text15(2) = "4" Then
   'modify by sonia 2018/12/24 不可判斷cp(109),因為新申請案之前都一定算過結餘
   'If Text15(2) = "4" And cp(109) = "" Then
   'Modify by Amy 2022/07/19 P案非台灣案收到專利權消滅函時由系統自動上結餘日期-雅娟
   'If Text15(2) = "4" And m_EndModCash = True Then
   If m_EndModCash = True Then
      Pub_UpdateEndModCash cp(1), cp(2), cp(3), cp(4)
   End If
   
   'Add by Morgan 2007/10/24 'Memo by Lydia 2024/01/17 大陸案沒有928重新委任
   strSql = "Update CaseProgress Set CP27=19221111,CP64=CP64||'" & ChangeTStringToTDateString(Label3(6)) & "'||'已收消滅函;' Where CP01='" & cp(1) & "' and CP02='" & cp(2) & "' and CP03='" & cp(3) & "' and CP04='" & cp(4) & "' and CP10='928' and CP27 is null and CP57 is null"
   cnnConnection.Execute strSql, intI
   
   'Add by Morgan 2010/6/7
   m_strHK(1) = ""
   m_bolHKletterAsk = False
   '若有香港案且第二階段未發文則email通知智權人員銷案
   'Modified by Morgan 2012/5/25
   '大陸案可能自撤,改判斷有已收未發程序時通知智權人員銷案,若無但有智權人員期限者通知結案,其他則通知智權人員系統自動結案(閉卷)
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
      strExc(3) = "" 'Added by Lydia 2017/04/07
      'Modified by Morgan 2014/10/22 香港案要抓大陸案件性質名稱
      strExc(0) = "select cpm04,cp10 from caseprogress,casepropertymap" & _
         " where cp01='" & m_strHK(1) & "' and cp02='" & m_strHK(2) & "' and cp03='" & m_strHK(3) & "' and cp04='" & m_strHK(4) & "'" & _
         " and cp09<'B' and cp27||cp57 is null and cpm01(+)=cp01 and cpm02(+)=cp10"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp("cp10") = "111" Then m_bolHKletterAsk = True
         strExc(3) = "大陸案" & strExc(1) & "已消滅,香港案" & strExc(2) & "(" & RsTemp("cpm04") & ")將無法辦理，請銷案。"
      Else
        'Added by Lydia 2017/04/07 要判斷香港第二階段111未收文(EX.香港案P98784,大陸案P95625)
        strExc(0) = "select cp10 from caseprogress where cp01='" & m_strHK(1) & "' and cp02='" & m_strHK(2) & "' and cp03='" & m_strHK(3) & "' and cp04='" & m_strHK(4) & "'" & _
           " and cp09<'B' and cp10 = '111' and cp159=0 "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 0 Then
        'end 2017/04/07 判斷香港第二階段111未收文
       'Add by Lydia 2014/10/27 大陸案專利權消滅時若有香港案關聯時，凡香港第二階段未收文不論下一階段是否有期限，系統將自動結案
'         strExc(0) = "select cpm04,np07 from nextprogress,casepropertymap" & _
'            " where np02='" & m_strHK(1) & "' and np03='" & m_strHK(2) & "' and np04='" & m_strHK(3) & "' and np05='" & m_strHK(4) & "'" & _
'            " and np06 is null and np07 is not null and cpm01(+)=np02 and cpm02(+)=np07" & strNpSqlOfNoSalesDuty
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If RsTemp("np07") = "111" Then m_bolHKletterAsk = True
'            strExc(3) = "大陸案" & strExc(1) & "已消滅,香港案" & strExc(2) & "(" & RsTemp("cpm04") & "將無法續行，請結案。"
'         Else
              strExc(3) = "大陸案" & strExc(1) & "已消滅,香港案" & strExc(2) & "已自動結案。"
              strSql = "update patent set PA57='Y',PA58=" & DBDATE(Label3(6)) & ", PA59='99',pa91=pa91||';相關大陸案" & strExc(1) & "已消滅本案自動上閉卷(" & DBDATE(Label3(6)) & ")。' where pa01='" & m_strHK(1) & "' and pa02='" & m_strHK(2) & "' and pa03='" & m_strHK(3) & "' and pa04='" & m_strHK(4) & "'"
              cnnConnection.Execute strSql, intI
'         End If

        'end. Add by Lydia 2014/10/27 大陸案專利權消滅時若有香港案關聯時，凡香港第二階段未收文不論下一階段是否有期限，系統將自動結案
         'add by sonia 2015/3/18 香港案產生一筆閉卷的進度 P-103232(大陸案P-101729)
         '先抓香港案下一程序之標準專利批准記錄請求未收文期限的總收文號
         str111np01 = "": str111np22 = ""
         If rsQuery.State = adStateOpen Then rsQuery.Close
         stSQL = "SELECT NP01,NP22 FROM NEXTPROGRESS WHERE NP02='" & m_strHK(1) & "' AND NP03='" & m_strHK(2) & "' AND NP04='" & m_strHK(3) & "' AND NP05='" & m_strHK(4) & "' AND NP07='111' AND NP06 IS NULL"
         rsQuery.CursorLocation = adUseClient
         rsQuery.Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
         If rsQuery.RecordCount > 0 Then
            str111np01 = "" & rsQuery.Fields(0)
            str111np22 = "" & rsQuery.Fields(1)
         End If
         If rsQuery.State = adStateOpen Then rsQuery.Close
         Set rsQuery = Nothing
         '同時更新下一程序之標準專利批准記錄請求未收文期限的續辦欄及解除期限日期,原因
         If str111np01 <> "" Then
            strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & strSrvDate(1) & ",NP12='99' WHERE NP01='" & str111np01 & "' AND NP22='" & str111np22 & "' AND NP02='" & m_strHK(1) & "' AND NP03='" & m_strHK(2) & "' AND NP04='" & m_strHK(3) & "' AND NP05='" & m_strHK(4) & "'"
            cnnConnection.Execute strSql, intI
         End If
         '產生閉卷進度
         m_HKNewCP09 = AutoNo("B", 6)
         'Modified by Morgan 2018/4/17 香港案都是內專處理(對香港代理人)，承辦人改掛程序2(原來掛操作人員)--玲玲
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
            ",'" & m_HKCloser & "','N','N'," & strSrvDate(1) & ",'N','" & str111np01 & "','" & str111np22 & "'," & strSrvDate(1) & ",'99','相關大陸案" & strExc(1) & "已消滅本案自動上閉卷(" & DBDATE(Label3(6)) & ")。') "
         'end 2018/4/17
         cnnConnection.Execute strSql, intI
         'end 2015/3/18
         
         'Added by Morgan 2018/4/17 EMail通知閉卷承辦人(同操作人時也要寄)--玲玲
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " values('" & strUserNum & "','" & m_HKCloser & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'香港案" & strExc(2) & "已閉卷請至待處理區上傳指示信','如旨')"
         cnnConnection.Execute strSql, intI
         'end 2018/4/17
        End If 'end 2017/04/07 判斷香港第二階段111未收文
      End If
      
      If strExc(3) <> "" Then 'Added by Lydia 2017/04/07 判斷有對香港案的處理,才發信
        strExc(4) = PUB_GetAKindSalesNo(m_strHK(1), m_strHK(2), m_strHK(3), m_strHK(4))
        strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
           " values('" & strUserNum & "','" & strExc(4) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
           ",'" & strExc(3) & "','如旨')"
        cnnConnection.Execute strSql, intI
      End If 'end 2017/04/07
   End If
   'end 2010/6/7
  
   'add by sonia 2015/9/16
   m_strMA(1) = ""
   '大陸案若有澳門發明案且未發文則email通知智權人員銷案
   strExc(0) = "select cm01,cm02,cm03,cm04,cpm04 from casemap,patent,caseprogress,casepropertymap" & _
      " where cm10='5' and cm05='" & pa(1) & "' and cm06='" & pa(2) & "'" & _
      " and cm07='" & pa(3) & "' and cm08='" & pa(4) & "'" & _
      " and pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 and pa09='044' and pa57 is null" & _
      " and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp09<'B' and cp27||cp57 is null and cp10='101'" & _
      " and cp01=cpm01(+) and cp10=cpm02(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_strMA(1) = RsTemp("cm01")
      m_strMA(2) = RsTemp("cm02")
      m_strMA(3) = RsTemp("cm03")
      m_strMA(4) = RsTemp("cm04")
      
      strExc(1) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
      strExc(2) = m_strMA(1) & "-" & m_strMA(2) & IIf(m_strMA(3) & m_strMA(4) = "000", "", "-" & m_strMA(3) & "-" & m_strMA(4))
      strExc(3) = "大陸案" & strExc(1) & "已消滅,澳門案" & strExc(2) & "(" & RsTemp("cpm04") & ")將無法辦理，請銷案。"
      strExc(4) = PUB_GetAKindSalesNo(m_strMA(1), m_strMA(2), m_strMA(3), m_strMA(4))
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
         " values('" & strUserNum & "','" & strExc(4) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & strExc(3) & "','如旨')"
      cnnConnection.Execute strSql, intI
   End If
   'end 2015/9/16
   
   'Add by Sindy 2016/10/5
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", m_NewCP09, "")
      'Modified by Lydia 2023/05/18 +不開啟附件, , , False
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010511_1", IIf(Pub_StrUserSt03 = "F22", m_NewCP09, ""), , , False
   End If
   '2016/10/5 END
   
   'Added by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail
   'Modified by Lydia 2023/05/26 已閉卷不通知
   'Move by Lydia 2023/05/26 從commit上方移過來,
   Dim bolFMP2mail As Boolean  'Added by Lydia 2023/05/26
   If m_bolFMP = True And m_bolFMP2 = True And pa(57) = "" Then
       'Modified by Lydia 2023/10/31 傳入C類收文號 m_NewCP09
       bolFMP2mail = Pub_SetFMP2toCMail(pa(1), pa(2), pa(3), pa(4), Ncp(10), strUserNum, m_NewCP09)
   End If
   'end 2023/05/17
   
   'Added by Morgan 2023/4/11
   'FMP案EMail通知承辦(智權人員)
   m_bolFMPNoPrint = False
   'Modified by Lydia 2023/05/26 排除-寰華案無期限之官方來函，系統自動發Mail => And bolFMP2mail = False
   If m_bolFMP And Left(Pub_StrUserSt03, 1) <> "F" And bolFMP2mail = False Then
         'Modified by Lydia 2024/01/17
         'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " select '" & strUserNum & "',cp13,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'FMP案'||m1.cpm04||'通知:'||c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04)||'案專利權已消滅，請參考卷宗區！','如旨',st52" & _
            " from caseprogress c1,casepropertymap m1,staff " & _
            " where c1.cp09='" & m_NewCP09 & "' and m1.cpm01(+)=c1.cp01 and m1.cpm02(+)=c1.cp10" & _
            " and st01(+)=cp13"
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " select '" & strUserNum & "',cp13,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'FMP案'||m1.cpm04||'通知:'||c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04)||'案" & IIf(Ncp(10) = "1610", "視為撤回", "專利權已消滅") & "，請參考卷宗區！','如旨',st52" & _
            " from caseprogress c1,casepropertymap m1,staff " & _
            " where c1.cp09='" & m_NewCP09 & "' and m1.cpm01(+)=c1.cp01 and m1.cpm02(+)=c1.cp10" & _
            " and st01(+)=cp13"
      cnnConnection.Execute strSql, intI
      m_bolFMPNoPrint = True
   End If
   'end 2023/4/11
   
   cnnConnection.CommitTrans
    FormSave = True
   
   '(列印接洽結案單)
   If FormSave = True Then
      If IsEmptyText(strProgressNo) = False Then
         g_PrtForm001.PrintForm strProgressNo, pa(1), pa(2), pa(3), pa(4)
      End If
   End If
Exit Function
ErrorHandler:
    If FormSave = False Then cnnConnection.RollbackTrans
End Function

Private Sub Form_Initialize()
   ReDim pa(TF_PA)
   ReDim cp(TF_CP)
End Sub

Private Sub Form_Load()
 Dim ret As Long
   
   MoveFormToCenter Me
   intWhere = 國內
   With frm04010511_2
      pa(1) = .Text2
      pa(2) = .Text3
      pa(3) = .Text4
      pa(4) = .Text5
      strReceiveNo = .Tag
      ReadPatent
   End With
   Combo1.ListIndex = 0
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010511_2.m_strIR01
   m_strIR02 = frm04010511_2.m_strIR02
   m_strIR03 = frm04010511_2.m_strIR03
   m_strIR04 = frm04010511_2.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   If pa(9) = 台灣國家代號 Then
      Text6 = Label3(6)
   Else
      'modify by Morgan 2010/1/8 來函日期改為 官方發文日期
      'Text6 = strSrvDate(2)
      Text6.MaxLength = 8
      'Text6 = strSrvDate(1)  '2010/4/8 CANCEL BY SONIA
   End If
   If pa(9) = 台灣國家代號 Then
      Me.Text26.Enabled = False
      Me.Text27(0).Enabled = False
      SendKeys "{Tab}"
'CANCEL BY SONIA 2014/12/30 智慧局函已無機關文號
'   Else
'      Me.Text9.Enabled = False
'END 2014/12/30
   End If
   
    'Remove by Morgan 2010/1/8 來函日期改為 官方發文日期
    ''Add By Cheng 2003/04/02
    ''若申請國家為大陸
    'If pa(9) = 大陸國家代號 Then
    '    '游標停在大陸專利權消滅函類別欄位
    '    SendKeys "{Tab}{Tab}{Tab}"
    'End If
    
    'Modified by Morgan 2019/8/29 業務員,業務區要用最新收文的判斷 Ex:P-103938
    'If Left(cp(12), 1) = "F" And pa(10) <> "000" Then
    m_CP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
    m_CP12 = GetSalesArea(m_CP13)
    If Left(m_CP12, 1) = "F" And pa(9) <> "000" Then
    'end 2019/8/29
      m_bolFMP = True
    Else
      m_bolFMP = False
    End If
    'Added by Lydia 2023/05/17 判斷寰華案
    m_bolFMP2 = False
    If m_bolFMP = True Then
       If PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4)) = True Then
          m_bolFMP2 = True
       End If
    End If
    'end 2023/05/17
   
   'Added by Morgan 2014/4/11 電子化-台灣案定稿要轉pdf故修改只能從定稿維護作業
   If pa(9) = "000" Then
      Text15(1).Enabled = False
   'Added by Morgan 2016/6/15  非臺灣案電子化
   ElseIf (內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F") Then
      Text15(1).Enabled = False
   'end 2016/6/15
   End If
   'end 2014/4/11
   
End Sub

Private Sub ReadPatent()
 Dim Lbl As LABEL, i As Integer
 Dim strTmp As String, rsTemp1 As New ADODB.Recordset, bolTmp As Boolean
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   
   Me.lblPA57.Caption = ""
   MPa9 = ""
   
   Label3(6) = frm04010511_1.Text5
   Label3(5) = strReceiveNo
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   
   If pa(1) = "P" Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         '申請日
         Label3(2) = pa(10)
         '申請案號
         Text1 = pa(11)
         '是否閉卷
         Text27(0) = pa(57)
         Me.lblPA57.Caption = pa(57)
         '申請國家
         MPa9 = pa(9)
      End If
      
   Else
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         Label3(2) = pa(10)
         Text1 = pa(11)
         Text27(0) = pa(15)
         Me.lblPA57.Caption = pa(15)
         MPa9 = pa(9)
      End If
   End If
   
   If pa(9) = 台灣國家代號 Then
      strTmp = "CPM03"
   Else
      strTmp = "CPM04"
   End If
   
   ' (機關文號設預設內容)
   If Len(strSrvDate(2)) = 6 Then
      strTemp = Left(strSrvDate(2), 2)
   Else
      strTemp = Left(strSrvDate(2), 3)
   End If
'CANCEL BY SONIA 2014/12/30 智慧局函已無機關文號
'   If pa(9) = 台灣國家代號 Then
'      Text9.Text = "（" & strTemp & "）智專一(一)權字第號"
'        'Add By Cheng 2003/03/26
'        '記錄機關文號的預設值
'        Me.Text9.Tag = Me.Text9.Text
'   End If
'
'Modified by Morgan 2016/1/22 還原,目前偶有電子公文會有機關文號
If m_DocNo <> "" Then
   Text9.Visible = True
   Label14.Visible = True
   
   'Added by Morgan 2014/1/14
   'Modified by Morgan 2014/4/17 +發文字
   If m_DocWord <> "" Then
      Text9 = m_DocWord & "字第" & m_DocNo & "號"
   ElseIf m_DocNo <> "" Then
      Text9 = Replace(Text9, "第號", "第" & m_DocNo & "號")
   End If
   'end 2014/1/14
End If
'end 2016/1/22
'END 2014/12/30
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(10) <> "" Then
         If pa(9) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(cp(1), cp(10), strExc(0), BolTmp) Then Label3(1) = strExc(0)
         If ClsPDGetCaseProperty(cp(1), cp(10), strExc(0), bolTmp) Then Label3(1) = strExc(0)
      End If

   End If
   '2005/9/21 ADD BY SONIA 全部預設閉卷
   If Text27(0) = "" Then Text27(0) = "Y"
   '2005/9/21 END
   
   'add by sonia 2018/12/24 判斷此案是否已上過可結餘日
   m_EndModCash = False
   'modify by sonia 2020/8/26 FMP案無CP16但有請款單的也要列入FCP-058979變更
   'strExc(0) = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
              " AND CP16>0 AND NVL(CP109,0)=0"
   strExc(0) = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
              " AND (CP16>0 or CP60||CP61 is not null) AND NVL(CP109,0)=0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then m_EndModCash = True
   'end 2018/12/24
End Sub

Private Function ChgType(i As Integer, Optional SstrKind As Integer) As Boolean
 Dim strTempName As String, bolTmp As Boolean
   ChgType = False
   If pa(9) = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   Select Case i
      Case 8
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(pA(1), Format(SstrKind), strTempName, BolTmp) Then
         If ClsPDGetCaseProperty(pa(1), Format(SstrKind), strTempName, bolTmp) Then
            Label3(3) = strTempName
            ChgType = True
         Else
            Label3(3) = ""
         End If
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2010/6/7
   Set frm04010511_3 = Nothing 'Removed by Morgan 2021/12/20 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub Text15_GotFocus(Index As Integer)
  TextInverse Text15(Index)
End Sub

Private Sub Text15_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
   Case 0 '是否列印客戶通知函
      If KeyAscii <> 78 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 1 '是否修改通知函內容
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 2 '大陸消滅函類別  'modify by sonia 2014/8/14 加入7復審結案
      If (KeyAscii < 49 Or KeyAscii > 55) And KeyAscii <> 8 Then
         KeyAscii = 0
      '2005/9/15 ADD BY SONIA
      Else
         If pa(9) = 台灣國家代號 Then   'ADD BY SONIA 2015/10/30 非台灣案之專利權消滅改為消滅
            Text29 = Replace(Text29, ",專利權消滅-逾期未辦", "")
            Text29 = Replace(Text29, ",專利權消滅-未領證", "")
            Text29 = Replace(Text29, ",專利權消滅-未續繳年費", "")
            Text29 = Replace(Text29, ",專利權消滅-屆滿", "")
            Text29 = Replace(Text29, ",專利權消滅-未繳實審", "")
            Text29 = Replace(Text29, ",專利權消滅-被主張國內優先權", "")
            Text29 = Replace(Text29, ",專利權消滅-復審結案", "")
            Select Case KeyAscii
               Case 49
                  Text29 = Text29 & ",專利權消滅-逾期未辦"
               Case 50
                  Text29 = Text29 & ",專利權消滅-未領證"
               Case 51
                  Text29 = Text29 & ",專利權消滅-未續繳年費"
               Case 52
                  Text29 = Text29 & ",專利權消滅-屆滿"
               Case 53
                  Text29 = Text29 & ",專利權消滅-未繳實審"
               Case 54
                  Text29 = Text29 & ",專利權消滅-被主張國內優先權"
               Case 55
                  Text29 = Text29 & ",專利權消滅-復審結案"
            End Select
         'ADD BY SONIA 2015/10/30 非台灣案之專利權消滅改為消滅
         Else
            'Modified by Lydia 2024/01/17
            'Text29 = Replace(Text29, ",消滅-逾期未辦", "")
            Text29 = Replace(Text29, ",視為撤回-逾期未辦", "")
            Text29 = Replace(Text29, ",消滅-未領證", "")
            Text29 = Replace(Text29, ",消滅-未續繳年費", "")
            Text29 = Replace(Text29, ",消滅-屆滿", "")
            'Modified by Lydia 2024/01/17
            'Text29 = Replace(Text29, ",消滅-未繳實審", "")
            'Text29 = Replace(Text29, ",消滅-被主張國內優先權", "")
            'Text29 = Replace(Text29, ",消滅-復審結案", "")
            Text29 = Replace(Text29, ",視為撤回-未繳實審", "")
            Text29 = Replace(Text29, ",視為撤回-被主張國內優先權", "")
            Text29 = Replace(Text29, ",視為撤回-復審結案", "")
            'end 2024/01/17
            Select Case KeyAscii
               Case 49
                  'Modified by Lydia 2024/01/17
                  'Text29 = Text29 & ",消滅-逾期未辦"
                  Text29 = Text29 & ",視為撤回-逾期未辦"
               Case 50
                  Text29 = Text29 & ",消滅-未領證"
               Case 51
                  Text29 = Text29 & ",消滅-未續繳年費"
               Case 52
                  Text29 = Text29 & ",消滅-屆滿"
               Case 53
                  'Modified by Lydia 2024/01/17
                  'Text29 = Text29 & ",消滅-未繳實審"
                  Text29 = Text29 & ",視為撤回-未繳實審"
               Case 54
                  'Modified by Lydia 2024/01/17
                  'Text29 = Text29 & ",消滅-被主張國內優先權"
                  Text29 = Text29 & ",視為撤回-被主張國內優先權"
               Case 55
                  'Modified by Lydia 2024/01/17
                  'Text29 = Text29 & ",消滅-復審結案"
                  Text29 = Text29 & ",視為撤回-復審結案"
            End Select
         End If
         'END 2015/10/30
      '2005/9/15 END
      End If
   Case 3 '國內其他來函類別
      If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End Select
   
End Sub

Private Sub Text15_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If Text15(Index) <> "" And Text15(Index) <> "N" Then
            MsgBox "是否列印客戶通知函，只可為空白或 N !", vbCritical
            Cancel = True
         End If
      Case 1
         If Text15(Index) <> "" And Text15(Index) <> "Y" Then
            MsgBox "是否修改通知函內容，只可為空白或 Y !", vbCritical
            Cancel = True
         End If
      Case 2
        '若申請國家非台灣者才要檢查
        If pa(9) <> 台灣國家代號 Then
            'modify by sonia 2014/8/14 加入7復審結案
            If Text15(Index) <> "1" And Text15(Index) <> "2" And Text15(Index) <> "3" And Text15(Index) <> "4" And Text15(Index) <> "5" And Text15(Index) <> "6" And Text15(Index) <> "7" Then
               MsgBox "大陸消滅函類別，只可為 1 ~ 7 !", vbCritical
               Cancel = True
            End If
            'add by sonia 2014/8/14 選7復審結案必須要有復審申請107那一道,且存檔時消滅函的相關總收文號要放復審申請
            If Text15(Index) = "7" Then
               strExc(0) = "SELECT MAX(CP09) FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                          " AND CP10='107' AND NVL(CP27,0)>0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If "" & RsTemp(0) = "" Then
                     MsgBox "消滅函類別選 7 復審結案, 必須要有復審申請的程序 !", vbCritical
                     Cancel = True
                  Else
                     cp(9) = RsTemp(0)
                  End If
               Else
                  MsgBox "消滅函類別選 7 復審結案, 必須要有復審申請的程序 !", vbCritical
                  Cancel = True
               End If
            End If
            'end 2014/8/14
        End If
   End Select
   If Cancel = True Then TextInverse Text15(Index)
End Sub

Private Sub Text26_GotFocus()
  TextInverse Text26
End Sub

Private Sub Text26_Validate(Cancel As Boolean)
   If Text26 = "" Then
      '若大陸專利權消滅函類別是2.未領證時專利權消滅日可空白
      '92.2.16 modify by sonia 1.逾期未辦也可不輸
      '2008/1/28 modify by sonia 5.未繳實審也可不輸
      '2014/8/14 modify by sonia 7.復審結案也可不輸
      If Me.Text15(2).Text <> "2" And Me.Text15(2).Text <> "1" And Me.Text15(2).Text <> "5" And Me.Text15(2).Text <> "6" And Me.Text15(2).Text <> "7" Then
         MsgBox "來函性質為消滅時，消滅日不可空白 !", vbCritical
         TextInverse Text26
         Cancel = True
      End If
   Else
      'Modify by Morgan 2005/4/29 大陸檢查西元年
'      If Not ChkDate(Text26) Then
'         MsgBox "日期不正確，請重新輸入 !", vbCritical
'         Cancel = True
'      End If
      If pa(9) <> "000" Then
         If Not CheckIsDate(Text26) Then
            TextInverse Text26
            Cancel = True
         '2005/9/15 ADD BY SONIA
         Else
            If Text26 >= strSrvDate(1) Then
               MsgBox "消滅日不可大於或等於系統日期 !", vbCritical
               TextInverse Text26
               Cancel = True
            End If
         '2005/9/15 END
         End If
      Else
         If Not CheckIsTaiwanDate(Text26) Then
            TextInverse Text26
            Cancel = True
         '2005/9/15 ADD BY SONIA
         Else
            'Modify by Morgan 2010/8/11 百年蟲
            'If Text26 >= ServerDate - 19110000 Then
            If Val(Text26) >= Val(strSrvDate(2)) Then
               MsgBox "消滅日不可大於或等於系統日期 !", vbCritical
               TextInverse Text26
               Cancel = True
            End If
         '2005/9/15 END
        End If
      End If
      '2005/4/29 end
      
   End If
End Sub

Private Sub Text27_GotFocus(Index As Integer)
  TextInverse Text27(Index)
End Sub

Private Sub Text27_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
   If Index = 0 Then
      If "1604" = 專利權消滅 Then
         If KeyAscii <> 89 Then
            MsgBox "來函性質為消滅時，是否閉卷欄必須為 Y !", vbCritical
            KeyAscii = 89
         End If
      End If
   End If
End Sub

Private Sub Text27_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      If Text27(0).Text = "Y" And lblPA57 = "" Then
         If MsgBox("是否確定閉卷 ?", vbYesNo + vbQuestion) = vbNo Then
            Cancel = True
            TextInverse Text27(0)
         End If
      End If
   End If
End Sub

Private Sub Text29_GotFocus()
  TextInverse Text29
End Sub

Private Sub Text6_GotFocus()
    TextInverse Text6
    CloseIme
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = "" Then
      MsgBox "官方發文日期不可空白 !", vbCritical
      Cancel = True
   Else
      If Text6.MaxLength = 8 And Len(Text6) < 8 Then
         MsgBox "非台灣官方發文日期請輸西元年 !", vbCritical
         Cancel = True
      ElseIf ChkDate(Text6) Then
         If Val(DBDATE(Text6)) > Val(strSrvDate(1)) Then
            MsgBox "官方發文日期不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text6
End Sub

'CANCEL BY SONIA 2014/12/30 智慧局函已無機關文號
'Private Sub Text9_GotFocus()
'Dim intPos As Integer
'   ''當來函性質為"1601"或"1604"時, 將游標設定在機關文號欄的"第"的後面, 其餘則放在"專"的後面
'   With Me.Text9
'      If Len("" & .Text) > 0 Then
'           '將游標停在第後面
'         intPos = InStr("" & .Text, "第")
'         If intPos > 0 Then
'            .SelStart = intPos
'            .SelLength = 0
'         End If
'      End If
'   End With
'End Sub
'
'Private Sub Text9_Validate(Cancel As Boolean)
'   'Modify by Morgan 2011/1/3 機關文號欄位改長度(百年問題)改抓MaxLength屬性控制
'   If CheckLengthIsOK(Text9, Text9.MaxLength) = False Then
'      Cancel = True
'   End If
'   If pa(9) = 台灣國家代號 Then
'      If Text9.Text = "" Then
'         MsgBox "申請國家為台灣時不得空白，請重新輸入 !", vbCritical
'         Cancel = True
'      End If
'   End If
'End Sub
'END 2014/12/30

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/20 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/20


For Each objTxt In Text15
   If objTxt.Enabled = True Then
      Cancel = False
      Text15_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Text15(objTxt.Index).SetFocus
         Exit Function
      End If
   End If
Next

If Me.Text26.Enabled = True Then
   Cancel = False
   Text26_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

For Each objTxt In Text27
   If objTxt.Enabled = True Then
      Cancel = False
      Text27_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

If Me.Text6.Enabled = True Then
   Cancel = False
   Text6_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'CANCEL BY SONIA 2014/12/30 智慧局函已無機關文號
'If Me.Text9.Enabled = True Then
'   Cancel = False
'   Text9_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If
'END 2014/12/30

'Added by Morgan 2014/4/11 電子化-檢查pdf檔
'Removed by Morgan 2014/5/15 消滅函,證書函不必檢查,程序輸入後才掃描--陳玲玲
'If pa(9) = "000" Then
'   If PUB_CheckPDF(pa(1), pa(2), pa(3), pa(4), 1, m_DocNo) = False Then
'      Exit Function
'   End If
'End If
'end 2014/4/11

   'Added by Morgan 2020/1/17
   If pa(9) = "000" Then
      m_bolRegMail = False
   Else
      '除屆滿外都掛號
      m_bolRegMail = IIf(Text15(2) = "4" Or pa(9) = "013", False, True)
   End If
   m_bolNoCP27 = False
   '大陸案,有通知函,非掛號
   'Removed by Morgan 2024/1/30 取消--郭
'   If pa(9) = "020" And Text15(0) <> "N" And m_bolRegMail = False Then
'      If PUB_GetCustomerValue(pa(26), "CU182") = "Y" Then
'         If MsgBox("請確認是否已收到公文正本？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
'            m_bolNoCP27 = True
'         End If
'      End If
'   End If
   'end 2020/1/17

TxtValidate = True
End Function
