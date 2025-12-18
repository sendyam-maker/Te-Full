VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm05010402_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "公開公告資料輸入"
   ClientHeight    =   5160
   ClientLeft      =   -840
   ClientTop       =   1248
   ClientWidth     =   8592
   ControlBox      =   0   'False
   LinkTopic       =   "Form25"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8592
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   11
      Left            =   5160
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   10
      Left            =   5460
      TabIndex        =   13
      Top             =   4500
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.CheckBox Check2 
      Caption         =   "隨函檢附公告資料影印本乙份"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   9
      Left            =   1050
      MaxLength       =   1
      TabIndex        =   12
      Top             =   4500
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   8
      Left            =   1950
      MaxLength       =   1
      TabIndex        =   9
      Top             =   3900
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   7
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   6
      Left            =   5160
      TabIndex        =   8
      Top             =   3570
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   3
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3570
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6432
      TabIndex        =   16
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5604
      TabIndex        =   15
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7656
      TabIndex        =   17
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   4
      Left            =   1104
      MaxLength       =   8
      TabIndex        =   10
      Top             =   4200
      Width           =   1092
   End
   Begin VB.TextBox txtCaseField 
      Height          =   270
      Index           =   5
      Left            =   5160
      MaxLength       =   30
      TabIndex        =   11
      Top             =   4200
      Width           =   1932
   End
   Begin VB.Frame fraOpen 
      BorderStyle     =   0  '沒有框線
      Height          =   915
      Left            =   120
      TabIndex        =   34
      Top             =   2220
      Width           =   8292
      Begin VB.TextBox txtCaseField 
         Height          =   264
         Index           =   12
         Left            =   6576
         MaxLength       =   8
         TabIndex        =   3
         Top             =   300
         Width           =   1092
      End
      Begin VB.CheckBox Check1 
         Caption         =   "隨函檢附公開資料影印本乙份"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtCaseField 
         Height          =   264
         Index           =   0
         Left            =   960
         MaxLength       =   8
         TabIndex        =   0
         Top             =   24
         Width           =   1092
      End
      Begin VB.TextBox txtCaseField 
         Height          =   264
         Index           =   2
         Left            =   960
         MaxLength       =   1
         TabIndex        =   2
         Top             =   300
         Width           =   492
      End
      Begin VB.TextBox txtCaseField 
         Height          =   264
         Index           =   1
         Left            =   5040
         MaxLength       =   30
         TabIndex        =   1
         Top             =   24
         Width           =   1932
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "檢索報告公開日："
         Height          =   180
         Left            =   5088
         TabIndex        =   45
         Top             =   342
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "公開日："
         Height          =   180
         Left            =   0
         TabIndex        =   37
         Top             =   72
         Width           =   720
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "公開號："
         Height          =   180
         Left            =   4200
         TabIndex        =   36
         Top             =   72
         Width           =   720
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "公開與否：           （1：公開   2：即將公開   3：檢索報告公開）"
         Height          =   180
         Left            =   24
         TabIndex        =   35
         Top             =   342
         Width           =   5028
      End
   End
   Begin MSForms.ComboBox cboCaseName 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   18
      Top             =   1080
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
   Begin VB.Label lblItemCnt 
      Caption         =   "項數："
      Height          =   255
      Left            =   4320
      TabIndex        =   44
      Top             =   3240
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label11 
      Caption         =   "泰國實審費："
      Height          =   255
      Left            =   4320
      TabIndex        =   43
      Top             =   4500
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "公告與否：           （1 : 公告, 2 : 即將公告） "
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   4500
      Width           =   4095
   End
   Begin VB.Label Label9 
      Caption         =   "是否修改通知函內容：           （Y:Word） "
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   3900
      Width           =   4095
   End
   Begin VB.Label Label7 
      Caption         =   "是否指定英國：           （Y/N） "
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   3240
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label5 
      Caption         =   "實審費："
      Height          =   255
      Left            =   4320
      TabIndex        =   39
      Top             =   3570
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "是否已提實審：           （Y/N） "
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   3570
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label32 
      Caption         =   "公告號："
      Height          =   255
      Left            =   4320
      TabIndex        =   27
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label31 
      Caption         =   "公告日："
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   33
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家："
      Height          =   255
      Left            =   4320
      TabIndex        =   32
      Top             =   1440
      Width           =   975
   End
   Begin MSForms.Label lblNation 
      Height          =   255
      Left            =   6120
      TabIndex        =   31
      Top             =   1440
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   255
      Left            =   1560
      TabIndex        =   30
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "專利種類："
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   25
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   24
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   23
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "櫃台收文日："
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "申請案號："
      Height          =   255
      Left            =   4320
      TabIndex        =   20
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   19
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "frm05010402_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/8 改成Form2.0 (cboCaseName,lblNation)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'2005/7/21整理
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,pa()存放基本資料檔
Dim cp() As String, pa() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer

'Add By Cheng 2002/02/15
Dim m_strCP09ByCheng As String

'Add by Morgan 2004/5/6
'是否已觸發 Form Active 事件
Dim bolActive As Boolean
'Add by Morgan 2007/4/26 香港案控制
Dim m_HKPA01 As String, m_HKPA02 As String, m_HKPA03 As String, m_HKPA04 As String '本所案號
Dim m_HKCP09 As String '收文號
Dim m_HKCP10 As String '案件性質
Dim m_HKCP14 As String '承辦人
Dim m_HKNP22 As String 'NP22
Dim m_HKNP08 As String 'NP08
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim m_bolAddLP As Boolean, m_strCP10 As String, m_strLD18 As String 'Added by Morgan 2018/7/9
Dim m_HK1913CP09 As String, m_HKNP01 As String, m_HKNP09 As String 'Added by Morgan 2018/7/10
Dim m_416NP09 As String 'Added by Morgan2024/12/6

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(1 To 5) As String, strTmp As String
   Dim ii As Integer
   
   ii = 0
   
   'Added by Morgan 2024/12/9
   '檢索報告公開
   If txtCaseField(2).Text = "3" Then
   
      EndLetter ET01, m_strLD18, ET03, strUserNum
   
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_strLD18 & "','" & ET03 & "','" & strUserNum & _
         "','檢索報告公開日','" & DBDATE(txtCaseField(12).Text) & "')"
         
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_strLD18 & "','" & ET03 & "','" & strUserNum & _
         "','法定期限','" & m_416NP09 & "')"
            
      strExc(0) = "select nvl(np23,np08) np23,np09 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06 is null and np07 in ('218','416','215') order by np09,np23"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_strLD18 & "','" & ET03 & "','" & strUserNum & _
            "','有未收文','♀')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_strLD18 & "','" & ET03 & "','" & strUserNum & _
            "','約定期限','" & RsTemp("np23") & "')"
      Else
         strExc(0) = "select cp07 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('218','416','215') and cp159=0 order by cp07"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_strLD18 & "','" & ET03 & "','" & strUserNum & _
               "','已收文','♀')"
         End If
      End If
      
   Else
   'end 2024/12/9
   
      EndLetter ET01, cp(9), ET03, strUserNum
   
      '若為通知公告
      'Modify by Morgan 2008/1/4
      'If Me.txtCaseField(4).Text <> "" Then
      If Me.txtCaseField(9).Text <> "" Then
         If Me.Check2.Value = vbChecked Then
            strTmp = "'，隨函檢附公告資料影印本乙份，敬請查收備存'"
         Else
            strTmp = "NULL"
         End If
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','是否隨函檢附公開資料影印本乙份'," & strTmp & ")"
            
         '公告與否
         If txtCaseField(9) = "1" Then
            strTmp = "已於"
         Else
            strTmp = "將於"
         End If
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','公開與否','" & strTmp & "')"
         
         If Me.txtCaseField(10).Text <> "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','實審費用','" & Me.txtCaseField(10).Text & "')"
         End If
         
      'end 2024/12/6
      '若為通知公開
      Else
         If Check1.Value = 1 Then
            strTmp = "'，隨函檢附公開資料影印本乙份，敬請存查'"
         Else
            strTmp = "NULL"
         End If
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','是否隨函檢附公開資料影印本乙份'," & strTmp & ")"
            
         '公開與否
         If txtCaseField(2) = "1" Then
            strTmp = "已於"
         Else
            strTmp = "將於"
         End If
         
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
            "','公開與否','" & strTmp & "')"
         
         If Me.txtCaseField(6).Text <> "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','實審費用','" & Me.txtCaseField(6).Text & "')"
         End If
         
         'Added by Morgan 2025/4/18
         If PUB_IsPCTByPass(pa(91)) Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','PCTbypss不印','♀')"
         End If
         'end 2025/4/18
      End If
   
      '2008/10/1 ADD BY SONIA 抓實審期限
      strExc(0) = "SELECT NP09 FROM NEXTPROGRESS WHERE np07=" & 實體審查 & " and np02=" + CNULL(cp(1)) + _
          " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) & " and np06 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) <> "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','實審期限','" & RsTemp.Fields(0) & "')"
         End If
      End If
      '2008/10/1 END
      
      'Added by Morgan 2013/5/7
      If pa(9) = "221" Then
         '有指定英國
         If Me.txtCaseField(7).Text = "Y" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & cp(9) & "','" & ET03 & "','" & strUserNum & _
               "','有指定英國要印','♀')"
         End If
      End If
      'end 2013/5/7
      
   End If 'Added by Morgan 2024/12/9
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim i As Integer, strTmp As String, bolChk1 As Boolean, bolChk2 As Boolean
   
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         For i = 0 To 5
            If i <> 3 Then
               If txtCaseField(i).Enabled Then
                  If CheckKeyIn(i) <> 1 Then
                     txtCaseField(i).SetFocus
                     txtCaseField_GotFocus (i)
                     Exit For
                  End If
               End If
            End If
         Next
         If i = 6 Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
            
            If SaveData Then
               'Added by Morgan 2024/12/6
               '檢索報告公開
               If txtCaseField(2).Text = "3" Then
                  strTmp = "00"
                  StartLetter "04", strTmp
                  NowPrint m_strLD18, "04", strTmp, IIf(Me.txtCaseField(8).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                  
                  If txtCaseField(8).Text = "Y" Then
                     frm1105_1.m_RecNo = m_strLD18
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & m_strCP10 & ".CUS.PDF"
                     frm1105_1.Show
                  End If
               Else
               'end 2024/12/6
               
                  'Add by Morgan 2007/4/26 通知香港案承辦人
                  If m_HKCP14 <> "" Then
                     Call PUB_SendMail(strUserNum, m_HKCP14, m_HKCP09, "香港的關聯案(" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & ")已" & IIf(m_HKCP10 = "110", "公開", "公告") & "，香港案(" & m_HKPA01 & "-" & m_HKPA02 & "-" & m_HKPA03 & "-" & m_HKPA04 & ")的[" & IIf(m_HKCP10 = "110", "標準專利記錄請求", "標準專利批准記錄請求") & "]可以處理！", "如旨")
                  End If
                  'end 2007/4/26
   
                  '通知公告
                  'Modify by Morgan 2008/1/4
                  'If Not IsEmptyText(txtCaseField(4)) Then
                  If txtCaseField(9).Text <> "" Then
                     '若有輸入泰國實審費
                     If Me.txtCaseField(10).Text <> "" Then
                        strTmp = "21"
                     Else
                        strTmp = "20"
                     End If
                  
                  '通知公開
                  Else
                     bolChk1 = False
                     '92.4.10 MODIFY BY SONIA
                     'strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                     '   " AND CP10='" & 實體審查 & "' AND CP27 IS NOT NULL"
                     'Modify by Morgan 2009/10/27 +判斷未取消收文
                     strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                        " AND CP10='" & 實體審查 & "' and cp57 is null"
                     '92.4.10 END
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 1 Then If RsTemp.Fields(0) > 0 Then bolChk1 = True
                     'modify by sonia 90.10.10有無提供前案資料
                     bolChk2 = False
                     '92.4.10 MODIFY BY SONIA
                     'strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                     '   " AND CP10='" & 提供前案資料 & "' AND CP27 IS NOT NULL"
                     'Modify by Morgan 2009/10/27 +判斷未取消收文
                     strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                        " AND CP10='" & 提供前案資料 & "' and cp57 is null"
                     '92.4.10 END
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 1 Then If RsTemp.Fields(0) > 0 Then bolChk2 = True
                     
                     '一般國家公開
                     strTmp = "00"
                     
                     'Added by Morgan 2025/4/21 PCT案
                     If pa(46) = "Y" Then
                        strTmp = "12"
                     Else
                     'end 2025/4/21
                     
                        Select Case pa(9)
                           Case "019" '泰國
                              If pa(8) = "1" And bolChk1 = False Then
                                 strTmp = "02"
                              Else
                                 strTmp = "01"
                              End If
                           Case "203", "202", "225", "231" '法國, 德國
                              strTmp = "03"
                           '2005/7/21 MODIFY BY SONIA 新法拆開日,韓
                           Case "011" '日本
                              '發明
                              If pa(8) = "1" Then
                                 If bolChk1 = False Then
                                    strTmp = "04"
                                 Else
                                    strTmp = "00"
                                 End If
                              '新型
                              ElseIf pa(8) = "2" Then
                                 strTmp = "05"
                              End If
                           Case "012" '韓國
                              '發明
                              If pa(8) = "1" Then
                                 If bolChk1 = False Then
                                    strTmp = "15"
                                 Else
                                    strTmp = "00"
                                 End If
                              '新型
                              ElseIf pa(8) = "2" Then
                                 strTmp = "05"
                              End If
                           '2005/7/21 END
                           Case "017" '印尼
                              '發明
                              If pa(8) = "1" Then
                                 'Modified by Morgan 2024/8/29 印尼請求實審的期限和案件公開無關關,通知公開函就單純通知公開--玫音
                                 'If bolChk1 = True Then
                                 '   strTmp = "06"
                                 'Else
                                 '   strTmp = "07"
                                 'End If
                                 strTmp = "07"
                                 'end 2024/8/298
                              Else
                                 strTmp = "00"
                              End If
                           Case "201" '英國
                              '發明
                              If pa(8) = "1" Then
                                 If bolChk1 = False Then
                                    strTmp = "08"
                                 Else
                                    strTmp = "09"
                                 End If
                              Else
                                 strTmp = "09"
                              End If
                           Case "221" 'EPC
                              'Modified by Morgan 2013/5/7 定稿合併(用例外欄位控制)
                              'If bolChk1 = True Then
                              '   If Me.txtCaseField(7).Text = "Y" Then
                              '      strTmp = "10"
                              '   Else
                              '      strTmp = "11"
                              '   End If
                              'Else
                              '   If Me.txtCaseField(7).Text = "Y" Then
                              '      strTmp = "12"
                              '   Else
                              '      strTmp = "13"
                              '   End If
                              'End If
                              strTmp = "11"
                              'end 2013/5/7
                           '92.4.10 add by sonia
                           Case "101"  '美國
                              strTmp = "14"
                           '92.4.10 end
                           'Add by Morgan 2006/8/30
                           Case "042"
                              '新型
                              If pa(8) = "2" Then
                                 strTmp = "16"
                              End If
                           'Add by Morgan 2009/7/29
                           Case "056" 'PCT
                              strTmp = "17"
                        End Select
                     
                     End If
                  End If
                  
                  StartLetter "04", strTmp
                  NowPrint cp(9), "04", strTmp, IIf(Me.txtCaseField(8).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                  
                  'Added by Morgan 2018/7/3 CFP電子化
                  If m_bolAddLP And txtCaseField(8).Text = "Y" Then
                     If m_HKNP22 <> "" And m_HKCP10 = "111" Then
                        MsgBox "為配合轉PDF檔至卷宗區，香港案[" & m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "]的標準專利批准紀錄請求通知函請到定稿維護修改內容!!", vbInformation, "CFP電子化"
                        txtCaseField(8).Text = ""
                     End If
                     frm1105_1.m_RecNo = m_strLD18
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & m_strCP10 & ".CUS.PDF"
                     frm1105_1.Show
                  End If
                  'end 2018/7/3
                  
                  'Add by Morgan 2007/4/30 若有未收文的香港案的"標準專利批准紀錄請求"時要出定稿
                  If m_HKNP22 <> "" And m_HKCP10 = "111" Then
                     RunHKInform IIf(Me.txtCaseField(8).Text = "Y", True, False)
                  End If
                  'end 2007/4/30
               
                  'Add by Morgan 2007/3/22
                  strExc(1) = ""
                  If (txtCaseField(0).Text <> "" And txtCaseField(0).Text <> txtCaseField(0).Tag) Then
                     strExc(1) = DBDATE(txtCaseField(0))
                     If strExc(1) > strSrvDate(1) Then
                        strExc(1) = ""
                     End If
                  End If
                  If strExc(1) <> "" Then
                     '公開
                     PUB_SameCaseCheck1 cp(), 3, strExc(1)
                  Else
                     If (txtCaseField(4).Text <> "" And txtCaseField(4).Text <> txtCaseField(4).Tag) Then
                        strExc(1) = DBDATE(txtCaseField(4))
                        If strExc(1) > strSrvDate(1) Then
                           strExc(1) = ""
                        End If
                     End If
                     If strExc(1) <> "" Then
                        '公告
                        PUB_SameCaseCheck1 cp(), 4, strExc(1)
                     End If
                  End If
                  'end 2007/3/22
               
               End If 'Added by Morgan 2024/12/6
               
               'Add By Sindy 2016/10/7
               If Me.m_strIR01 <> "" Then
                  bolLeave = True
                  Unload frm05010402_1
                  Unload Me
                  'Modify By Sindy 2022/5/20
                  'frm04010519.GoNext
                  Forms(0).Tmpfrm04010519.GoNext
                  Set Forms(0).Tmpfrm04010519 = Nothing
                  '2022/5/20 END
               Else
               '2016/10/7 END
                  bolLeave = True
                  intLeaveKind = 1
                  Unload Me
               End If
            '911202 nick
            Else
                MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
            End If
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         If Index = 2 Then
            intLeaveKind = 0
         Else
            intLeaveKind = 2
         End If
         Unload Me
   End Select
End Sub

Private Function SaveData() As Boolean
Dim strDateS(0 To 5) As String
Dim strCaseProperty As String
Dim strTxt(1 To 30) As String, iStep As Integer
Dim strTemp As String, strTemp1 As String, strTemp2 As String, dobDateAdd As Double
Dim strDate As String, strDate1 As String, strStartDate As String
Dim lMax As Long
Dim varTemp As Variant
Dim bolNP22 As Boolean, NP22(1 To 3) As String, iNP22 As Integer
Dim i As Integer
'Add by Morgan 2004/2/6
Dim stCP12 As String, stCP13 As String
Dim m_strPA04 As String   'add by sonia 記錄EPC之土耳其子案案號
 
 '911105 nick
 SaveData = True
 On Error GoTo CheckingErr
 
 '911105 nick transation
 cnnConnection.BeginTrans
 
   pa(12) = txtCaseField(0)
   pa(13) = txtCaseField(1)
   pa(14) = txtCaseField(4)
   pa(15) = txtCaseField(5)
   cp(14) = strUserNum
   
   'Modify by Morgan 2007/12/27
   'If txtCaseField(4) <> "" Then
   If txtCaseField(9).Text <> "" Then
      strCaseProperty = 通知公告
   'Added by Morgan 2024/12/6
   '檢索報告公開
   ElseIf txtCaseField(2).Text = "3" Then
      strCaseProperty = "1238"
   'end 2024/12/6
   Else
      strCaseProperty = 通知公開
   End If
   
   strTxt(1) = GetPASQL(pa())
   
   '911106 nick transation
   cnnConnection.Execute strTxt(1)
   
   'edit by nickc 2007/02/02
   'Dim strDataTemp(1 To T_CP) As String
   Dim strDataTemp() As String
   ReDim strDataTemp(1 To TF_CP) As String
   
   strDataTemp(1) = pa(1)
   strDataTemp(2) = pa(2)
   strDataTemp(3) = pa(3)
   strDataTemp(4) = pa(4)
   '91.12.16 MODIFY BY SONIA
   'strDataTemp(5) = TransDate((lblCaseField(4)), 2)
   strDataTemp(5) = strSrvDate(1)
   '91.12.16 END
   strDataTemp(9) = 主管機關來函
   strDataTemp(10) = strCaseProperty
   '2009/12/30 MODIFY BY SONIA
   'strDataTemp(12) = cp(12)
   'strDataTemp(13) = cp(13)
   strDataTemp(13) = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
   strDataTemp(12) = GetSalesArea(strDataTemp(13))
   '2009/12/30 END
   strDataTemp(14) = strUserNum
   strDataTemp(27) = strSrvDate(1)
   strDataTemp(20) = "N"
   '92.5.8 ADD BY SONIA
   strDataTemp(26) = "N"
   '92.5.8 END
   strDataTemp(32) = "N"
   strDataTemp(43) = cp(9)
   '2008/8/26 modify by sonia 櫃台收文日改存 cp119
   '2008/10/24 MODIFY BY SONIA CP64仍存
   strDataTemp(64) = "櫃台收文日：" & lblCaseField(4)
   'Added by Morgan 2024/12/6 檢索報告公開
   If txtCaseField(2).Text = "3" Then
      strDataTemp(115) = DBDATE(txtCaseField(12))
   End If
   'end 2024/12/6
   strDataTemp(119) = ChangeTStringToWString(lblCaseField(4))
   '2008/8/26 end
   
   strTxt(2) = GetCPSQL(strDataTemp(), False)
   '911106 nick transation
   cnnConnection.Execute strTxt(2)
   
   'Add by Morgan 2004/11/30 抓最新的AB類發文代理人更新
   Pub_UpdateFromMaxCP27 pa(1), pa(2), pa(3), pa(4)
   
   'Added by Morgan 2024/12/6
   '檢索報告公開
   m_416NP09 = ""
   If strCaseProperty = "1238" Then
      m_416NP09 = CompDate(1, 6, txtCaseField(12))
      '下一程序上Y
      strSql = "update nextprogress set np06='Y' where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06 is null and np07='" & strCaseProperty & "'"
      cnnConnection.Execute strSql, intI
      '更新回覆檢索報告、實審、指定費期限
      'Modified by Morgan 2025/11/6 還原，不可取消，下面不會Run到
      PUB_UpdExamDate cp(1), cp(2), cp(3), cp(4), cp(9) 'Removed by Morgan 2025/7/24 取消，下面有呼叫
   Else
   'end 2024/12/6
   
      iStep = 3
      
   '2009/12/30 CANCEL BY SONIA
   '    'Add By Cheng 2003/04/03
   '    '智權人員存最近收文A類接洽記錄單的智權人員
   '    'Modify by Morgan 2004/2/6
   '    'strTxt(iStep) = "Update Caseprogress Set CP13='" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "' Where CP09='" & strDataTemp(9) & "' "
   '    stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   '    stCP12 = GetSalesArea(stCP13)
   '    strTxt(iStep) = "Update Caseprogress Set CP12='" & stCP12 & "',CP13='" & stCP13 & "' Where CP09='" & strDataTemp(9) & "' "
   '    'Modify end 2004/2/6
   '
   '    cnnConnection.Execute strTxt(iStep)
   '    iStep = iStep + 1
   '2009/12/30 END
   
      lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
      bolNP22 = False
      iNP22 = 1
      
      strTemp = ""
      strTemp1 = ""
      
'Removed by Morgan 2025/7/24 和下面呼叫的 PUB_UpdExamDate 作用重複
'         'edit by nickc 2007/02/02 不用 dll 了
'         'If objPublicData.GetNationTaxEx(Val(pA(8)) + 3, pA(9), strTemp, strTemp1, , , False) = 0 Then
'         If ClsPDGetNationTaxEx(Val(pa(8)) + 3, pa(9), strTemp, strTemp1, , , False) = 0 Then
'            If Val(strTemp) = 公開日 Then
'               dobDateAdd = Val(strTemp1)
'               strStartDate = GetStartDate(strTemp, strDataTemp(), pa())
'               If strStartDate <> "" Then
'                  strStartDate = CompDate(1, dobDateAdd, strStartDate)
'                  strTemp = strStartDate
'                  strDateS(1) = pa(1)
'                  strDateS(2) = pa(9)
'                  strDateS(3) = TransDate(strTemp, 2)
'                  GetCtrlDT strDateS
'                  strTemp1 = strDateS(0)
'                  '92.4.3 end
'
'                  '92.6.8 ADD BY SONIA
'                  '至案件進度檔中找是否已收文實體審查
'                  strExc(0) = "SELECT CP09,CP27 FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
'                     " AND CP10=" & 實體審查 & ""
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     If IsNull(RsTemp.Fields("CP27")) Then
'                          'Modify By Cheng 2003/12/08
'                          '若本所期限非工作天則抓最近的工作天
'      '                  strTxt(iStep) = "update caseprogress set cp06=" + strTemp1 + ",cp07=" + strTemp + " WHERE cp10=" & 實體審查 & _
'      '                    " and cp01=" + CNULL(cp(1)) + _
'      '                    " and cp02=" + CNULL(cp(2)) + " and cp03=" + CNULL(cp(3)) + " and cp04=" + CNULL(cp(4))
'                        strTxt(iStep) = "update caseprogress set cp06=" + PUB_GetWorkDay1(strTemp1, True) + ",cp07=" + strTemp + " WHERE cp10=" & 實體審查 & _
'                          " and cp01=" + CNULL(cp(1)) + _
'                          " and cp02=" + CNULL(cp(2)) + " and cp03=" + CNULL(cp(3)) + " and cp04=" + CNULL(cp(4))
'                        cnnConnection.Execute strTxt(iStep)
'                        iStep = iStep + 1
'                     End If
'                  Else
'                     '實體審查未收文
'                     strExc(0) = "SELECT NP22 FROM NEXTPROGRESS WHERE np07=" & 實體審查 & " and np02=" + CNULL(cp(1)) + _
'                         " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) & " and np06 is null"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'                     If intI = 0 Then
'                        'Modify By Cheng 2003/04/03
'                        '智權人員存最近收文A類接洽記錄單的智權人員
'                          'Modify By Cheng 2003/12/08
'                          '若本所期限非工作天則抓最近的工作天
'      '                  strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" + _
'      '                      CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
'      '                      "," + CNULL(strDataTemp(4)) + "," + 實體審查 + "," & Val(strTemp1) & "," & Val(strTemp) & _
'      '                      "," + CNULL(PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))) + "," & lMax & ")"
'                        strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" + _
'                            CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
'                            "," + CNULL(strDataTemp(4)) + "," + 實體審查 + "," & Val(PUB_GetWorkDay1(strTemp1, True)) & "," & Val(strTemp) & _
'                            "," + CNULL(PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))) + "," & lMax & ")"
'                        '911105 nick transation
'                        cnnConnection.Execute strTxt(iStep)
'
'                        iStep = iStep + 1
'                        lMax = lMax + 1
'                     Else
'                          'Modify By Cheng 2003/12/08
'                          '若本所期限非工作天則抓最近的工作天
'      '                  strTxt(iStep) = "update nextprogress set np08=" + strTemp1 + ",np09=" + strTemp + " WHERE np07=" & 實體審查 & _
'      '                      " and np02=" + CNULL(strDataTemp(1)) + _
'      '                      " and np03=" + CNULL(strDataTemp(2)) + " and np04=" + CNULL(strDataTemp(3)) + " and np05=" + CNULL(strDataTemp(4)) & " and np06 is null"
'                        strTxt(iStep) = "update nextprogress set np08=" + PUB_GetWorkDay1(strTemp1, True) + ",np09=" + strTemp + " WHERE np07=" & 實體審查 & _
'                            " and np02=" + CNULL(strDataTemp(1)) + _
'                            " and np03=" + CNULL(strDataTemp(2)) + " and np04=" + CNULL(strDataTemp(3)) + " and np05=" + CNULL(strDataTemp(4)) & " and np06 is null"
'                        '911105 nick transation
'                        cnnConnection.Execute strTxt(iStep)
'
'                        iStep = iStep + 1
'                     End If
'                  End If
'               End If
'            End If
'         End If
'      End If
'end 2025/7/24
      
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetNationTaxEx(Val(pA(8)), pA(9), strTemp, strTemp1, 年費, strTemp2) = 0 Then
      'Modified by Morgan 2013/10/23
      'If ClsPDGetNationTaxEx(Val(pa(8)), pa(9), strTemp, strTemp1, 年費, strTemp2) = 0 Then
      If ClsPDGetNationTaxEx(Val(pa(8)), pa(9), strTemp, strTemp1, 年費, strTemp2, , , pa(10), pa(21), pa(72)) = 0 Then
         'Modified by Morgan 2015/7/8
         'If Val(strTemp) = 公開日 Then
         If Val(strTemp) = 公開日 Or Val(strTemp) = 公告日 Then
         'end 2015/7/8
            'edit by nickc 2007/02/02
            'Dim np(1 To T_NP) As String
            Dim np() As String
            ReDim np(1 To TF_NP) As String
            
            varTemp = Split(strTemp1, ",")
            i = GetMoneyYears(pa(72))
            'Modify by Morgan 2005/3/21
            'If i > UBound(varTemp) Then GoTo NextStep
            If i > UBound(varTemp) + 1 Then GoTo Nextstep
            
            dobDateAdd = varTemp(i - 1)
            'Add by Morgan 2005/3/21 年費期限要減一年計算
            If Not GetNP07(pa(9), pa(8), np(7)) Then Exit Function
            If np(7) = "605" Then dobDateAdd = dobDateAdd - 1
            '2005/3/21 end
            strStartDate = GetStartDate(strTemp, cp(), pa())
            ' 90.12.18 modify by louis (抓不到日期則離開)
            If strStartDate = "" Then GoTo Nextstep
         
            If strStartDate <> "" Then
               ' 91.09.28 modify by louis (加年用yyyy)
               'strStartDate = DateAdd("Y", dobDateAdd, ChangeWStringToWDateString(strStartDate))
               strStartDate = DateAdd("yyyy", dobDateAdd, ChangeWStringToWDateString(strStartDate))
               'Remove by Morgan 2005/3/21 不必減一天
               'strStartDate = DateAdd("d", -1, strStartDate)
               strDate = ChangeWDateStringToWString(strStartDate)
            End If
            '92.4.3 modify by sonia
            'strDate1 = ChangeWDateStringToWString(DateAdd("M", -1, ChangeWStringToWDateString(strDate)))
            strDateS(1) = pa(1)
            strDateS(2) = pa(9)
            strDateS(3) = TransDate(strDate, 2)
            GetCtrlDT strDateS
            strDate1 = strDateS(0)
            '92.4.3 end
            np(2) = cp(1)
            np(3) = cp(2)
            np(4) = cp(3)
            np(5) = cp(4)
            'If Not GetNP07(pa(9), pa(8), np(7)) Then Exit Function 'Modify by Morgan 2005/3/21 上移
            '911105 nick 重新抓
            lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
            
            strExc(0) = "SELECT NP22 FROM NEXTPROGRESS WHERE np07=" + CNULL(np(7)) + " and np02=" + CNULL(np(2)) + _
               " and np03=" + CNULL(np(3)) + " and np04=" + CNULL(np(4)) + " and np05=" + CNULL(np(5)) & " and np06 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               'Modify By Cheng 2003/09/05
               '判斷是否要更新下一程序
               If blnUpdateNP(cp(1) & cp(2) & cp(3) & cp(4)) = True Then
                   'Modify By Cheng 2003/04/03
                   '智權人員存最近收文A類接洽記錄單的智權人員
                       'Modify By Cheng 2003/12/08
                       '若本所期限非工作天則抓最近的工作天
   '                 strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" + _
   '                    CNULL(strDataTemp(9)) + "," + CNULL(cp(1)) + "," + CNULL(cp(2)) + "," + CNULL(cp(3)) + _
   '                    "," + CNULL(cp(4)) + "," + CNULL(np(7)) + "," & Val(strDate1) & "," & Val(strDate) & _
   '                    "," + CNULL(PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))) + "," & lMax & ")"
                    strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" + _
                       CNULL(strDataTemp(9)) + "," + CNULL(cp(1)) + "," + CNULL(cp(2)) + "," + CNULL(cp(3)) + _
                       "," + CNULL(cp(4)) + "," + CNULL(np(7)) + "," & Val(PUB_GetWorkDay1(strDate1, True)) & "," & Val(strDate) & _
                       "," + CNULL(PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))) + "," & lMax & ")"
                    '911105 nick transation
                    cnnConnection.Execute strTxt(iStep)
               End If
               bolNP22 = True
               NP22(iNP22) = lMax
               iNP22 = iNP22 + 1
               
               '911105 nick 重新抓
               lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
               
               iStep = iStep + 1
               lMax = lMax + 1
            Else
               'Modify By Cheng 2003/12/08
               '若本所期限非工作天則抓最近的工作天
   '            strTxt(iStep) = "update nextprogress set np08=" + strDate1 + ",np09=" + strDate + " where np07=" + CNULL(np(7)) + " and np02=" + CNULL(np(2)) + _
   '            " and np03=" + CNULL(np(3)) + " and np04=" + CNULL(np(4)) + " and np05=" + CNULL(np(5)) & " and np06 is null"
               strTxt(iStep) = "update nextprogress set np08=" + PUB_GetWorkDay1(strDate1, True) + ",np09=" + strDate + " where np07=" + CNULL(np(7)) + " and np02=" + CNULL(np(2)) + _
               " and np03=" + CNULL(np(3)) + " and np04=" + CNULL(np(4)) + " and np05=" + CNULL(np(5)) & " and np06 is null"
               
               '911105 nick transation
               cnnConnection.Execute strTxt(iStep)
      
               iStep = iStep + 1
            End If
         End If
      End If
       
Nextstep:
   
      '92.6.8 ADD BY SONIA 更新下一程序公開期限
      If strCaseProperty = 通知公開 Then
         strTxt(iStep) = "update nextprogress set np06='Y' where np07='999' and np02=" + CNULL(cp(1)) + _
              " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) & " and np06 is null"
         cnnConnection.Execute strTxt(iStep)
         iStep = iStep + 1
      End If
      '92.6.8 END
      
   'Modify by Morgan 2009/7/3 更新實審期限合併到 PUB_UpdExamDate
   If strCaseProperty <> 通知公告 Then 'Added by Morgan 2025/7/24 因中間來所案件輸公告可能會誤管制 Ex:CFP-035038
      PUB_UpdExamDate cp(1), cp(2), cp(3), cp(4), cp(9)
   End If
      
   '   '2008/10/30 整合 modify by sonia 以公開日加某段時間去更新下一程序檔之實體審查之期限
   '   'Add By Cheng 2002/04/11 英國201發明案, 匈牙利219發明或新型案, EPC221發明案 時, 以公開日加半年
   '   '2003.3.23 Add By sonia  泰國019發明案, 以公開日加5年
   '   '2008/10/30 add by sonia 菲律賓030發明,印尼017發明, 以公開日加半年
   '   '                        土耳其235發明, 以公開日加3個月
   '   If (Me.lblCaseField(3).Caption = "201" And lblCaseField(2) = "1") _
   '   Or (Me.lblCaseField(3).Caption = "219" And lblCaseField(2) <> "3") _
   '   Or (Me.lblCaseField(3).Caption = "221" And lblCaseField(2) = "1") _
   '   Or (Me.lblCaseField(3).Caption = "030" And lblCaseField(2) = "1") _
   '   Or (Me.lblCaseField(3).Caption = "017" And lblCaseField(2) = "1") _
   '   Or (Me.lblCaseField(3).Caption = "253" And lblCaseField(2) = "1") _
   '   Or (Me.lblCaseField(3).Caption = "019" And lblCaseField(2) = "1") Then
   '      If Me.lblCaseField(3).Caption = "019" And lblCaseField(2) = "1" Then   '泰國發明案時, 以公開日加5年
   '         strTemp = CompDate(0, 5, ChangeTStringToWString(Val(Me.txtCaseField(0).Text)))
   '      ElseIf Me.lblCaseField(3).Caption = "253" And lblCaseField(2) = "1" Then   '土耳其發明案時, 以公開日加3個月
   '         strTemp = CompDate(1, 3, ChangeTStringToWString(Val(Me.txtCaseField(0).Text)))
   '      Else                                                                   '以公開日加其他半年
   '         '2008/10/30 MODIFY BY SONIA
   '         'strTemp = DateAdd("m", 6, (ChangeWStringToWDateString(Val(Me.txtCaseField(0).Text) + 19110000)))
   '         'strTemp = ChangeWDateStringToWString(strTemp)
   '         strTemp = CompDate(1, 6, ChangeTStringToWString(Val(Me.txtCaseField(0).Text)))
   '         '2008/10/30 END
   '      End If
   '
   '      strDates(1) = pa(1)
   '      strDates(2) = pa(9)
   '      strDates(3) = TransDate(strTemp, 2)
   '      GetCtrlDT strDates
   '      strTemp1 = strDates(0)
   '
   '      '先抓案件進度檔, 若有收文則更新, 若未收文才判斷新增或更新下一程序檔
   '      strExc(0) = "SELECT CP09,CP27 FROM CASEPROGRESS WHERE CP10=" & 實體審查 & " and CP01=" + CNULL(cp(1)) + _
   '         " and CP02=" + CNULL(cp(2)) + " and CP03=" + CNULL(cp(3)) + " and CP04=" + CNULL(cp(4))
   '      intI = 1
   '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '      If intI = 1 Then
   '         '未發文才更新案件進度檔之期限
   '         If IsNull(RsTemp.Fields("CP27")) Then
   '            strTxt(iStep) = "Update CaseProgress Set CP06=" & Val(PUB_GetWorkDay1(strTemp1, True)) & ",CP07=" & Val(strTemp) & " Where " & _
   '               "CP09 ='" & (RsTemp.Fields("CP09")) & "'"
   '               cnnConnection.Execute strTxt(iStep)
   '            iStep = iStep + 1
   '         End If
   '      Else
   '         '未收文才判斷新增或更新下一程序檔
   '         strExc(0) = "SELECT NP22 FROM NEXTPROGRESS WHERE np07=" & 實體審查 & " and np02=" + CNULL(cp(1)) + _
   '            " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) & " and np06 is null"
   '         intI = 1
   '         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '         If intI = 1 Then
   '            strTxt(iStep) = "Update NextProgress Set NP08=" & Val(PUB_GetWorkDay1(strTemp1, True)) & ",NP09=" & Val(strTemp) & " Where " & _
   '               "NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP06 IS NULL AND NP07=" & 實體審查
   '            cnnConnection.Execute strTxt(iStep)
   '            iStep = iStep + 1
   '         Else
   '            lMax = ClsLawGetMax
   '            strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" + _
   '               CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
   '               "," + CNULL(strDataTemp(4)) + "," + 實體審查 + "," & Val(PUB_GetWorkDay1(strTemp1, True)) & "," & Val(strTemp) & _
   '               "," + CNULL(PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))) + "," & lMax & ")"
   '            cnnConnection.Execute strTxt(iStep)
   '            iStep = iStep + 1
   '         End If
   '      End If
   '   End If
   '   '2008/10/30 end
      
   '   'Add by Morgan 2004/9/2 EPC發明更新指定費215
   '   If Me.lblCaseField(3).Caption = "221" And lblCaseField(2) = "1" Then
   '      strExc(0) = "SELECT CP09,CP27 FROM CASEPROGRESS WHERE CP10=215 and CP01=" + CNULL(cp(1)) + _
   '         " and CP02=" + CNULL(cp(2)) + " and CP03=" + CNULL(cp(3)) + " and CP04=" + CNULL(cp(4))
   '      intI = 1
   '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   '      If intI = 1 Then
   '         '未發文才更新案件進度檔之期限
   '         If IsNull(RsTemp.Fields("CP27")) Then
   '            strTxt(iStep) = "Update CaseProgress Set CP06=" & Val(PUB_GetWorkDay1(strTemp1, True)) & ",CP07=" & Val(strTemp) & " Where " & _
   '               "CP09 ='" & (RsTemp.Fields("CP09")) & "'"
   '               cnnConnection.Execute strTxt(iStep)
   '            iStep = iStep + 1
   '         End If
   '      Else
   '         '未收文才判斷新增或更新下一程序檔
   '         strExc(0) = "SELECT NP22 FROM NEXTPROGRESS WHERE np07=215 and np02=" + CNULL(cp(1)) + _
   '            " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) & " and np06 is null"
   '         intI = 1
   '         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   '         If intI = 1 Then
   '            strTxt(iStep) = "Update NextProgress Set NP08=" & Val(PUB_GetWorkDay1(strTemp1, True)) & ",NP09=" & Val(strTemp) & " Where " & _
   '               "NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP06 IS NULL AND NP07=215"
   '            cnnConnection.Execute strTxt(iStep)
   '            iStep = iStep + 1
   '         Else
   '            lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
   '            strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" + _
   '               CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
   '               "," + CNULL(strDataTemp(4)) + ", 215 ," & Val(PUB_GetWorkDay1(strTemp1, True)) & "," & Val(strTemp) & _
   '               "," + CNULL(stCP13) + "," & lMax & ")"
   '            cnnConnection.Execute strTxt(iStep)
   '            iStep = iStep + 1
   '         End If
   '      End If
   '   End If
   '   '2004/9/2 END
      
   'end 2009/7/3
      
      'Add by Morgan 2007/4/26
      'EPC或英國案須檢查是否有香港案
      m_HKCP14 = "": m_HKCP09 = "": m_HKCP10 = "": m_HKNP22 = ""
      If pa(9) = "221" Or pa(9) = "201" Then
         '有香港案
         If ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), m_HKPA01, m_HKPA02, m_HKPA03, m_HKPA04) = True Then
            '法限=公開/公告日+6個月
            '公告
            If txtCaseField(4) <> "" Then
               m_HKCP10 = "111"
               strDateS(3) = CompDate(1, 6, txtCaseField(4))
            '公開
            ElseIf txtCaseField(0) <> "" Then
               m_HKCP10 = "110"
               strDateS(3) = CompDate(1, 6, txtCaseField(0))
            End If
            If m_HKCP10 <> "" Then
               strDateS(0) = ""
               strDateS(1) = m_HKPA01
               strDateS(2) = "013"
               GetCtrlDT strDateS
               '所限
               strDateS(4) = PUB_GetWorkDay1(strDateS(0), True)
               
               strExc(0) = "select cp09,cp14,cp27,EP06,cf04 from patent,caseprogress,engineerprogress,casefee" & _
                  " where pa01='" & m_HKPA01 & "' and pa02='" & m_HKPA02 & "' and pa03='" & m_HKPA03 & "' and pa04='" & m_HKPA04 & "' and pa57 is null" & _
                  " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='" & m_HKCP10 & "' and cp57 is null" & _
                  " and ep02(+)=cp09 and cf01(+)=cp01 and cf02(+)='013' and cf03(+)=cp10"
            
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               '已收文
               If intI = 1 Then
                  With RsTemp
                  '未發文
                  If IsNull(.Fields("cp27")) Then
                     m_HKCP09 = "" & .Fields("cp09")
                     '未齊備
                     If IsNull(.Fields("EP06")) Then
                        m_HKCP14 = "" & .Fields("cp14")
                        
                        '更新齊備日
                        strSql = "update engineerprogress set ep06=" & strSrvDate(1) & " where ep02='" & m_HKCP09 & "' and ep06 is null"
                        cnnConnection.Execute strSql, intI
                        
                        If PUB_IfSetCP48(m_HKCP09) Then  'Add by Morgan 2010/10/4
                           '承辦期限
                           'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
                           'strDates(5) = CompWorkDay(Val("" & .Fields("cf04")), strSrvDate(1))
                           strDateS(5) = Pub_GetHandleDay(m_HKPA01, "013", m_HKCP10, , , m_HKCP09)
                           'end 2007/10/11
                           
                           '更新承辦期限
                           strSql = "Update CaseProgress Set CP48=" & strDateS(5) & " Where CP09='" & m_HKCP09 & "' AND ( CP48 IS NULL OR CP48>" & strDateS(5) & ")"
                           cnnConnection.Execute strSql, intI
                        End If
                     End If
                     '更新期限
                     strSql = "Update CaseProgress Set CP06=" & strDateS(4) & ",CP07=" & strDateS(3) & " Where CP09='" & m_HKCP09 & "' and (cp07 is null or CP07>" & strDateS(3) & ") and cp27 is null"
                     cnnConnection.Execute strSql, intI
                  End If
                  End With
               '未收文
               ElseIf m_HKCP10 = "111" Then
                  m_HKNP08 = strDateS(4)
                  m_HKNP09 = strDateS(3) 'Added by Morgan 2018/7/12
                  strExc(0) = "select np22,np01 from nextprogress where  np02='" & m_HKPA01 & "' and np03='" & m_HKPA02 & "' and np04='" & m_HKPA03 & "' and np05='" & m_HKPA04 & "' and np07='" & m_HKCP10 & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     m_HKNP01 = RsTemp.Fields("np01")  'Added by Morgan 2018/7/10
                     m_HKNP22 = RsTemp.Fields("np22")
                     strSql = "update nextprogress set np08=" & strDateS(4) & ",np09=" & strDateS(3) & " Where np22=" & m_HKNP22 & " and np01='" & RsTemp.Fields("np01") & "'"
                  Else
                     m_HKNP01 = strDataTemp(9)  'Added by Morgan 2018/7/10
                     m_HKNP22 = GetNextProgressNo
                     strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22)" & _
                              " VALUES ('" & strDataTemp(9) & "','" & m_HKPA01 & "','" & m_HKPA02 & "','" & m_HKPA03 & "','" & m_HKPA04 & "'" & _
                              ",'" & m_HKCP10 & "'," & strDateS(4) & "," & strDateS(3) & ",'" & PUB_GetAKindSalesNo(m_HKPA01, m_HKPA02, m_HKPA03, m_HKPA04) & "'," & m_HKNP22 & ")"
                  End If
                  cnnConnection.Execute strSql, intI
               End If
            End If
         End If
      End If
      'end 2007/4/26
      
      '2010/7/16 ADD BY SONIA 以色列發明公告日加半年先掛年費期限,發證再更新
      If pa(9) = "027" And pa(8) = "1" And txtCaseField(9).Text <> "" Then
         strTemp = CompDate(1, 6, DBDATE(txtCaseField(4)))
         strDateS(1) = pa(1)
         strDateS(2) = pa(9)
         strDateS(3) = strTemp
         GetCtrlDT strDateS
         strTemp1 = strDateS(0)
         '至案件進度檔中找是否已收文年費
         strExc(0) = "SELECT CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
            " AND CP27 IS NULL AND CP57 IS NULL AND CP10=" & 年費 & ""
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strTxt(iStep) = "update caseprogress set cp06=" + PUB_GetWorkDay1(strTemp1, True) + ",cp07=" + strTemp + " WHERE cp09=' " & RsTemp.Fields(0) & "'"
            cnnConnection.Execute strTxt(iStep)
            iStep = iStep + 1
         Else
            '年費未收文
            strExc(0) = "SELECT NP22 FROM NEXTPROGRESS WHERE np07=" & 年費 & " and np02=" + CNULL(cp(1)) + _
                " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) & " and np06 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" + _
                   CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
                   "," + CNULL(strDataTemp(4)) + "," + 年費 + "," & Val(PUB_GetWorkDay1(strTemp1, True)) & "," & Val(strTemp) & _
                   "," + CNULL(PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))) + "," & lMax & ")"
               cnnConnection.Execute strTxt(iStep)
               iStep = iStep + 1
               lMax = lMax + 1
            Else
               strTxt(iStep) = "update nextprogress set np08=" + PUB_GetWorkDay1(strTemp1, True) + ",np09=" + strTemp + " WHERE np07=" & 年費 & _
                   " and np02=" + CNULL(strDataTemp(1)) + _
                   " and np03=" + CNULL(strDataTemp(2)) + " and np04=" + CNULL(strDataTemp(3)) + " and np05=" + CNULL(strDataTemp(4)) & " and np06 is null"
               cnnConnection.Execute strTxt(iStep)
               iStep = iStep + 1
            End If
         End If
      End If
      '2010/7/16 END
      
      'add by sonia 2017/12/26土耳其235發明案公告時更新下一程序"商業使用聲明"為公告日+3年為法限,本所=法定-2月
      'modify by sonia 2020/4/4 +pa(4)="00"即EPC子案輸公告不可更新CFP-029945-0-39
      'modify by sonia 2020/7/24 土耳其加新型案
      If txtCaseField(4) <> "" And (pa(8) = "1" Or pa(8) = "2") And pa(9) = "235" And pa(4) = "00" Then
         strExc(1) = CompDate(0, 3, txtCaseField(4))   '法限
         strExc(2) = CompDate(1, -2, strExc(1))        '本所
         strExc(2) = PUB_GetWorkDay1(strExc(2), True)
         strSql = "select np22 from nextprogress" & _
            " where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'" & _
            " and np07='930' and np06 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & _
               " where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np22=" & RsTemp.Fields("np22")
         Else
            strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
               "NP07,NP08,NP09,NP10,NP22) select '" & strDataTemp(9) & "','" & pa(1) & "'" & _
               ",'" & pa(2) & "','" & pa(3) & "','" & pa(4) & "',930," & strExc(2) & "," & strExc(1) & _
               "," & CNULL(PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))) & ",NP22 from dual,(select nvl(max(np22),0)+1 NP22 from nextprogress)"
         End If
         cnnConnection.Execute strSql, intI
      End If
      'end 2017/12/26
      
   End If 'Added by Morgan 2024/12/6
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm05010402_1"
   End If
   '2016/10/7 END
   
   'Added by Morgan 2018/7/9 CFP電子化
   If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
      m_strLD18 = strDataTemp(9)
      m_strCP10 = strDataTemp(10)
      strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), m_strCP10, pa(9), cp(10))
      PUB_AddLetterProgress m_strLD18, 1 + IIf(Check1.Value + Check2.Value > 0, 1, 0), True, strExc(1), False, pa(26), m_strCP10, pa(75)
      If m_HKNP22 <> "" Then
         strExc(0) = "select pa26,pa75 from patent where " & ChgPatent(m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If PUB_AddCP1913(m_HKPA01, m_HKPA02, m_HKPA03, m_HKPA04, m_HKNP08, m_HKNP09, m_HKNP01, m_HKNP22, "013", "" & RsTemp.Fields("pa26"), m_HK1913CP09, "" & RsTemp.Fields("pa75"), , True) = False Then
            Err.Raise 999, , "新增進度檔【通知期限】失敗！作業中斷！"
         End If
      End If
      m_bolAddLP = True
   End If
   'end 2018/7/9
   
   cnnConnection.CommitTrans
   
   'Add by Morgan 2007/4/30 印結案單
   If m_HKNP22 <> "" Then
      MsgBox "請更換紙張！", , "列印接洽單！"
      g_PrtForm001.PrintForm m_HKNP22, m_HKPA01, m_HKPA02, m_HKPA03, m_HKPA04
   End If
    
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   SaveData = False
End Function
Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String
'Add By Cheng 2002/08/08
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

On Error GoTo HndErr
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetReceiveCode(frm05010402_1.txtSystem, frm05010402_1.txtCode(0), _
   IIf(frm05010402_1.txtCode(1) = "", "0", frm05010402_1.txtCode(1)), _
   IIf(frm05010402_1.txtCode(2) = "", "00", frm05010402_1.txtCode(2)), strTemp) Then
If ClsPDGetReceiveCode(frm05010402_1.txtSystem, frm05010402_1.txtCode(0), _
   IIf(frm05010402_1.txtCode(1) = "", "0", frm05010402_1.txtCode(1)), _
   IIf(frm05010402_1.txtCode(2) = "", "00", frm05010402_1.txtCode(2)), strTemp) Then
   
   'Modify by Morgan 2006/10/19 改不Call Dll
   'If objPublicData.ReadAllData(strTemp, cp(), pA(), intCaseKind, intPWhere) Then
   ReDim cp(TF_CP) As String
   cp(9) = strTemp
   If PUB_ReadAllData(cp(), pa(), intCaseKind, intPWhere) Then
   'end 2006/10/19
      lblCaseField(0) = pa(1) + " - " + pa(2) + _
         IIf(pa(4) = "00" And pa(3) = "0", "", " - " + pa(3)) + _
         IIf(pa(4) = "00", "", " - " + pa(4))
      lblCaseField(1) = pa(11)
      lblCaseField(2) = pa(8)
      lblCaseField(3) = pa(9)
      lblCaseField(4) = frm05010402_1.txtReceivedDay
      SetNameToCombo cboCaseName, pa(5), pa(6), pa(7)
      txtCaseField(0) = TransDate(pa(12), 1)
      txtCaseField(1) = pa(13)
      txtCaseField(4) = TransDate(pa(14), 1)
      txtCaseField(4).Tag = txtCaseField(4).Text 'Add by Morgan 2007/3/22
      txtCaseField(5) = pa(15)
      'Modify by Morgan 2004/5/6
      'If pa(16) = "1" Or pa(20) <> "" Then
      If pa(16) = "1" And pa(20) <> "" Then
         '公開日為空白
         If txtCaseField(0).Text = "" Then
            txtCaseField(4).SetFocus: txtCaseField_GotFocus (4)
         '若有公開日
         Else
            'Me.txtCaseField(0).Enabled = False
            '92.10.2 MODIFY BY SONIA 核准後不可改公開號, 但核駁可
            'If txtCaseField(1).Text <> "" Then
            If pa(16) = "1" And txtCaseField(1).Text <> "" Then
               Me.txtCaseField(1).Enabled = False
            End If
         End If
      End If
   End If
   'Add By Cheng 2002/08/08
   '若申請國家為英國(201)或日本(011)或韓國(012)或印尼(017)或EPC(221)或泰國(019)時, 依據專利種類判斷是否需提實審
   'Modified by Morgan 2024/8/29 取消印尼(017)--玫音
   If pa(9) = "201" Or pa(9) = "011" Or pa(9) = "012" Or pa(9) = "221" Or pa(9) = "019" Then
      StrSQLa = "Select * From Nation Where NA01='" & pa(9) & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         If pa(8) = "1" Then
            If Not IsNull(rsA.Fields("NA27")) Then
               Me.Label1.Visible = True
               Me.txtCaseField(3).Visible = True
               Me.Label5.Visible = True
               Me.txtCaseField(6).Visible = True
            End If
         ElseIf pa(8) = "2" Then
            If Not IsNull(rsA.Fields("NA29")) Then
               Me.Label1.Visible = True
               Me.txtCaseField(3).Visible = True
               Me.Label5.Visible = True
               Me.txtCaseField(6).Visible = True
            End If
         ElseIf pa(8) = "3" Then
            If Not IsNull(rsA.Fields("NA31")) Then
               Me.Label1.Visible = True
               Me.txtCaseField(3).Visible = True
               Me.Label5.Visible = True
               Me.txtCaseField(6).Visible = True
            End If
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   '若申請國家為EPC(221)時
   If pa(9) = "221" Then
      Me.Label7.Visible = True
      Me.txtCaseField(7).Visible = True
      'Add by Morgan 2005/3/14 EPC通知公開時,是否指定英國預設Y --禧佩
      txtCaseField(7) = "Y"
   End If
   '若申請國家為泰國(019)且專利種類為新型(2)時
   If pa(9) = "019" And pa(8) = "2" Then
      Me.Label11.Visible = True
      Me.txtCaseField(10).Visible = True
   End If
   '預設實審費
   If Me.txtCaseField(6).Visible Then
      StrSQLa = "SELECT NVL(YF07,0)+NVL(YF06,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='" & Left(cp(44) & "00000000", 9) & "' AND YF04='416' AND YF05=1 "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
'      20080925 取消實審費用 Modify by Toni
'         Me.txtCaseField(6).Text = "" & rsA.Fields(0).Value
     End
      Else
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         StrSQLa = "SELECT NVL(YF07,0)+NVL(YF06,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y00000000' AND YF04='416' AND YF05=1 "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            '20080925 取消實審費用 Modify by Toni
            'Me.txtCaseField(6).Text = "" & rsA.Fields(0).Value
            '20080925 End
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   '預設泰國實審費
   If Me.txtCaseField(10).Visible Then
      StrSQLa = "SELECT NVL(YF07,0)+NVL(YF06,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='" & Left(cp(44) & "00000000", 9) & "' AND YF04='416' AND YF05=1 "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         Me.txtCaseField(10).Text = "" & rsA.Fields(0).Value
      Else
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         StrSQLa = "SELECT NVL(YF07,0)+NVL(YF06,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y00000000' AND YF04='416' AND YF05=1 "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            Me.txtCaseField(10).Text = "" & rsA.Fields(0).Value
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   '預設是否已提實審
   If Me.txtCaseField(3).Visible Then
      '92.4.9 MODIFY BY SONIA
      'strSQLA = "SELECT * FROM CASEPROGRESS WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10='416' AND CP27 IS NOT NULL "
      '2009/3/27 MODIFY BY SONIA
      'StrSQLa = "SELECT * FROM CASEPROGRESS WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10='416'"
      StrSQLa = "SELECT * FROM CASEPROGRESS WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10='416' AND CP57 IS NULL "
      '92.4.9 END
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         Me.txtCaseField(3).Text = "Y"
      Else
         Me.txtCaseField(3).Text = "N"
      End If
      Me.txtCaseField(3).Enabled = False
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
End If

'Add by Morgan 2010/6/17
If pa(9) = "012" Then
   lblItemCnt.Visible = txtCaseField(6).Visible
   txtCaseField(11).Visible = txtCaseField(6).Visible
End If
'end 2010/6/17

Screen.MousePointer = varSaveCursor
Exit Sub
HndErr:
ErrorMsg
Screen.MousePointer = varSaveCursor
End Sub

Private Sub Form_Activate()
   'Add by Morgan 2004/5/6
   'If bolActive = True Then Exit Sub
   bolActive = True
   
   ReadAllData
End Sub
Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   Me.Caption = frm05010402_1.Caption
   'Add by Morgan 2004/5/6
   bolActive = False
   
   'Add By Sindy 2017/12/28
   m_strIR01 = frm05010402_1.m_strIR01
   m_strIR02 = frm05010402_1.m_strIR02
   m_strIR03 = frm05010402_1.m_strIR03
   m_strIR04 = frm05010402_1.m_strIR04
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
   'Add By Sindy 2016/10/13
   If Me.m_strIR01 = "" Then
   '2016/10/13 END
      If intLeaveKind = 1 Then
         frm05010402_1.Show
         frm05010402_1.Clear
      ElseIf intLeaveKind = 0 Then
         Unload frm05010402_1
      ElseIf intLeaveKind = 2 Then
         frm05010402_1.Show
      End If
   End If
   'Add By Cheng 2002/07/18
   Set frm05010402_2 = Nothing
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String

Select Case Index
   Case 2
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , pA(9)) Then
      If ClsPDGetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , pa(9)) Then
         lblTrademarkKind = strTemp
      End If
   Case 3
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetNation(lblCaseField(Index), strTemp) Then
      If ClsPDGetNation(lblCaseField(Index), strTemp) Then
         lblNation.Caption = strTemp
      End If
End Select
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 3, 7
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> 89 And KeyAscii <> 78 Then
      KeyAscii = 0
   End If
Case 8
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> 89 Then
      KeyAscii = 0
   End If
Case 9
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
      KeyAscii = 0
   End If

'Added by Morgan 2024/12/9
Case 2
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 Then
      KeyAscii = 0
   ElseIf KeyAscii = 51 Then
      txtCaseField(12).TabStop = True
   Else
      txtCaseField(12).TabStop = False
   End If
   
End Select
End Sub

Private Sub txtCaseField_LostFocus(Index As Integer)
'Add By Cheng 2002/08/08
Select Case Index
Case 2 '公開與否
   If Me.txtCaseField(Index).Text <> "" Then
      If Me.txtCaseField(0).Text = "" Then
         MsgBox "請輸入公開日!!!", vbExclamation + vbOKOnly
         If Me.txtCaseField(0).Enabled = True Then
            Me.txtCaseField(0).SetFocus
            TextInverse Me.txtCaseField(0)
         Else
            Me.txtCaseField(2).SetFocus
            TextInverse Me.txtCaseField(2)
         End If
         Exit Sub
      Else
         If Me.txtCaseField(Index).Text = "1" Then
            'Modified by Morgan 2025/6/16 Ex:CFP-034072--禧佩
            'If Val(Me.txtCaseField(0).Text) >= Val(strSrvDate(2)) Then
            '   MsgBox "公開日必須小於系統日!!!", vbExclamation + vbOKOnly
            If Val(Me.txtCaseField(0).Text) > Val(strSrvDate(2)) Then
               MsgBox "公開日不可大於系統日!!!", vbExclamation + vbOKOnly
               If Me.txtCaseField(0).Enabled = True Then
                  Me.txtCaseField(0).SetFocus
                  TextInverse Me.txtCaseField(0)
               Else
                  Me.txtCaseField(2).SetFocus
                  TextInverse Me.txtCaseField(2)
               End If
               Exit Sub
            End If
         ElseIf Me.txtCaseField(Index).Text = "2" Then
            'Modified by Morgan 2025/6/16
            'If Val(Me.txtCaseField(0).Text) < Val(strSrvDate(2)) Then
            '   MsgBox "公開日不可小於系統日!!!", vbExclamation + vbOKOnly
            If Val(Me.txtCaseField(0).Text) <= Val(strSrvDate(2)) Then
               MsgBox "公開日必須大於系統日!!!", vbExclamation + vbOKOnly
               If Me.txtCaseField(0).Enabled = True Then
                  Me.txtCaseField(0).SetFocus
                  TextInverse Me.txtCaseField(0)
               Else
                  Me.txtCaseField(2).SetFocus
                  TextInverse Me.txtCaseField(2)
               End If
               Exit Sub
            End If
         End If
      End If
   End If
Case 9 '公告與否
   If Me.txtCaseField(Index).Text <> "" Then
      If Me.txtCaseField(4).Text = "" Then
         MsgBox "請輸入公告日!!!", vbExclamation + vbOKOnly
         If Me.txtCaseField(4).Enabled = True Then
            Me.txtCaseField(4).SetFocus
            TextInverse Me.txtCaseField(4)
         Else
            Me.txtCaseField(9).SetFocus
            TextInverse Me.txtCaseField(9)
         End If
         Exit Sub
      Else
         If Me.txtCaseField(Index).Text = "1" Then
            'Modified by Morgan 2025/6/16
            'If Val(Me.txtCaseField(4).Text) >= Val(strSrvDate(2)) Then
            '   MsgBox "公告日必須小於系統日!!!", vbExclamation + vbOKOnly
            If Val(Me.txtCaseField(4).Text) > Val(strSrvDate(2)) Then
               MsgBox "公告日不可大於系統日!!!", vbExclamation + vbOKOnly
               If Me.txtCaseField(4).Enabled = True Then
                  Me.txtCaseField(4).SetFocus
                  TextInverse Me.txtCaseField(4)
               Else
                  Me.txtCaseField(9).SetFocus
                  TextInverse Me.txtCaseField(9)
               End If
               Exit Sub
            End If
         ElseIf Me.txtCaseField(Index).Text = "2" Then
            'Modified by Morgan 2025/6/16
            'If Val(Me.txtCaseField(4).Text) < Val(strSrvDate(2)) Then
            '   MsgBox "公告日不可小於系統日!!!", vbExclamation + vbOKOnly
            If Val(Me.txtCaseField(4).Text) <= Val(strSrvDate(2)) Then
               MsgBox "公告日必須大於系統日!!!", vbExclamation + vbOKOnly
               If Me.txtCaseField(4).Enabled = True Then
                  Me.txtCaseField(4).SetFocus
                  TextInverse Me.txtCaseField(4)
               Else
                  Me.txtCaseField(9).SetFocus
                  TextInverse Me.txtCaseField(9)
               End If
               Exit Sub
            End If
         End If
      End If
   End If

End Select

End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
CheckKeyIn = -1
Select Case intIndex
             Case 0
                  If txtCaseField(2) = "" And txtCaseField(intIndex) = "" Then
                     CheckKeyIn = 1
                     Exit Function
                  End If
                  If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                     If Me.txtCaseField(2).Text = "1" Then
                        'Modified by Morgan 2025/6/16 Ex:CFP-034072--禧佩
                        'If Val(txtCaseField(intIndex).Text) >= Val(strSrvDate(2)) Then
                        '   MsgBox "公開日必須小於系統日，請重新輸入 !", vbCritical
                        If Val(txtCaseField(intIndex).Text) > Val(strSrvDate(2)) Then
                           MsgBox "公開日不可大於系統日，請重新輸入 !", vbCritical
                        Else
                           CheckKeyIn = 1
                        End If
                     ElseIf Me.txtCaseField(2).Text = "2" Then
                        'Modified by Morgan 2025/6/16
                        'If Val(txtCaseField(intIndex).Text) < Val(strSrvDate(2)) Then
                        '   MsgBox "公開日不可小於系統日，請重新輸入 !", vbCritical
                        If Val(txtCaseField(intIndex).Text) <= Val(strSrvDate(2)) Then
                           MsgBox "公開日必須大於系統日，請重新輸入 !", vbCritical
                        Else
                           CheckKeyIn = 1
                        End If
                     Else
                        CheckKeyIn = 1
                     End If
                  End If
                  
             Case 4
                        If txtCaseField(intIndex).Text = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           If Me.txtCaseField(9).Text = "1" Then
                              'Modified by Morgan 2025/6/16
                              'If Val(txtCaseField(intIndex).Text) >= Val(strSrvDate(2)) Then
                              '   MsgBox "公告日必須小於系統日，請重新輸入 !", vbCritical
                              If Val(txtCaseField(intIndex).Text) > Val(strSrvDate(2)) Then
                                 MsgBox "公告日不可大於系統日，請重新輸入 !", vbCritical
                              Else
                                 CheckKeyIn = 1
                              End If
                           ElseIf Me.txtCaseField(9).Text = "2" Then
                              'Modified by Morgan 2025/6/16
                              'If Val(txtCaseField(intIndex).Text) < Val(strSrvDate(2)) Then
                              '   MsgBox "公告日不可小於系統日，請重新輸入 !", vbCritical
                              If Val(txtCaseField(intIndex).Text) <= Val(strSrvDate(2)) Then
                                 MsgBox "公告日必須大於系統日，請重新輸入 !", vbCritical
                              Else
                                 CheckKeyIn = 1
                              End If
                           Else
                              CheckKeyIn = 1
                           End If
                        End If
             Case 2
                        If txtCaseField(intIndex).Text = "" Then
                           CheckKeyIn = 1
                           Exit Function
                        End If
                        'Modified by Morgan 2024/12/9 +3
                        If txtCaseField(intIndex) = "1" Or txtCaseField(intIndex) = "2" Or txtCaseField(intIndex) = "3" Then
                           CheckKeyIn = 1
                        Else
                           'Modified by Morgan 2024/12/9
                           'ShowMsg MsgText(9196)
                           ShowMsg "請輸入1-3！"
                        End If
             'Add by Morgan 2010/6/17
             Case 11
                  CheckKeyIn = 1
                  If txtCaseField(11).Tag <> txtCaseField(11) Then
                     If Val(txtCaseField(11)) > 0 Then
                        Alert416Fee txtCaseField(11), pa(9), pa(157), pa(8), cp(44)
                     End If
                  End If
                  txtCaseField(11).Tag = txtCaseField(11)
             Case Else
                        CheckKeyIn = 1
End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
'儲存未修改前之值至Tag中,供再確認時使用
txtCaseField(Index).Tag = txtCaseField(Index)
'Add by Morgan 2010/6/17
'項數離開都提示金額
If Index = 11 Then
   txtCaseField(Index).Tag = ""
End If
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   
   For Each objTxt In Me.txtCaseField
      If objTxt.Enabled = True And objTxt.Visible = True Then
         Cancel = False
         txtCaseField_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   'Add by Morgan 2008/1/4
   If txtCaseField(2).Text = "" And txtCaseField(9).Text = "" Then
      MsgBox "請輸入公開或公告與否！", vbExclamation
      txtCaseField(2).SetFocus
      Exit Function
   ElseIf txtCaseField(2).Text <> "" And txtCaseField(9).Text <> "" Then
      MsgBox "公開或公告與否只能擇一輸入！", vbExclamation
      txtCaseField(2).SetFocus
      txtCaseField_GotFocus 2
      Exit Function
   End If
   
   'Add By Cheng 2002/08/08
   If Me.txtCaseField(2).Text <> "" Then
      If Me.txtCaseField(0).Text = "" Then
         MsgBox "請輸入公開日!!!", vbExclamation + vbOKOnly
         If Me.txtCaseField(0).Enabled = True Then
            Me.txtCaseField(0).SetFocus
            TextInverse Me.txtCaseField(0)
         Else
            Me.txtCaseField(2).SetFocus
            TextInverse Me.txtCaseField(2)
         End If
         Exit Function
      Else
         If Me.txtCaseField(2).Text = "1" Then
            'Modified by Morgan 2025/6/16 Ex:CFP-034072--禧佩
            'If Val(Me.txtCaseField(0).Text) >= Val(strSrvDate(2)) Then
            '   MsgBox "公開日必須小於系統日!!!", vbExclamation + vbOKOnly
            If Val(Me.txtCaseField(0).Text) > Val(strSrvDate(2)) Then
               MsgBox "公開日不可大於系統日!!!", vbExclamation + vbOKOnly
               If Me.txtCaseField(0).Enabled = True Then
                  Me.txtCaseField(0).SetFocus
                  TextInverse Me.txtCaseField(0)
               Else
                  Me.txtCaseField(2).SetFocus
                  TextInverse Me.txtCaseField(2)
               End If
               Exit Function
            End If
         ElseIf Me.txtCaseField(2).Text = "2" Then
            'Modified by Morgan 2025/6/16
            'If Val(Me.txtCaseField(0).Text) < Val(strSrvDate(2)) Then
            '   MsgBox "公開日不可小於系統日!!!", vbExclamation + vbOKOnly
            If Val(Me.txtCaseField(0).Text) <= Val(strSrvDate(2)) Then
               MsgBox "公開日必須大於系統日!!!", vbExclamation + vbOKOnly
               If Me.txtCaseField(0).Enabled = True Then
                  Me.txtCaseField(0).SetFocus
                  TextInverse Me.txtCaseField(0)
               Else
                  Me.txtCaseField(2).SetFocus
                  TextInverse Me.txtCaseField(2)
               End If
               Exit Function
            End If
         'Added by Morgan 2024/12/9
         ElseIf txtCaseField(2).Text = "3" Then
            If txtCaseField(0).Text = "" Then
               MsgBox "本案尚未公開!!!", vbExclamation + vbOKOnly
               Exit Function
            ElseIf txtCaseField(12).Text = "" Then
               MsgBox "請輸入檢索報告公開日!!!", vbExclamation + vbOKOnly
               If txtCaseField(12).Enabled Then
                  txtCaseField(12).SetFocus
               End If
               Exit Function
            
            ElseIf CheckIsTaiwanDate(txtCaseField(12).Text) = False Then
               If txtCaseField(12).Enabled Then
                  txtCaseField(12).SetFocus
               End If
               Exit Function
            ElseIf Val(txtCaseField(12).Text) > Val(strSrvDate(2)) Then
               MsgBox "檢索報告公開日不可大於系統日!!!", vbExclamation + vbOKOnly
               If txtCaseField(12).Enabled Then
                  txtCaseField(12).SetFocus
               End If
               Exit Function
            ElseIf Val(txtCaseField(12).Text) < Val(txtCaseField(0).Text) Then
               MsgBox "檢索報告公開日不可小於公開日!!!", vbExclamation + vbOKOnly
               If txtCaseField(12).Enabled Then
                  txtCaseField(12).SetFocus
               End If
               Exit Function
               
            End If
         'end 2024/12/9
         End If
      End If
   End If
   
   
   If Me.txtCaseField(9).Text <> "" Then
      'Added by Morgan 2012/4/26
      If pa(8) = "1" And pa(9) = "211" Then
         MsgBox "西班牙發明案通知核准公告請到『一般來函輸入』輸核准!!"
         Exit Function
      End If
      'end 2012/4/26
   
      If Me.txtCaseField(4).Text = "" Then
         MsgBox "請輸入公告日!!!", vbExclamation + vbOKOnly
         If Me.txtCaseField(4).Enabled = True Then
            Me.txtCaseField(4).SetFocus
            TextInverse Me.txtCaseField(4)
         Else
            Me.txtCaseField(9).SetFocus
            TextInverse Me.txtCaseField(9)
         End If
         Exit Function
      Else
         If Me.txtCaseField(9).Text = "1" Then
            'Modified by Morgan 2025/6/16
            'If Val(Me.txtCaseField(4).Text) >= Val(strSrvDate(2)) Then
            '   MsgBox "公告日必須小於系統日!!!", vbExclamation + vbOKOnly
            If Val(Me.txtCaseField(4).Text) > Val(strSrvDate(2)) Then
               MsgBox "公告日不可大於系統日!!!", vbExclamation + vbOKOnly
               If Me.txtCaseField(4).Enabled = True Then
                  Me.txtCaseField(4).SetFocus
                  TextInverse Me.txtCaseField(4)
               Else
                  Me.txtCaseField(9).SetFocus
                  TextInverse Me.txtCaseField(9)
               End If
               Exit Function
            End If
         ElseIf Me.txtCaseField(9).Text = "2" Then
            'Modified by Morgan 2025/6/16
            'If Val(Me.txtCaseField(4).Text) < Val(strSrvDate(2)) Then
            '   MsgBox "公告日不可小於系統日!!!", vbExclamation + vbOKOnly
            If Val(Me.txtCaseField(4).Text) <= Val(strSrvDate(2)) Then
               MsgBox "公告日必須大於系統日!!!", vbExclamation + vbOKOnly
               If Me.txtCaseField(4).Enabled = True Then
                  Me.txtCaseField(4).SetFocus
                  TextInverse Me.txtCaseField(4)
               Else
                  Me.txtCaseField(9).SetFocus
                  TextInverse Me.txtCaseField(9)
               End If
               Exit Function
            End If
         End If
      End If
   End If
   'Modified by Morgan 2021/4/15 加判斷通知公告不用檢查--禧佩 Ex:CFP-031927
   If Me.txtCaseField(6).Visible And txtCaseField(9) = "" Then
      If Me.txtCaseField(3).Text = "N" Then
         'Add by Morgan 2010/6/17
         If txtCaseField(11).Visible = True And txtCaseField(11) = "" Then
            If MsgBox("是否要輸入項數以便估算實審費用？", vbYesNo + vbDefaultButton1) = vbYes Then
               txtCaseField(11).SetFocus
               Exit Function
            End If
         End If
         'end 2010/6/17
         
         If pa(9) <> "221" Then 'Added by Morgan 2015/3/13 EPC 公開不檢查實審費--甄妮
            If Not IsNumeric(Me.txtCaseField(6).Text) Then
               MsgBox "實審費輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.txtCaseField(6).SetFocus
               TextInverse Me.txtCaseField(6)
               Exit Function
            End If
         End If 'Added by Morgan 2015/3/13
      End If
   End If
   If txtCaseField(4) <> "" And Me.txtCaseField(10).Visible Then
      If IsNumeric(Me.txtCaseField(10).Text) = False Then
         MsgBox "泰國實審費輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.txtCaseField(10).SetFocus
         TextInverse Me.txtCaseField(10)
         Exit Function
      End If
   End If
   
   TxtValidate = True
   
   'add by sonia 2018/10/18 泰國發明案輸入公開日號若有實審已收文未發文要提醒
   If txtCaseField(0) <> "" And pa(9) = "019" And pa(8) = "1" Then
      strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " AND CP10='" & 實體審查 & "' and cp57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "本案件已公開，請確認實審程序是否可發文!!!", vbExclamation
      End If
   End If
   'end 2018/10/18
   
End Function

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
If (Index = 0 And fraOpen.Enabled = True) Or (Index <> 0) Then
   If CheckKeyIn(Index) = -1 Then
      Cancel = True
      txtCaseField_GotFocus (Index)
      txtCaseField(Index).SetFocus
   End If
End If
If Cancel Then txtCaseField_GotFocus (Index)

End Sub

'Add By Cheng 2003/09/05
Private Function blnUpdateNP(strCaseNo As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strPA08 As String '專利種類
Dim strPA09 As String '申請國家
Dim strPA16 As String '目前准駁
Dim strPA20 As String '准駁通知日

blnUpdateNP = False
StrSQLa = "Select * From Patent Where " & ChgPatent(strCaseNo)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    strPA08 = "" & rsA("PA08").Value
    strPA09 = "" & rsA("PA09").Value
    strPA16 = "" & rsA("PA16").Value
    strPA20 = "" & rsA("PA20").Value
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    StrSQLa = "Select * From Nation Where NA01='" & strPA09 & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        Select Case strPA08
        Case "1" '發明
            '若非准後繳費者
            If "" & rsA("NA56").Value <> "Y" Then
                '新增下一程序
                blnUpdateNP = True
            Else
                '若有准駁通知日且目前准駁為准的案件
                If "" & strPA20 <> "" And "" & strPA16 = "1" Then
                    '新增下一程序
                    blnUpdateNP = True
                Else
                    '不新增下一程序
                    blnUpdateNP = False
                End If
            End If
        Case "2" '新型
            '若非准後繳費者
            If "" & rsA("NA57").Value <> "Y" Then
                '新增下一程序
                blnUpdateNP = True
            Else
                '若有准駁通知日且目前准駁為准的案件
                If "" & strPA20 <> "" And "" & strPA16 = "1" Then
                    '新增下一程序
                    blnUpdateNP = True
                Else
                    '不新增下一程序
                    blnUpdateNP = False
                End If
            End If
        Case "3" '設計
            '若非准後繳費者
            If "" & rsA("NA58").Value <> "Y" Then
                '新增下一程序
                blnUpdateNP = True
            Else
                '若有准駁通知日且目前准駁為准的案件
                If "" & strPA20 <> "" And "" & strPA16 = "1" Then
                    '新增下一程序
                    blnUpdateNP = True
                Else
                    '不新增下一程序
                    blnUpdateNP = False
                End If
            End If
        End Select
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function
'Add by Morgan 2007/4/30
Private Sub RunHKInform(p_bolEdit As Boolean)
   Dim strTxt(1 To 2) As String
   EndLetter "08", m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "&000", "12", strUserNum
   strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('08','" & m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "&000" & "','12','" & strUserNum & _
               "','本所期限','" & m_HKNP08 & "')"
   strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('08','" & m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "&000" & "','12','" & strUserNum & _
       "','下一程序','" & m_HKCP10 & "')"
   If Not ClsLawExecSQL(2, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   NowPrint m_HKPA01 & m_HKPA02 & m_HKPA03 & m_HKPA04 & "&000", "08", "12", p_bolEdit, strUserNum, , , , , , , , , , , , , m_HK1913CP09
End Sub
