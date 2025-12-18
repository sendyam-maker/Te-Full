VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010306_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-退費"
   ClientHeight    =   4788
   ClientLeft      =   -36
   ClientTop       =   2376
   ClientWidth     =   7884
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4788
   ScaleWidth      =   7884
   Begin VB.TextBox Text12 
      Height          =   270
      Left            =   5148
      TabIndex        =   49
      Top             =   4140
      Width           =   264
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   4644
      TabIndex        =   48
      Top             =   4140
      Width           =   264
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   675
      MaxLength       =   1
      TabIndex        =   5
      Top             =   4440
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   1
      Left            =   6615
      TabIndex        =   7
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   0
      Left            =   3825
      TabIndex        =   6
      Top             =   4440
      Width           =   1800
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1200
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2376
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2376
      Width           =   300
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   18
      Top             =   540
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   17
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   16
      Top             =   540
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   15
      Top             =   540
      Width           =   375
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4650
      MaxLength       =   7
      TabIndex        =   4
      Top             =   2688
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   3
      Top             =   2688
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   6888
      TabIndex        =   11
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   4836
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   5664
      TabIndex        =   10
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "繳費記錄(&R)"
      Height          =   405
      Index           =   3
      Left            =   3612
      TabIndex        =   8
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   1020
      Left            =   6240
      TabIndex        =   2
      Top             =   2430
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;1799"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   19
      Top             =   1200
      Width           =   6540
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "11536;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "7. 一案兩請新型預繳年費退費 ( 退第　　－　　年年費)"
      Height          =   180
      Index           =   7
      Left            =   1710
      TabIndex        =   47
      Top             =   4170
      Width           =   4350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收據:   　(1.附件 2.遺失 3.已作帳)"
      Height          =   180
      Index           =   5
      Left            =   240
      TabIndex        =   46
      Top             =   4470
      Width           =   2595
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "6. 預繳移作次年"
      Height          =   180
      Index           =   6
      Left            =   1710
      TabIndex        =   45
      Top             =   3945
      Width           =   1260
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "5. 預繳退費"
      Height          =   180
      Index           =   5
      Left            =   1725
      TabIndex        =   44
      Top             =   3705
      Width           =   900
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人:"
      Height          =   180
      Left            =   5250
      TabIndex        =   43
      Top             =   2430
      Width           =   945
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "4. 非年費退費"
      Height          =   180
      Index           =   4
      Left            =   1725
      TabIndex        =   42
      Top             =   3465
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   180
      X2              =   7740
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   7740
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "退費金額:"
      Height          =   180
      Index           =   2
      Left            =   5850
      TabIndex        =   41
      Top             =   4485
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "收據號碼:"
      Height          =   180
      Index           =   1
      Left            =   3060
      TabIndex        =   40
      Top             =   4485
      Width           =   765
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "3. 移做次年"
      Height          =   180
      Index           =   3
      Left            =   1725
      TabIndex        =   39
      Top             =   3225
      Width           =   900
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "2. 取消繳納"
      Height          =   180
      Index           =   2
      Left            =   1725
      TabIndex        =   38
      Top             =   2985
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   240
      TabIndex        =   37
      Top             =   2430
      Width           =   945
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容:          (Y:WORD)"
      Height          =   180
      Index           =   1
      Left            =   2235
      TabIndex        =   36
      Top             =   2430
      Width           =   2925
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   3195
      TabIndex        =   35
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3195
      TabIndex        =   34
      Top             =   1980
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   240
      TabIndex        =   33
      Top             =   1980
      Width           =   945
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   0
      Left            =   4080
      TabIndex        =   32
      Top             =   540
      Width           =   2700
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "4762;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3195
      TabIndex        =   31
      Top             =   1620
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   240
      TabIndex        =   30
      Top             =   1620
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   29
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   28
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Left            =   3195
      TabIndex        =   27
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   26
      Top             =   1260
      Width           =   765
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   25
      Top             =   900
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3440;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   2
      Left            =   4080
      TabIndex        =   24
      Top             =   930
      Width           =   2670
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "4710;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   4
      Left            =   1080
      TabIndex        =   23
      Top             =   1620
      Width           =   1980
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3492;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   5
      Left            =   4080
      TabIndex        =   22
      Top             =   1620
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "2646;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   6
      Left            =   1320
      TabIndex        =   21
      Top             =   1980
      Width           =   1710
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "3016;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   180
      Index           =   7
      Left            =   4080
      TabIndex        =   20
      Top             =   1980
      Width           =   3480
      VariousPropertyBits=   27
      Caption         =   "Label12"
      Size            =   "6138;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "下次繳費日:"
      Height          =   180
      Left            =   3675
      TabIndex        =   14
      Top             =   2730
      Width           =   945
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請書類別:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   2730
      Width           =   945
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "1. 重複繳納退費"
      Height          =   180
      Index           =   0
      Left            =   1725
      TabIndex        =   12
      Top             =   2730
      Width           =   1260
   End
End
Attribute VB_Name = "frm04010306_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/10 改成Form2.0 (Combo1,Label12,lstNameAgent)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'整理 by Morgan 2005/7/29
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/7/29 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String
Dim cp() As String 'Added by Morgan 2023/8/23
Dim m_CP110 As String
Dim m_CP22 As String
Dim intWhere As Integer
'Add by Morgan 2010/1/21
Dim m_lngRefund As Long '未退預繳年費金額
Dim m_strFromYear As String, m_strToYear As String '預繳年費起迄
Dim m_CP53 As String, m_CP54 As String '繳年費起迄年
Dim m_CP27 As String '發文日
Dim m_CP10 As String '案件性質 Add by Amy 2014/08/14
Dim m_DaulAppInvPA11 As String 'Added by Morgan 2018/1/29 一案兩請發明案申請號
'Added by Morgan 2023/8/23
Public m_CP118isY As Boolean '是否為電子送件申請書:True.是
Dim m_CaseNo As String
'end 2023/8/23

Private Function StartLetter(ByVal ET01 As String, ByVal ET03 As String) As Boolean
   Dim strTxt(200) As String, ii As Integer
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   'Added by Morgan 2023/8/23
   If m_CP118isY Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
         
      Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa())
      Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 1, m_CP110, ET01, strReceiveNo, ET03, ii, strTxt())
      
      ii = ii + 1
      strExc(0) = "重覆繳納" '3,5,6也用這個再自行修改--玲玲
      If Text6.Text = "2" Then
         strExc(0) = "取消繳納"
      ElseIf Text6.Text = "4" Then
         strExc(0) = "非年費退費"
      ElseIf Text6.Text = "7" Then
         strExc(0) = "一案兩請新型案預繳退費"
      End If
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strExc(0) & "','♀')"
            
      '辦理依據
      If Left(cp(43), 1) = "C" Then
         strExc(0) = "select * from caseprogress,edocument where cp09='" & cp(43) & "' and ed11(+)=cp09 and ed01 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文日期','" & ChangeWStringToTDateString(RsTemp("ed08")) & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','智專字','" & Mid(RsTemp("ed17"), InStr(RsTemp("ed17"), "智專") + 2) & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發文號','" & RsTemp("ed01") & "')"
         End If
      End If
   End If
   'end 2023/8/23
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','年費收據編號'," & CNULL(Text8(0).Text) & ")"
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','退費金額'," & CNULL(Text8(1).Text) & ")"
   
   
   'Added by Morgan 2018/1/29
   If Text6.Text = "7" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明案申請號','" & m_DaulAppInvPA11 & "')"
         
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','預繳起迄年','" & IIf(Text11 = Text12, Text11, Text11 & "年至第" & Text12) & "')"
   End If
   'end 2018/1/29
   
   'Add by Morgan 2010/1/29
   '附件
   If Text10 = "1" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件說明','附件：" & Text8(0).Text & "號收據正本乙紙。" & vbCrLf & "')"
   '遺失
   ElseIf Text10 = "2" Then
      ii = ii + 1
      If Text6.Text = "5" Or Text6.Text = "6" Then         '2010/5/28 ADD BY SONIA
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件說明','說明：惟因遺失　鈞局第 " & Text8(0).Text & " 號收據正本，謹請　鈞局准予退還上項費用予申請人，至感德便。" & vbCrLf & "')"
      '2010/5/28 ADD BY SONIA
      ElseIf Text6.Text = "2" Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件說明','備註：　鈞局第 " & Text8(0).Text & " 號收據正本已遺失。" & vbCrLf & "')"
      Else
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件說明','備註：惟因遺失　鈞局第 " & Text8(0).Text & " 號收據正本，謹請　鈞局准予退還上項費用予申請人，至感德便。" & vbCrLf & "')"
      End If
      '2010/5/28 END
   '已作帳
   ElseIf Text10 = "3" Then
      ii = ii + 1
      If Text6.Text = "5" Or Text6.Text = "6" Then     '2010/5/28 ADD BY SONIA
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件說明','說明：惟因　鈞局第 " & Text8(0).Text & " 號收據已作帳不易取回，謹請　鈞局准予退還上項費用予申請人，至感德便。" & vbCrLf & "')"
      '2010/5/28 ADD BY SONIA
      ElseIf Text6.Text = "2" Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件說明','備註：　鈞局第 " & Text8(0).Text & " 號收據已作帳不易取回。" & vbCrLf & "')"
      Else
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附件說明','備註：惟因　鈞局第 " & Text8(0).Text & " 號收據已作帳不易取回，謹請　鈞局准予退還上項費用予申請人，至感德便。" & vbCrLf & "')"
      End If
      '2010/5/28 END
   End If
   'end 2010/1/29
   
   If m_lngRefund > 0 Then
      If m_strFromYear = m_strToYear Then
         strExc(1) = m_strFromYear
      Else
         strExc(1) = m_strFromYear & "年至第" & m_strToYear
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','預繳起迄年','" & strExc(1) & "')"
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','可退金額','" & m_lngRefund & "')"
         
      If Val(Text8(1)) = 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','整行不印','♀')"
      End If
         
      If m_CP53 <> "" Then
         If m_CP53 = m_CP54 Then
            strExc(1) = m_CP53
         Else
            strExc(1) = m_CP53 & "年至第" & m_CP54
         End If
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','年費起迄年','" & strExc(1) & "')"
      End If
   End If
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter = True
   End If
End Function

Private Function FormSave() As Boolean
   
On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
   
   'Added by Morgan 2011/11/4 若期限有變更時先將原期限上續辦,然後新增期限
   If Text9 <> Text9.Tag Then
      'modify by sonia 2022/7/18 原上Y改上N
      strSql = "UPDATE NEXTPROGRESS SET NP06='N' WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07=" & 年費 & " AND NP06 IS NULL"
      cnnConnection.Execute strSql, intI
   'end 2011/11/4
   
      If Text9 <> "" Then
         strExc(2) = TransDate(Text9, 2)
         'Added by Morgan 2014/10/28
         If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
            strExc(3) = PUB_GetOurDeadline(strExc(2))
         Else
         'end 2014/10/28
            strExc(3) = CompDate(2, -2, strExc(2))
            strExc(3) = PUB_GetWorkDay1(strExc(3), True) 'Add by Morgan 2011/5/30
         End If 'Added by Morgan 2014/10/28
         
         strExc(4) = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4)) 'Added by Morgan 2012/10/19
         'Removed by Morgan 2011/11/4
         'strSql = "UPDATE NEXTPROGRESS SET NP08=" & strExc(3) & "," & _
         '   "NP09=" & strExc(2) & " WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07=" & 年費 & " AND NP06 IS NULL"
         'cnnConnection.Execute strSql, intI
         ''Add by Morgan 2009/11/30
         'If intI = 0 Then
         
            'Modify by Morgan 2011/11/4
            'strSql = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
               " select NP01,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
               ",'" & 年費 & "'," & strExc(3) & "," & strExc(2) & _
               ",cp13,NP22 from caseprogress,(select max(cp09) NP01 from caseprogress" & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp10 in('601','605') and cp27>0 and cp57 is null) X" & _
               ",(select max(np22)+1 NP22 from nextprogress) Y where cp09='" & strReceiveNo & "'"
            strSql = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
               " select cp09,cp01,cp02,cp03,cp04,'" & 年費 & "'," & strExc(3) & "," & strExc(2) & _
               ",'" & strExc(4) & "',NP22 from caseprogress,(select max(np22)+1 NP22 from nextprogress) Y" & _
               " where cp09='" & strReceiveNo & "'"
            'end 2011/11/4
            cnnConnection.Execute strSql, intI
            
         'End If
         ''end 2009/11/30
         
      'Removed by Morgan 2011/11/4
      'Else
      '   strSql = "DELETE FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      '      " AND NP07=" & 年費 & " AND NP06 IS NULL"
      '   cnnConnection.Execute strSql, intI
      'end 2011/11/4
      
      End If
      
   End If 'Added by Morgan 2011/11/4
   
   If lstNameAgent.Visible = True Then
      'Modified by Morgan 2023/8/23
      'strSql = " UPDATE CASEPROGRESS SET cp22=" & CNULL(m_CP22) & ",cp110=" & CNULL(m_CP110) & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cp(110) = m_CP110
      strSql = " UPDATE CASEPROGRESS SET cp22=" & CNULL(m_CP22) & ",cp110=" & CNULL(m_CP110) & IIf(m_CP118isY, ",cp118='A',cp160=''", "") & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql
   End If
   
    'Add by Amy 2014/08/14 P台灣案電子化
   If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
   If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
        '新增申請書轉檔記錄
        PUB_AddAppForm strReceiveNo
   End If
   End If
   'end 2014/08/14
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.NUMBER <> 0 Then
    cnnConnection.RollbackTrans
   End If
End Function

Private Sub cmdOK_Click(Index As Integer)
   Dim bolChk As Boolean, strTxt(1 To 2) As String, strTmp As String
   Dim lRtn As Long
   Dim strFolder As String, strFileName As String 'Added by Morgan 2023/8/23
   
   Select Case Index
      Case 0 '確定
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         If Text8(0) = "" And Text6 <> "3" Then
            MsgBox "請輸入年費收據編號 !", vbInformation
            Text8(0).SetFocus
            Exit Sub
         End If
         If Text8(1) = "" And Text6 <> "1" Then
            MsgBox "請輸入退費金額 !", vbInformation
            Text8(1).SetFocus
            Exit Sub
         End If
         
         'Add by Morgan 2010/1/29
         'Modified by Morgan 2018/1/29 +7
         If Text6 = "5" Or Text6 = "6" Or Text6 = "7" Then
            If Text10 = "" And Val(Text8(1)) Then
               MsgBox "請輸入收據選項！"
               Text10.SetFocus
               Exit Sub
            End If
            If Text6 = "6" Then
               strExc(0) = "select lastyear(pa72) from patent where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If m_CP54 <> "" & RsTemp(0) Then
                     MsgBox "最後已繳年度錯誤(應為" & m_CP54 & ")，請確認是否繳費紀錄是否正確！"
                     Exit Sub
                  End If
               Else
                  If MsgBox("無法判斷最後已繳年度，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                     Exit Sub
                  End If
               End If
            End If
         End If
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
                           
                           
         'Added by Morgan 2023/8/23
         '電子送件
         If m_CP118isY Then
            m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
            strFolder = PUB_Getdesktop & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            
            If StartLetter("01", "08") = False Then Exit Sub
            NowPrint strReceiveNo, "01", "08", False, strUserNum, , , True, strExc(9)
            
            strFileName = strFolder & "\" & m_CaseNo & ".data"
            StartLetterPA_EData "01", "14", strReceiveNo, pa, cp, False, True
            NowPrint strReceiveNo, "01", "14", False, strUserNum, , , True, strExc(10)
            Call PUB_MakeDoc(strExc(9) & Chr(12) & strExc(10), strFileName, False)
         Else
         'end 2023/8/23
                           
             If Text7 = "Y" Then
                bolChk = True
             Else
                bolChk = False
             End If
             
            '申請書類別
            Select Case Text6.Text
                Case "1"
                    '重覆繳納 01
                    strTmp = "01"
                Case "2"
                    '取消繳納 02
                    strTmp = "02"
                Case "3"
                    '移做次年 03
                    strTmp = "03"
                'Add By Cheng 2003/03/27
                Case "4"
                    '非年費退費 04
                    strTmp = "04"
                
                Case "5"
                    '預繳退費 06
                    strTmp = "06"
                    
                Case "6"
                    '預繳移做次年 07
                    strTmp = "07"
                
                Case "7"
                    '一案兩請新型案預繳退費
                    strTmp = "05"
                    
            End Select
             StartLetter "01", strTmp
             strLetterDate = Text5.Text
            'Modify by Amy 2014/08/14 +傳strLetterRecNo,修改改frm1105_1開
             NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
             If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
             If bolChk = True Then
                 frm1105_1.m_RecNo = strReceiveNo
                 'Modify By Sindy 2022/5/11 流水號要足6碼
                 frm1105_1.m_PdfName = Text1 & Text2 & IIf(Text3 & Text4 = "000", "", "-" & Text3 & "-" & Text4) & "." & m_CP10 & ".DATA.PDF"
                 frm1105_1.Show
             End If
             End If
             'end 2014/08/14
             'Add by Morgan 2010/3/30
             If m_CP27 = "" Then
                lRtn = MsgBox("是否要自動發文？", vbYesNo + vbDefaultButton1)
                If lRtn = vbYes Then
                   If PUB_CheckFormExist("frm040104_1") Then
                      MsgBox "發文畫面已開啟，請手動發文！"
                   Else
                      frm040104_1.Show
                      With frm040104_1
                         .Option1(1).Value = True
                         .Text5.Text = strReceiveNo
                         .Command1.Value = True
                      End With
                      If PUB_CheckFormExist("frm040104_3") Then
                         frm040104_3.Text7(21) = "N"
                         frm040104_3.cmdOK(0).Value = True
                         If Not PUB_CheckFormExist("frm040104_3") Then
                            Unload frm040104_1
                         End If
                      End If
                   End If
                End If
             End If
          
         End If 'Added by Morgan 2023/8/23
         
         frm040103_1.Show
         ' 90.08.27 modify by louis
         frm040103_1.ClearForm
         Unload Me
      Case 1 '回前畫面
         frm040103_1.Show
         Unload Me
      Case 2 '結束
         Unload frm040103_1
         Unload Me
      Case 3 '繳費記錄
         If Text6 = "" Then
            MsgBox "請先選擇申請書類別！", vbExclamation
            Text6.SetFocus
            Exit Sub
         ElseIf Text6 = "7" Then
            If MsgBox("修改繳費記錄會影響下次預設的退費年度與金額，是否確認要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
               Text6.SetFocus
               Exit Sub
            End If
         End If
         
         Set frm060104_b.oParent = Me 'Add by Morgan 2011/10/5
         frm060104_b.LoadMe pa(1), pa(2), pa(3), pa(4), 3
         Me.Hide
   End Select
End Sub

Private Sub Form_Load()
   'Add by Morgan 2010/1/21
   Dim bolDiscount As Boolean '年費是否可減免
   Dim lngYearFee As Long '年費規費
   
   MoveFormToCenter Me
   intWhere = 國內
   With frm040103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/7/29
   ReDim pa(TF_PA)
   ReDim cp(TF_CP) 'Added by Morgan 2023/8/23
   ReadPatent
   
   'Added by Morgan 2023/8/23
   '電子送件
   If m_CP118isY Then
      '收據一律為附件
      Text10 = "1"
      Text10.Locked = True
      '不顯示修改申請書
      Label18(1).Visible = False
      Text7.Visible = False
   End If
   'end 2023/8/23
   
   'Add by Morgan 2005/7/29
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   lstNameAgent.Clear
   If pa(9) = "000" Then
      PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True 'Modified by Morgan 2021/12/10 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/7/29 END
   
   Combo1.ListIndex = 0
   Text5 = strSrvDate(2)
   
   'Add by Morgan 2010/1/21
   '台灣年費或退費發文檢查是否有預繳年費可退
   If PUB_ChkRefund(pa, m_lngRefund, m_strFromYear, m_strToYear, True) Then
      '預繳退費
      If m_CP53 = "" Then
         Text6 = "5"
         Text8(1) = m_lngRefund
      '預繳移作次年
      Else
         Text6 = "6"
         If PUB_GetCaseDiscStat(pa(1) & pa(2) & pa(3) & pa(4)) = "Y" Then
            bolDiscount = True
         Else
            bolDiscount = False
         End If
         'Modified by Morgan 2013/9/13
         'lngYearFee = PUB_GetYearFee(pa(8), Val(m_CP53), Val(m_CP54), bolDiscount)
         PUB_GetPatentYearFee pa(9), pa(8), "Y00000001", "605", m_CP53, m_CP54, IIf(DBDATE(Text9) < strSrvDate(1), True, False), IIf(bolDiscount, "Y", ""), pa(14), Text5, strExc(1)
         lngYearFee = Val(strExc(1))
         
         If m_lngRefund < lngYearFee Then
            MsgBox "本案有可退預繳年費共 " & Format(m_lngRefund, DDollar) & " 元" & _
            "，現欲繳年費計 " & Format(lngYearFee, DDollar) & " 元" & _
            "，尚缺 " & Format(lngYearFee - m_lngRefund, DDollar) & " 元，是否請智權人員改收文【年費】並向客戶補收差額！"
         Else
            Text8(1) = m_lngRefund - lngYearFee
         End If
      End If
   End If
   Call Text6_Validate(False) 'Add By Sindy 2020/9/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010306_1 = Nothing
End Sub

Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset
 Dim Lbl As Object
 
   For Each Lbl In Label12
      Lbl.Caption = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label12(1) = pa(11)
      Label12(2) = pa(22)
      AddCboName Combo1, pa(5), pa(6), pa(6)
   End If
   
   'Added by Morgan 2023/8/23
   cp(9) = strReceiveNo
   Call PUB_ReadCaseProgressDatabase(cp(), intWhere)
   'end 2023/8/23
   
   'Modify by Amy 2014/08/14 +CP10
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2," & _
      "cp43,CP110,cp53,cp54,cp27,CP10 from caseprogress,casepropertymap,staff,staff staff1 where " & _
      "cp09='" & strReceiveNo & "' and cp01=cpm01(+) and cp10=cpm02(+) and " & _
      "cp14=staff.st01(+) and cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP10 = .Fields("CP10") 'Add by Amy 2014/08/14
      m_CP27 = "" & .Fields("cp27") 'Add by Morgan 2010/3/30
      m_CP110 = "" & .Fields("CP110")
      'Add by Morgan 2010/1/21
      m_CP53 = "" & .Fields("CP53")
      m_CP54 = "" & .Fields("CP54")
      'end 2010/1/21
      If Not IsNull(.Fields(0)) Then Label12(0) = .Fields(0)
      If Not IsNull(.Fields(1)) Then Label12(4) = .Fields(1)
      If Not IsNull(.Fields(2)) Then Label12(5) = .Fields(2)
      If Not IsNull(.Fields(3)) Then
         strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(rsTemp1.Fields(0)) Then Label12(6) = TransDate(rsTemp1.Fields(0), 1)
            If Not IsNull(rsTemp1.Fields(1)) Then Label12(7) = rsTemp1.Fields(1)
         End If
      End If
   End If
   End With
   '2006/9/13 MODIFY BY SONIA 加NP06條件
   'strExc(0) = "SELECT NP09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pA(1) & pA(2) & pA(3) & pA(4)) & " AND NP07=" & 年費
   strExc(0) = "SELECT NP09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07=" & 年費 & " AND NP06 IS NULL "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then If Not IsNull(RsTemp.Fields(0)) Then Text9 = TransDate(RsTemp.Fields(0), 1): Text9.Tag = Text9
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
End Sub

'Added by Morgan 2018/1/29
Private Sub Text11_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

'Added by Morgan 2018/1/29
Private Sub Text12_KeyPress(KeyAscii As Integer)
   If Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Morgan 2018/1/31
Private Sub Text12_Validate(Cancel As Boolean)
   If Text6 = "7" And Val(Text11) > 0 And Val(Text12) >= Val(Text11) Then
      SetYearPrice
   End If
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   'Modify by Morgan 2010/1/21 +5,6
   'Modify by Morgan 2018/1/29 +7
   If (KeyAscii < Asc("1") Or KeyAscii > Asc("7")) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
   Text8(1) = "" 'Add by Morgan 2010/3/17 改其他選項時金額要重輸 -- 玲玲
   
   'Added by Morgan 2018/1/29
   If KeyAscii = Asc("7") Then
      If Not CheckChoice Then
         KeyAscii = 0
         Beep
      End If
   Else
      Text11 = ""
      Text12 = ""
   End If
   'end 2018/1/29
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
Dim strTmpCP64 As String
   
   'Add By Sindy 2020/9/21 抓取收據號碼
   If Text6 <> "3" Then
      '要去抓代辦退費掛的相關總收文號進度備註裡的收據號碼
      strExc(0) = "select cp09||cp64 from caseprogress where CP09 in(select c.cp43 from caseprogress c where c.CP09='" & strReceiveNo & "' AND c.CP43 is not null)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTmpCP64 = RsTemp.Fields(0)
      End If
'   ElseIf Text6 = "6" Then
'      '要去抓發明申請進度備註裡的收據號碼
'      strExc(0) = "select cp09||cp64 from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10 in(" & NewCasePtyList & ") order by cp05 asc"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         strTmpCP64 = RsTemp.Fields(0)
'      End If
'   End If
'   If strTmpCP64 = "" And Text6 = "5" Then
'      '實審、再審退費
'      '要去抓代辦退費掛的相關總收文號進度備註裡的收據號碼
'      '若實審的代辦退費抓不到收據號碼則抓(911)補收款(掛實體審查的相關總收文號)
'      strExc(0) = "select cp09||cp64,cp43 from caseprogress where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='911' AND cp43 is not null AND exists (select c.cp09 from caseprogress c where c.cp09=cp43 and c.cp10='416')"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         strTmpCP64 = RsTemp.Fields(0)
'      End If
   End If
   If InStr(strTmpCP64, "收據號碼:") > 0 Then
      Text8(0) = Mid(strTmpCP64, InStr(strTmpCP64, "收據號碼:") + 5, 11)
   End If
End Sub

Private Sub Text7_GotFocus()
  TextInverse Text7
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text8_GotFocus(Index As Integer)
  TextInverse Text8(Index)
  CloseIme
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
 Static strDateOld As String
   If Text9 <> "" Then
      If Not ChkDate(Text9) Then
         MsgBox "下次繳費日期不正確，請重新輸入 !", vbCritical
         Cancel = True
         TextInverse Text9
      Else
         If Text9 <> strDateOld Then
            If MsgBox("是否確定修改 ?", vbQuestion + vbYesNo) = vbNo Then
               Cancel = True
               TextInverse Text9
            End If
         End If
      End If
   End If
   If Cancel = False Then strDateOld = Text9
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False
   If Me.Text5.Enabled = True Then
      Cancel = False
      Text5_Validate Cancel
      If Cancel = True Then
         Me.Text5.SetFocus
         Text5_GotFocus
         Exit Function
      End If
   End If
   
   Call Text6_Validate(False) 'Add By Sindy 2020/9/21
   
   'Add by Amy 2014/08/14 解 申請書類別為空導致新增ET03會錯
   If Trim(Text6) = MsgText(601) Then
       MsgBox Label17(0).Caption & "不可為空", vbCritical
       Me.Text6.SetFocus
       Text6_GotFocus
       Exit Function
   End If
   'end 2014/08/14
   
   If Me.Text9.Enabled = True Then
      Cancel = False
      Text9_Validate Cancel
      If Cancel = True Then
         Me.Text9.SetFocus
         Text9_GotFocus
         Exit Function
      End If
   End If

   'Add by Morgan 2005/7/29
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   'Added by Morgan 2018/1/29
   If Text6.Text = "7" Then
      If Val(Text11) = 0 Then
         MsgBox "請輸入退費起年！"
         Text11.SetFocus
         Exit Function
      ElseIf Val(Text12) = 0 Then
         MsgBox "請輸入退費迄年！"
         Text12.SetFocus
         Exit Function
      ElseIf Val(Text11) > Val(Text12) Then
         MsgBox "退費起年不可大於迄年！"
         Text11.SetFocus
         Exit Function
      End If
   End If
   'end 2018/1/29
   
   TxtValidate = True
End Function

'Add by Morgan 2005/7/29
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Morgan 2021/12/10 Forms2.0 改用模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      m_CP22 = ""
   Else
      m_CP22 = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

'Added by Morgan 2018/1/29
'設定一案兩請新型案預繳退費起迄年度及金額
Private Function CheckChoice() As Boolean
   Dim str605NP09 As String, iYear As Integer, iFromYear As Integer, iToYear As Integer, strInvPA14 As String
   
   'Added by Morgan 2018/1/29
   If pa(8) <> "2" Then
      MsgBox "本案非新型案，不可選 7！"
      Text6.SetFocus
      Exit Function
   End If
      
   strExc(0) = "select pa11,pa14 from (select cm05,cm06,cm07,cm08 from casemap where cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "' and cm10='3'" & _
      " union select cm01,cm02,cm03,cm04 from casemap where cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "' and cm10='3'),patent" & _
      " where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 and pa08='1'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If IsNull(RsTemp("pa14")) Then
         MsgBox "一案兩請發明案尚未公告！"
         Text6.SetFocus
         Exit Function
      Else
         strInvPA14 = "" & RsTemp("pa14")
         m_DaulAppInvPA11 = "" & RsTemp("pa11")
      End If
   Else
      MsgBox "一案兩請發明案申請號讀取失敗！"
      Text6.SetFocus
      Exit Function
   End If
   
   CheckChoice = True
   
   strExc(0) = "select lastyear(pa72) Yr from patent where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      iToYear = Val("" & RsTemp("Yr"))
      If iToYear > 0 Then
         str605NP09 = CompDate(0, iToYear, pa(14))
         str605NP09 = CompDate(2, -1, str605NP09) 'Added by Morgan 2022/5/18 新型專利權於發明公告之日須維持存續狀態 Ex:P-126314 (此處規則要與 frm04010505_2 證書號數輸入同步修改)
         iYear = (str605NP09 - strInvPA14) \ 10000 '退X年
         If iYear > 0 Then
            iFromYear = iToYear - iYear + 1  '求退費起始年
            Text11 = iFromYear
            Text12 = iToYear
            SetYearPrice
         Else
            Text11.SetFocus
            MsgBox "繳費紀錄錯誤，無法預設退費年度！", vbCritical
         End If
      End If
   End If
End Function
'計算預繳年費退費金額
Private Sub SetYearPrice()
   strExc(1) = PUB_GetCP81(pa(), strExc(2)) '是否減免/發文日
   PUB_GetPatentYearFee pa(9), pa(8), "Y00000001", 年費, Text11, Text12, False, strExc(1), pa(14), strExc(2), strExc(3)
   Text8(1) = Val(strExc(3))
End Sub
