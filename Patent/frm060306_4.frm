VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060306_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "請款通知函-讓與、合併、授權、變更、自撤、退費核准函"
   ClientHeight    =   5340
   ClientLeft      =   1185
   ClientTop       =   855
   ClientWidth     =   7035
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7035
   Begin VB.OptionButton Option1 
      Caption         =   "寄C/N"
      Height          =   255
      Index           =   1
      Left            =   1035
      TabIndex        =   2
      Top             =   4980
      Width           =   870
   End
   Begin VB.OptionButton Option1 
      Caption         =   "寄支票"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   4980
      Width           =   870
   End
   Begin VB.TextBox txtCheckNo 
      Height          =   270
      Left            =   5610
      TabIndex        =   4
      Top             =   4965
      Width           =   1365
   End
   Begin VB.TextBox txtCheckUSD 
      Height          =   270
      Left            =   3495
      TabIndex        =   3
      Top             =   4965
      Width           =   1005
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   3495
      TabIndex        =   5
      Top             =   4620
      Width           =   3465
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   " 回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4860
      TabIndex        =   7
      Top             =   12
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   48
      TabIndex        =   9
      Top             =   444
      Width           =   6975
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   7
         Left            =   1230
         TabIndex        =   38
         Top             =   2175
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   6
         Left            =   1230
         TabIndex        =   37
         Top             =   1935
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   5
         Left            =   1230
         TabIndex        =   36
         Top             =   1710
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   4
         Left            =   1230
         TabIndex        =   35
         Top             =   1470
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   3
         Left            =   1230
         TabIndex        =   34
         Top             =   1215
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   2
         Left            =   1230
         TabIndex        =   33
         Top             =   975
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   1
         Left            =   1230
         TabIndex        =   32
         Top             =   375
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   0
         Left            =   1230
         TabIndex        =   31
         Top             =   120
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   285
         Left            =   1230
         TabIndex        =   29
         Top             =   600
         Width           =   5475
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "9657;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   18
         Top             =   132
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "請款函日期"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   17
         Top             =   372
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利名稱"
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   16
         Top             =   612
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(中)"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   15
         Top             =   972
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(日)"
         Height          =   180
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   1452
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(英)"
         Height          =   180
         Index           =   5
         Left            =   60
         TabIndex        =   13
         Top             =   1212
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(中)"
         Height          =   180
         Index           =   6
         Left            =   60
         TabIndex        =   12
         Top             =   1692
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(日)"
         Height          =   180
         Index           =   7
         Left            =   60
         TabIndex        =   11
         Top             =   2172
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(英)"
         Height          =   180
         Index           =   8
         Left            =   60
         TabIndex        =   10
         Top             =   1932
         Width           =   936
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4080
      TabIndex        =   6
      Top             =   12
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6048
      TabIndex        =   8
      Top             =   12
      Width           =   800
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Left            =   1416
      MaxLength       =   1
      TabIndex        =   0
      Top             =   4635
      Width           =   255
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   1980
      TabIndex        =   43
      Top             =   3840
      Width           =   4965
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "8758;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   11
      Left            =   1290
      TabIndex        =   42
      Top             =   3600
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   1290
      TabIndex        =   41
      Top             =   3360
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   1290
      TabIndex        =   40
      Top             =   3120
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   8
      Left            =   1290
      TabIndex        =   39
      Top             =   2880
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   540
      Index           =   0
      Left            =   1260
      TabIndex        =   30
      Top             =   4050
      Width           =   5715
      VariousPropertyBits=   -1466939365
      MaxLength       =   2000
      ScrollBars      =   3
      Size            =   "10081;952"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "支票號碼"
      Height          =   180
      Index           =   18
      Left            =   4770
      TabIndex        =   28
      Top             =   5010
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "支票/CN美金"
      Height          =   180
      Index           =   17
      Left            =   2295
      TabIndex        =   27
      Top             =   5010
      Width           =   1005
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "請款單印表機"
      Height          =   180
      Index           =   11
      Left            =   2295
      TabIndex        =   26
      Top             =   4680
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "相關總收文號案件性質"
      Height          =   180
      Index           =   15
      Left            =   96
      TabIndex        =   25
      Top             =   3840
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶案件案號"
      Height          =   180
      Index           =   9
      Left            =   96
      TabIndex        =   24
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼此案號"
      Height          =   180
      Index           =   10
      Left            =   96
      TabIndex        =   23
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P. S."
      Height          =   180
      Index           =   11
      Left            =   96
      TabIndex        =   22
      Top             =   4080
      Width           =   312
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本收受人"
      Height          =   180
      Index           =   12
      Left            =   96
      TabIndex        =   21
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "副本聯絡人"
      Height          =   180
      Index           =   13
      Left            =   96
      TabIndex        =   20
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "是否修改請款函        (Y)"
      Height          =   180
      Index           =   14
      Left            =   90
      TabIndex        =   19
      Top             =   4680
      Width           =   1860
   End
End
Attribute VB_Name = "frm060306_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/16 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim intWhere As Integer, strReceiveNo As String
Dim m_PA22 As String, m_PA14 As String
Const ET01 As String = "09"
'Add By Cheng 2003/03/07
'本所案號變數
Dim m_PA01 As String, m_PA02 As String, m_PA03 As String, m_PA04 As String
Dim m_PA08 As String 'Add by Morgan 2006/7/6
Dim m_LetterLanguage As String 'Add by Morgan 2006/11/20 加日文定稿

'Add by Morgan 2009/10/13
Dim m_CP43 As String '相關收文號
Dim m_RefCP10 As String '相關收文號案件性質
Dim strPrinter As String
Dim m_stRefund As String '退費金額
Dim m_stOsDnNo As String '未收款單號
Dim m_stOsDnNt As String '未收款台幣
Dim m_stOsDnUs As String '未收款美金
Dim m_stChkNt As String '支票台幣
Dim m_stSFee As String '退費服務費


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt(15) As String, i As Integer, j As Integer, strTmp As String
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   i = 0
   '請款函日期
   If frm060306.Text5.Text <> "" Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函日期','" & DBDATE(frm060306.Text5.Text) & "')"
   End If
   If Text1(0).Text <> "" Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','請款函備註','P.S. " & ChgSQL(Text1(0).Text) & "')"
   End If
   If ET03 = "03" Or ET03 = "04" Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','下次繳年費日','" & GetEngMMDD(CompDate(2, -1, m_PA14)) & "')"
   End If
   If ET03 = "03" Then
      '至下一程序檔中找下一程序代號是繳年費及是否續辦為空，是則一般，若空的則是最後一次年費
      strExc(0) = "SELECT np09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(frm060306.Text1.Text & frm060306.Text2.Text & frm060306.Text3.Text & frm060306.Text4.Text) & _
         " AND NP07=" & 年費 & " AND NP06 IS NULL"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) <> "" Then
            i = i + 1
            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','年費法定期限'," & CNULL(DBDATE(RsTemp.Fields(0))) & ")"
         End If
      End If
   End If
   
   'Add by Morgan 2009/10/13
   If Val(m_stRefund) > 0 Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','退費金額','" & Val(m_stRefund) & "')"
   End If
   If Val(m_stSFee) > 0 Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','退費服務費','" & Val(m_stSFee) & "')"
   End If
   If m_stOsDnNo <> "" Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','未收款單號','" & m_stOsDnNo & "')"
   End If
   If Val(m_stOsDnNt) > 0 Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','未收款台幣','" & Val(m_stOsDnNt) & "')"
   End If
   If Val(m_stOsDnUs) > 0 Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','未收款美金','" & Val(m_stOsDnUs) & "')"
   End If
   If txtCheckNo <> "" Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','支票號碼','" & txtCheckNo & "')"
   End If
   If Val(m_stChkNt) > 0 Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','支票台幣','" & Val(m_stChkNt) & "')"
   End If
   If Val(txtCheckUSD) > 0 Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','支票美金','" & Val(txtCheckUSD) & "')"
   End If
   'end 2009/10/13
   
   'Added by Morgan 2014/3/26
   If Option1(1).Value = True Then
      i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','C/N要印','♀')"
         i = i + 1
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','C/N不印','♀')"
   End If
   'end 2014/3/26
   
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim bolChk As Boolean
   Dim strTmp As String
   Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
   Dim stContent As String
   Dim stDnNo As String, stCnNo As String 'Add by Morgan 2009/10/13
   Dim prnPrint As Printer
   'Added by Morgan 2014/6/3
   Dim bolDNEmail As Boolean, bolDNPlusPaper As Boolean
   Dim stCnAmt As String 'Added by Morgan 2017/8/4 折讓金額
   
   Select Case Index
      Case 1, 2
          '回前畫面
         If Index = 1 Then
            frm060306.Show
         '結束
         Else
            Unload frm060306
         End If
         Unload Me
         
      Case 0
         Screen.MousePointer = vbHourglass
         '是否修改請款函
         If Text2.Text = "Y" Then bolChk = True
         
         '定稿語文
         m_LetterLanguage = PUB_GetLanguage(m_PA01, m_PA02, m_PA03, m_PA04)
         'Add by Morgan 2008/3/24 判斷是否產生電子檔
         bolEmail = PUB_GetEMailFlag(m_PA01 & m_PA02 & m_PA03 & m_PA04, , , bolPlusPaper)
         
         'Added by Morgan 2014/6/3
         If bolEmail = False Then
            bolDNEmail = PUB_GetEMailFlag(m_PA01 & m_PA02 & m_PA03 & m_PA04, , , bolDNPlusPaper, , True)
         Else
            bolDNEmail = bolEmail
            bolDNPlusPaper = bolPlusPaper
         End If
         'end 2014/6/3

         'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
         If bolPlusPaper Then
            iCopy = 0
         Else
            iCopy = 1
         End If
         'end 2009/10/20
                        
         'Add by Morgan 2009/10/13 +退費核准函
         If m_RefCP10 = "908" Then
            m_stOsDnNo = "": m_stOsDnNt = 0: m_stOsDnUs = 0: m_stChkNt = 0
            stCnNo = "": stDnNo = ""
            strExc(0) = "select c2.cp19,c2.cp60,c2.cp86,k2.a1k01,k2.a1k08,k2.a1k11,k2.a1k29,k1.a1k11 SFee,k2.a1k06,k2.a1k07 from caseprogress c1,caseprogress c2,caseprogress c3,acc1k0 k1,acc1k0 k2" & _
               " where c1.cp09='" & strReceiveNo & "' and c2.cp09(+)=c1.cp43 and k1.a1k01(+)=c2.cp60 and c3.cp09(+)=c2.cp43 and k2.a1k01(+)=c3.cp60"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_stRefund = "" & RsTemp.Fields("cp19")
               m_stSFee = "" & RsTemp.Fields("SFee")
               '未收款
               If IsNull(RsTemp.Fields("a1k29")) Then
                  'Modified by Morgan 2014/3/26 +日文定稿
                  If m_LetterLanguage = "3" Then
                     strTmp = "11"
                  Else
                     strTmp = "08"
                  End If
                  m_stOsDnNo = "" & RsTemp.Fields("a1k01")
                  m_stOsDnNt = "" & RsTemp.Fields("a1k11")
                  m_stOsDnUs = "" & RsTemp.Fields("a1k08")
                  stDnNo = "" & RsTemp.Fields("cp60")
                  stCnNo = m_stOsDnNo
                  stCnAmt = Val("" & RsTemp.Fields("a1k06")) 'Added by Morgan 2017/8/4 折讓金額
               'Added by Morgan 2014/3/26
               ElseIf (Option1(0).Value Or Option1(1).Value) = False Then
                  MsgBox "請選擇 寄支票 或 寄C/N !!", vbInformation
                  Screen.MousePointer = vbDefault
                  Exit Sub
               ElseIf txtCheckUSD = "" Then
                  MsgBox "請輸入美金金額 !!"
                  txtCheckUSD.SetFocus
                  Screen.MousePointer = vbDefault
                  Exit Sub
               ElseIf Option1(0).Value = True And txtCheckNo = "" Then
                  MsgBox "請輸入支票號碼 !!", vbInformation
                  txtCheckNo.SetFocus
                  Screen.MousePointer = vbDefault
                  Exit Sub
               'end 2014/3/26
               '無相反指示
               ElseIf IsNull(RsTemp.Fields("cp86")) Then
                  'Modified by Morgan 2014/3/26 +日文定稿
                  If m_LetterLanguage = "3" Then
                     strTmp = "09"
                  Else
                     strTmp = "06"
                  End If
                  m_stChkNt = Val(m_stRefund) - Val(m_stSFee)
               '有相反指示
               Else
                  'Modified by Morgan 2014/3/26 +日文定稿
                  If m_LetterLanguage = "3" Then
                     strTmp = "10"
                  Else
                     strTmp = "07"
                  End If
                  m_stChkNt = m_stRefund
                  stDnNo = "" & RsTemp.Fields("cp60")
               End If
            End If
            
            If stDnNo <> "" Then
               If MsgBox("準備列印請款單，請更換紙張！", vbYesNo + vbDefaultButton1) = vbYes Then
                  Load Frmacc2480
                  With Frmacc2480
                     .Text1.Text = stDnNo
                     .Text2.Text = stDnNo
                     .Combo1.Text = Me.Combo2.Text
                     .m_bBeCalled = True
                     .m_CallPrevForm = Me.Name  'Added by Lydia 2020/01/06 呼叫請款單的程式名稱
                     'Modified by Morgan 2014/6/3
                     '.m_bEMail = bolEmail
                     '.m_bPaper = bolPlusPaper
                     .m_bEMail = bolDNEmail
                     .m_bPaper = bolDNPlusPaper
                     'end 2014/6/3
                     .Command2_Click: DoEvents
                  End With
                  Unload Frmacc2480: DoEvents
               End If
            End If
            
            'Removed by Morgan 2019/4/18 流程已改為固定由承辦提供折讓單(單號財務處給)，程序不必再輸折讓--敏莉,婉莘
            'If stCnNo <> "" Then
            '   'Added by Morgan 2017/8/4 有輸折讓財要印折讓單,否則彈訊息
            '   If Val(stCnAmt) = 0 Then
            '      MsgBox "尚未輸入折讓，請承辦提供折讓單！", vbInformation
            '   Else
            '   'end 2017/8/4
            '      If MsgBox("準備列印折讓單，請更換紙張！", vbYesNo + vbDefaultButton1) = vbYes Then
            '         For Each prnPrint In Printers
            '            If prnPrint.DeviceName = Combo2 Then
            '               Set Printer = prnPrint
            '               Exit For
            '            End If
            '         Next
            '         Load Frmacc24h0
            '         With Frmacc24h0
            '            .Option1.Value = True
            '            .Text1.Text = stCnNo
            '            .m_iCopy = 2
            '            .Command2_Click: DoEvents
            '         End With
            '         Unload Frmacc24h0: DoEvents
            '      End If
            '   End If
            'End If
            'end 2019/4/18
         Else
         'end 2009/10/13
         
            If m_LetterLanguage = "3" Then
               strTmp = "05"
            Else
               '無專利權號數
               If m_PA22 = "" Then
                  strTmp = "02"
               '有專利權號數號
               Else
                  strTmp = "03"
                  '判斷下一程序是否有年費期限
                  If getNP605(m_PA01, m_PA02, m_PA03, m_PA04) = False Then strTmp = "04"
               End If
            End If
         End If
         
         StartLetter ET01, strTmp
         
         '英文有專利號數才印專利證書譯文
         If m_LetterLanguage = "2" And m_PA22 <> "" Then
            '要修改
            If bolChk Then
               '要傳Save2File參數來控制是否含信頭
               NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum, , , True, stContent, , , , bolEmail
               
               If Val(m_PA14) > 0 And Val(m_PA14) < 930701 Then
                  '沒有年費期限的不印期限表
                  If strTmp = "04" Then
                     NowPrint strReceiveNo, "07", "10", bolChk, strUserNum, , stContent, , , , , , bolEmail
                  Else
                     NowPrint strReceiveNo, "07", "10", bolChk, strUserNum, , stContent, True, stContent, , , , bolEmail
                  End If
               Else
                  If m_PA08 = "2" Then
                     '沒有年費期限的不印期限表
                     If strTmp = "04" Then
                        NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "16", bolChk, strUserNum, , stContent, , , , , , bolEmail
                     Else
                        NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "16", bolChk, strUserNum, , stContent, True, stContent, , , , bolEmail
                     End If
                  Else
                     '沒有年費期限的不印期限表
                     If strTmp = "04" Then
                        NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "15", bolChk, strUserNum, , stContent, , , , , , bolEmail
                     Else
                        NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "15", bolChk, strUserNum, , stContent, True, stContent, , , , bolEmail
                     End If
                  End If
               End If
               '年費期限表(目前與證書函共用)
               If strTmp <> "04" Then
                  NowPrint strReceiveNo, "07", "12", bolChk, strUserNum, , stContent, , , , , , bolEmail
               End If
            '不修改
            Else
               '不EMail
               If Not bolEmail Then
                  NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum
                  'Modify by Morgan 2006/7/6 '公告日分新舊法
                  If Val(m_PA14) > 0 And Val(m_PA14) < 930701 Then
                      NowPrint strReceiveNo, "07", "10", bolChk, strUserNum
                  Else
                      If m_PA08 = "2" Then
                         NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "16", bolChk, strUserNum
                      Else
                         NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "15", bolChk, strUserNum
                      End If
                  End If
                  '年費期限表(目前與證書函共用)
                  If strTmp <> "04" Then
                     NowPrint strReceiveNo, "07", "12", bolChk, strUserNum
                  End If
               '要EMail
               Else
                  NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum, , , , , iCopy
                  NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum, , , True, stContent, , , , bolEmail
                  'Modify by Morgan 2006/7/6 '公告日分新舊法
                  If Val(m_PA14) > 0 And Val(m_PA14) < 930701 Then
                     NowPrint strReceiveNo, "07", "10", bolChk, strUserNum, , , , , iCopy
                     '沒有年費期限的不印期限表
                     If strTmp = "04" Then
                        NowPrint strReceiveNo, "07", "10", bolChk, strUserNum, , stContent, , , , , True, bolEmail
                     Else
                        NowPrint strReceiveNo, "07", "10", bolChk, strUserNum, , stContent, True, stContent, , , , bolEmail
                     End If
                  Else
                     If m_PA08 = "2" Then
                        NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "16", bolChk, strUserNum, , , , , iCopy
                        '沒有年費期限的不印期限表
                        If strTmp = "04" Then
                           NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "16", bolChk, strUserNum, , stContent, , , , , True, bolEmail
                        Else
                           NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "16", bolChk, strUserNum, , stContent, True, stContent, , , , bolEmail
                        End If
                     Else
                        NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "15", bolChk, strUserNum, , , , , iCopy
                        '沒有年費期限的不印期限表
                        If strTmp = "04" Then
                           NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "15", bolChk, strUserNum, , stContent, , , , , True, bolEmail
                        Else
                           NowPrint m_PA01 & m_PA02 & m_PA03 & m_PA04 & "&1603", "07", "15", bolChk, strUserNum, , stContent, True, stContent, , , , bolEmail
                        End If
                     End If
                  End If
                  '年費期限表(目前與證書函共用)
                  If strTmp <> "04" Then
                     NowPrint strReceiveNo, "07", "12", bolChk, strUserNum, , , , , iCopy
                     NowPrint strReceiveNo, "07", "12", bolChk, strUserNum, , stContent, , , , , True, bolEmail
                  End If
                  If bolChk = False Then
                     MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_PA01) & " ]！"
                  End If
               End If
            End If
            
         Else
            '要存檔
            If bolEmail Then
               '要傳Save2File參數來控制是否含信頭
               NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum, , , , , iCopy, , True, True
               If bolChk = False Then
                  MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_PA01) & " ]！"
               End If
            '不存檔
            Else
               NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum
            End If
         End If
         
         If Not bolEmail Or bolPlusPaper Then
'            'Add By Sindy 2015/9/21 日文定稿才要印地址條
'            If m_LetterLanguage = "3" Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'            '2015/9/21 END
            'Add By Sindy 2017/3/20 日文定稿才要印地址條
            If frm060306.m_FCna01 = "101" Or m_LetterLanguage = "3" Then '美國 或 日文定稿才要印地址條
            '2017/3/20 END
               '新增地址條列表資料
               pub_AddressListSN = pub_AddressListSN + 1
               PUB_AddNewAddressList strUserNum, frm060306.Text1.Text, frm060306.Text2.Text, frm060306.Text3.Text, frm060306.Text4.Text, "" & pub_AddressListSN, "0"
            End If
         End If
         
         frm060306.Show
         frm060306.Clear
         Screen.MousePointer = vbDefault
         Unload Me
   End Select
End Sub
'92.1.17 add by sonia
Public Function GetEngMMDD(ByVal strValue As String) As String
Dim strTmp As String
Dim ii As Integer
Dim arrTmp
   
GetEngMMDD = ""
'若有傳入值
If strValue <> "" Then
    arrTmp = Split(strValue, "; ")
    For ii = 0 To UBound(arrTmp)
        Select Case Mid(arrTmp(ii), 5, 2)
           Case "01": strTmp = "January "
           Case "02": strTmp = "February "
           Case "03": strTmp = "March "
           Case "04": strTmp = "April "
           Case "05": strTmp = "May "
           Case "06": strTmp = "June "
           Case "07": strTmp = "July "
           Case "08": strTmp = "August "
           Case "09": strTmp = "September "
           Case "10": strTmp = "October "
           Case "11": strTmp = "November "
           Case "12": strTmp = "December "
        End Select
        GetEngMMDD = GetEngMMDD & strTmp & Right(strValue, 2)
        '93.3.25 ADD BY SONIA
        If GetEngMMDD = "February 29" Then
           GetEngMMDD = "February 28"
        End If
        '93.3.25 END
    Next ii
Else
   GetEngMMDD = ""
End If
End Function
'92.1.17 end

Private Sub Form_Unload(Cancel As Integer)
   '若請款單印表機變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   Set frm060306_4 = Nothing
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   ReadPatent
   'Add by Morgan 2009/10/13
   '設定請款單印表機
   PUB_SetPrinter Me.Name, Combo2, strPrinter
End Sub

Private Sub ReadPatent()
 'edit by nickc 2007/02/02
 'Dim lbl As Label, pA(1 To T_PA) As String, i As Integer, strTmp As String
 Dim Lbl As Object, pa() As String, i As Integer, strTmp As String
 'add by nickc 2007/02/02
 ReDim pa(1 To TF_PA) As String
 
   For Each Lbl In Label2
      Lbl = ""
   Next
   strReceiveNo = frm060306.Tag
   pa(1) = frm060306.Text1.Text
   pa(2) = frm060306.Text2.Text
   pa(3) = frm060306.Text3.Text
   pa(4) = frm060306.Text4.Text
    'Add By Cheng 2003/03/07
    '記錄本所案號
   m_PA01 = frm060306.Text1.Text
   m_PA02 = frm060306.Text2.Text
   m_PA03 = frm060306.Text3.Text
   m_PA04 = frm060306.Text4.Text
   Label2(0).Caption = GiveSymbol(pa(1), pa(2), pa(3), pa(4))
   Label2(1).Caption = frm060306.Text5.Text
   SetComboToCombo Combo1, frm060306.Combo1
   
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa, intWhere) Then  'edit by nickc 2007/02/02 不用 dll 了  If objPublicData.ReadPatentDatabase(pA, intWhere) Then
            If PA51CU58FA07(pa) Then
               For i = 1 To 6
                  Label2(i + 1) = pa(50 + i)
               Next
            End If
            Label2(8) = pa(48)
            Label2(9) = pa(77)
            If pa(86) <> "" Then
               'edit by nickc 2007/02/05 不用 dll 了
               'If objLawDll.LawGetName(pa(86), strTmp) Then Label2(10) = strTmp
               If ClsLawLawGetName(pa(86), strTmp) Then Label2(10) = strTmp
            End If
            Label2(11) = pa(87)
         End If
      Case "FG"
         If ClsPDReadServicePracticeDatabase(pa, intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadServicePracticeDatabase(pA, intWhere) Then
            If PA51CU58FA07(pa) Then Label2(2) = pa(30)
            Label2(8) = pa(29)
            Label2(9) = pa(27)
            'edit by nickc 2007/02/05 不用 dll 了
            'If objLawDll.LawGetName(pa(35), strTmp) Then Label2(10) = strTmp
            If ClsLawLawGetName(pa(35), strTmp) Then Label2(10) = strTmp
            Label2(11) = pa(36)
         End If
   End Select
   
   m_PA22 = "": m_PA14 = ""
    'Modify By Cheng 2003/01/22
'   If IsNull(pa(22)) Then
'      m_PA22 = "N"  '無專利號數
'   End If
    m_PA22 = "" & pa(22)
   If pa(14) <> "" Then
      m_PA14 = pa(14)  '公告日期
   End If
   
   m_PA08 = pa(8) 'Add by Morgan 2006/7/6
   
   strExc(0) = "SELECT CPM03,CP09,CP10 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP09 IN " & _
      "(SELECT CP43 FROM CASEPROGRESS WHERE CP09='" & strReceiveNo & "') AND CP01=CPM01(+) AND CP10=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Label2(12) = RsTemp.Fields(0)
      'Add by Morgan 2009/10/13
      m_CP43 = RsTemp.Fields("cp09")
      m_RefCP10 = RsTemp.Fields("cp10")
   End If
End Sub

'Add By Cheng 2003/03/07
'判斷下一程序605, 且NP06 IS NULL
Private Function getNP605(strNP02 As String, strNP03 As String, strNP04 As String, strNP05 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

getNP605 = False
StrSQLa = "Select * From NextProgress Where NP02='" & strNP02 & "' And NP03='" & strNP03 & "' And NP04='" & strNP04 & "' And NP05='" & strNP05 & "' " & _
                " And NP07='605' And NP06 IS NULL "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    getNP605 = True
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function
