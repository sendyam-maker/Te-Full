VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060306_7 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "½Ð´Ú³qª¾¨ç-¹êÅé¼f¬d¡B¥[µù±M§QÅv©µªø¡B´£¦­¤½¶}½Ð´Ú¨ç"
   ClientHeight    =   5472
   ClientLeft      =   756
   ClientTop       =   1548
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5472
   ScaleWidth      =   7140
   Begin VB.TextBox txtPAID 
      Height          =   270
      Left            =   4140
      MaxLength       =   1
      TabIndex        =   1
      Top             =   5055
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   " ¦^«eµe­±(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4890
      TabIndex        =   3
      Top             =   36
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1455
      MaxLength       =   1
      TabIndex        =   0
      Top             =   4680
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6084
      TabIndex        =   4
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4110
      TabIndex        =   2
      Top             =   36
      Width           =   756
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   105
      TabIndex        =   5
      Top             =   468
      Width           =   6975
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   7
         Left            =   1200
         TabIndex        =   33
         Top             =   2145
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   6
         Left            =   1200
         TabIndex        =   32
         Top             =   1905
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   5
         Left            =   1200
         TabIndex        =   31
         Top             =   1680
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   4
         Left            =   1200
         TabIndex        =   30
         Top             =   1440
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   3
         Left            =   1200
         TabIndex        =   29
         Top             =   1185
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   2
         Left            =   1200
         TabIndex        =   28
         Top             =   945
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   1
         Left            =   1200
         TabIndex        =   27
         Top             =   405
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   0
         Left            =   1200
         TabIndex        =   26
         Top             =   150
         Width           =   5565
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9816;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Top             =   600
         Width           =   5475
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "9657;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(­^)"
         Height          =   180
         Index           =   8
         Left            =   48
         TabIndex        =   14
         Top             =   1932
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(¤é)"
         Height          =   180
         Index           =   7
         Left            =   48
         TabIndex        =   13
         Top             =   2172
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(¤¤)"
         Height          =   180
         Index           =   6
         Left            =   48
         TabIndex        =   12
         Top             =   1692
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(­^)"
         Height          =   180
         Index           =   5
         Left            =   48
         TabIndex        =   11
         Top             =   1212
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(¤é)"
         Height          =   180
         Index           =   4
         Left            =   48
         TabIndex        =   10
         Top             =   1452
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(¤¤)"
         Height          =   180
         Index           =   3
         Left            =   48
         TabIndex        =   9
         Top             =   972
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "±M§Q¦WºÙ"
         Height          =   180
         Index           =   2
         Left            =   48
         TabIndex        =   8
         Top             =   612
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "½Ð´Ú¨ç¤é´Á"
         Height          =   180
         Index           =   1
         Left            =   48
         TabIndex        =   7
         Top             =   372
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥»©Ò®×¸¹"
         Height          =   180
         Index           =   0
         Left            =   48
         TabIndex        =   6
         Top             =   132
         Width           =   720
      End
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   11
      Left            =   1350
      TabIndex        =   37
      Top             =   3600
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   1350
      TabIndex        =   36
      Top             =   3360
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   1350
      TabIndex        =   35
      Top             =   3120
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   8
      Left            =   1350
      TabIndex        =   34
      Top             =   2880
      Width           =   5685
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10028;317"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1455
      TabIndex        =   25
      Top             =   5040
      Width           =   255
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "450;503"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   840
      Index           =   0
      Left            =   1350
      TabIndex        =   24
      Top             =   3810
      Width           =   5685
      VariousPropertyBits=   -1466939365
      MaxLength       =   2000
      ScrollBars      =   3
      Size            =   "10028;1482"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblPAID 
      AutoSize        =   -1  'True
      Caption         =   "¤w¦¬´Ú:           (1-¤£±HD/N, 2-±HD/N)"
      Height          =   180
      Left            =   3480
      TabIndex        =   22
      Top             =   5100
      Width           =   2700
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "·s¼W½Ð´Ú³æ¡G             (Y:·s¼W½Ð´Ú³æ)"
      Height          =   180
      Index           =   9
      Left            =   165
      TabIndex        =   21
      Top             =   5100
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_­×§ï½Ð´Ú¨ç         (Y)"
      Height          =   180
      Index           =   14
      Left            =   135
      TabIndex        =   20
      Top             =   4710
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "°Æ¥»Ápµ¸¤H"
      Height          =   180
      Index           =   13
      Left            =   168
      TabIndex        =   19
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "°Æ¥»¦¬¨ü¤H"
      Height          =   180
      Index           =   12
      Left            =   168
      TabIndex        =   18
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P. S."
      Height          =   180
      Index           =   11
      Left            =   180
      TabIndex        =   17
      Top             =   3780
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "©¼¦¹®×¸¹"
      Height          =   180
      Index           =   10
      Left            =   168
      TabIndex        =   16
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "«È¤á®×¥ó®×¸¹"
      Height          =   180
      Index           =   9
      Left            =   168
      TabIndex        =   15
      Top             =   2880
      Width           =   1080
   End
End
Attribute VB_Name = "frm060306_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/10/26 ¤é¤å¤w§ï©ñ©w½Z¤º
'Memo By Sindy 2021/7/16 Form2.0¤w­×§ï
'Memo By Morgan 2012/12/10 ´¼Åv¤H­ûÄæ¤w­×§ï
'2010/12/6 memo by sonia ­û¤u½s¸¹Äæ¤w­×§ï
'Memo by Morgan2010/8/13 ¤é´ÁÄæ¤w­×§ï
'Modify by Morgan 2009/4/7 ¤£¥i¦AÂI¥D°Ê­×¥¿(­Y¦³¥D°Ê­×¥¿®É·|¥Ñ¤uµ{®v½Ð´Ú)--David
Option Explicit

Dim intWhere As Integer, strReceiveNo As String
Dim m_CP10 As String, m_A1J17 As Long, m_A1J17_1 As Long
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
Dim strPrinter As String
Const ET01 As String = "09"
'Add By Cheng 2003/06/26
Dim m_blnContinuous As Boolean '¬O§_Ä~Äò
'Add by Morgan 2004/12/29
Dim m_iLanguage As Integer '©w½Z»y¤å
Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean, m_iCopy As Integer
'Add by Morgan 2010/8/5 ­n¦P®É½Ð´Úªº¦¬¤å¸¹,³W¶O
Dim m_RefCP09s As String, m_lngFee938 As Long, m_lngFee939 As Long
'Added by Morgan 2014/6/3
Dim m_bolDNEmail As Boolean, m_bolDNPlusPaper As Boolean
Dim m_bSpecial As Boolean 'Added by Morgan 2015/6/8
Public m_CallName As String 'Added by Lydia 2020/08/17 ©I¥sªºªí³æ¦WºÙ
'Added by Lydia 2021/01/21
Dim m_CP60 As String '½Ð´Ú³æ¸¹


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(1 To 10) As String, i As Integer, j As Integer, strTmp As String
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   i = 1
    'Add By Cheng 2003/02/24
   '½Ð´Ú¨ç¤é´Á
   If frm060306.Text5.Text <> "" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç¤é´Á','" & DBDATE(frm060306.Text5.Text) & "')"
      i = i + 1
   End If
   '½Ð´Ú¨ç³Æµù
   If Text1(0).Text <> "" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç³Æµù','P.S. " & ChgSQL(Text1(0).Text) & "')"
      i = i + 1
   'Add By Sindy 2017/3/30 µL½Ð´Ú¨ç³Æµù¤~­n¦L
   Else
      'Added by Lydia 2020/08/17 ¤w¦¬´Ú¤£±HD/N=¦³½Ð´Ú¨ç³Æµù®É¤£¦L
      If Me.txtPAID = "1" Then
      Else
      'end 2020/08/17
            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªA°È¶O´£¿ô','¡ð')"
            i = i + 1
      End If 'Added by Lydia 2020/08/17
   '2017/3/30 END
   End If
   
   'Added by Lydia 2020/08/17 ¤é¤å©w½Z¦³"¤w¦¬´Ú¤£±HD/N",¤~­n¦b¤º¤å®³±¼±b³æ³Æµù
   If Me.txtPAID = "1" And m_iLanguage = 3 Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤w¦¬´Ú¤£±HDN','¡ð')"
      i = i + 1
   End If
   
   'Modified by Morgan 2013/10/17
'   'Add by Morgan 2010/6/28
'   If Not PUB_ChkCPExist(pa, "422") Then
'      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥[³t¼f¬d´£¿ô','¡ð')"
'      i = i + 1
'   End If
   strExc(0) = "select cp10 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('422','431') and cp57 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥[³t©Î°ª³t¼f¬d','" & IIf(RsTemp(0) = "422", "AEP", "PPH") & "')"
      i = i + 1
      'Modified by Morgan 2022/10/26
      'strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥[³t©Î°ª³t¼f¬d/¤é','" & IIf(RsTemp(0) = "422", "¥[³t¼f¬dÇU½Ð¨D¤â“dþà(AEP)", "°ª³t¼f¬dÇU½Ð¨D¤â“dþà(PPH)") & "')"
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & IIf(RsTemp(0) = "422", "¥[³t¼f¬d­n¦L", "°ª³t¼f¬d­n¦L") & "','¡ð')"
      i = i + 1
   Else
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥[³t¼f¬d´£¿ô','¡ð')"
      i = i + 1
   End If
   'end 2013/10/17
   
   'Add by Morgan 2011/7/20
   'Modified by Morgan 2013/10/18 §ï©w½Z+X55340--David
   'If m_CP10 = "416" And InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X21199") > 0 Then
   'Modified by Morgan 2014/1/28 +Y45268
   'Modified by Morgan 2014/2/17 +FCP-43077
   'Modified by Morgan 2014/8/20 +FCP-50499
   'Modified by Morgan 2015/2/6 +X59055
   'Modified by Morgan 2015/8/13 +X74121,X74123
   'Modified by Morgan 2016/4/26 +Y30131020
   'Modified by Morgan 2017/1/20 +FCP-050258 --Anny
   'Modified by Morgan 2019/2/15 +Y53480 Kateeva, Inc. --Joseph
   If m_CP10 = "416" And (InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X21199") > 0 _
      Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X55340") > 0 _
      Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X59055") > 0 _
      Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X74121") > 0 _
      Or InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X74123") > 0 _
      Or pa(75) = "Y30131020" Or Left(pa(75), 6) = "Y45268" Or Left(pa(75), 6) = "Y53480" _
      Or InStr("FCP043077000,FCP050499000,FCP050258000", pa(1) & pa(2) & pa(3) & pa(4)) > 0) Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','³qª¾¤À³Î®É¶¡','¡ð')"
      i = i + 1
   End If
   
   'Added by Morgan 2016/10/13
   If ChangeCustomerL(pa(75)) = "Y54391000" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µo²¼','µo²¼')"
      i = i + 1
   End If
   'end 2016/10/13
   
   If i <> 1 Then
       'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
       'If Not objLawDll.ExecSQL(i - 1, strTxt) Then
       If Not ClsLawExecSQL(i - 1, strTxt) Then
           MsgBox "Àx¦s¨Ò¥~Äæ¦ì¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
       End If
   End If
End Sub

'Add By Sindy 2017/3/20
'Private Sub cmdok_Click(Index As Integer)
Public Sub cmdOK_Click(Index As Integer)
'2017/3/20 END
Dim bolChk As Boolean
'92.11.26 ADD BY SONIA
Dim Cancel As Boolean
Dim stET03 As String 'Add by Morgan 2004/10/13
'Added by Lydia 2021/01/21
Dim iCopy As Integer, m_NationID As String
Dim strCaseNo As String, strFileName As String, strFullFileName As String

   Select Case Index
      Case 2 'µ²§ô
         Unload frm060306
         Unload Me
      Case 0 '½T©w
        Screen.MousePointer = vbHourglass
        '92.11.26 ADD BY SONIA
         Text1_Validate 1, Cancel
         If Cancel = True Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
        '92.11.26 END
         If Text2.Text = "Y" Then bolChk = True
         '¨ú±o©w½Z»y¤å
         '½Ð¦A°Ï¤À­^¤å02©Î¤é¤å03
         
         'Modify by Morgan 2004/12/29
         'StartLetter ET01, "02"
         'NowPrint strReceiveNo, ET01, "02", bolChk, strUserNum, 0
         'Modify by Morgan 2006/6/2
         'm_iLanguage = GetLetterLanguage(pA(1), pA(2), pA(3), pA(4))
         m_iLanguage = PUB_GetLanguage(pa(1), pa(2), pa(3), pa(4))
         stET03 = "00" '¹w³]
         Select Case m_iLanguage
            'Added by Morgan 2016/8/18
            Case 1 '¤¤¤å
               stET03 = "01"
            'end 2016/8/18
            Case 3 '¤é¤å
               Select Case m_CP10
                  Case "416" '¹êÅé¼f¬d
                     stET03 = "03"
                  Case "417" '´£¦­¤½¶}
                     stET03 = "01"
               End Select
            Case Else
               Select Case m_CP10
                  Case "416" '¹êÅé¼f¬d
                     stET03 = "02"
                  Case "608", "417" '¥[µù±M§QÅv©µªø608¡B´£¦­¤½¶}417
                     stET03 = "00"
               End Select
         End Select
         '2004/12/29 end
         StartLetter ET01, stET03
         'Add by Morgan 2008/3/31 §PÂ_¬O§_²£¥Í¹q¤lÀÉ
         m_bolEmail = PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4), , , m_bolPlusPaper)
         'Added by Morgan 2014/6/3
         If m_bolEmail = False Then
            m_bolDNEmail = PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4), , , m_bolDNPlusPaper, , True)
         Else
            m_bolDNEmail = m_bolEmail
            m_bolDNPlusPaper = m_bolPlusPaper
         End If
         'end 2014/6/3
         
         'Added by Lydia 2021/01/21 ¹ê¼f(¹êÅé¼f¬d)µo¤å¡G³B²z©w½Z¡B±b³æ
         'Modified by Morgan 2024/11/21 +447¦A¼f¬d¥[³t¼f¬d
         If m_CP10 = "416" Or m_CP10 = "447" Then
            m_iCopy = 0 'E¤Æ¤£¥X¯È¥»
            If m_bolEmail = False Or m_bolPlusPaper = True Then
                m_iCopy = 1
                '«DE¤Æ¥X¯È¥»¤@¥÷; ­Y¦³«ü©w¥÷¼Æ>1«h´î¤@¥÷,«ü©w¥÷¼Æ=1«h¤£¥X¯È¥»
                PUB_GetCopySetting iCopy, pa(1), pa(2), pa(3), pa(4), m_CP10, ET01
                '­Y¬OE+±H¦L¤@¥÷,©T©w³£­n¦L; ex.Y53715000(E+±H)¥u¦L¤@¥÷
                If m_bolEmail = True Or m_bolPlusPaper = True Then
                    If iCopy = 0 Then
                       m_iCopy = 1
                    ElseIf iCopy > 1 Then
                       m_iCopy = iCopy - 1
                    End If
                Else
                    If iCopy = 1 Then
                        m_iCopy = 0
                    ElseIf iCopy > 1 Then
                        m_iCopy = iCopy - 1
                    End If
                End If
            End If
            '³]©wDN¥X¯È¥»¡A¤~¯à¦P®É²£¥ÍPDFÀÉ©M¯È¥»
            If m_bolEmail = False And m_iCopy > 0 Then
                m_bolDNEmail = True
                m_bolDNPlusPaper = True
            End If
            '§t¯È¥», E+±H(¥X¯È¥»)
            If m_iCopy > 0 Then
                NowPrint strReceiveNo, ET01, stET03, bolChk, strUserNum, , , , , m_iCopy, , True, True
                If bolChk = False Then
                   MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(pa(1)) & " ]¡I"
                End If
            Else
                'e¤Æ=>¦sPDF
                NowPrint strReceiveNo, ET01, stET03, bolChk, strUserNum, , , , , , , , True
            End If
            '¤é¥»°Ï®×¥ó¥t¦sPDF¦bTyping2
            If pa(75) <> "" Then
                m_NationID = GetPrjNationNumber(ChangeCustomerL(pa(75)))
                If Left(m_NationID, 3) = "011" Then
                     strCaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
                     strFileName = PUB_GetEFilePath(pa(1)) & "\" & pa(1) & "\" & Left(pa(2), 3) & "\" & strCaseNo & "\" & strCaseNo & "_" & strSrvDate(1) & "." & m_CP10 & ".CUS.PDF"
                     Call PUB_PrintLetter(strReceiveNo, , , True, strFullFileName, False, True)
                     DoEvents
                     Do While Dir(strFullFileName) <> ""
                        Exit Do
                     Loop
                     FileCopy strFullFileName, strFileName
                     DoEvents
                     Do While Dir(strFileName) <> ""
                        Kill strFullFileName
                        Exit Do
                     Loop
                     DoEvents
                End If
            End If
            '©ñ©w½Z¦C¦L
            If m_iCopy > 0 Then
                PUB_SetOsDefaultPrinter strPrinter   '¹w³]OS¦Lªí¾÷
                PUB_SetWordActivePrinter
                PUB_PrintLetter strReceiveNo
            End If
         Else
         'end 2021/01/21
                'Add by Morgan 2009/10/20 +§PÂ_¬O§_EMail¦P®É±H¯È¥»
                If m_bolPlusPaper Then
                   m_iCopy = 0
                Else
                   m_iCopy = 1
                End If
                'end 2009/10/20
                If m_bolEmail Then
                   NowPrint strReceiveNo, ET01, stET03, bolChk, strUserNum, , , , , m_iCopy, , True, True
                   If bolChk = False Then
                      MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(pa(1)) & " ]¡I"
                   End If
                Else
                   NowPrint strReceiveNo, ET01, stET03, bolChk, strUserNum
                End If
                
                'Added by Lydia 2020/08/20 ¹w³]Word¦Lªí¾÷
                PUB_SetOsDefaultPrinter strPrinter   '¹w³]OS¦Lªí¾÷
                PUB_SetWordActivePrinter
                'end 2020/08/20
                PUB_PrintLetter strReceiveNo 'Add By Sindy 2017/3/30 ª½±µ¦C¦L©w½Z
         End If 'Added by Lydia 2021/01/21
         
         If Not m_bolEmail Or m_bolPlusPaper Then
'            'Add By Sindy 2015/9/21 ¤é¤å©w½Z¤~­n¦L¦a§}±ø
'            If m_iLanguage = "3" Or Val(¥~±M¶}µ¡«H¨ç±Ò¥Î¤é) >= Val(strSrvDate(1)) Then
'            '2015/9/21 END
            'Add By Sindy 2017/3/20 ¤é¤å©w½Z¤~­n¦L¦a§}±ø
            If frm060306.m_FCna01 = "101" Or m_iLanguage = "3" Then '¬ü°ê ©Î ¤é¤å©w½Z¤~­n¦L¦a§}±ø
            '2017/3/20 END
               '·s¼W¦a§}±ø¦Cªí¸ê®Æ
               pub_AddressListSN = pub_AddressListSN + 1
               PUB_AddNewAddressList strUserNum, frm060306.Text1.Text, frm060306.Text2.Text, frm060306.Text3.Text, frm060306.Text4.Text, "" & pub_AddressListSN, "0"
            End If
         End If

         'Added by Lydia 2021/01/21 ¥i­«¦L½Ð´Ú³æ
         If m_CP60 <> "" And Text1(1) = "Y" Then
            Call ProcessPrint2(m_CP60)
         Else
         'end 2021/01/21
            '·s¼W¨Ã¦C¦L½Ð´Ú³æ
            '92.7.20 ADD BY SONIA
            If Text1(1) = "Y" Then
            '92.7.20 END
               If ProcessPrint = False Then
                  MsgBox "·s¼W½Ð´Ú³æ¸ê®Æ¿ù»~ !", vbCritical
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
         End If 'Added by Lydia 2021/01/21
         
         frm060306.Show
         frm060306.Clear
         Screen.MousePointer = vbDefault
         Unload Me
      Case 1 '¦^«eµe­±
         frm060306.Show
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
    'Add By Cheng 2003/06/29
    '­Y¤£Ä~Äò
    If m_blnContinuous = False Then
        cmdOK_Click (1)
    '­YÄ~Äò
    Else
        Me.Visible = True
    End If
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
   'Add by Morgan 2006/10/20 ²Î¤@¦bUnload³]©w
   '­Y½Ð´Ú³æ¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
   'Mark by Lydia 2020/08/20 ²Î¤@¥Î¨t²Î¹w³]¦Lªí¾÷,¨ú®ø¦Lªí¾÷¿ï³æ
   'If Me.Combo2.Text <> Me.Combo2.Tag Then
   '    PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   'End If
   'end 2020/08/20
   Set frm060306_7 = Nothing
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   
   Select Case Index
   Case 1 '·s¼W½Ð´Ú³æ
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         MsgBox "·s¼W½Ð´Ú³æ¥u¯à¿é¤J Y ©ÎªÅ¥Õ !!!", vbExclamation + vbOKOnly
         KeyAscii = 0
      End If
   End Select
End Sub

'92.11.25 add by sonia
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   Select Case Index
      Case 1 '·s¼W½Ð´Ú³æ
         'If Text1(1) = "Y" Then 'Remove by Lydia 2021/01/21
            'ÀË¬d¬O§_¤w½Ð´Ú
            'Modified by Lydia 2021/01/21 *=> CP60,CP148
            StrSQLa = "Select CP60,CP148 From Caseprogress Where CP09='" & strReceiveNo & "' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               'Added by Lydia 2021/01/21 §ï§PÂ_
               m_CP60 = "" & rsA.Fields("cp60")
               If m_CallName = "frm060104_3" Then
                    '¥i­«¦L½Ð´Ú³æ
               ElseIf Text1(1) = "Y" Then
               'end 2021/01/21
                    'Modified by Lydia 2021/01/21
                    'If "" & rsA("CP60").Value <> "" Then
                    '   MsgBox "¦¹µ§¸ê®Æ¤w½Ð´Ú¡A¤£¥i¦A½Ð´Ú!!!", vbExclamation + vbOKOnly
                    If "" & rsA("CP60").Value <> "" Or rsA.Fields("cp148") = "Y" Then
                       If "" & rsA("CP60").Value <> "" Then
                           MsgBox "¦¹µ§¸ê®Æ¤w½Ð´Ú¡A¤£¥i¦A½Ð´Ú!!!", vbExclamation + vbOKOnly
                       ElseIf "" & rsA("CP148").Value <> "" Then
                           MsgBox "¦¹µ§¸ê®Æ¬°¯S®í½Ð´Ú¡A¤£¥i¦b¦¹·s¼W½Ð´Ú³æ!!!", vbExclamation + vbOKOnly
                       End If
                    'end 2021/01/21
                       If Text1(1).Visible = True Then
                          Text1(1).SetFocus
                          Text1_GotFocus (1)
                       End If
                       Cancel = True
                    End If
               End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         'End If 'Remove by Lydia 2021/01/21
      Case Else
   End Select
End Sub
'92.11.25 end
Private Sub Text1_GotFocus(Index As Integer)
   InverseTextBox Text1(Index)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Form_Load()
Dim ii As Integer
'Add By Cheng 2003/02/05
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   MoveFormToCenter Me
   intWhere = °ê¥~_FC
   ReadPatent
   
'Modify by Morgan 2011/3/15 §ï¦@¥Î¥B¤£­n±Æ°£¹w³]¦Lªí¾÷
'Modified by Lydia 2020/08/20 ²Î¤@¥Î¨t²Î¹w³]¦Lªí¾÷,¨ú®ø¦Lªí¾÷¿ï³æ
'   PUB_SetPrinter Me.Name, Combo2, strPrinter
''end 2011/3/15
   strPrinter = PUB_GetOsDefaultPrinter
   
'   '¥ý§ì©T©w½Ð´Úª÷ÃB
'   strExc(0) = "SELECT A1J17 FROM ACC1J0 WHERE A1J01='" & pa(1) & "' AND A1J02='" & m_CP10 & "'"
'   intI = 1
'   Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If IsNull(rsTemp.Fields(0)) Then
'         MsgBox "¥¼«Ø¥ß©T©w½Ð´Úª÷ÃB¡A½Ð¥ý«Ø¸ê®Æ¤~¯à²£¥Í½Ð´Ú³æ!!!"
'         cmdok_Click (1)
'         Exit Sub
'      Else
'         m_A1J17 = rsTemp.Fields(0)
'      End If
'   Else
'      MsgBox "¥¼«Ø¥ß©T©w½Ð´Úª÷ÃB¡A½Ð¥ý«Ø¸ê®Æ¤~¯à²£¥Í½Ð´Ú³æ!!!"
'      cmdok_Click (1)
'      Exit Sub
'   End If
'   '¥ý§ì©T©w½Ð´Úª÷ÃB-³W¶O
'   strExc(0) = "SELECT A1J17 FROM ACC1J0 WHERE A1J01='" & pa(1) & "' AND A1J02='" & m_CP10 & "99" & "'"
'   intI = 1
'   Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If IsNull(rsTemp.Fields(0)) Then
'         MsgBox "¥¼«Ø¥ß©T©w½Ð´Úª÷ÃB-³W¶O¡A½Ð¥ý«Ø¸ê®Æ¤~¯à²£¥Í½Ð´Ú³æ!!!"
'         cmdok_Click (1)
'         Exit Sub
'      Else
'         m_A1J17_1 = rsTemp.Fields(0)
'      End If
'   Else
'      MsgBox "¥¼«Ø¥ß©T©w½Ð´Úª÷ÃB-³W¶O¡A½Ð¥ý«Ø¸ê®Æ¤~¯à²£¥Í½Ð´Ú³æ!!!"
'      cmdok_Click (1)
'      Exit Sub
'   End If
    Me.Visible = False
    m_blnContinuous = True
    StrSQLa = "Select * From CaseProgress Where CP09='" & strReceiveNo & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        m_A1J17 = Val("" & rsA("CP16").Value) - Val("" & rsA("CP17").Value) 'ªA°È¶O
        m_A1J17_1 = Val("" & rsA("CP17").Value) '³W¶O
        
         'Added by Morgan 2015/6/8
         m_bSpecial = False
         'SYNGENTA ¹ê¼f©T©w½Ð´Úª÷ÃB USD$65 (ªA°È¶O)
         If IsNull(rsA("CP60")) And m_CP10 = "416" And Left(pa(26), 6) = "X48310" Then
            m_bSpecial = True
            strExc(2) = PUB_GetUSXRate_1(strSrvDate(2), "USD")
            m_A1J17 = Round(65 * Val(strExc(2)))
         End If
         'end 2015/6/8
        
         'Add by Morgan 2010/8/6
         m_RefCP09s = ""
         m_lngFee938 = 0 '¶W­¶¶O
         m_lngFee939 = 0 '¶W¶µ¶O
         
         If m_CP10 <> "447" Then 'Added by Morgan 2024/11/21
         
            strExc(0) = "select cp10,cp17,cp09 from caseprogress where cp01='" & rsA("cp01") & "' and cp02='" & rsA("cp02") & "' and cp03='" & rsA("cp03") & "' and cp04='" & rsA("cp04") & "' and cp27>0 and cp60 is null and cp10 in ('938','939')"
            strExc(1) = ""
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Do While Not RsTemp.EOF
                  m_RefCP09s = m_RefCP09s & ",'" & RsTemp("cp09") & "'"
                  If RsTemp("cp10") = "938" Then
                     m_lngFee938 = m_lngFee938 + Val("" & RsTemp("cp17"))
                  Else
                     m_lngFee939 = m_lngFee939 + Val("" & RsTemp("cp17"))
                  End If
                  RsTemp.MoveNext
               Loop
               strExc(1) = "¡F" & vbCrLf & "¨Ã¥t¦³"
               If m_lngFee938 > 0 Then
                  strExc(1) = strExc(1) & "¶W­¶¶O " & Format(m_lngFee938, "#,##0") & " ¤¸"
               End If
               If m_lngFee939 > 0 Then
                  If m_lngFee938 > 0 Then strExc(1) = strExc(1) & "¤Î"
                  strExc(1) = strExc(1) & "¶W¶µ¶O " & Format(m_lngFee939, "#,##0") & " ¤¸"
               End If
               strExc(1) = strExc(1) & "¡C"
            End If
            'end 2010/8/6
            
         End If 'Added by Morgan 2024/11/21
        
        'Modify by Morgan 2004/12/29
        'If MsgBox("¦¹µ§¹êÅé¼f¬dªºªA°È¶O¬°" & Format(m_A1J17, "#,##0") & "¡A³W¶O¬°" & Format(m_A1J17_1, "#,##0") & vbCrLf & "¬O§_­nÄ~Äò§@·~? ", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
        'Add By Sindy 2017/3/20 + if ¦]¹êÅé¼f¬dµo¤å·|Call¦¹§@·~
        'Modified by Morgan 2024/11/21 447¦A¼f¬d¥[³t¼f¬dµo¤å¤]·|©I¥s,§ï§PÂ_¦³«ü©w®×¥ó©Ê½è
        'If Not frm060306.m_quy416 = True Then
        If frm060306.m_quyAnyCP10 = "" Then
        '2017/3/20 END
            If MsgBox("¦¹µ§" & GetCaseTypeName(pa(1), m_CP10) & "ªºªA°È¶O¬°" & Format(m_A1J17, "#,##0") & "¡A³W¶O¬°" & Format(m_A1J17_1, "#,##0") & strExc(1) & vbCrLf & "¬O§_­nÄ~Äò§@·~? ", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               m_blnContinuous = False
            End If
        End If
    Else
        m_blnContinuous = False
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    
   'Added by Lydia 2020/08/17 ¹ê¼f: ¼W¥[¤w¦¬´Ú
   lblPAID.Visible = False: txtPAID.Visible = False
   'Mark by Lydia 2020/08/17 ¦]¬°¹ê¼f±Ä¥Îµo¤å®É,¦Û°Ê²£¥Í½Ð´Ú³æ;©Ò¥H³æµ§¶]¨ü­­©ó¡u·s¼W½Ð´Ú³æ¡v
   'If pa(1) = "FCP" And m_CP10 = "416" Then
   '    lblPAID.Visible = True: txtPAID.Visible = True
   'End If
   'end 2020/08/17
   
End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, i As Integer, strTmp As String
   For Each Lbl In Label2
      Lbl = ""
   Next
   strReceiveNo = frm060306.Tag
   pa(1) = frm060306.Text1.Text
   pa(2) = frm060306.Text2.Text
   pa(3) = frm060306.Text3.Text
   pa(4) = frm060306.Text4.Text
   Label2(0).Caption = GiveSymbol(pa(1), pa(2), pa(3), pa(4))
   Label2(1).Caption = frm060306.Text5.Text
   SetComboToCombo Combo1, frm060306.Combo1
   
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa, intWhere) Then  'edit by nickc 2007/02/02 ¤£¥Î dll ¤F  If objPublicData.ReadPatentDatabase(pA, intWhere) Then
            If PA51CU58FA07(pa) Then
               For i = 1 To 6
                  Label2(i + 1) = pa(50 + i)
               Next
            End If
            Label2(8) = pa(48)
            Label2(9) = pa(77)
            If pa(86) <> "" Then
               'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
               'If objLawDll.LawGetName(pa(86), strTmp) Then Label2(10) = strTmp
               If ClsLawLawGetName(pa(86), strTmp) Then Label2(10) = strTmp
            End If
            Label2(11) = pa(87)
         End If
      Case "FG"
         If ClsPDReadServicePracticeDatabase(pa, intWhere) Then 'edit by nickc 2007/02/02 ¤£¥Î dll ¤F If objPublicData.ReadServicePracticeDatabase(pA, intWhere) Then
            If PA51CU58FA07(pa) Then Label2(2) = pa(30)
            Label2(8) = pa(29)
            Label2(9) = pa(27)
            'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
            'If objLawDll.LawGetName(pa(35), strTmp) Then Label2(10) = strTmp
            If ClsLawLawGetName(pa(35), strTmp) Then Label2(10) = strTmp
            Label2(11) = pa(36)
         End If
   End Select
   pa(1) = "": m_CP10 = ""
   strExc(0) = "SELECT CP25,CP01,CP10,GetEmailFlag(CP01||CP02||CP03||CP04) FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP09='" & strReceiveNo & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      pa(1) = "" & RsTemp.Fields(1)
      m_CP10 = "" & RsTemp.Fields(2)
   End If
   
End Sub

Private Function ProcessPrint() As Boolean
Dim m_strSerialNo As String '½Ð´Ú³æ¸¹
Dim strAgentNo As String '¥N²z¤H½s¸¹
Dim strPrintCust  As String '¬O§_¦C¦L¥Ó½Ð¤H
Dim dblUSRate As Double '¬üª÷¶×²v
Dim strA1K27 As String '¦C¦L¹ï¶H
Dim strA1K28 As String '½Ð´Ú¹ï¶H
Dim strA1K05 As String '½Ð´Ú³æ³Æµù Add by Morgan 2011/4/18

   On Error GoTo ErrorHandler
   
   ProcessPrint = False
   cnnConnection.BeginTrans
   
   '¶}©l·s¼W°ê¥~½Ð´Ú¸ê®Æ
   '1:¥ý¥H"X"§ìACC1R0¤§°ê¥~½Ð´Ú³æªº¦Û°Ê½s¸¹, ¨Ã§ó·s¨ä¬y¤ô¸¹
   m_strSerialNo = AccAutoNo(MsgText(815), 5)
   AccSaveAutoNo MsgText(815), Right(m_strSerialNo, 5)
   '2:·s¼WACC1K0
'   strAgentNo = GetAgentNO '¥N²z¤H½s¸¹
   strAgentNo = PUB_GetA1K03(pa(1), pa(2), pa(3), pa(4))
   'dblUSRate = GetUSRate '¬üª÷¶×²v
        
    strA1K27 = PUB_GetA1K27(pa(1), pa(2), pa(3), pa(4), m_CP10)
    If strA1K27 = "" Then strA1K27 = strAgentNo
    strA1K28 = PUB_GetA1K28(pa(1), pa(2), pa(3), pa(4), m_CP10)
    If strA1K28 = "" Then strA1K28 = strAgentNo
    
'   strPrintCust = GetPrintCust '¬O§_¦C¦L¥Ó½Ð¤H
   'Modify by Morgan 2004/12/16 §ï³W«h
   'strPrintCust = PUB_GetA1K04(pa(1), pa(2), pa(3), pa(4))
   strPrintCust = PUB_GetA1K04(pa(1), pa(2), pa(3), pa(4), strA1K28, m_CP10)
   '2004/12/16
   
    'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
     Dim strA1K33 As String, strA1K18 As String
     'Modify By Sindy 2016/11/30
     'strA1K33 = PUB_GetInitCurrPrintType(pa(1), strA1K28, strA1K18, dblUSRate)
     'Modified by Morgan 2018/4/27 +strA1K27
     strA1K33 = PUB_GetInitCurrPrintType(pa(1), strA1K28, strA1K18, dblUSRate, pa(2), pa(3), pa(4), strA1K27)
     '2016/11/30 END
     
   'Added by Morgan 2015/6/8
   If m_bSpecial Then
      strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K05,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21,A1K33) " & _
               "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",'" & strA1K05 & "',0,NULL," & (m_A1J17_1 + m_lngFee938 + m_lngFee939) & "," & dblUSRate & "," & (m_A1J17 + m_A1J17_1 + m_lngFee938 + m_lngFee939) & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & strA1K18 & "',0, 0,'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'),'" & strUserNum & "','" & strA1K33 & "')"
      cnnConnection.Execute strSql
      
      strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
               "VALUES  ('" & m_strSerialNo & "','FCP','',0,'001','" & m_CP10 & "'," & m_A1J17 & "," & strSrvDate(2) & ",to_char(sysdate,'HH24miss'),'" & strUserNum & "')"
      cnnConnection.Execute strSql
   Else
   'end 2015/6/8
     
   '   strSQL = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21) " & _
   '            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,NULL,0," & dblUSRate & "," & Val(m_A1J17) + Val(m_A1J17_1) & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','USD',0, " & IIf(dblUSRate = 0, Val(m_A1J17) + Val(m_A1J17_1), ((Val(m_A1J17) + Val(m_A1J17_1)) / dblUSRate)) & ",'" & strAgentNo & "','" & strAgentNo & "','" & strAgentNo & "','" & strPrintCust & "'," & (ServerDate - 19110000) & "," & ServerTime & ",'" & strUserNum & "')"
       Dim strDisc As String '§é¦©
       strDisc = 1 - (PUB_GetA1L07Disc(pa(1), pa(2), pa(3), pa(4), m_CP10, strSrvDate(2)) / 100)
       'Modified by Morgan 2015/8/5 +¥Ó½Ð¤H
       strDisc = 1 - PUB_GetDiscX(strAgentNo, pa(1), m_CP10, 1 - Val(strDisc), pa(26)) 'Added by Morgan 2013/5/2
       
      'Add by Morgan 2011/4/18
      'Modified by Morgan 2014/12/3
      'strA1K05 = PUB_GetDNRemark(strA1K28)
      strA1K05 = PUB_GetDNRemark(strA1K28, pa(1), pa(2), pa(3), pa(4))
      'end 2014/12/3
      
       'Modify By Cheng 2004/01/07
       'A1K11­n¥ý¦©°£§é¦©¤~¦sÀÉ
   '   strSQL = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21) " & _
   '            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,NULL," & Val(m_A1J17_1) & "," & dblUSRate & "," & Val(m_A1J17) + Val(m_A1J17_1) & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','USD',0, " & IIf(dblUSRate = 0, Val(m_A1J17) + Val(m_A1J17_1), ((Val(m_A1J17) + Val(m_A1J17_1)) / dblUSRate)) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & (ServerDate - 19110000) & "," & ServerTime & ",'" & strUserNum & "')"
       'Modify By Cheng 2004/04/26
       '¬üª÷¨ú¾ã¼Æ¦ì(µL±ø¥ó±Ë¥h)
   '   strSQL = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21) " & _
   '            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,NULL," & Val(m_A1J17_1) & "," & dblUSRate & "," & Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc)) & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','USD',0, " & IIf(dblUSRate = 0, (Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc))), ((Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc))) / dblUSRate)) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & (ServerDate - 19110000) & "," & ServerTime & ",'" & strUserNum & "')"
      'Modify by Morgan 2010/8/6 +¶W­¶¶W¶µ¶O
      'strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21) " & _
               "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,NULL," & Val(m_A1J17_1) & "," & dblUSRate & "," & Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc)) & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','USD',0, " & Fix(Val("" & IIf(dblUSRate = 0, (Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc))), ((Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc))) / dblUSRate)))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'),'" & strUserNum & "')"
      'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
   '   strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K05,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21) " & _
               "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",'" & strA1K05 & "',0,NULL," & Val(m_A1J17_1) + m_lngFee938 + m_lngFee939 & "," & dblUSRate & "," & Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc)) + m_lngFee938 + m_lngFee939 & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','USD',0, " & Fix(Val("" & IIf(dblUSRate = 0, (Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc)) + m_lngFee938 + m_lngFee939), ((Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc)) + m_lngFee938 + m_lngFee939) / dblUSRate)))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'),'" & strUserNum & "')"
       strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K05,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21,A1K33) " & _
               "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",'" & strA1K05 & "',0,NULL," & (Val(m_A1J17_1) + m_lngFee938 + m_lngFee939) & "," & dblUSRate & "," & Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc)) + m_lngFee938 + m_lngFee939 & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & strA1K18 & "',0, " & Fix(Val("" & IIf(dblUSRate = 0, (Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc)) + m_lngFee938 + m_lngFee939), ((Val(m_A1J17) + Val(m_A1J17_1) - Val(Val(m_A1J17) * Val(strDisc)) + m_lngFee938 + m_lngFee939) / dblUSRate)))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'),'" & strUserNum & "','" & strA1K33 & "')"
     
      'end 2010/8/6
       'End
      cnnConnection.Execute strSql
      '3:·s¼WACC1L0
   '    Dim strDisc As String '§é¦©
   '    strDisc = 1 - (PUB_GetA1L07Disc(pa(1), pa(2), pa(3), pa(4), m_CP10, strSrvDate(2)) / 100)
      strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
               "VALUES  ('" & m_strSerialNo & "','FCP',''," & Val(m_A1J17) * Val(strDisc) & ",'001','" & m_CP10 & "'," & Val(m_A1J17) & "," & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'),'" & strUserNum & "')"
      cnnConnection.Execute strSql
   
   End If 'Added by Morgan 2015/6/8
   
   If Val(m_A1J17_1) > 0 Then 'Added by Morgan 2024/11/21
      strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
               "VALUES  ('" & m_strSerialNo & "','FCP','',0 ,'002','" & m_CP10 & "99" & "'," & Val(m_A1J17_1) & "," & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'),'" & strUserNum & "')"
      cnnConnection.Execute strSql
   End If
   
   'Add by Morgan 2010/8/6 ¶W­¶¶W¶µ¶O
   If m_RefCP09s <> "" Then
      strExc(3) = "003"
      If m_lngFee938 > 0 Then
         strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
            "VALUES  ('" & m_strSerialNo & "','FCP','',0 ,'" & strExc(3) & "','93899" & "'," & m_lngFee938 & "," & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'),'" & strUserNum & "')"
         cnnConnection.Execute strSql
         strExc(3) = "004"
      End If
      If m_lngFee939 > 0 Then
         strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
            "VALUES  ('" & m_strSerialNo & "','FCP','',0 ,'" & strExc(3) & "','93999" & "'," & m_lngFee939 & "," & strSrvDate(2) & ",to_char(sysdate, 'HH24MISS'),'" & strUserNum & "')"
         cnnConnection.Execute strSql
      End If
      strSql = "INSERT INTO ACC1W0 " & _
            " select '" & m_strSerialNo & "',cp09 from caseprogress where cp09 in (" & Mid(m_RefCP09s, 2) & ")"
      cnnConnection.Execute strSql, intI
   
      strSql = "UPDATE CASEPROGRESS SET CP60='" & m_strSerialNo & "' WHERE CP09 in (" & Mid(m_RefCP09s, 2) & ")"
      cnnConnection.Execute strSql, intI
   End If
   'end 2010/8/5
   
   PUB_UpdateA1k08 m_strSerialNo 'Added by Morgan 2012/11/2 §ó·s½Ð´Ú³æ¥~¹ôª÷ÃB
   
   '4:·s¼WACC1W0
   strSql = "INSERT INTO ACC1W0 " & _
            "VALUES  ('" & m_strSerialNo & "','" & strReceiveNo & "')"
   cnnConnection.Execute strSql
   '5:§ó·s½Ð´Ú³æ¸¹
'   strSQL = "UPDATE CASEPROGRESS SET CP16='" & m_A1J17 + m_A1J17_1 & "', CP17='" & m_A1J17_1 & "',CP18='" & Round(m_A1J17 / 1000) & "',CP60='" & m_strSerialNo & "' WHERE CP09='" & strReceiveNo & "'"
   'Modified by Morgan 2015/6/8 +¦]¦³¯S®í½Ð´Úª÷ÃB,§ï­n¦^¼g CP16,CP17,CP18
   strSql = "UPDATE CASEPROGRESS SET CP16=" & (m_A1J17 + m_A1J17_1) & ",CP17=" & m_A1J17_1 & ",CP18=" & Round(m_A1J17 / 1000, 1) & ",CP60='" & m_strSerialNo & "' WHERE CP09='" & strReceiveNo & "'"
   cnnConnection.Execute strSql
   
   PUB_PointAutoassign m_strSerialNo, True 'Add by Morgan 2010/4/21 ¦Û°Ê¤À°tÂI¼Æ
   
   cnnConnection.CommitTrans
   ProcessPrint = True
   
    'Added by Lydia 2016/11/17 ¥H½Ð´Ú¹ï¶HÀË¬d¬O§_¦s¦b©ó°ê¥~©T©w±H¶Ê´Ú³æ¥N²z¤HÀÉ(ACC225)¥B¤U¦¸±Hµo¤é´Á¡Ö¨t²Î¤é¡A­Y¦s¦b«hÅã¥Ü°T®§´£¿ô¾Þ§@¤H­û
    If m_strSerialNo <> "" And strA1K28 <> "" Then
       If PUB_ChkAcc225MsgList(m_strSerialNo, strA1K28, pa(1), pa(2), pa(3), pa(4)) Then
       End If
    End If
    'end 2016/11/17
    
   'Added by Lydia 2020/08/17  ¼f¹êµo¤å->(·s¼W½Ð´Ú³æ)¤w¦¬´Ú,¦Û°ÊEmailµ¹°]°È³B
   'Modified by Moragn 2024/11/21 +447¦A¼f¬d¥[³t¼f¬d
   If m_strSerialNo <> "" And pa(1) = "FCP" And (m_CP10 = "416" Or m_CP10 = "447") And Me.txtPAID.Text <> "" And m_CallName = "frm060104_3" Then
      'Modifie by Lydia 2024/01/29 §ï¦¨¯S®í³]©w
      'strExc(0) = "A2008" '°]°È¹w³]¦¬¥ó¤H, CCµ¹¾Þ§@ªÌ
      strExc(0) = Pub_GetSpecMan("¥~±M½Ð´Ú³æ¤w¦¬´Ú³qª¾¤H­û")
      If strExc(0) <> "" Then
      'end 2024/01/29
         strExc(1) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & "¹êÅé¼f¬d¤w¦¬´Ú¡A½Ð´Ú³æ½s¸¹¬°" & m_strSerialNo
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            "values('" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & ChgSQL(strExc(1)) & "','¦P¥D¦®','" & strUserNum & "' )"
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2020/08/17
   
   Screen.MousePointer = vbHourglass
   'Mark by Lydia 2021/01/21 ¥t¥~¤À¦¨¼Ò²Õ
'   Load Frmacc2480
'   With Frmacc2480
'      .text1.Text = m_strSerialNo
'      .Text2.Text = m_strSerialNo
'      'Modified by Lydia 2020/08/20
'      '.Combo1.Text = Me.Combo2.Text
'      .Combo1.Text = strPrinter
'      'Add by Morgan 2008/5/23 +¶Ç¬O§_¦s¹q¤lÀÉ°Ñ¼Æ
'      .m_bBeCalled = True
'      .m_CallPrevForm = Me.Name  'Added by Lydia 2020/01/06 ©I¥s½Ð´Ú³æªºµ{¦¡¦WºÙ
'      'Modified by Morgan 2014/6/3
'      '.m_bEMail = m_bolEmail
'      '.m_bPaper = m_bolPlusPaper
'      .m_bEMail = m_bolDNEmail
'      .m_bPaper = m_bolDNPlusPaper
'      'end 2014/6/3
'      'end 2008/5/23
'      'Added by Lydia 2020/08/17 ¼f¹êµo¤å->¤w¦¬´Ú
'      .m_bPAID = False
'      If pa(1) = "FCP" And m_CP10 = "416" And Me.txtPAID.Text <> "" Then
'         .m_bPAID = True
'      End If
'      'end 2020/08/17
'      .Command2_Click: DoEvents
'   End With
'   Unload Frmacc2480: DoEvents
   Call ProcessPrint2(m_strSerialNo)
   'end 2021/01/21
   Exit Function

ErrorHandler:
      
      cnnConnection.RollbackTrans
      ProcessPrint = False
End Function

'Added by Lydia 2021/01/21 ¦C¦L½Ð´Ú³æ
Private Sub ProcessPrint2(ByVal m_KeyNO As String)
    
   If m_KeyNO = "" Then Exit Sub
    
   Load Frmacc2480
   With Frmacc2480
      .Text1.Text = m_KeyNO
      .Text2.Text = m_KeyNO
      .Combo1.Text = strPrinter
      .m_bBeCalled = True
      .m_CallPrevForm = Me.Name
      .m_bEMail = m_bolDNEmail
      .m_bPaper = m_bolDNPlusPaper
      .m_bPAID = False
      'Modified by Morgan 2024/11/21 +447
      If pa(1) = "FCP" And (m_CP10 = "416" Or m_CP10 = "447") And Me.txtPAID.Text <> "" Then
         .m_bPAID = True
      End If
      .Command2_Click: DoEvents
   End With
   Unload Frmacc2480: DoEvents
End Sub

Private Function GetUSRate() As Double
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetUSRate = 0
'strSQLA = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & (ServerDate - 19110000) & " AND ROWNUM = 1 ORDER BY USXR01 "
StrSQLa = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & strSrvDate(2) & " ORDER BY USXR01 DESC "
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   GetUSRate = rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Function GetAgentNO() As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetAgentNO = ""
'¨ú±o±M§Q°ò¥»ÀÉªº"©T©w½Ð´Ú¹ï¶H"
StrSQLa = "SELECT PA88 FROM PATENT WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' AND PA88 IS NOT NULL "
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   GetAgentNO = rsA.Fields(0).Value
Else
   '¨ú±o±M§Q°ò¥»ÀÉªº"FC¥N²z¤H"
   StrSQLa = "Select PA75 From PATENT WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' AND PA75 IS NOT NULL "
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      GetAgentNO = rsA.Fields(0).Value
      '¨ú±o°ê¥~¥N²z¤HÀÉªº"©T©w½Ð´Ú¹ï¶H"
      StrSQLa = "Select FA30 From FAGENT WHERE SUBSTR('" & GetAgentNO & "',1,8)=FA01 AND SUBSTR('" & GetAgentNO & "',9,1)=FA02 AND FA30 IS NOT NULL "
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         GetAgentNO = rsA.Fields(0).Value
      End If
   Else
      '¨ú±o±M§Q°ò¥»ÀÉªº"¥Ó½Ð¤H1"
      StrSQLa = "SELECT PA26 FROM PATENT WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' AND PA26 IS NOT NULL "
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         GetAgentNO = rsA.Fields(0).Value
         '¨ú±o«È¤á°ò¥»ÀÉªº"©T©w½Ð´Ú¹ï¶H"
         StrSQLa = "SELECT CU57 FROM CUSTOMER WHERE CU01=SUBSTR('" & GetAgentNO & "',1,8) AND CU02=SUBSTR('" & GetAgentNO & "',9,1) AND CU57 IS NOT NULL "
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            GetAgentNO = rsA.Fields(0).Value
         End If
      End If
   End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Private Function GetPrintCust() As String
'Dim rsA As New ADODB.Recordset
'Dim strSQLA As String
'
'GetPrintCust = ""
''¨ú±o±M§Q°ò¥»ÀÉªº"D/N¬O§_¦C¦L¥Ó½Ð¤H"
'strSQLA = "SELECT PA78 FROM PATENT WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' AND PA78 IS NOT NULL "
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'rsA.CursorLocation = adUseClient
'rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   GetPrintCust = rsA.Fields(0).Value
'Else
'   '¨ú±o°ê¥~¥N²z¤HÀÉªº"D/N¬O§_¦C¦L¥Ó½Ð¤H"
'   strSQLA = "SELECT FA44 FROM PATENT, FAGENT WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' " & _
'            " AND SUBSTR(PA75,1,8)=FA01 AND SUBSTR(PA75,9,1)=FA02 AND FA44 IS NOT NULL "
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   rsA.CursorLocation = adUseClient
'   rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      GetPrintCust = rsA.Fields(0).Value
'   Else
'      '¨ú±o«È¤á°ò¥»ÀÉªº"D/N¬O§_¦C¦L¥Ó½Ð¤H"
'      strSQLA = "SELECT CU77 FROM PATENT, CUSTOMER WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' " & _
'               " AND SUBSTR(PA26,1,8)=CU01 AND SUBSTR(PA26,9,1)=CU02 AND CU77 IS NOT NULL "
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      rsA.CursorLocation = adUseClient
'      rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         GetPrintCust = rsA.Fields(0).Value
'      End If
'   End If
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'
'End Function

'Added by Lydia 2020/08/17
Private Sub txtPAID_GotFocus()
   TextInverse txtPAID
End Sub

Private Sub txtPAID_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
End Sub
