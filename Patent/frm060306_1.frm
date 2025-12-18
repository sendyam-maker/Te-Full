VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060306_1 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "½Ð´Ú³qª¾¨ç-¤¤»¡½Ð´Ú¨ç"
   ClientHeight    =   6510
   ClientLeft      =   1080
   ClientTop       =   960
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7080
   Begin VB.CheckBox Check1 
      Caption         =   "µo©ú¤H¦WµL¯S®í¦r"
      Height          =   180
      Index           =   20
      Left            =   1635
      TabIndex        =   73
      Top             =   4410
      Width           =   1995
   End
   Begin VB.CheckBox Check1 
      Caption         =   "¥þ³¡µo©ú¤H¦W"
      Height          =   180
      Index           =   19
      Left            =   105
      TabIndex        =   72
      Top             =   4410
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "½Ð¨D¸ê°T¤£¤½¶}¤§Án©ú®Ñ"
      Height          =   180
      Index           =   18
      Left            =   4830
      TabIndex        =   71
      Top             =   4200
      Width           =   2355
   End
   Begin VB.CheckBox Check1 
      Caption         =   "¥xÄyµo©ú¤HID½X"
      Height          =   180
      Index           =   17
      Left            =   3120
      TabIndex        =   20
      Top             =   4200
      Width           =   1635
   End
   Begin VB.CheckBox Check1 
      Caption         =   "­^¤åºK­n"
      Height          =   180
      Index           =   16
      Left            =   1635
      TabIndex        =   19
      Top             =   4200
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "¥Nªí¤H"
      Height          =   180
      Index           =   15
      Left            =   105
      TabIndex        =   18
      Top             =   4200
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "¥Ó½Ð¤H°êÄy"
      Height          =   180
      Index           =   14
      Left            =   4815
      TabIndex        =   17
      Top             =   3990
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "¥Ó½Ð¤H¤¤Ä¶¦W"
      Height          =   180
      Index           =   13
      Left            =   3120
      TabIndex        =   16
      Top             =   3990
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "µo©ú¤H°êÄy"
      Height          =   180
      Index           =   12
      Left            =   1635
      TabIndex        =   15
      Top             =   3990
      Width           =   1425
   End
   Begin VB.CheckBox Check1 
      Caption         =   "µo©ú¤H¤¤Ä¶¦W"
      Height          =   180
      Index           =   11
      Left            =   105
      TabIndex        =   14
      Top             =   3990
      Width           =   1425
   End
   Begin VB.CheckBox Check2 
      Caption         =   "IPO¸É¥ó¨ç"
      Height          =   180
      Index           =   6
      Left            =   4860
      TabIndex        =   33
      Top             =   5550
      Width           =   1320
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   5
      Left            =   5460
      MaxLength       =   1
      TabIndex        =   24
      Top             =   4890
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   264
      Left            =   5745
      TabIndex        =   2
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   4
      Left            =   4350
      MaxLength       =   1
      TabIndex        =   1
      Top             =   2880
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   264
      Left            =   3870
      TabIndex        =   29
      Top             =   5265
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   3
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   22
      Top             =   4605
      Width           =   255
   End
   Begin VB.TextBox txtDate 
      Height          =   270
      Left            =   960
      MaxLength       =   7
      TabIndex        =   0
      Top             =   2880
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   " µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   5820
      TabIndex        =   40
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox Text3 
      Height          =   264
      Left            =   1215
      TabIndex        =   25
      Top             =   5070
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   2
      Left            =   2175
      MaxLength       =   1
      TabIndex        =   23
      Top             =   4890
      Width           =   255
   End
   Begin VB.CheckBox Check2 
      Caption         =   "¹ê¼f½Ð¨D®Ñ"
      Height          =   180
      Index           =   4
      Left            =   3105
      TabIndex        =   32
      Top             =   5550
      Width           =   1365
   End
   Begin VB.CheckBox Check2 
      Caption         =   "¤w­×¥¿¤§»¡©ú®Ñ/Claims"
      Height          =   180
      Index           =   3
      Left            =   105
      TabIndex        =   31
      Top             =   5550
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      Caption         =   "¤¤»¡"
      Height          =   180
      Index           =   2
      Left            =   3105
      TabIndex        =   28
      Top             =   5340
      Width           =   735
   End
   Begin VB.CheckBox Check2 
      Caption         =   "¦¬¾Ú¼v¥»"
      Height          =   180
      Index           =   1
      Left            =   1545
      TabIndex        =   27
      Top             =   5340
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Debit Note"
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   26
      Top             =   5340
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Ä~©ÓÃÒ©ú"
      Height          =   180
      Index           =   10
      Left            =   4815
      TabIndex        =   13
      Top             =   3765
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "¤½ÃÒ¤§¹µ¶Ä«´¬ù©Î»{ÃÒ¤§Åý´ç¤å¥ó"
      Height          =   180
      Index           =   9
      Left            =   105
      TabIndex        =   11
      Top             =   3765
      Width           =   3015
   End
   Begin VB.CheckBox Check1 
      Caption         =   "¦º¤`ÃÒ©ú"
      Height          =   180
      Index           =   8
      Left            =   3120
      TabIndex        =   12
      Top             =   3765
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "µo©ú¤H©ÚÃ±¤§¤Áµ²®Ñ"
      Height          =   180
      Index           =   7
      Left            =   4815
      TabIndex        =   10
      Top             =   3570
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "°êÄyÃÒ©ú"
      Height          =   180
      Index           =   6
      Left            =   3120
      TabIndex        =   9
      Top             =   3570
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ªk¤H¦a¦ìÃÒ©ú"
      Height          =   180
      Index           =   5
      Left            =   1635
      TabIndex        =   8
      Top             =   3570
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "°ê¥~±H¦sÃÒ©ú"
      Height          =   180
      Index           =   4
      Left            =   105
      TabIndex        =   7
      Top             =   3570
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Àu¥ýÅvÃÒ©ú"
      Height          =   180
      Index           =   3
      Left            =   4815
      TabIndex        =   6
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "©e¥ô®Ñ"
      Height          =   180
      Index           =   2
      Left            =   3120
      TabIndex        =   5
      Top             =   3360
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Åý´ç®Ñ"
      Height          =   180
      Index           =   1
      Left            =   1635
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "°ê¤º±H¦sÃÒ©ú"
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   3360
      Width           =   1425
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   1
      Left            =   2175
      MaxLength       =   1
      TabIndex        =   21
      Top             =   4605
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3840
      TabIndex        =   36
      Top             =   30
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   " ¦^«eµe­±(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4620
      TabIndex        =   38
      Top             =   30
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   0
      Left            =   1455
      MaxLength       =   1
      TabIndex        =   34
      Top             =   6270
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   2436
      Left            =   60
      TabIndex        =   35
      Top             =   390
      Width           =   6792
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   7
         Left            =   1260
         TabIndex        =   70
         Top             =   2160
         Width           =   5385
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9499;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   6
         Left            =   1260
         TabIndex        =   69
         Top             =   1920
         Width           =   5385
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9499;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   5
         Left            =   1260
         TabIndex        =   68
         Top             =   1695
         Width           =   5385
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9499;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   4
         Left            =   1260
         TabIndex        =   67
         Top             =   1455
         Width           =   5385
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9499;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   3
         Left            =   1260
         TabIndex        =   66
         Top             =   1200
         Width           =   5385
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9499;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   2
         Left            =   1260
         TabIndex        =   65
         Top             =   975
         Width           =   5385
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9499;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   1
         Left            =   1260
         TabIndex        =   64
         Top             =   390
         Width           =   5385
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9499;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   0
         Left            =   1260
         TabIndex        =   63
         Top             =   135
         Width           =   5385
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9499;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   285
         Left            =   1260
         TabIndex        =   61
         Top             =   630
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
         Left            =   84
         TabIndex        =   47
         Top             =   1968
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(¤é)"
         Height          =   180
         Index           =   7
         Left            =   84
         TabIndex        =   46
         Top             =   2208
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(¤¤)"
         Height          =   180
         Index           =   6
         Left            =   84
         TabIndex        =   45
         Top             =   1728
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(­^)"
         Height          =   180
         Index           =   5
         Left            =   84
         TabIndex        =   44
         Top             =   1248
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(¤é)"
         Height          =   180
         Index           =   4
         Left            =   84
         TabIndex        =   43
         Top             =   1488
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(¤¤)"
         Height          =   180
         Index           =   3
         Left            =   84
         TabIndex        =   42
         Top             =   1008
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "±M§Q¦WºÙ"
         Height          =   180
         Index           =   2
         Left            =   84
         TabIndex        =   41
         Top             =   648
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "½Ð´Ú¨ç¤é´Á"
         Height          =   180
         Index           =   1
         Left            =   84
         TabIndex        =   39
         Top             =   408
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥»©Ò®×¸¹"
         Height          =   180
         Index           =   0
         Left            =   84
         TabIndex        =   37
         Top             =   168
         Width           =   720
      End
   End
   Begin VB.CheckBox Check2 
      Caption         =   "±M§Q¥Ó½Ð®Ñ"
      Height          =   180
      Index           =   5
      Left            =   4860
      TabIndex        =   30
      Top             =   5340
      Width           =   1275
   End
   Begin MSForms.TextBox Text1 
      Height          =   450
      Index           =   0
      Left            =   480
      TabIndex        =   62
      Top             =   5760
      Width           =   6285
      VariousPropertyBits=   -1466939365
      MaxLength       =   2000
      ScrollBars      =   3
      Size            =   "11086;794"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_¥[¦L©T©w½Ð´Ú¹ï¶H«H¨ç         (Y:¬O)"
      Height          =   180
      Index           =   21
      Left            =   3240
      TabIndex        =   60
      Top             =   4890
      Width           =   3030
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥÷¼Æ"
      Height          =   180
      Index           =   20
      Left            =   5295
      TabIndex        =   59
      Top             =   2925
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_¥u¦L Cover Page         (Y:¬O)"
      Height          =   180
      Index           =   19
      Left            =   2685
      TabIndex        =   58
      Top             =   2925
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥÷"
      Height          =   180
      Index           =   18
      Left            =   4305
      TabIndex        =   57
      Top             =   5340
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_½Ð¨D¹ê¼f                  (Y:¬O)"
      Height          =   180
      Index           =   17
      Left            =   3240
      TabIndex        =   56
      Top             =   4650
      Width           =   2355
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥÷"
      Height          =   180
      Index           =   15
      Left            =   1650
      TabIndex        =   55
      Top             =   5100
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ªþ¥ó(¥i½Æ¿ï)"
      Height          =   180
      Index           =   13
      Left            =   105
      TabIndex        =   54
      Top             =   5100
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_­×¥¿»¡©ú®Ñ/Claims             (Y:¬O)"
      Height          =   180
      Index           =   12
      Left            =   105
      TabIndex        =   53
      Top             =   4650
      Width           =   2835
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_ÀËµø¤Î­×¥¿¤j³°¤¤»¡         (Y:¬O)"
      Height          =   180
      Index           =   16
      Left            =   105
      TabIndex        =   52
      Top             =   4890
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¸É¥ó´Á­­"
      Height          =   180
      Index           =   9
      Left            =   105
      TabIndex        =   51
      Top             =   2925
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¸É¥ó¤º®e(¥i½Æ¿ï)"
      Height          =   180
      Index           =   10
      Left            =   105
      TabIndex        =   50
      Top             =   3180
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P. S."
      Height          =   180
      Index           =   11
      Left            =   105
      TabIndex        =   49
      Top             =   5820
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_­×§ï½Ð´Ú¨ç        (Y)"
      Height          =   180
      Index           =   14
      Left            =   105
      TabIndex        =   48
      Top             =   6300
      Width           =   1860
   End
End
Attribute VB_Name = "frm060306_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/16 Form2.0¤w­×§ï
'Memo by Morgan 2022/10/26 ¤é¤å¤w§ï§ìTable
'Memo By Morgan 2012/12/10 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo by Morgan2010/12/27 ¥Ó½Ð®×¸¹Äæ¤w­×§ï
'2010/12/6 memo by sonia ­û¤u½s¸¹Äæ¤w­×§ï
'Memo by Morgan2010/8/13 ¤é´ÁÄæ¤w­×§ï
Option Explicit

Dim intWhere As Integer, strReceiveNo As String
Dim m_PA11 As String, m_PriDate As String

Const ET01 As String = "09"
Dim m_PA10 As String 'Add by Morgan 2004/9/8 ¥Ó½Ð¤é
Public m_CP10 As String 'Add by Morgan 2004/10/7 ®×¥ó©Ê½è
Dim m_LetterLanguage As String
Dim m_JPPriData As String 'Add by Morgan 2004/10/8 ¤é¤åÀu¥ýÅv¸ê®Æ
Dim m_CNPriData As String 'Add by Morgan 2016/10/12 ¤¤¤åÀu¥ýÅv¸ê®Æ
Dim m_PA75 As String 'Add by Morgan 2007/4/14 ¥N²z¤H½s¸¹
Dim m_NP09 As String
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String
Dim m_bolAMD As Boolean, m_AMD05 As String 'Added by Lydia 2015/04/27 +¤¤»¡½Ð´Ú­×¥¿
Dim m_strOtherDoc As String 'Added by Morgan 2022/2/10

'Removed by Morgan 2022/2/10 ¨S¦³¥Î¤F
'Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
' Dim strTxt(1 To 50) As String, j As Integer, strTmp As String, oChk As CheckBox
'
'   EndLetter ET01, strReceiveNo, ET03, strUserNum
'   j = 1
'    'Add By Cheng 2003/02/24
'   '½Ð´Ú¨ç¤é´Á
'   If frm060306.Text5.Text <> "" Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç¤é´Á','" & DBDATE(frm060306.Text5.Text) & "')"
'      j = j + 1
'   End If
'   '½Ð´Ú¨ç³Æµù
'   If Text1(0) <> "" Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç³Æµù','P.S. " & ChgSQL(Text1(0).Text) & "')"
'      j = j + 1
'   End If
'   '¬O§_¸É¤å¥ó
'   If txtDate <> "" Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¨ä¥L¤é´Á'," & DBDATE(txtDate) & ")"
'      j = j + 1
'   End If
'      '¬O§_¸É¤å¥ó
'   If m_NP09 <> "" Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªk©w´Á­­'," & DBDATE(m_NP09) & ")"
'      j = j + 1
'   End If
'    '«Å»}®Ñ
'   If Check1(0).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 1','* Original Oath')"
'      j = j + 1
'   End If
'    'Åý´ç®Ñ
'   If Check1(1).Value = 1 Then
'      strTmp = "* Original Assignment"
'      'Added by Morgan 2012/9/27
'      If Val(m_NP09) > 20130101 Then
'         strTmp = strTmp & vbCrLf & "According to the current Patent Act, the Assignment is required. |#(©³½u)However, since the" & _
'         " deadline falls after January 1, 2013 when the new Patent Act becomes effective," & _
'         " this document is allowed to be omitted.#|"
'      End If
'      'end 2012/9/27
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 2','" & strTmp & "')"
'      j = j + 1
'   End If
'    '©e¥ôª¬
'   If Check1(2).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 3','* Power of Attorney signed by the applicant')"
'      j = j + 1
'   End If
'    'Àu¥ýÅvÃÒ©ú
'   If Check1(3).Value = 1 Then
'      'Modify by Morgan 2004/7/29
'      'David »¡¼È®É¤£§ïµ¥´¼¼z§½¥¿¦¡¤½§G
'      'Modify by Morgan 2004/7/28
'      'Modify by Morgan 2004/9/8 ½T©w¥Î¥Ó½Ð¤é§PÂ_
'      'Modified by Morgan 2012/9/27 ªk­­>=20130101 ¤£±anon-extendable...
'      If Val(m_PA10) >= 930701 And Val(m_NP09) < 20130101 Then
'         strTmp = "* Certified copy of Priority Document(s)(non-extendable deadline " & ChgEngDate(DBDATE(txtDate)) & ")"
'      Else
'         strTmp = "* Certified copy of Priority Document(s)"
'      End If
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 4','" & strTmp & "')"
'      j = j + 1
'   End If
'    '­ì¥Ó½Ð¤é¡B¸¹
'   If Check1(4).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 5','* The corresponding filing date and filing number of first/basic application')"
'      j = j + 1
'   End If
'    'ªk¤H¦a¦ìÃÒ©ú
'   If Check1(5).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 6','* Certificate of Corporation duly notarized')"
'      j = j + 1
'   End If
'    '°êÄyÃÒ©ú
'   If Check1(6).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 7','* Certificate of Nationality duly notarized')"
'      j = j + 1
'   End If
'    'µo©ú¤H©ÚÃ±¤§¤Áµ²®Ñ
'   If Check1(7).Value = 1 Then
'        'Modify By Cheng 2004/02/04
''      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
''         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 8','* Notarized Declaration Stating ownership based on Employment Contract')"
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 8','* Notarized Declaration Stating ownership')"
'        'End
'      j = j + 1
'   End If
'    '¤½ÃÒ¤§¹µ¶Ä«´¬ù©Î»{ÃÒ¤§Åý´ç¤å¥ó
'    'Modify By Cheng 2004/02/04
''   If Check1(8).Value = 1 Then
'   If Check1(9).Value = 1 Then
'    'End
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 9','* Notarized Employment Contract or Certified worldwide patent assignment')"
'      j = j + 1
'   End If
'    '¦º¤`ÃÒ©ú
'    'Modify By Cheng 2004/02/04
''   If Check1(9).Value = 1 Then
'   If Check1(8).Value = 1 Then
'    'End
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 10','* Notarized Certificate of Death')"
'      j = j + 1
'   End If
'    'Ä~©ÓÃÒ©ú
'   If Check1(10).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 11','* Notarized Certificate of Inheritance')"
'      j = j + 1
'   End If
'   '¬O§_­×¥¿Claims
'   If Text2(1) = "Y" Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 21','" & vbCrLf & "    As per your instructions, we have amended the specification/claims. ')"
'      j = j + 1
'   End If
'   '¬O§_½Ð¨D¹ê¼f
'   If Text2(3) = "Y" Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 22','" & vbCrLf & "    We have also filed a request for Substantive Examination. ')"
'      j = j + 1
'   End If
'   '¬O§_ÀËµø¤Î­×¥¿¤j³°¤¤»¡
'   If Text2(2) = "Y" Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 23','" & vbCrLf & "    We have reviewed and amended the Chinese text to comply with the R.O.C. Taiwanese formality requirements. ')"
'      j = j + 1
'   End If
'
'   'ªþ¥ó
'   intI = 0
'   For Each oChk In Check2
'      If oChk.Value = 1 Then
'         intI = 1
'         Exit For
'      End If
'   Next
'
'   If intI = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¤å¥ó V 24','" & vbCrLf & "    Please find enclosed :')"
'      j = j + 1
'   End If
'   If Check2(0).Value = 1 Then
'        strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'           "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','Dobit Note','* Debit Note ')"
'        j = j + 1
'        'Add By Cheng 2003/01/29
'        If Val(Text3.Text) > 1 Then
'           strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'              "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó´X¥÷','( " & Val(Text3.Text) & " copies)')"
'           j = j + 1
'        End If
'   End If
'   If Check2(1).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦¬¾Ú¼v¥»','* Copy of filing receipt ')"
'      j = j + 1
'        'Add By Cheng 2003/01/29
'        If Val(Text3.Text) > 1 Then
'           strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'              "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó´X¥÷1','( " & Val(Text3.Text) & " copies)')"
'           j = j + 1
'        End If
'   End If
'
'   'Modify by Morgan 2004/10/5 ¤¤»¡¥÷¼Æ³æ¿W³]©w
'   If Check2(2).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤¤»¡','* Copy of the Chinese specification ')"
'      j = j + 1
'        If Val(Text4.Text) > 1 Then
'           strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'              "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó´X¥÷2','( " & Val(Text4.Text) & " copies)')"
'           j = j + 1
'        End If
'   End If
'   If Check2(3).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤w­×¥¿¤§»¡©ú®Ñ/Claims','* Amended specification/claims ')"
'      j = j + 1
'        If Val(Text3.Text) > 1 Then
'           strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'              "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó´X¥÷3','( " & Val(Text3.Text) & " copies)')"
'           j = j + 1
'        End If
'   End If
'   If Check2(4).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¹ê¼f³qª¾®Ñ','* Copy of request for Substantive Examination ')"
'      j = j + 1
'        If Val(Text3.Text) > 1 Then
'           strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'              "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó´X¥÷4','( " & Val(Text3.Text) & " copies)')"
'           j = j + 1
'        End If
'   End If
'
'   'Add by Morgan 2010/3/23
'   If Check2(5).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','±M§Q¥Ó½Ð®Ñ','* Copy of Written Application ')"
'      j = j + 1
'        If Val(Text3.Text) > 1 Then
'           strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'              "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó´X¥÷5','( " & Val(Text3.Text) & " copies)')"
'           j = j + 1
'        End If
'   End If
'   If Check2(6).Value = 1 Then
'      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','IPO¸É¥ó¨ç','* Taiwan IPO''s Notification ')"
'      j = j + 1
'        If Val(Text3.Text) > 1 Then
'           strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'              "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó´X¥÷6','( " & Val(Text3.Text) & " copies)')"
'           j = j + 1
'        End If
'   End If
'   'end 2010/3/23
'
'   If Not ClsLawExecSQL(j - 1, strTxt) Then
'      MsgBox "Àx¦s¨Ò¥~Äæ¦ì¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
'   End If
'End Sub

Private Sub StartLetter1(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(1 To 50) As String, j As Integer, strTmp As String
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   j = 1
   '½Ð´Ú¨ç¤é´Á
   If frm060306.Text5.Text <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç¤é´Á','" & DBDATE(frm060306.Text5.Text) & "')"
      j = j + 1
   End If
   'Àu¥ýÅv
   If m_JPPriData <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','Àu¥ýÅv¥D±i','" & m_JPPriData & "')"
      j = j + 1
   End If
   
   'Add by Morgan 2005/3/31
   '­×¥¿¹ê¼f
   If Text2(1) <> "" Or Text2(3) <> "" Then
      strTmp = ""
      If Text2(3) <> "" Then
         'Modified by Morgan 2022/10/24
         'strTmp = strTmp & vbCrLf & "¡@¡¯ üÁÊ^¼f¬dÇU½Ð¨D "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "¹ê¼f½Ð¨D")
      End If
      If Text2(1) <> "" Then
         'Modify by Morgan 2008/4/22
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ¦Ûúú¸É¥¿ÇU´£¥X "
         'Modified by Morgan 2022/10/24
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ©ú²Ó®ÑÇU¸É¥¿ "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "»¡©ú®Ñ­×¥¿")
      End If
      If strTmp <> "" Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','­×¥¿¹ê¼f','" & strTmp & "')"
         j = j + 1
      End If
   End If
   '¬O§_¸É¤å¥ó
   If txtDate.Text <> "" Then
      strTmp = ""
      'Åý´ç®Ñ
      If Check1(1).Value = 1 Then
         'Modified by Morgan 2022/10/24
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ¥Ó½Ð“¸’ð´çµý©ú®Ñ "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "Åý´ç®Ñ")
      End If
      '©e¥ôª¬
      If Check1(2).Value = 1 Then
         'Modified by Morgan 2022/10/26
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ©e¥ôûì "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "©e¥ôª¬")
      End If
      'Àu¥ýÅvÃÒ©ú
      If Check1(3).Value = 1 Then
         'Modified by Morgan 2022/10/26
         'strTmp = strTmp & vbCrLf & "¡@¡¯ Àu¥ý“¸µý©ú®Ñ "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "Àu¥ýÅvÃÒ©ú")
      End If
      If strTmp <> "" Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¥¿´Á­­'," & DBDATE(txtDate.Text) & ")"
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¸É¥ó','" & strTmp & "')"
         j = j + 1
      End If
   End If
   'ªþ¥ó
   If Val(Text3.Text) > 0 Then
      strTmp = ""
      'Debit Note
      If Check2(0).Value = 1 Then
         'Modified by Morgan 2022/10/26
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ¹ú©ÒÇU½Ð¨D®Ñ "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "½Ð´Ú³æ")
      End If
      '¦¬¾Ú¼v¥»
      If Check2(1).Value = 1 Then
         'Modified by Morgan 2022/10/26
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ¥XÄ@»â¦¬®ÑÇU‡Àþê "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "¦¬¾Ú¼v¥»")
      End If
      '¤¤»¡
      If Check2(2).Value = 1 And Val(Text4.Text) > 0 Then
         'Modified by Morgan 2022/10/26
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ¤¤üÂ»y©ú²Ó®Ñ "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "¤¤»¡")
      End If
      '¤w­×¥¿¤§»¡©ú®Ñ/Claims
      If Check2(3).Value = 1 And Val(Text4.Text) > 0 Then
         'Modify by Morgan 2007/7/4
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ¦Û¥D¸É¥¿ÇU±±Æî "
         'Modify by Morgan 2008/4/22
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ¦Ûúú¸É¥¿ÇU±±Æî "
         'Modified by Morgan 2022/10/26
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ¸É¥¿¤º®eÇU±±Æî "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "¤w­×¥¿¤§»¡©ú®Ñ")
         'end 2007/7/4
      End If
      '¹ê¼f½Ð¨D®Ñ
      If Check2(4).Value = 1 And Val(Text4.Text) > 0 Then
         'Modified by Morgan 2022/10/26
         'strTmp = strTmp & vbCrLf & "¡@¡¯ üÁÊ^¼f¬dÇU¥Ó½Ð®Ñ "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "¹ê¼f½Ð¨D®Ñ")
      End If
          
      'Added by Morgan 2013/1/28
      '±M§Q¥Ó½Ð®Ñ
      If Check2(5).Value = 1 Then
         'Modified by Morgan 2022/10/26
         'strTmp = strTmp & vbCrLf & "¡@¡¯ ¥XÄ@Ä@®Ñ "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "±M§Q¥Ó½Ð®Ñ")
      End If
      
      'IPO¸É¥ó¨ç
      If Check2(6).Value = 1 Then
         'Modified by Morgan 2022/10/26
         'strTmp = strTmp & vbCrLf & "¡@¡¯ §½¨ü²z³qª¾®Ñ "
         strTmp = strTmp & vbCrLf & "¡@¡¯ " & PUB_GetUniText(Me.Name, "IPO¸É¥ó¨ç")
      End If
      'end 2013/1/28
   
      If strTmp <> "" Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó','" & strTmp & "')"
         j = j + 1
      End If
      
   End If
   'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
   'If Not objLawDll.ExecSQL(j - 1, strTxt) Then
   If Not ClsLawExecSQL(j - 1, strTxt) Then
      MsgBox "Àx¦s¨Ò¥~Äæ¦ì¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
   End If
End Sub
'Added by Morgan 2012/11/30 ½Æ»s StartLetter ¨Ó­×§ï
Private Sub StartLetter2(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(1 To 50) As String, j As Integer, strTmp As String, oChk As CheckBox
 'Dim strDateList1 As String, strDateList2 As String 'Removed by Morgan 2022/4/22 ¨S¥Î¤F
 
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   j = 1
   
   If m_PA11 = "N" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES" & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µL¥Ó½Ð®×¸¹','¡ð')"
      j = j + 1
   End If
   
   If Left(m_PA75, 6) = "Y45148" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','Nikon®×¥ó','¡ð')"
      j = j + 1
   End If
   
   '½Ð´Ú¨ç¤é´Á
   If frm060306.Text5.Text <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç¤é´Á','" & DBDATE(frm060306.Text5.Text) & "')"
      j = j + 1
   End If
   '½Ð´Ú¨ç³Æµù
   If Text1(0) <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç³Æµù','P.S. " & ChgSQL(Text1(0).Text) & "')"
      j = j + 1
   End If
   
   '¦³¸É¤å¥ó
   intI = 0
   For Each oChk In Check1
      If oChk.Value = 1 Then
         intI = 1
         Exit For
      End If
   Next
   If intI = 1 Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES" & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦³¸É¤å¥ó','¡ð')"
      j = j + 1
   Else
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES" & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¨S¦³¸É¤å¥ó','¡ð')"
      j = j + 1
   End If
   
   'Removed by Morgan 2022/4/22 ¨S¥Î¤F
   'strDateList1 = ""
   'strDateList2 = ""
   'end 2022/4/22
   
   '°ê¤º±H¦sÃÒ©ú
   If Check1(0).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'°ê¤º±H¦sÃÒ©ú')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','°ê¤º±H¦sÃÒ©ú¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','°ê¤º±H¦sÃÒ©ú¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','°ê¤º±H¦sÃÒ©úªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
      End If
   End If
   '°ê¥~±H¦sÃÒ©ú
   If Check1(4).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'°ê¥~±H¦sÃÒ©ú')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','°ê¥~±H¦sÃÒ©ú¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','°ê¥~±H¦sÃÒ©ú¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','°ê¥~±H¦sÃÒ©úªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   '©e¥ô®Ñ
   If Check1(2).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and (instr(np15,'©e¥ô®Ñ')>0 or instr(np15,'©e¥ôª¬')>0) and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','©e¥ô®Ñ¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','©e¥ô®Ñ¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','©e¥ô®Ñªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   'Àu¥ýÅvÃÒ©ú
   If Check1(3).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'Àu¥ýÅvÃÒ©ú')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','Àu¥ýÅvÃÒ©ú¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','Àu¥ýÅvÃÒ©ú¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','Àu¥ýÅvÃÒ©úªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   'ªk¤H¦a¦ìÃÒ©ú
   If Check1(5).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'ªk¤H¦a¦ìÃÒ©ú')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªk¤H¦a¦ìÃÒ©ú¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªk¤H¦a¦ìÃÒ©ú¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªk¤H¦a¦ìÃÒ©úªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   '°êÄyÃÒ©ú
   If Check1(6).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'°êÄyÃÒ©ú')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','°êÄyÃÒ©ú¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','°êÄyÃÒ©ú¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','°êÄyÃÒ©úªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   'Add By Sindy 2021/7/20
   'µo©ú¤H¤¤Ä¶¦W
   If Check1(11).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'µo©ú¤H¤¤Ä¶¦W')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µo©ú¤H¤¤Ä¶¦W¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µo©ú¤H¤¤Ä¶¦W¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µo©ú¤H¤¤Ä¶¦Wªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   'µo©ú¤H°êÄy
   If Check1(12).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'µo©ú¤H°êÄy')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µo©ú¤H°êÄy¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µo©ú¤H°êÄy¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µo©ú¤H°êÄyªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   '¥Ó½Ð¤H¤¤Ä¶¦W
   If Check1(13).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'¥Ó½Ð¤H¤¤Ä¶¦W')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥Ó½Ð¤H¤¤Ä¶¦W¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥Ó½Ð¤H¤¤Ä¶¦W¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥Ó½Ð¤H¤¤Ä¶¦Wªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   '¥Ó½Ð¤H°êÄy
   If Check1(14).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'¥Ó½Ð¤H°êÄy')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥Ó½Ð¤H°êÄy¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥Ó½Ð¤H°êÄy¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥Ó½Ð¤H°êÄyªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   '¥Nªí¤H
   If Check1(15).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'¥Nªí¤H')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥Nªí¤H¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥Nªí¤H¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥Nªí¤Hªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   '­^¤åºK­n
   If Check1(16).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07 in (202,231) and instr(np15,'­^¤åºK­n')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','­^¤åºK­n¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         Else
         '2021/7/20 END
            strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','­^¤åºK­n¸É¥ó´Á­­','" & RsTemp("np08") & "')"
         End If
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','­^¤åºK­nªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = RsTemp("np08")
         '   End If
         '   strDateList2 = RsTemp("np09")
         'Else
         '   'Add by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
         '   If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         '      strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   Else
         '   '2021/7/20 END
         '      strDateList1 = strDateList1 & "," & RsTemp("np08")
         '   End If
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   '2021/7/20 END
   
   'Added by Morgan 2021/10/12
   '¥xÄyµo©ú¤HID½X
   If Check1(17).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07=202 and instr(np15,'¥xÄyµo©ú¤HID½X')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥xÄyµo©ú¤HID½X¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥xÄyµo©ú¤HID½Xªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   strDateList1 = RsTemp("np23")
         '   strDateList2 = RsTemp("np09")
         'Else
         '   strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   'end 2021/10/12
   
   'Added by Morgan 2022/1/13
   '½Ð¨D¸ê°T¤£¤½¶}¤§Án©ú®Ñ
   If Check1(18).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07=202 and instr(np15,'½Ð¨D¸ê°T¤£¤½¶}¤§Án©ú®Ñ')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð¨D¸ê°T¤£¤½¶}¤§Án©ú®Ñ¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð¨D¸ê°T¤£¤½¶}¤§Án©ú®Ñªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   strDateList1 = RsTemp("np23")
         '   strDateList2 = RsTemp("np09")
         'Else
         '   strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   'end 2022/1/13
   
   'Added by Morgan 2022/1/13
   '¥þ³¡µo©ú¤H¦W
   If Check1(19).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07=202 and instr(np15,'¥þ³¡µo©ú¤H¦W')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥þ³¡µo©ú¤H¦W¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¥þ³¡µo©ú¤H¦Wªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   strDateList1 = RsTemp("np23")
         '   strDateList2 = RsTemp("np09")
         'Else
         '   strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   'µo©ú¤H¦W¦³¯S®í¦r
   If Check1(20).Value = 1 Then
      strExc(0) = "select np08,np09,np23 from nextprogress where np02='" & m_CP01 & "' and np03='" & m_CP02 & "' and np04='" & m_CP03 & "' and np05='" & m_CP04 & "' and np06 is null" & _
         " and np07=202 and instr(np15,'µo©ú¤H¦W¦³¯S®í¦r')>0 and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µo©ú¤H¦W¦³¯S®í¦r¸É¥ó´Á­­','" & RsTemp("np23") & "')"
         
         j = j + 1
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            " values('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µo©ú¤H¦W¦³¯S®í¦rªk©w´Á­­','" & RsTemp("np09") & "')"
         j = j + 1
         
         'Removed by Morgan 2022/4/22 ¨S¥Î¤F
         'If strDateList1 = "" Then
         '   strDateList1 = RsTemp("np23")
         '   strDateList2 = RsTemp("np09")
         'Else
         '   strDateList1 = strDateList1 & "," & RsTemp("np23")
         '   strDateList2 = strDateList2 & "," & RsTemp("np09")
         'End If
         'end 2022/4/22
         
      End If
   End If
   If m_strOtherDoc <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¨ä¥L¸É¤å¥ó','" & ChgSQL(m_strOtherDoc) & "')"
      j = j + 1
   End If
   'end 2022/1/13
   
   'Removed by Morgan 2022/4/22 ¨S¥Î¤F
   'If strDateList1 <> "" Then
   '   strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
   '      "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¨ä¥L¤é´Á','" & strDateList1 & "')"
   '   j = j + 1
   'End If
   'If strDateList2 <> "" Then
   '   strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
   '      "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªk©w´Á­­','" & strDateList2 & "')"
   '   j = j + 1
   'End If
   'end 2022/4/22
   
   '¬O§_­×¥¿Claims
   If Text2(1) = "Y" Then
      'Modified by Lydia 2015/04/27 ¤¤»¡½Ð´Ú
     ' strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦³­×¥¿»¡©ú®Ñ','¡ð')"
      If m_bolAMD Then
        strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
           "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦³­×¥¿»¡©ú®Ñ','" & ChgSQL(m_AMD05) & "')"
      Else '­ì¤å
        strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
           "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦³­×¥¿»¡©ú®Ñ','As per your instructions, we have amended the specification/claims.')"
      End If
      
      j = j + 1
   End If
   '¬O§_½Ð¨D¹ê¼f
   If Text2(3) = "Y" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦³½Ð¨D¹ê¼f','¡ð')"
      j = j + 1
   End If
   '¬O§_ÀËµø¤Î­×¥¿¤j³°¤¤»¡
   If Text2(2) = "Y" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦³ÀËµø¤Î­×¥¿¤j³°¤¤»¡','¡ð')"
      j = j + 1
   End If
   
   'ªþ¥ó
   intI = 0
   For Each oChk In Check2
      If oChk.Value = 1 Then
         intI = 1
         Exit For
      End If
   Next
   
   If intI = 1 Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦³ªþ¥ó','¡ð')"
      j = j + 1
   End If
   
   If Check2(2).Value = 1 Then
      'Modified by Lydia 2015/04/27 ¤¤»¡½Ð´Ú
      'strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤¤»¡','¡ð')"
      If m_bolAMD Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                 "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤¤»¡','specification in Chinese incorporated with the amendments')"
      Else  '­ì¤å
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                 "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤¤»¡','Chinese specification')"
      End If
      
      j = j + 1
      If Val(Text4.Text) > 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤¤»¡¥÷¼Æ','" & Val(Text4.Text) & "')"
         j = j + 1
      End If
   End If
   
   'Modified by Lydia 2015/04/27 ­ìDobit -> Debit
   If Check2(0).Value = 1 Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','Debit Note','¡ð')"
      j = j + 1
      If Val(Text3.Text) > 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','Debit Note¥÷¼Æ','" & Val(Text3.Text) & "')"
         j = j + 1
      End If
   End If
   If Check2(1).Value = 1 Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦¬¾Ú¼v¥»','¡ð')"
      j = j + 1
      If Val(Text3.Text) > 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦¬¾Ú¼v¥»¥÷¼Æ','" & Val(Text3.Text) & "')"
         j = j + 1
      End If
   End If
   If Check2(3).Value = 1 Then
      'Modified by Lydia 2015/04/27 ¤¤»¡½Ð´Ú
     ' strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤w­×¥¿¤§»¡©ú®Ñ/Claims','¡ð')"
      If m_bolAMD Then
        strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
           "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤w­×¥¿¤§»¡©ú®Ñ/Claims','pages of the specification in English')"
      Else '­ì¤å
        strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
           "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤w­×¥¿¤§»¡©ú®Ñ/Claims','specification/claims')"
      End If
      
      j = j + 1
      If Val(Text3.Text) > 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤w­×¥¿¤§»¡©ú®Ñ/Claims¥÷¼Æ','" & Val(Text3.Text) & "')"
         j = j + 1
      End If
   End If
   If Check2(4).Value = 1 Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¹ê¼f³qª¾®Ñ','¡ð')"
      j = j + 1
      If Val(Text3.Text) > 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¹ê¼f³qª¾®Ñ¥÷¼Æ','" & Val(Text3.Text) & "')"
         j = j + 1
      End If
   End If
   If Check2(5).Value = 1 Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','±M§Q¥Ó½Ð®Ñ','¡ð')"
      j = j + 1
      If Val(Text3.Text) > 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','±M§Q¥Ó½Ð®Ñ¥÷¼Æ','" & Val(Text3.Text) & "')"
         j = j + 1
      End If
   End If
   If Check2(6).Value = 1 Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','IPO¸É¥ó¨ç','¡ð')"
      j = j + 1
      If Val(Text3.Text) > 1 Then
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','IPO¸É¥ó¨ç¥÷¼Æ','" & Val(Text3.Text) & "')"
         j = j + 1
      End If
   End If
   
   If Not ClsLawExecSQL(j - 1, strTxt) Then
      MsgBox "Àx¦s¨Ò¥~Äæ¦ì¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
   End If
End Sub

'Added by Morgan 2016/10/12
Private Sub StartLetter3(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(1 To 50) As String, j As Integer, strTmp As String
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   j = 1
   '½Ð´Ú¨ç¤é´Á
   If frm060306.Text5.Text <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç¤é´Á','" & DBDATE(frm060306.Text5.Text) & "')"
      j = j + 1
   End If
   
   '½Ð´Ú¨ç³Æµù
   If Text1(0) <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç³Æµù','³Æµù¡G" & ChgSQL(Text1(0).Text) & "')"
      j = j + 1
   End If
   
   strTmp = ""
   '¹ê¼f
   If Text2(3) <> "" Then
      strTmp = "©ó¥Ó½Ð®É¤@¨Ö´£¥X¹êÅé¼f¬d¤§½Ð¨D"
   End If
   
   'Àu¥ýÅv
   If m_CNPriData <> "" Then
      If strTmp <> "" Then
         strTmp = strTmp & "¡A¨Ã"
      End If
      strTmp = strTmp & "¥D±i" & m_CNPriData & "Àu¥ýÅv"
   End If
   
   If strTmp <> "" Then
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¹ê¼f»PÀu¥ýÅv','" & strTmp & "')"
      j = j + 1
   End If
   
   '¬O§_¸É¤å¥ó --¥¼´£¨Ñ½d¥»
   'ªþ¥ó --´£¨Ñ½d¥»¤£§¹¾ã
   If Val(Text3.Text) > 0 Then
      strTmp = ""
      '¤¤»¡
      If Check2(2).Value = 1 And Val(Text4.Text) > 0 Then
         strTmp = "¥æ§½¤§¤¤¤å»¡©ú®Ñ"
      End If
      'Debit Note
      If Check2(0).Value = 1 Then
         strTmp = strTmp & IIf(strTmp = "", "", "¡B") & "¥»©Ò±b³æ"
      End If

      '¤w­×¥¿¤§»¡©ú®Ñ/Claims
      If Check2(3).Value = 1 And Val(Text4.Text) > 0 Then
      End If
      '¹ê¼f½Ð¨D®Ñ
      If Check2(4).Value = 1 And Val(Text4.Text) > 0 Then

      End If

      If Check2(5).Value = 1 Then

      End If
      
      If Check2(6).Value = 1 Then

      End If
   
      If strTmp <> "" Then
         If InStr(strTmp, "¡B") > 0 Then
            strTmp = Left(strTmp, InStrRev(strTmp, "¡B") - 1) & "¤Î" & Mid(strTmp, InStrRev(strTmp, "¡B") + 1)
         End If
         
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó','" & strTmp & "')"
         j = j + 1
      End If
      
   End If
   If Not ClsLawExecSQL(j - 1, strTxt) Then
      MsgBox "Àx¦s¨Ò¥~Äæ¦ì¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
   End If
End Sub
Private Sub Check2_Click(Index As Integer)
   If Index = 2 Then
      If Check2(Index).Value = 1 Then
         Text5.Text = Text4.Text
      Else
         Text5.Text = ""
      End If
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim bolChk As Boolean
   Dim strTmp As String
   Dim Cancel As Boolean
   Dim stLetter As String
   Dim strTmp1 As String
   Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
 
   Select Case Index
      Case 2
         Unload frm060306
         Unload Me
      Case 0
        
         'Add by Morgan 2004/10/7
         Cancel = False
         Text3_Validate Cancel
         If Cancel = True Then Exit Sub
         Cancel = False
         Text4_Validate Cancel
         If Cancel = True Then Exit Sub
         '2004/10/7 end
         
         Screen.MousePointer = vbHourglass
         If Text2(0).Text = "Y" Then bolChk = True
         
         'Add by Morgan 2004/11/3
         '¤é¥»
         If m_LetterLanguage = 3 Then
            strTmp = "30"
            strTmp1 = "98"
            If Text2(4).Text <> "Y" Then StartLetter1 ET01, strTmp
         
         'Added by Morgan 2016/10/12 ¤¤¤å
         ElseIf m_LetterLanguage = 1 Then
            strTmp = "13"
            strTmp1 = "97"
            If Text2(4).Text <> "Y" Then StartLetter3 ET01, strTmp
         'end 2016/10/12
         Else
         '2004/11/3 end
         
            'Added by Morgan 2012/11/30
            '²Î¤@¥Î·s©w½Z
            If strSrvDate(1) >= "20130101" Then
               'Modify By Sindy 2021/7/21 Elaine:Nikonhªº­^¤å©w½Z²Î¤@¨Ï¥Î·s©w½Z§Y¥i¡C
'               If Left(m_PA75, 6) = "Y45148" Then
'                  strTmp = "25"
'               Else
                  strTmp = "24"
'               End If
               strTmp1 = "97"
               If Text2(4).Text <> "Y" Then StartLetter2 ET01, strTmp
            Else
            'end 2012/11/30
            
'Removed by Morgan 2018/6/20 §R°£¤£¦A¨Ï¥Îªº©w½Z
'
'               Select Case m_PA11
'                  Case ""  '¦³¥Ó½Ð®×¸¹
'                     'Modified by Morgan 2012/8/30
'                     '¦³µLÀu¥ýÅv©w½Z¦X¨Ö
'                     'Select Case m_PriDate
'                     '   Case ""    '¦³Àu¥ýÅv
'                           If txtDate.Text <> "" Then
'                              strTmp = "22"   '¦³¸É¥ó´Á­­
'                              'Add by Morgan 2007/4/14 Nikon(Y45148)©w½Z¯S§O
'                              If Left(m_PA75, 6) = "Y45148" Then
'                                 strTmp = "31"
'                              End If
'                              'end 2007/4/14
'                           Else
'                              strTmp = "23"   'µL¸É¥ó´Á­­
'                           End If
'                     '   Case "N"   'µLÀu¥ýÅv
'                     '      If txtDate.Text <> "" Then
'                     '         strTmp = "24"   '¦³¸É¥ó´Á­­
'                     '         'Add by Morgan 2007/4/14 Nikon(Y45148)©w½Z¯S§O
'                     '         If Left(m_PA75, 6) = "Y45148" Then
'                     '            strTmp = "32"
'                     '         End If
'                     '         'end 2007/4/14
'                     '      Else
'                     '         strTmp = "25"   'µL¸É¥ó´Á­­
'                     '      End If
'                     'End Select
'                     'end 2012/8/30
'
'                  Case "N" 'µL¥Ó½Ð®×¸¹
'                     'Modified by Morgan 2012/8/30
'                     '¦³µLÀu¥ýÅv©w½Z¦X¨Ö
'                     'Select Case m_PriDate
'                     '   Case ""    '¦³Àu¥ýÅv
'                           If txtDate.Text <> "" Then
'                              strTmp = "26"   '¦³¸É¥ó´Á­­
'                              'Add by Morgan 2007/4/14 Nikon(Y45148)©w½Z¯S§O
'                              If Left(m_PA75, 6) = "Y45148" Then
'                                 strTmp = "33"
'                              End If
'                              'end 2007/4/14
'                           Else
'                              strTmp = "27"   'µL¸É¥ó´Á­­
'                           End If
'                     '   Case "N"   'µLÀu¥ýÅv
'                     '      If txtDate.Text <> "" Then
'                     '         strTmp = "28"   '¦³¸É¥ó´Á­­
'                     '         'Add by Morgan 2007/4/14 Nikon(Y45148)©w½Z¯S§O
'                     '         If Left(m_PA75, 6) = "Y45148" Then
'                     '            strTmp = "34"
'                     '         End If
'                     '         'end 2007/4/14
'                     '      Else
'                     '         strTmp = "29"   'µL¸É¥ó´Á­­
'                     '      End If
'                     'End Select
'                     'end 2012/8/30
'               End Select
'               strTmp1 = "97"
'               If Text2(4).Text <> "Y" Then StartLetter ET01, strTmp
'end 2018/6/20
               
            End If 'Added by Morgan 2012/11/30
         End If
         
         'Add by Morgan 2008/3/24 §PÂ_¬O§_²£¥Í¹q¤lÀÉ
         bolEmail = PUB_GetEMailFlag(m_CP01 & m_CP02 & m_CP03 & m_CP04, , , bolPlusPaper)
         'Add by Morgan 2009/10/20 +§PÂ_¬O§_EMail¦P®É±H¯È¥»
         If bolPlusPaper Then
            iCopy = 0
         Else
            iCopy = 1
         End If
         'end 2009/10/20
         'Add by Morgan 2004/10/7 ·í³]©w¥u¦L Cover Page ®É±±¨î¤£¦L«ü¥Ü«H
         'Modify by Morgan 2008/4/24 ¹q¤lÀÉ¤£­n«Ê­±--ÀRªÚ
         If bolChk Then
            '¥u§ï«Ê­±
            If Text2(4).Text = "Y" Then
               NowPrint strReceiveNo, ET01, strTmp1, bolChk, strUserNum
            Else
               'Add by Morgan 2008/3/31
               If bolEmail Then
                  NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum, , , , , iCopy, , True, True
                  'Add by Morgan 2009/10/19 +§PÂ_¬O§_EMail¦P®É±H¯È¥»
                  If Val(Text5.Text) > 0 And bolPlusPaper Then
                     NowPrint strReceiveNo, ET01, strTmp1, bolChk, strUserNum
                  End If
               Else
                  '¥u§ï«ü¥Ü«H
                  If Val(Text5.Text) = 0 Then
                     NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum
                  Else
                     NowPrint strReceiveNo, ET01, strTmp1, bolChk, strUserNum, , , True, stLetter
                     NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum, , stLetter
                  End If
               End If
            End If
         Else
            If Text2(4).Text <> "Y" Then
               'Add by Morgan 2008/3/31
               If bolEmail Then
                  NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum, , , , , iCopy, , True, True
                  MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(m_CP01) & " ]¡I"
               Else
               'end 2008/3/31
                  NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum
               End If
            End If
            
            '«Ê­±
            If Val(Text5.Text) > 0 Then
               If Not bolEmail Or bolPlusPaper Then
                  NowPrint strReceiveNo, ET01, strTmp1, bolChk, strUserNum, , , , , Val(Text5.Text)
               End If
            End If
         End If
         '2004/10/7 end
         
         If Not bolEmail Or bolPlusPaper Then
'            'Add By Sindy 2015/9/21 ¤é¤å©w½Z¤~­n¦L¦a§}±ø
'            If m_LetterLanguage = "3" Or Val(¥~±M¶}µ¡«H¨ç±Ò¥Î¤é) >= Val(strSrvDate(1)) Then
'            '2015/9/21 END
            'Add By Sindy 2017/3/20 ¤é¤å©w½Z¤~­n¦L¦a§}±ø
            If frm060306.m_FCna01 = "101" Or m_LetterLanguage = "3" Then '¬ü°ê ©Î ¤é¤å©w½Z¤~­n¦L¦a§}±ø
            '2017/3/20 END
               '·s¼W¦a§}±ø¦Cªí¸ê®Æ
               pub_AddressListSN = pub_AddressListSN + 1
               PUB_AddNewAddressList strUserNum, frm060306.Text1.Text, frm060306.Text2.Text, frm060306.Text3.Text, frm060306.Text4.Text, "" & pub_AddressListSN, IIf(Me.Check2(2).Value = vbChecked, IIf(Me.Text3.Text = "", "1", Me.Text3.Text), "0")
            End If
         End If
         
         frm060306.Show
         frm060306.Clear
         Screen.MousePointer = vbDefault
         Unload Me
      Case 1
         frm060306.Show
         Unload Me
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm060306_1 = Nothing
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = °ê¥~_FC
   txtDate.Locked = True 'Add by Morgan 2007/4/14 ¤£¦A´£¨Ñ¿é¤J¥H§K¸òNP¤£¤@­P
   ReadPatent
   'Add by Morgan 2004/10/7
   If InStr("301,302,303,304,305,306,307", m_CP10) > 0 Then
      Text2(4).Text = "Y": Text5.Text = "2"
   Else
      Text2(4).Text = ""
   End If
   '2004/10/7 end

End Sub

Private Sub ReadPatent()
 'edit by nickc 2007/02/02
 'Dim lbl As Label, pA(1 To T_PA) As String, i As Integer, strTmp As String
 Dim Lbl As Object, pa() As String, i As Integer, strTmp As String
 'add by nickc 2007/02/02
 ReDim pa(1 To TF_PA) As String
 Dim oChk As CheckBox 'Added by Morgan 2012/11/30
 
   'Added by Morgan 2022/2/10 ±q¤U­±²¾¤W¨Ó
   m_strOtherDoc = ""
   Label1(9).Visible = False
   txtDate.Visible = False
   For Each oChk In Check1
      oChk.Enabled = False
   Next
   'end 2022/2/10
   
   For Each Lbl In Label2
      Lbl = ""
   Next
   strReceiveNo = frm060306.Tag
   pa(1) = frm060306.Text1.Text
   pa(2) = frm060306.Text2.Text
   pa(3) = frm060306.Text3.Text
   pa(4) = frm060306.Text4.Text
   m_CP01 = pa(1): m_CP02 = pa(2): m_CP03 = pa(3): m_CP04 = pa(4) 'Add by Morgan 2008/3/31
   
   'Modify by Morgan 2006/6/2
   'm_LetterLanguage = GetLetterLanguage(pA(1), pA(2), pA(3), pA(4)) 'Add by Morgan 2004/10/7
   m_LetterLanguage = PUB_GetLanguage(pa(1), pa(2), pa(3), pa(4))
   
   Label2(0).Caption = GiveSymbol(pa(1), pa(2), pa(3), pa(4))
   Label2(1).Caption = frm060306.Text5.Text
    'Modify By Cheng 2002/12/16
'   SetComboToCombo frm060306.Combo1, Combo1
   SetComboToCombo Combo1, frm060306.Combo1
   
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa, intWhere) Then  'edit by nickc 2007/02/02 ¤£¥Î dll ¤F  If objPublicData.ReadPatentDatabase(pA, intWhere) Then
            If PA51CU58FA07(pa) Then
               For i = 1 To 6
                  Label2(i + 1) = pa(50 + i)
               Next
            End If
            m_PA75 = pa(75) 'Add by Morgan 2007/4/14
         End If
      Case "FG"
         If ClsPDReadServicePracticeDatabase(pa, intWhere) Then 'edit by nickc 2007/02/02 ¤£¥Î dll ¤F If objPublicData.ReadServicePracticeDatabase(pA, intWhere) Then
            If PA51CU58FA07(pa) Then Label2(2) = pa(30)
         End If
   End Select
      
   '¸É¥ó´Á­­±a¥»©Ò´Á­­
   'Modify by Morgan 2004/12/10 ¥¼¤WÄò¿ì¤~±a´Á­­
   'strExc(0) = "SELECT NVL(NP08,''),NP15 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07=" & ¸É¤å¥ó
   'Modified by Morgan 2012/11/30 +231 ±H¦sÃÒ©ú
   'Modify By Sindy 2021/7/21 + ,NVL(NP23,'')
   strExc(0) = "SELECT NVL(NP08,'') NP08,NP15,NP09,NVL(NP23,'') NP23 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07  in (202,231) and np06 is null "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modify by Sindy 2021/7/20 §ï§ì¬ù©w´Á­­
      If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         Me.txtDate.Text = ChangeWStringToTString("" & RsTemp.Fields(3))
      Else
      '2021/7/20 END
         Me.txtDate.Text = ChangeWStringToTString("" & RsTemp.Fields(0))
      End If
      m_NP09 = "" & RsTemp.Fields("NP09") 'Add by Morgan 2007/4/14
      strTmp = ""
      Do While Not RsTemp.EOF
         If Not IsNull(RsTemp.Fields(1)) Then
            'Modified by Morgan 2022/2/10
            'strTmp = strTmp & RsTemp.Fields(1)
            strTmp = "" & RsTemp.Fields(1)
            If InStr(strTmp, "°ê¤º±H¦sÃÒ©ú") > 0 Then
               Check1(0).Enabled = True: Check1(0).Value = 1
            ElseIf InStr(strTmp, "©e¥ô®Ñ") > 0 Or InStr(strTmp, "©e¥ôª¬") > 0 Then
               Check1(2).Enabled = True: Check1(2).Value = 1
            ElseIf InStr(strTmp, "Àu¥ýÅvÃÒ©ú") > 0 Then
               Check1(3).Enabled = True: Check1(3).Value = 1
            ElseIf InStr(strTmp, "°ê¥~±H¦sÃÒ©ú") > 0 Then
               Check1(4).Enabled = True: Check1(4).Value = 1
            ElseIf InStr(strTmp, "ªk¤H¦a¦ìÃÒ©ú") > 0 Then
               Check1(5).Enabled = True: Check1(5).Value = 1
            ElseIf InStr(strTmp, "°êÄyÃÒ©ú") > 0 Then
               Check1(6).Enabled = True: Check1(6).Value = 1
            ElseIf InStr(strTmp, "µo©ú¤H¤¤Ä¶¦W") > 0 Then
               Check1(11).Enabled = True: Check1(11).Value = 1
            ElseIf InStr(strTmp, "µo©ú¤H°êÄy") > 0 Then
               Check1(12).Enabled = True: Check1(12).Value = 1
            ElseIf InStr(strTmp, "¥Ó½Ð¤H¤¤Ä¶¦W") > 0 Then
               Check1(13).Enabled = True: Check1(13).Value = 1
            ElseIf InStr(strTmp, "¥Ó½Ð¤H°êÄy") > 0 Then
               Check1(14).Enabled = True: Check1(14).Value = 1
            ElseIf InStr(strTmp, "¥Nªí¤H") > 0 Then
               Check1(15).Enabled = True: Check1(15).Value = 1
            ElseIf InStr(strTmp, "­^¤åºK­n") > 0 Then
               Check1(16).Enabled = True: Check1(16).Value = 1
            ElseIf InStr(strTmp, "¥xÄyµo©ú¤HID½X") > 0 Then
               Check1(17).Enabled = True: Check1(17).Value = 1
            ElseIf InStr(strTmp, "½Ð¨D¸ê°T¤£¤½¶}¤§Án©ú®Ñ") > 0 Then
               Check1(18).Enabled = True: Check1(18).Value = 1
            ElseIf InStr(strTmp, "¥þ³¡µo©ú¤H¦W") > 0 Then
               Check1(19).Enabled = True: Check1(19).Value = 1
            ElseIf InStr(strTmp, "µo©ú¤H¦W¦³¯S®í¦r") > 0 Then
               Check1(20).Enabled = True: Check1(20).Value = 1
            Else
               'Modified by Morgan 2022/4/22 +ªk©w´Á­­
               m_strOtherDoc = m_strOtherDoc & vbCrLf & strTmp & ", ªk©w´Á­­:" & RsTemp.Fields("NP09") & ", ¸É¥ó´Á­­:" & RsTemp.Fields("NP23")
            End If
         End If
         RsTemp.MoveNext
      Loop
   End If

'Removed by Morgan 2022/2/10 ­n¦Ò¼{¦³¥¼¦C¥Xªº¸É¤å¥ó¡A¤ñ·Ó§i¥Ó¤é¸¹©w½Z¤]­n±a¥X(¨ä¥L¸É¤å¥ó)
'   'Added by Morgan 2012/11/30
'   '102.1.1 °_¹w³]¸É¥ó¤º®e¨Ã¨ú®ø³¡¤À¤£¦A¨Ï¥Îªº¶µ¥Ø
'   If strSrvDate(1) >= "20130101" Then
'      Label1(9).Visible = False
'      txtDate.Visible = False
'      For Each oChk In Check1
'         oChk.Enabled = False
'      Next
'
'      If InStr(strTmp, "°ê¤º±H¦sÃÒ©ú") > 0 Then Check1(0).Enabled = True: Check1(0).Value = 1
'      If InStr(strTmp, "°ê¥~±H¦sÃÒ©ú") > 0 Then Check1(4).Enabled = True: Check1(4).Value = 1
'      If InStr(strTmp, "©e¥ô®Ñ") > 0 Or InStr(strTmp, "©e¥ôª¬") > 0 Then Check1(2).Enabled = True: Check1(2).Value = 1
'      If InStr(strTmp, "Àu¥ýÅvÃÒ©ú") > 0 Then Check1(3).Enabled = True: Check1(3).Value = 1
'      If InStr(strTmp, "ªk¤H¦a¦ìÃÒ©ú") > 0 Then Check1(5).Enabled = True: Check1(5).Value = 1
'      If InStr(strTmp, "°êÄyÃÒ©ú") > 0 Then Check1(6).Enabled = True: Check1(6).Value = 1
'
'      'Add By Sindy 2021/7/20
'      If InStr(strTmp, "µo©ú¤H¤¤Ä¶¦W") > 0 Then Check1(11).Enabled = True: Check1(11).Value = 1
'      If InStr(strTmp, "µo©ú¤H°êÄy") > 0 Then Check1(12).Enabled = True: Check1(12).Value = 1
'      If InStr(strTmp, "¥Ó½Ð¤H¤¤Ä¶¦W") > 0 Then Check1(13).Enabled = True: Check1(13).Value = 1
'      If InStr(strTmp, "¥Ó½Ð¤H°êÄy") > 0 Then Check1(14).Enabled = True: Check1(14).Value = 1
'      If InStr(strTmp, "¥Nªí¤H") > 0 Then Check1(15).Enabled = True: Check1(15).Value = 1
'      If InStr(strTmp, "­^¤åºK­n") > 0 Then Check1(16).Enabled = True: Check1(16).Value = 1
'      '2021/7/20 END
'      'Added by Morgan 2021/10/12
'      If InStr(strTmp, "¥xÄyµo©ú¤HID½X") > 0 Then Check1(17).Enabled = True: Check1(17).Value = 1
'      'end 2021/10/12
'      'Added by Morgan 2022/1/13
'      If InStr(strTmp, "½Ð¨D¸ê°T¤£¤½¶}¤§Án©ú®Ñ") > 0 Then Check1(18).Enabled = True: Check1(18).Value = 1
'      'end 2022/1/13
'   End If
'end 2022/2/10

   m_PA11 = "": m_PriDate = ""
   '¦³µL¥Ó½Ð®×¸¹
   If pa(11) = "" Then m_PA11 = "N"
   '¦³µLÀu¥ýÅv¸ê®Æ
   'Modify by Morgan 2004/10/7
   '¦P®É§ì°ê®a
   'strExc(0) = "SELECT * FROM PRIDATE WHERE " & ChgPriDate(pa(1) & pa(2) & pa(3) & pa(4))
   strExc(0) = "SELECT PD05,PD06,NA03,NA04,pd07 FROM PRIDATE,NATION WHERE " & ChgPriDate(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NA01(+)=PD07"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
   'Modify by Morgan 2004/10/7
   '­Y¦³¥D±iÀu¥ýÅv®É¥B¬°¤é¤å©w½Z®É
   'If intI = 0 Then m_PriDate = "N"
   If intI = 0 Then
      m_PriDate = "N"
   ElseIf m_LetterLanguage = 3 Then
      m_JPPriData = ""
      Do While Not RsTemp.EOF
         'Modify by Morgan 2008/5/14 §ï§ì­^¤å--·¶ªÚ
         'm_JPPriData = m_JPPriData & vbCrLf & "¡@¡¯ " & Format("" & RsTemp.Fields("PD05"), "####/##/##") & "¡@" & RsTemp.Fields("NA03") & "¡@" & RsTemp.Fields("PD06")
         m_JPPriData = m_JPPriData & vbCrLf & "¡@¡¯ " & Format("" & RsTemp.Fields("PD05"), "####/##/##") & "¡@" & RsTemp.Fields("NA04") & "¡@" & RsTemp.Fields("PD06")
         RsTemp.MoveNext
      Loop
   'Added by Morgan 2016/10/12
   ElseIf m_LetterLanguage = 1 Then
      m_CNPriData = RsTemp.Fields("NA03") & IIf(RsTemp("pd07") = "012", "(¹q¤l¥æ´«§Î¦¡)", "")
      RsTemp.MoveNext
      Do While Not RsTemp.EOF
         m_CNPriData = m_CNPriData & "¡B" & RsTemp.Fields("NA03") & IIf(RsTemp("pd07") = "012", "(¹q¤l¥æ´«§Î¦¡)", "")
         RsTemp.MoveNext
      Loop
      
      If InStr(m_CNPriData, "¡B") > 0 Then
         m_CNPriData = Left(m_CNPriData, InStrRev(m_CNPriData, "¡B") - 1) & "¤Î" & Mid(m_CNPriData, InStrRev(m_CNPriData, "¡B") + 1)
      End If
   'end 2016/10/12
   End If
   'Add by Morgan 2004/9/8
   m_PA10 = pa(10)
   
   'Add by Morgan 2010/3/23
   If m_PA75 = "Y34232" Then
      Check2(5).Value = vbChecked
      Check2(5).Enabled = False
   'Modify by Morgan 2011/6/1 Y34210¶}ÀY¤£ºÞ±M§QºØÃþ³£¾A¥Î--§d­Yªâ,³¯·¶ªÚ
   'ElseIf m_CP10 = "101" And m_PA75 = "Y34210" Then
   ElseIf Left(m_PA75, 6) = "Y34210" Then
      Check2(5).Value = vbChecked
      Check2(5).Enabled = False
      Check2(6).Value = vbChecked
      Check2(6).Enabled = False
   End If
   
   'Added by Lydia 2015/04/27 §ì®×¥ó¶i«×¤¤¦³201,209,210,235 ¤¤»¡½Ð´Ú,¨Ã¥B¦³­×¥¿©w½Z¤å¦r
   'Added by Lydia 2015/06/25 +¥D°Ê­×¥¿203
    m_bolAMD = False

    strExc(0) = "SELECT AMD05 FROM AmendedText,caseprogress WHERE AMD01=" & CNULL(pa(1)) & " and AMD02=" & CNULL(pa(2)) & " and AMD03=" & CNULL(pa(3)) & " and AMD04=" & CNULL(pa(4)) & _
        "and AMD01=CP01(+) and AMD02=CP02(+) and AMD03=CP03(+) and AMD04=CP04(+) AND AMD09=CP09(+) AND CP10 IN ('201','209','210','235','203') and cp57 is null"
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
       m_AMD05 = "" & RsTemp(0)
       If Len(m_AMD05) > 0 Then
          Text2(1).Text = "Y": Check2(2).Value = 1: Check2(3).Value = 1: m_bolAMD = True
       End If
    End If
      
   'Add By Sindy 2022/5/12 ¹ê¼f¤wµo¤å,¡¨¬O§_½Ð¨D¹ê¼f¡¨´N¤WY
   If PUB_ChkCPExist(pa, "416", 2) = True Then
      Text2(3).Text = "Y"
   End If
   '2022/5/12 END
End Sub

Private Sub Text2_GotFocus(Index As Integer)
   CloseIme
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
   Text4.Text = Text3.Text
   If Check2(2).Value = 1 Then
      Text5.Text = Text4.Text
   End If
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
   If Text3 <> "" Then
      If Val(Text3.Text) = 0 Then
         MsgBox "¥÷¼Æ¤£¥i¬°0¡I"
         Cancel = True
         Text3.SetFocus
         Text3_GotFocus
      End If
   End If
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
   If Check2(2).Value = 1 Then
      Text5.Text = Text4.Text
   End If
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
   If Check2(2).Value = 1 Then
      If Val(Text4.Text) < 1 Then
         MsgBox "¥÷¼Æ¥²¶·¦Ü¤Ö¬°1¡I"
         Cancel = True
         Text4.SetFocus
         Text4_GotFocus
      End If
   End If
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text4
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If (KeyAscii > 57 Or KeyAscii < 48) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtDate_GotFocus()
    'Add By Cheng 2002/12/17
    TextInverse Me.txtDate
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
    'Add By Cheng 2002/12/17
    With Me.txtDate
        If .Text <> "" Then
            If ChkDate(.Text) = False Then
                Cancel = True
                txtDate_GotFocus
            End If
        End If
    End With
End Sub
