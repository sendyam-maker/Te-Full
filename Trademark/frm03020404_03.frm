VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020404_03 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "°Ó¼Ðµoµù¥UÃÒ¿é¤J"
   ClientHeight    =   5750
   ClientLeft      =   3350
   ClientTop       =   2760
   ClientWidth     =   9130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   9130
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   1260
      TabIndex        =   56
      Top             =   4830
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "¤å¨ì¦¸¤é"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   13
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "¤å¨ì·í¤é"
         Height          =   180
         Index           =   0
         Left            =   144
         TabIndex        =   12
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   4140
      TabIndex        =   55
      Top             =   4830
      Width           =   4215
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   17
         Top             =   128
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   15
         Top             =   128
         Width           =   375
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   19
         Top             =   128
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                      ¤é"
         Height          =   225
         Index           =   2
         Left            =   2520
         TabIndex        =   18
         Top             =   180
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "        ¤ë"
         Height          =   180
         Index           =   1
         Left            =   1560
         TabIndex        =   16
         Top             =   180
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "¤å¨ì          ¤Ñ"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   180
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "°Ó«~¤ÎªA°È¸ê®Æ¬d¸ß(&I)"
      Height          =   400
      Index           =   6
      Left            =   3780
      TabIndex        =   20
      Top             =   60
      Width           =   1935
   End
   Begin VB.TextBox textNP09 
      Height          =   285
      Left            =   5970
      MaxLength       =   7
      TabIndex        =   11
      Top             =   4530
      Width           =   2292
   End
   Begin VB.TextBox textNP08 
      Height          =   285
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   10
      Top             =   4530
      Width           =   2292
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   9
      Top             =   4200
      Width           =   492
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5700
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3120
      Width           =   1092
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   5700
      TabIndex        =   6
      Top             =   3420
      Width           =   2895
   End
   Begin VB.TextBox textPrtTrans 
      Height          =   285
      Left            =   6300
      MaxLength       =   1
      TabIndex        =   8
      Top             =   3840
      Width           =   372
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8040
      TabIndex        =   23
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5760
      TabIndex        =   21
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "¦^«eµe­±(&U)"
      Height          =   400
      Left            =   6780
      TabIndex        =   22
      Top             =   60
      Width           =   1212
   End
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3840
      Width           =   732
   End
   Begin VB.TextBox textCreFee 
      Height          =   285
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   5
      Top             =   3480
      Width           =   492
   End
   Begin VB.TextBox textTM14 
      Height          =   285
      Left            =   1260
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox textTM21 
      Height          =   285
      Left            =   1260
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3120
      Width           =   1092
   End
   Begin VB.TextBox textTM22 
      Height          =   264
      Left            =   2700
      MaxLength       =   8
      TabIndex        =   3
      Top             =   3120
      Width           =   1092
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textTM27 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2292
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2412
   End
   Begin VB.TextBox textTM15 
      Height          =   285
      Left            =   5700
      MaxLength       =   20
      TabIndex        =   1
      Top             =   2760
      Width           =   2532
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1260
      TabIndex        =   61
      Top             =   944
      Width           =   7485
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13203;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1260
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   1304
      Width           =   7485
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "13203;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5760
      TabIndex        =   59
      Top             =   600
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LabNP07 
      Height          =   255
      Left            =   8400
      TabIndex        =   58
      Top             =   4980
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label32 
      Caption         =   "¨Ó¨ç´Á­­:"
      Height          =   255
      Left            =   180
      TabIndex        =   57
      Top             =   5010
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¤l®×·sªk©w´Á­­ :"
      Height          =   180
      Index           =   17
      Left            =   4560
      TabIndex        =   54
      Top             =   4560
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¤l®×·s¥»©Ò´Á­­ :"
      Height          =   180
      Index           =   18
      Left            =   180
      TabIndex        =   53
      Top             =   4560
      Width           =   1350
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3840
      TabIndex        =   52
      Top             =   652
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "¬O§_§ó§ïÃÒ®Ñ :"
      Height          =   255
      Index           =   6
      Left            =   180
      TabIndex        =   51
      Top             =   4200
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "(Y:¤º³¡¦¬¤å§ó§ï)"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   50
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "ÃÒ®Ñ¤é´Á :"
      Height          =   255
      Left            =   4710
      TabIndex        =   49
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "¦Lªí¾÷ :"
      Height          =   255
      Left            =   4740
      TabIndex        =   48
      Top             =   3450
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "(N:¤£¦L)"
      Height          =   252
      Left            =   6780
      TabIndex        =   47
      Top             =   3840
      Width           =   852
   End
   Begin VB.Label Label4 
      Caption         =   "¬O§_¦C¦LÂ½Ä¶¨ç :"
      Height          =   252
      Left            =   4740
      TabIndex        =   46
      Top             =   3840
      Width           =   1452
   End
   Begin VB.Label Label22 
      Caption         =   "¦C¦L©w½Z :"
      Height          =   252
      Left            =   180
      TabIndex        =   45
      Top             =   3840
      Width           =   972
   End
   Begin VB.Label Label23 
      Caption         =   "(N:¤£¦L)"
      Height          =   252
      Left            =   2100
      TabIndex        =   44
      Top             =   3840
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "(Y:²£¥Í)"
      Height          =   252
      Index           =   5
      Left            =   2460
      TabIndex        =   43
      Top             =   3480
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "¬O§_²£¥Í½Ð´Ú¸ê®Æ :"
      Height          =   252
      Index           =   3
      Left            =   180
      TabIndex        =   42
      Top             =   3480
      Width           =   1572
   End
   Begin VB.Label Label10 
      Caption         =   "¤½§i¤é :"
      Height          =   252
      Left            =   180
      TabIndex        =   41
      Top             =   2760
      Width           =   732
   End
   Begin VB.Label Label14 
      Caption         =   "±M¥Î´Á­­ :"
      Height          =   252
      Left            =   180
      TabIndex        =   40
      Top             =   3120
      Width           =   972
   End
   Begin VB.Line Line1 
      X1              =   2460
      X2              =   2580
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label27 
      Caption         =   "¥Ó½Ð®×¸¹ :"
      Height          =   255
      Left            =   4740
      TabIndex        =   39
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "°Ó¼ÐºØÃþ :"
      Height          =   252
      Index           =   2
      Left            =   180
      TabIndex        =   37
      Top             =   1680
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "¥»©Ò®×¸¹ :"
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   35
      Top             =   616
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "®×¥ó¦WºÙ :"
      Height          =   252
      Left            =   180
      TabIndex        =   34
      Top             =   960
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "¥Ó½Ð¤H :"
      Height          =   252
      Left            =   180
      TabIndex        =   33
      Top             =   1320
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "¥¿°Ó¼Ð¸¹¼Æ :"
      Height          =   252
      Index           =   4
      Left            =   4740
      TabIndex        =   32
      Top             =   1680
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "°Ó«~Ãþ§O :"
      Height          =   252
      Index           =   7
      Left            =   180
      TabIndex        =   31
      Top             =   2040
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "¨Ó¨ç¦¬¤å¤é :"
      Height          =   252
      Index           =   10
      Left            =   180
      TabIndex        =   30
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "´¼Åv¤H­û :"
      Height          =   252
      Index           =   11
      Left            =   4740
      TabIndex        =   29
      Top             =   616
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "¼f©w¸¹¼Æ :"
      Height          =   255
      Left            =   4740
      TabIndex        =   28
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "frm03020404_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/13 §ï¦¨Form2.0 ; cmbTM05¡BtextTM23¡BtextCP13
'Memo By Sindy 2012/12/4 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/16 SQLDate¤wÀË¬d
'Memo By Sindy 2010/11/29 ­û¤u½s¸¹Äæ¤w­×§ï
'Memo By Sindy 2010/8/11 ¤é´ÁÄæ¤w­×§ï
Option Explicit

' ¥»©Ò®×¸¹
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' ¨Ó¨ç¦¬¤å¤é
Dim m_CP05 As String
' ¦¬¤å¸¹
Dim m_CP09 As String
' ­ì®×¥ó©Ê½è
Dim m_CP10 As String
' ­ì·~°È°Ï
Dim m_CP12 As String
' ­ì´¼Åv¤H­û¥N¸¹
Dim m_CP13 As String
' °ê®a¥N½X
Dim m_TM10 As String
' ¥¿°Ó¼Ð¸¹¼Æ
Dim m_TM27 As String
' °Ó«~Ãþ§O
Dim m_TM08 As String
' ·s¼Wªº¦¬¤å¸¹
Dim strCP09 As String
Dim NowCP09 As String 'Added by Lydia 2020/03/09 ·s¼W¤§µù¥UÃÒ1701¦¬¤å¸¹
Dim strCP05 As String
Dim strCP27 As String
Dim ii As Integer
Dim rsTmp As New ADODB.Recordset

Dim m_CurrSel As Integer
'Add By Cheng 2002/06/06
Dim m_strSerialNo As String '½Ð´Ú³æ¸¹
Public adoacc1k0 As New ADODB.Recordset
'Public adoacc1l0 As New ADODB.Recordset
'Public adoadodc1 As New ADODB.Recordset
'Public adoaccsum As New ADODB.Recordset
'Public adoaccmax As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
'Modify By Cheng 2002/12/13
'Public adocheck As New ADODB.Recordset
'Public adoselect As New ADODB.Recordset
Dim strSql As String
Dim strNo As String
Dim lngAmount As Long
Dim douAmount As Double
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim douUSDollar As Double
Dim strLanguage As String
Dim strMaxNo As String
Dim strDiscount As String
Private Const intDefault As Integer = 500
Private Const intTop As Integer = 1000
Dim strNewPage As String
Dim prnPrint As Printer
Dim strPrint As String
'Add By Cheng 2003/02/19
Dim m_TM67 As String '©ñ±ó±M¥ÎÅv
'Add By Cheng 2003/02/27
Dim m_blnPrintAddress As Boolean '¬O§_­n¦C¦L¦a±ø
'Add By Cheng 2003/12/23
Dim m_TM11 As String '¥Ó½Ð¤é
Dim m_TM14 As String '¤½§i¤é
Dim m_TM58 As String '®×¥ó³Æµù
'ADD BY NICK 2004/08/17
Dim Is716Have As Boolean
'add by nickc 2006/08/04
Public UpForm As Form
Dim m_MonTM01 As String     '¬ö¿ý¤À³Î¥À®×®×¸¹
Dim m_MonTM02 As String
Dim m_MonTM03 As String
Dim m_MonTM04 As String
Public m_MonCP09 As String  '¶Ç¤J¤À³Î¥À®×¦¬¤å¸¹
Dim m_MonNP08 As String
Dim m_MonNP09 As String
'add by nickc 2007/03/08 ¥[¤J¦P·N®Ñ°Ó¼Ð¸¹¼Æ
Dim m_TM118 As String
'End
'92.04.16 nick ¬ö¿ý§@¥Î«öÁä
Public cmdState As Integer
'add by nick 2004/10/05 ÀË¬d¬O§_¤w¸g¦³°Ó«~¤ÎªA°È
Public ChkTG As Boolean
Dim m_blnReceiveSecond As Boolean '§PÂ_¤À³Î¥À®×¬O§_¦¬²Ä¤G´Áµù¥U¶O '2011/9/22 add by sonia
Dim strRvType As String 'Add By Sindy 2012/5/18
Dim m_TM13 As String 'Add By Sindy 2012/12/19 ¼f©w¨Ó¨ç¤é
'Added by Morgan 2017/6/14 ¹q¤l¤½¤å
Public m_DocWord As String
Public m_DocNo As String
Public m_DocPdf As String
Public m_DocPdfDate As String
Public m_DocPdfTime As String
'end 2017/6/14
Dim m_NA85 As String 'Added by Lydia 2019/11/13 ­pºâ°Ó¼Ð±M¥Î´Á¬O§_´î1¤Ñ
Dim m_NA86 As String 'Added by Sindy 2020/4/24 ¬O§_°±¤î¶l°È
Dim m_TM136 As String 'Added by Lydia 2023/02/24 µù¥UÃÒ§Î¦¡
Dim strFN03 As String  'Added by Lydia 2023/06/05 (±qPrintLetterNew²¾¹L¨Ó)ÃÒ®ÑÀÉ¦W

' ­ì¸ê®Æ¬O§_¦³¹ê»Úµ²ªG
Private Sub cmdCancel_Click()
'add by nickc 2008/01/23 ¥[¤J¥i¥H¨ú®ø
If UpForm Is Nothing Or Me.Visible = False Then
   Unload Me
   frm03020404_02.Show
Else
    'add by nickc 2008/01/23 ¥[¤J¥i¥H¨ú®ø
    If UpForm Is frm02010401_6 Then
        frm02010401_6.m_IsCancal = True
        Unload Me
    End If
End If
End Sub

Private Sub cmdExit_Click()
   Unload frm03020404_02
   Unload frm03020404_01
   Unload Me
End Sub

Public Sub cmdok_Click(Index As Integer)
'92.04.16 nick ¬ö¿ý§@¥Î«öÁä
cmdState = Index
PubShowNextData
Exit Sub
End Sub

'Add By Sindy 2009/05/14
Public Sub PubShowNextData()
Select Case cmdState
Dim strFilePath As String 'Added by Lydia 2020/03/09 ±½ºËÀÉªº¸ô®|

Case 0
   If CheckDataValid = True Then
        'Add By Cheng 2002/05/23
        '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
        If TxtValidate = False Then Exit Sub
         If m_DocNo = "" Then 'Added by Morgan 2023/1/17 «D¹q¤l¤½¤å¤~­n
            'Added by Lydia 2020/03/09 ¿éµù¥UÃÒ­Y¯ÊÀÉ«h´£¿ô¤£¥i¿é¤J¡A¤£¯Ê«h¦Û°ÊÂk¤Jµù¥UÃÒ¨º¹D¤§¨÷©v°Ï¡C
            If PUB_FCTCheckPDF(m_TM01, m_TM02, m_TM03, m_TM04, "1701", , strFilePath) = False Then
               Exit Sub
            End If
            'end 2020/03/09
         End If 'Added by Morgan 2023/1/17
        
        'add by nickc 2006/08/04
        If UpForm Is Nothing Or Me.Visible = False Then
            ' ³]©w·Æ¹«´å¼Ð¬°µ¥«Ýª¬ºA
            Screen.MousePointer = vbHourglass
            ' Àx¦s¸ê®Æ
          'edit by  nick 2004/11/03
          'OnSaveData
          If OnSaveData = False Then MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
            'Add By Cheng 2003/02/27
            '·s¼W¦a§}±ø¦Cªí¸ê®Æ
            'Modify By Sindy 2025/10/2 ¨ú®ø¦a§}±ø
'            If m_blnPrintAddress = True Then
'                pub_AddressListSN = pub_AddressListSN + 1
'                PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
'            End If
            ' ³]©w·Æ¹«´å¼Ð¬°¹w³]
            Screen.MousePointer = vbDefault
        
            'Add By Cheng 2003/02/18
            '­Y¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
            If Me.Combo2.Text <> Me.Combo2.Tag Then
                PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
            End If
            
            'Added by Lydia 2020/03/09 FCT®×¿é¤Jµù¥UÃÒ©Î§ó¥¿®Ö­ã(µù¥UÃÒ)«e¡A¥ý±½ºËµù¥UÃÒ¦Ü©T©w¸ê®Æ§¨¡A¿éµù¥UÃÒ­Y¯ÊÀÉ«h´£¿ô¤£¥i¿é¤J¡A¤£¯Ê«h¦Û°ÊÂk¤Jµù¥UÃÒ¨º¹D¤§¨÷©v°Ï¡C
            If strFilePath <> "" Then
                If Pub_AutoSavePdf2_FCT(m_TM01, m_TM02, m_TM03, m_TM04, NowCP09, "1701", strFilePath) = False Then
                    Exit Sub
                End If
            End If
            'end 2020/03/09
            
            If textPrint <> "N" And strFN03 <> "" Then 'Added by Morgan 2025/10/2 ¨S¥X©w½Z¤]¤£¥Î¤U¸üÃÒ®Ñ(ÅÜ¼Æ¨S³]©w¤]·|¿ù) --´ð¼_
            
               'Added by Lydia 2023/06/05 ¹q¤l©Î¯È¥»ÃÒ®Ñ²Î¤@¦b³Ì«á¤U¸ü¨÷©v°ÏªºÃÒ®ÑPDF: ¯È¥»¦bPrintLetterNew¨S¦³¥i¤U¸üªºÀÉ®×; ex.FCT-049497
               'Modified by Morgan 2025/3/28 +CPP19
               strSql = "select cpp14,cpp19 From casepaperpdf where cpp01='" & NowCP09 & "' and instr(upper(cpp02),upper('." & IIf(m_TM136 = "1", "CERT", "1701") & ".PDF'))>0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  If PUB_GetFtpFile("" & RsTemp.Fields("cpp14"), Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04) & "\" & strFN03, "Casepaperpdf", , , "" & RsTemp.Fields("cpp19") <> "") = True Then
                  End If
               End If
               'end 2023/06/05
               
            End If
        End If
        
        If UpForm Is Nothing Then
            'Added by Morgan 2023/1/17
            If m_DocNo <> "" Then
               frm02010412.m_TM14 = textTM14.Text 'Added by Morgan 2023/6/15
               Unload Me
               Unload frm03020404_01
               frm02010412.GoNext
            Else
            'end 2023/1/17
               'Add By Sindy 2019/7/22
               frm03020404_01.m_TM14 = textTM14.Text
               Unload Me
               Unload frm03020404_02
               '2019/7/22 END
               frm03020404_01.Show
               
            End If 'Added by Morgan 2023/1/17
        ElseIf UpForm Is frm02010401_6 Then
          '­Y¬Oµe­±¦³¥X²{¥i¥H¿é¸ê®Æ¡A­n±N¸ê®Æ¥á¦^«e­±¦s
          If Me.Visible = True Then
            frm02010401_6.PutSeekData01 = textTM14
            frm02010401_6.PutSeekData02 = textTM15
            frm02010401_6.PutSeekData03 = textTM21
            frm02010401_6.PutSeekData04 = textTM22
            frm02010401_6.PutSeekData05 = Text1
            frm02010401_6.PutSeekData06 = textCreFee
            frm02010401_6.PutSeekData07 = textPrint
            frm02010401_6.PutSeekData08 = textPrtTrans
            frm02010401_6.PutSeekData09 = Text2
            frm02010401_6.PutSeekData10 = textNP08
            frm02010401_6.PutSeekData11 = textNP09
          End If
          Unload Me
       End If
        
   End If
    
'add by nick 2004/10/05
Case 6
    'frm03010303_04.Hide 'Modify By Sindy 2009/09/17
    Set frm03010303_04.UpForm = Me
    frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 'textTMKey 'lbl1(0).Caption
    frm03010303_04.AllClass = textTM09 'txt1(0).Text
    frm03010303_04.cmdOK(0).Visible = False
    frm03010303_04.cmd.Visible = False
    frm03010303_04.cmd2.Visible = False
    frm03010303_04.txt2(0).Visible = False
    frm03010303_04.Line1.Visible = False
    frm03010303_04.txt2(1).Visible = False
    frm03010303_04.txt2(2).Visible = False
    frm03010303_04.txt2(3).Visible = False
    frm03010303_04.Caption = "°Ó«~¤ÎªA°È¸ê®Æ"
    'edit by nickc 2008/02/12 §ï¦¨¥i¥H½Æ»s
    'frm03010303_04.TXT1(0).Enabled = False
    'frm03010303_04.TXT1(1).Enabled = False
    'frm03010303_04.TXT1(2).Enabled = False
    frm03010303_04.TXT1(0).Locked = True
    frm03010303_04.TXT1(1).Locked = True
    frm03010303_04.TXT1(2).Locked = True
    frm03010303_04.Label2.Visible = False
    'Me.Hide 'Modify By Sindy 2009/09/17
    frm03010303_04.QueryData
    frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 §ï¬°±j¨î¦^À³ªí³æ
End Select
End Sub

Private Sub Form_Load()
   
    ' ³]©w±±¨î¶µªº­I´ºÃC¦â
    textTMKey.BackColor = &H8000000F
    textTM08.BackColor = &H8000000F
    textTM09.BackColor = &H8000000F
    textTM12.BackColor = &H8000000F
    textTM23.BackColor = &H8000000F
    textTM27.BackColor = &H8000000F
    textCP05S.BackColor = &H8000000F
    textCP13.BackColor = &H8000000F
    
    MoveFormToCenter Me
    
    PUB_SetPrinter Me.Name, Combo2, strPrint 'Modified by Morgan 2017/11/21 ³]©w¦Lªí¾÷§ï©I¥s¤½¥Î¨ç¼Æ,­ìµ{¦¡²¾°£
    
    'Add By Cheng 2003/02/27
    '¹w³]¤£¦C¦L¦a§}±ø
    m_blnPrintAddress = False
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' ²M°£·j´MªºKey
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' ¥»©Ò®×¸¹ Äæ¦ì1
      Case 0: m_TM01 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì2
      Case 1: m_TM02 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì3
      Case 2: m_TM03 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì4
      Case 3: m_TM04 = strData
      ' ¨Ó¨ç¦¬¤å¤é
      Case 4: m_CP05 = strData
      ' ¦¬¤å¸¹
      Case 5: m_CP09 = strData
      'Add By Sindy 2019/7/22 ¼È¦s¤½§i¤é
      Case 6: m_TM14 = strData: textTM14.Text = strData
   End Select
End Sub

' Åª¨ú°Ó¼Ð°ò¥»ÀÉ
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
    'Add By Cheng 2002/12/09
    Dim StrSQLa As String
    Dim rsA As New ADODB.Recordset
      
   m_blnReceiveSecond = False '2011/9/19 add by sonia
   ' ¨ú±o°Ó¼Ð°ò¥»ÀÉªº¬ÛÃö¶µ¥Ø
   'Modified by Lydia 2019/11/13 +Nation
   'Modify by Sindy 2020/4/24 ¬O§_°±¤î¶l°È
   strSql = "SELECT x.*,y.NA85,y.NA86 FROM TradeMark x, Nation y " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' AND TM10=NA01(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' ¥Ó½Ð°ê®a
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         m_NA85 = "" & rsTmp.Fields("NA85") 'Added by Lydia 2019/11/13 ­pºâ°Ó¼Ð±M¥Î´Á¬O§_´î1¤Ñ
      End If
      ' ¥Ó½Ð®×¸¹
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' °Ó¼Ð¦WºÙ(¤¤)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' °Ó¼Ð¦WºÙ(­^)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' °Ó¼Ð¦WºÙ(¤é)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' Åã¥Ü°Ó¼Ð¦WºÙ
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' °Ó¼ÐºØÃþ
      If IsNull(rsTmp.Fields("TM08")) = False Then
         m_TM08 = rsTmp.Fields("TM08")
         If m_TM10 < "010" Then
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
      End If
      ' °Ó«~Ãþ§O
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      
      'Add By Sindy 2012/12/19
      ' ¼f©w¨Ó¨ç¤é
      If IsNull(rsTmp.Fields("TM13")) = False Then
         m_TM13 = rsTmp.Fields("TM13")
      Else
         m_TM13 = strSrvDate(1)
      End If
      '2012/12/19 End
      
      ' ¤½§i¤é
      If IsNull(rsTmp.Fields("TM14")) = False Then
         'edit by nick 2004/10/06
         'textTM14 = TAIWANDATE(rsTmp.Fields("TM14"))
         textTM14 = DBDATE(rsTmp.Fields("TM14"))
      End If
      ' ¼f©w¸¹¼Æ
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' ±M¥Î´Á­­(°_)
      If IsNull(rsTmp.Fields("TM21")) = False Then
         'edit by nick 2004/10/06
         'textTM21 = TAIWANDATE(rsTmp.Fields("TM21"))
         textTM21 = DBDATE(rsTmp.Fields("TM21"))
      End If
      ' ±M¥Î´Á­­(¨´)
      If IsNull(rsTmp.Fields("TM22")) = False Then
         'edit by nick 2004/10/06
         'textTM22 = TAIWANDATE(rsTmp.Fields("TM22"))
         textTM22 = DBDATE(rsTmp.Fields("TM22"))
      End If
      ' ¥Ó½Ð¤H
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      
      ' ¥¿°Ó¼Ð¸¹¼Æ
      If IsNull(rsTmp.Fields("TM27")) = False Then
         m_TM27 = rsTmp.Fields("TM27")
         textTM27 = rsTmp.Fields("TM27")
      End If
        'Add By Cheng 2003/02/19
        '©ñ±ó±M¥ÎÅv
        m_TM67 = "" & rsTmp("TM67").Value
        'Add By Cheng 2003/12/23
        '¥Ó½Ð¤é
        m_TM11 = "" & rsTmp("TM11").Value
        '®×¥ó³Æµù
        m_TM58 = "" & rsTmp("TM58").Value
        'End
      'add by nickc 2006/05/29 ¥[¤J³¬¨÷´£¥Ü
      If IsNull(rsTmp.Fields("tm29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "¤w³¬¨÷"
      End If
      'add by nickc 2007/03/08
      m_TM118 = "" & rsTmp("tm118").Value
      '2011/9/22 ADD BY SONIA
      If InStr("" & rsTmp.Fields("TM58"), "²Ä¤G´Á") > 0 Then
         m_blnReceiveSecond = True
      End If
      '2011/9/22 end
      m_TM136 = "" & rsTmp.Fields("TM136") 'Added by Lydia 2023/02/24 µù¥UÃÒ§Î¦¡
   End If
   rsTmp.Close
   Set rsTmp = Nothing
    'Add By Cheng 2002/12/09
    '­Y¦³¥¿°Ó¼Ð¸¹¼Æ
    'If "" & m_TM27 <> "" Then
    '    '­Y°Ó¼ÐºØÃþ¬°2,3«h§ì1; ­Y¬°5,6«h§ì4
    '    If m_TM08 = "2" Or m_TM08 = "3" Then
    '        strSQLA = "Select   TM21,TM22 From TradeMark Where TM15 = '" & m_TM27 & "' And TM08 = '1' "
    '    ElseIf m_TM08 = "5" Or m_TM08 = "6" Then
    '        strSQLA = "Select   TM21,TM22 From TradeMark Where TM15 = '" & m_TM27 & "' And TM08 = '4' "
    '    Else
    '        strSQLA = "Select   TM21,TM22 From TradeMark Where TM15 = '" & m_TM27 & "' "
    '    End If
    '    rsA.CursorLocation = adUseClient
    '    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
    '    If rsA.RecordCount > 0 Then
    '        textTM21 = TAIWANDATE(rsTmp.Fields("TM21"))
    '        textTM22 = TAIWANDATE(rsTmp.Fields("TM22"))
    '    End If
    '    If rsA.RecordCount > 0 Then rsA.Close
    '    Set rsA = Nothing
    'End If

End Sub

' Åª¨ú®×¥ó¶i«×ÀÉ
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   '2012/11/1 add by sonia ¥ý§ì¥Ó½Ð101©Î¤À³Î308,­YµL101¤Î308¤~¥ý§ìAÃþ¦¬¤å (T-179141¼f©w«á¥¼¦A¦¬¤åµù¥U¶O©Î¨ä¥LAÃþ¬G·|§ì¨ì¤À³Î308)
   strSql = "SELECT * FROM CaseProgress WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "' AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' AND " & _
                  "CP10 = '101' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GoTo DisplayData
   End If
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   Set rsTmp = Nothing
   strSql = "SELECT * FROM CaseProgress WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "' AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' AND " & _
                  "CP10 = '308' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GoTo DisplayData
   End If
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   Set rsTmp = Nothing
   '2012/11/1 end
   
   ' ¨ú±o®×¥ó¶i«×ÀÉÀÉ®×¤¤Äæ¦ì
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' AND " & _
                  "CP09 LIKE 'A%' AND " & _
                  "CP05 IN (SELECT MAX(CP05) FROM CaseProgress " & _
                           "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                 "CP02 = '" & m_TM02 & "' AND " & _
                                 "CP03 = '" & m_TM03 & "' AND " & _
                                 "CP04 = '" & m_TM04 & "' AND " & _
                                 "CP09 LIKE 'A%') "
            
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04
   If rsTmp.RecordCount > 0 Then
DisplayData:
     rsTmp.MoveFirst
      ' Á`¦¬¤å¸¹
      If IsNull(rsTmp.Fields("CP09")) = False Then
         m_CP09 = rsTmp.Fields("CP09")
      End If
      'add by nickc 2006/10/18
      m_CP10 = CheckStr(rsTmp.Fields("CP10"))
      ' ·~°È°Ï
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' ´¼Åv¤H­û
      'Modified by Lydia 2021/08/03 §ï¥ÑPUB_GetFCTSalesNo±a¥X©M²£¥ÍªºCÃþ¦¬¤å¤@­P
      'If IsNull(rsTmp.Fields("CP13")) = False Then
      '   m_CP13 = rsTmp.Fields("CP13")
      '   textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      'End If
      m_CP13 = Empty
      m_CP13 = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      textCP13 = GetStaffName(m_CP13)
      'end 2021/08/03
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Public Sub QueryData()
   ' ¨Ó¨ç¦¬¤å¤é
   'add by nickc 2006/08/14
   If UpForm Is frm02010401_6 Then
        textCP05S = TAIWANDATE(UpForm.oStrCDate)
   Else
        textCP05S = m_CP05
   End If
   ' ¥»©Ò®×¸¹
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   m_TM11 = ""
   m_TM58 = ""
   m_TM13 = Empty 'Add By Sindy 2012/12/19 ¼f©w¨Ó¨ç¤é
   
   ' Åª¨ú°Ó¼Ð°ò¥»ÀÉ
   QueryTradeMark
   
   ' Åª¨ú®×¥ó¶i«×ÀÉ
   QueryCaseProgress
   
   'Add By Sindy 2019/7/22 ¹w³]«e¤@µ§¿é¤J¤§¤½§i¤é
   textTM14.Text = m_TM14: Call textTM14_Validate(False)
   m_TM14 = Empty
   
   'add by nickc 2006/10/02
   If UpForm Is frm02010401_6 Then
      QueryMonTradeMark
   End If
   
   'Add by Sindy 2020/4/24 ¬O§_°±¤î¶l°È
   Call GetPrjPeopleNum6(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04, "NA86", m_NA86)
   
   'add by nick 2004/09/24 92.11.28 ¥H«á¥Ó½Ðªº®×¥ó±Hµù¥UÃÒ®É¤£½Ð´Ú
   If DBDATE(Val(m_TM11)) >= 20031128 Then
      textCreFee.Locked = True
   End If
   
   Call ChgType 'Add By Sindy 2012/5/18 Åª¨ú¨Ó¨ç´Á­­
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strSql As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP09 As String
   Dim strNP22 As String
   '93.6.11 ADD BY SONIA
   Dim strCP06 As String
   Dim strCP07 As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   '93.6.11 END
   Dim strCP118 As String 'Add by Amy 2023/02/06 ¬O§_¹q¤l°e¥ó
   
'add by nickc 2006/08/11
If Me.Visible = True Then
     '911107 nick transation
    On Error GoTo CheckingErr
    cnnConnection.BeginTrans
End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' §ó·s°Ó¼Ð°ò¥»ÀÉªºµoÃÒ¤é¤Î±M¥Î´Á¶¡
   'Modify By Cheng 2002/07/11
   '±M¥ÎÅv¬O§_¦s¦b³]¬°"Y"
'   strSQL = "UPDATE TradeMark SET TM20 = " & DBDATE(m_CP05) & ", " & _
'                                 "TM21 = " & DBDATE(textTM21) & ", " & _
'                                 "TM22 = " & DBDATE(textTM22) & " " & _
'            "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                  "TM02 = '" & m_TM02 & "' AND " & _
'                  "TM03 = '" & m_TM03 & "' AND " & _
'                  "TM04 = '" & m_TM04 & "' "
    'Modify By Cheng 2004/04/21
    '§ó·sµù¥U¤½§i¤é, ¤Îµù¥U¸¹
'   strSQL = "UPDATE TradeMark SET TM17='Y', TM20 = " & DBDATE(m_CP05) & ", " & _
'                                 "TM21 = " & DBDATE(textTM21) & ", " & _
'                                 "TM22 = " & DBDATE(textTM22) & " " & _
'            "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                  "TM02 = '" & m_TM02 & "' AND " & _
'                  "TM03 = '" & m_TM03 & "' AND " & _
'                  "TM04 = '" & m_TM04 & "' "
   '2008/10/24 modify by sonia µù¥U¤À³Î¤l®×¦P®É±N¥À®×¥Ó½Ð®×¸¹§ó·s¦Ü¤l®×,TM13¼f©w¨Ó¨ç¤é¤W¨Ó¨ç¦¬¤å¤é,TM16­ã»éÄæ¤W­ã,T-137268
   'strSQL = "UPDATE TradeMark SET TM14=" & DBDATE(Me.textTM14.Text) & ", TM15='" & Me.textTM15.Text & "', TM17='Y', TM20 = " & DBDATE(m_CP05) & ", " & _
                                 "TM21 = " & DBDATE(textTM21) & ", " & _
                                 "TM22 = " & DBDATE(textTM22) & " " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
   If m_CP10 = "308" Then
      '2011/9/22 modify by sonia ¥[¤£ºÞ¨î²Ä¤G´Á³Æµù
      strSql = "UPDATE TradeMark SET TM14=" & DBDATE(Me.textTM14.Text) & ", TM15='" & Me.textTM15.Text & "', TM16='1', TM17='Y', TM20 = " & DBDATE(m_CP05) & ", " & _
                                    "TM12 = '" & textTM12 & "', TM13 = " & DBNullDate(m_CP05) & ", " & _
                                    "TM21 = " & DBDATE(textTM21) & ", " & _
                                    "TM22 = " & DBDATE(textTM22) & ", " & _
                                    "TM58 = " & IIf(m_blnReceiveSecond, "decode(tm58,null,'¤£ºÞ¨î²Ä¤G´Á;','¤£ºÞ¨î²Ä¤G´Á;'||tm58) ", "tm58") & " " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "' "
   Else
      strSql = "UPDATE TradeMark SET TM14=" & DBDATE(Me.textTM14.Text) & ", TM15='" & Me.textTM15.Text & "', TM17='Y', TM20 = " & DBDATE(m_CP05) & ", " & _
                                    "TM21 = " & DBDATE(textTM21) & ", " & _
                                    "TM22 = " & DBDATE(textTM22) & " " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "' "
   End If
   '2008/10/24 END
   'End
   cnnConnection.Execute strSql
   
   'add by nickc 2006/08/14
   If UpForm Is Nothing Then
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '  ·s¼W¸ê®Æ¨ì®×¥ó¶i«×ÀÉ
       ' ¦¬¤å¸¹
       strCP09 = Empty
       strCP09 = AutoNo("C", 6)
       NowCP09 = strCP09 'Added by Lydia 2020/03/09
       
       ' ®×¥ó©Ê½è¬°µù¥UÃÒ
       strCP10 = "1701"
       ' ·~°È°Ï§O 91.8.26 MODIFY BY SONIA
       'strCP12 = GetStaffDepartment(m_CP13)
       ' 91.10.2 MODIFY BY SONIA cp20¦snull¦]¬°­n½Ð´Ú
       'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32) " & _
       '         "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
       '                 "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
       '                 "'" & "N" & "','" & "N" & "'," & DBDATE(SystemDate()) & ",'" & "N" & "') "
        'Modify By Cheng 2003/04/07
        '´¼Åv¤H­û¦s³Ìªñ¦¬¤åAÃþ±µ¬¢°O¿ý³æªº´¼Åv¤H­û
        'Modify By Cheng 2003/09/05
    '   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP32) " & _
    '            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
    '                    "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
    '                    "'" & "N" & "'," & DBDATE(SystemDate()) & ",'" & "N" & "') "
        'Modify By Cheng 2003/10/08
        '©Ó¿ì¤H§ìFCTSales
    '   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP32) " & _
    '            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
    '                    "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
    '                    "'" & "N" & "'," & DBDATE(SystemDate()) & ",'" & "N" & "') "
    'edit by nick 2004/09/24 92.11.28 ¥H«á¤§¤£½Ð´Ú
    '   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP32) " & _
                "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                        "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                        "'" & "N" & "'," & DBDATE(SystemDate()) & ",'" & "N" & "') "
        '2009/9/23 modify by sonia cp14§ï¬°¾Þ§@¤H­û
'        If DBDATE(Val(m_TM11)) >= 20031128 Then
           strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP32,cp20) " & _
                "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                        "'" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                        "'" & "N" & "'," & DBDATE(SystemDate()) & ",'" & "N" & "','N') "
'        Else
'           strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP32) " & _
'                "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                        "'" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
'                        "'" & "N" & "'," & DBDATE(SystemDate()) & ",'" & "N" & "') "
'        End If
        
       cnnConnection.Execute strSql
        'add by nickc 2007/03/06 ¥Ó½Ð°ê®a¬O¥xÆW®É¡A±N715©Î717µo¤åªº¡A¤Wcp24='1'¡Acp25=¨Ó¨ç¦¬¤å¤é¡A¨Ã±N npªº 305 np06¤W Y
        'modify by sonia 2022/10/6 +301ÅÜ§ó,302§ó¥¿ BY ªü½¬
        If m_TM10 = "000" Then
            'modify by sonia 2022/10/6 +301ÅÜ§ó,302§ó¥¿ BY ªü½¬
            strSql = "update caseprogress set cp24='1' ,cp25=" & DBDATE(m_CP05) & " where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10 in ('715','717','301','302') and cp27 is not null "
            cnnConnection.Execute strSql
            'modify by sonia 2022/10/6 +301ÅÜ§ó,302§ó¥¿ BY ªü½¬¡A¦P®É§ó·sNP15
            'strSql = "update nextprogress set np06='Y' where np06 is null and np07=305 and np01 in (select cp09 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10 in ('715','717') and cp27 is not null ) "
            strSql = "update nextprogress set np06='Y',np15='¦]µoµù¥UÃÒ¤WÄò¿ìY;'||NP15 where np06 is null and np07=305 and np01 in (select cp09 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10 in ('715','717','301','302') and cp27 is not null ) "
            cnnConnection.Execute strSql
            'Add By Sindy 2013/8/5
            '¤º°ÓªºT¥xÆW®×¤Î¥~°ÓFCT, ¦sÀÉ®É­Y¸Ó®×¸¹ªº¤U¤@µ{§ÇÀÉ¦³NP06 IS NULLªº 717(µù¥U¶O)´Á­­®É, ½Ð¤@¨Ö§ó·s.
            If m_TM01 = "FCT" Then
               strSql = "update nextprogress set np06='N',np11=" & strSrvDate(1) & ",NP12='10' " & _
                         "where np06 is null and np07='717' " & _
                           "and NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "'"
               cnnConnection.Execute strSql
            End If
            '2013/8/5 END
        End If
    End If
    'Add By Cheng 2003/09/03
    '·s¼W¤º³¡¦¬¤å
    If Me.Text2.Text <> "" Then
        'Modify By Cheng 2003/09/05
'        strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP32,CP64) " & _
'                        "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & _
'                        "'" & AutoNo("B", 6) & "','302','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
'                        "'N'," & strSrvDate(1) & ",'N','§ó§ïµù¥UÃÒ') "
        'Modify By Cheng 2003/10/08
        '©Ó¿ì¤H§ìFCTSales
'        strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP32, CP43, CP64) " & _
'                        "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & _
'                        "'" & AutoNo("B", 6) & "','302','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
'                        "'N'," & strSrvDate(1) & ",'N','" & strCP09 & "', '§ó§ïµù¥UÃÒ') "
        '2009/3/13 modify by sonia ¨ú®øµo¤å¤é, ¦]¬°°t¦Xµo¤å«Ç¹q¸£¤ÆÀ³©óªü½¬§Pµo®É¤~¤Wµo¤å¤é
        'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP27,CP32, CP43, CP64,CP20) " & _
                        "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & _
                        "'" & AutoNo("B", 6) & "','302','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                        "'N'," & strSrvDate(1) & ",'N','" & strCP09 & "', '§ó§ïµù¥UÃÒ','N') "
        '2009/9/23 modify by sonia CP14§ï¬°¾Þ§@¤H­û
        '2017/1/11 modify by sonia CP26§ï¬°­n­p¥ó
        'Modify by Amy 2023/02/06 +CP118 ¬O§_¹q¤l°e¥ó
        strCP118 = IIf(Pub_GetField("TradeMark", "tm01||tm02||tm03||tm04='" & m_TM01 & m_TM02 & m_TM03 & m_TM04 & "'", "TM136") = "1", "Y", "")
        strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP26,CP32, CP43, CP64,CP20,CP118) " & _
                        "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & _
                        "'" & AutoNo("B", 6) & "','302','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                        "'','N','" & strCP09 & "', '§ó§ïµù¥UÃÒ','N'," & CNULL(ChgSQL(strCP118)) & " ) "
        cnnConnection.Execute strSql
    End If
    'add by nickc 2006/08/14
    If UpForm Is Nothing Then
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' ·s¼W¸ê®Æ¨ì¤U¤@µ{§ÇÀÉ
       ' ¤U¤@µ{§Ç¬°©µ®i
       strNP07 = "102"
       ' §Ç¸¹
       strNP22 = GetNextProgressNo()
       ' ªk©w´Á­­¬°±M¥Î´Á­­¤î¤é
       strNP09 = DBDATE(textTM22)
       ' ¥»©Ò´Á­­¬°ªk©w´Á­­-2¤Ñ
        'Modify By Cheng 2003/09/02
    '   strNP08 = DBDATE(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)), Val(DBDAY(strNP09)) - 2))
       'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
       If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
          strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
       Else
       '2014/10/6 END
          strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
       End If
       strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ

       ' ²Õ¦¨SQL»yªk
       '91.12.12 modify by sonia
       'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
       '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
       '                  strNP08 & "," & strNP09 & ",'" & m_CP13 & "'," & strNP22 & ")"
        'Modify By Cheng 2003/04/07
        '´¼Åv¤H­û¦s³Ìªñ¦¬¤åAÃþ±µ¬¢°O¿ý³æªº´¼Åv¤H­û
        'Modify By Cheng 2003/09/05
    '   strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
    '            "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
    '                     strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
       strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                         strNP08 & "," & strNP09 & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
       '91.12.12 end
       cnnConnection.Execute strSql
       ' ©µ®i, ¨Ï¥Î«Å»}, ¥Zµn¼s§i, Ãº¦~¶O, ¶Ê¼f, ´£¥Ó, ¦¬¹F¤£¦L±µ¬¢µ²®×³æ
       Select Case strNP07
          Case "102", "105", "702", "708", "305", "998", "997":
          Case Else:
            'Modify By Cheng 2002/12/05
            '«ì´_¦C¦L±µ¬¢µ²®×³æ
    '            'Modify By Cheng 2002/01/15
    '            '¨ú®ø¥~°ÓFCT¦C¦L±µ¬¢µ²®×³æ
             ' ¦C¦L°ê¤º®×¥ó±µ¬¢¤Îµ²®×°O¿ý³æ
    '         g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
                'Modify By Cheng 2003/06/26
                '¨ú®ø¦C¦L±µ¬¢µ²®×³æ
    '            'Add By Cheng 2003/06/23
    '            '·s¼W¦C¦L±µ¬¢µ²®×³æ¸ê®Æ
    '            pub_AddressListSN = pub_AddressListSN + 1
    '            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
       End Select
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       '93.6.11 ADD BY SONIA ±¾²Ä¤G´Áµù¥U¶O´Á­­
       'ADD BY NICK 2004/08/17
       Is716Have = True
       
       'edit by nick  2004/12/21 ¥[¥Ó½Ð¤é¦b 92/11/28 «e¡A¥B¤½§i¤é¦b 92/9/1(§t)«á¡A­Y np ¨S¦³ 716 ´N·s¼W
       'If DBDATE(textTM21) > 20031128 Then
       If (DBDATE(textTM21) >= 20031128) Or (DBDATE(m_TM11) <= 20030901 And DBDATE(textTM21) < 20031128 And Trim(textTM14) <> "") Then
         'Add By Sindy 2012/12/19 101¦~7¤ë°Ó¼Ð·s­×ªk¼o°£¤G´Áµù¥U¶OÃº¶O¨î«× +if
         If Val(m_TM13) < 20120701 Then
            'add by nick 2004/08/17
            '¥ýÀË¬d¬O§_¦³ 717
            StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='717' and cp05 is not null and cp57 is null "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
            Else
               Set rsA = New ADODB.Recordset
               'ªk©w´Á­­
               strCP07 = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 3, ChangeWStringToWDateString(DBDATE(textTM21)))))
               '¥»©Ò´Á­­
               'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
               If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                  strCP06 = PUB_GetOurDeadline(DBDATE(strCP07))
               Else
               '2014/10/6 END
                  strCP06 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strCP07))))
               End If
               strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
               
               StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='716' "
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               '­Y¦³¦¬¤å²Ä¤G´Áµù¥U¶O, §ó·s¶i«×ÀÉ
               If rsA.RecordCount > 0 Then
                   StrSQLa = "Update CaseProgress Set CP06=" & strCP06 & ", CP07=" & strCP07 & " Where CP09='" & rsA("CP09").Value & "' "
                   cnnConnection.Execute StrSQLa
               '­Y¥¼¦¬¤å²Ä¤G´Áµù¥U¶O, ·s¼W¤U¤@µ{§ÇÀÉ
               Else
                    Is716Have = False
                    'add by nick 2004/08/17
                    ' ÀË¬d¤U¤@µ{§Ç¦³µL 716
                    Set rsA = New ADODB.Recordset
                    StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And np07=716 "
                    rsA.CursorLocation = adUseClient
                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                    If rsA.RecordCount > 0 Then
                        strSql = "update NextProgress set np08=" & DBDATE(strCP06) & ",np09=" & DBDATE(strCP07) & " where " & ChgNextProgress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And np07=716 "
                        cnnConnection.Execute strSql
                    Else
                       If m_blnReceiveSecond = False Then '2011/9/22 add by sonia­Y®×¥ó³Æµù¤£ºÞ¨î«h¤£·s¼W
                          strNP07 = "716"
                          strNP22 = GetNextProgressNo()
                          strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                                          "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                                          DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
                          cnnConnection.Execute strSql
                       End If  '2011/9/22 end
                    End If
               End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         End If '2012/12/19 End
       End If
   End If
   '93.6.11 END
   '911107 nick ²¾¨ì¤U­±
   ' ¦C¦L©w½Z
   'If textPrint <> "N" Then
   '   PrintLetter
   'End If
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Modify By Cheng 2003/02/18
   '­Y³]©w²£¥Í½Ð´Ú¸ê®Æ
   'If Me.textCreFee.Visible And Me.Combo2.Visible Then
   If Me.textCreFee.Text = "Y" Then
       'Modify By Cheng 2003/02/27
       '³]©w­n¦C¦L¦a§}±ø
   '    'Add By Cheng 2003/02/17
   '    '·s¼W¦a§}±ø¦Cªí¸ê®Æ
   '    pub_AddressListSN = pub_AddressListSN + 1
   '    PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
       'edit by nick 2004/11/24
       'm_blnPrintAddress = True
      '·s¼W°ê¥~½Ð´Ú¸ê®Æ
      Dim strAgentNo As String '¥N²z¤H½s¸¹
      Dim strPrintCust  As String '¬O§_¦C¦L¥Ó½Ð¤H
      Dim dblUSRate As Double '¬üª÷¶×²v
       Dim strDisc As String '§é¦©
       Dim strA1K27 As String '¦C¦L¹ï¶H
       Dim strA1K28 As String '½Ð´Ú¹ï¶H
      
      '1:¥ý¥H"X"§ìACC1R0¤§°ê¥~½Ð´Ú³æªº¦Û°Ê½s¸¹, ¨Ã§ó·s¨ä¬y¤ô¸¹
      m_strSerialNo = AccAutoNo(MsgText(815), 5)
      AccSaveAutoNo MsgText(815), Right(m_strSerialNo, 5)
      '2:·s¼WACC1K0
   '   strAgentNo = GetAgentNO
      strAgentNo = PUB_GetA1K03(m_TM01, m_TM02, m_TM03, m_TM04)
      strPrintCust = PUB_GetA1K04(m_TM01, m_TM02, m_TM03, m_TM04)
     ' dblUSRate = GetUSRate
        
       strA1K27 = PUB_GetA1K27(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10)
       If strA1K27 = "" Then strA1K27 = strAgentNo
       strA1K28 = PUB_GetA1K28(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10)
       If strA1K28 = "" Then strA1K28 = strAgentNo
       
       'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
       Dim strA1K33 As String, strA1K18 As String
       'Modify By Sindy 2016/11/30
       'strA1K33 = PUB_GetInitCurrPrintType(m_TM01, strA1K28, strA1K18, dblUSRate)
       'Modified by Morgan 2018/4/27 +strA1K27
       strA1K33 = PUB_GetInitCurrPrintType(m_TM01, strA1K28, strA1K18, dblUSRate, m_TM02, m_TM03, m_TM04, strA1K27)
       '2016/11/30 END
         
       strDisc = 1 - (PUB_GetA1L07Disc(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, strSrvDate(2)) / 100)
       'Modify By Cheng 2002/12/13
   '   strSQL = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K17,A1K18,A1K25,A1K26,A1K29,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04) " & _
   '            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,0,0," & dblUSRate & ",3500,0,'" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','','USD','','','',0," & IIf(dblUSRate = 0, 0, Format(3500 / dblUSRate, "##0.00")) & ",'" & strAgentNo & "','" & strAgentNo & "','" & strAgentNo & "','" & strPrintCust & "' )"
       'Modify By Cheng 2002/12/24
       '§éÅý¤é´Á¦sNULL, §@¼o¤é´Á¦sNULL
   '   strSQL = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K17,A1K18,A1K19,A1K20,A1K21,A1K25,A1K26,A1K29,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04) " & _
   '            "VALUES  ('" & m_strSerialNo & "'," & (ServerDate - 19110000) & ",0,0,0," & dblUSRate & ",3500,0,'" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','','USD'," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "','','','',0," & IIf(dblUSRate = 0, 0, Format(3500 / dblUSRate, "##0.00")) & ",'" & strAgentNo & "','" & strAgentNo & "','" & strAgentNo & "','" & strPrintCust & "' )"
       'Modify By Cheng 2004/01/07
       'A1K11­n¥ý¦©°£§é¦©«á¤~¦sÀÉ
   '   strSQL = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K17,A1K18,A1K19,A1K20,A1K21,A1K25,A1K26,A1K29,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04) " & _
   '            "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",0,NULL,0," & dblUSRate & ",3500,NULL,'" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','','USD'," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "','','','',0," & IIf(dblUSRate = 0, 0, Format(3500 / dblUSRate, "##0.00")) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "' )"
       'Modify By Cheng 2004/04/26
       '¬üª÷¨ú¦Ü¾ã¼Æ¦ì(µL±ø¥ó±Ë¥h)
   '   strSQL = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K17,A1K18,A1K19,A1K20,A1K21,A1K25,A1K26,A1K29,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04) " & _
   '            "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",0,NULL,0," & dblUSRate & "," & 3500 - (3000 * Val(strDisc)) & ",NULL,'" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','','USD'," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "','','','',0," & IIf(dblUSRate = 0, 3500 - (3000 * Val(strDisc)), Format((3500 - (3000 * Val(strDisc))) / dblUSRate, "##0.00")) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "' )"
     'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
'      strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K17,A1K18,A1K19,A1K20,A1K21,A1K25,A1K26,A1K29,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04) " & _
               "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",0,NULL,0," & dblUSRate & "," & 3500 - (3000 * Val(strDisc)) & ",NULL,'" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','','USD'," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "','','','',0," & Fix(Val("" & IIf(dblUSRate = 0, 3500 - (3000 * Val(strDisc)), (3500 - (3000 * Val(strDisc))) / dblUSRate))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "' )"
       strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K17,A1K18,A1K19,A1K20,A1K21,A1K25,A1K26,A1K29,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K33) " & _
               "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",0,NULL,0," & dblUSRate & "," & 3500 - (3000 * Val(strDisc)) & ",NULL,'" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','','" & strA1K18 & "'," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "','','','',0," & Fix(Val("" & IIf(dblUSRate = 0, 3500 - (3000 * Val(strDisc)), (3500 - (3000 * Val(strDisc))) / dblUSRate))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "','" & strA1K33 & "' )"
      
       'End
      cnnConnection.Execute strSql
      '3:·s¼W¨âµ§ACC1L0
   '    strDisc = 1 - (PUB_GetA1L07Disc(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, strSrvDate(2)) / 100)
       'Modify By Cheng 2002/12/13
   '   strSQL = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05) " & _
   '            "VALUES  ('" & m_strSerialNo & "','FCT','',0,'001','1701',3000 )"
      strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
               "VALUES  ('" & m_strSerialNo & "','FCT','' ," & 3000 * Val(strDisc) & ", '001', '1701', 3000, " & strSrvDate(2) & ", " & ServerTime & ", '" & strUserNum & "' )"
      cnnConnection.Execute strSql
       'Modify By Cheng 2002/12/13
   '   strSQL = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05) " & _
   '            "VALUES  ('" & m_strSerialNo & "','FCT','',0,'002','02',500 )"
      strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
               "VALUES  ('" & m_strSerialNo & "','FCT','',0 ,'002','02',500," & Val(ACDate(ServerDate)) & "," & ServerTime & ",'" & strUserNum & "' )"
      cnnConnection.Execute strSql
      
      PUB_UpdateA1k08 m_strSerialNo 'Added by Morgan 2012/11/2 §ó·s½Ð´Ú³æ¥~¹ôª÷ÃB
      
      '4:·s¼WACC1W0
      strSql = "INSERT INTO ACC1W0 (A1W01,A1W02) " & _
               "VALUES  ('" & m_strSerialNo & "','" & strCP09 & "')"
      cnnConnection.Execute strSql
      '5:§ó·s·s¼WªºCÃþ¦¬¤å¸¹
      strSql = "UPDATE CASEPROGRESS SET CP60='" & m_strSerialNo & "' WHERE CP09='" & strCP09 & "'"
      cnnConnection.Execute strSql
       'Moved By Cheng 2004/05/12
   '   '6:¦C¦L·s¼Wªº½Ð´Ú¸ê®Æ
   '   ProcessPrint
       'End
       PUB_PointAutoassign m_strSerialNo, True 'Add by Morgan 2010/4/21 ¦Û°Ê¤À°tÂI¼Æ
   End If

    Dim m_MonTM11 As String
    Dim m_MonTM14 As String
    Dim m_MonTM21 As String
    'add by nickc 2006/08/14
    If m_CP10 = "308" Then
      '·s¼W¤l®×®Ö­ã¨Ó¤å
      strCP09 = AutoNo("C", 6)
      strCP05 = DBDATE(UpForm.oStrCDate)
      strCP27 = DBDATE(SystemDate())
      ' ²Õ¦¨SQL»yªk
      strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26,cp27,   CP43) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "1001" & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "','" & "N" & "'," & strCP27 & ",'" & m_CP09 & "')"
      ' ·s¼W¸ê®Æ¨ì¸ê®Æ®w
      cnnConnection.Execute strSql
      
      'Added by Morgan 2017/6/14 ¹q¤l¤½¤å
      If m_DocNo <> "" Then
         '§ó·s¾÷Ãö¤å¸¹
         strSql = "update caseprogress set cp08='" & m_DocWord & "¦r²Ä" & PUB_GetEDocNo(m_DocNo) & "¸¹' where cp09='" & strCP09 & "'"
         cnnConnection.Execute strSql, intI
         '½Æ»s¥À®×¤½¤å¹q¤lÀÉ
         strExc(0) = PUB_GetEDocFileName(m_TM01, m_TM02, m_TM03, m_TM04, "1001")
         SaveAttFile_PDF strCP09, m_DocPdf, strExc(0), Format(m_DocPdfDate), Format(m_DocPdfTime), False, , , True
      End If
      'end 2017/6/14
      
      '§ó·s¤l®×®Ö­ã¤Îµ²ªG¤é
      strSql = "update caseprogress set cp24='1',cp25=" & strCP05 & " where cp09='" & m_CP09 & "' "
      cnnConnection.Execute strSql
      '2011/9/22 ADD BY SONIA ¥À®×¤Î¤l®×ªº¶Ê¼f´Á­­¤WY
      strSql = "update nextprogress set np06='Y' where np01='" & m_CP09 & "' and np07='305' and np06 is null"
      cnnConnection.Execute strSql
      strSql = "update nextprogress set np06='Y' where np02='" & m_MonTM01 & "' and np03='" & m_MonTM02 & "' and np04='" & m_MonTM03 & "' and np05='" & m_MonTM04 & "' and np01='" & frm02010401_6.oKey & "' and np07='305' and np06 is null"
      cnnConnection.Execute strSql
      '¦P®É¤l®×ºÞ¨î©µ®i´Á­­
      strNP07 = "102"
      If IsEmptyText(textTM22) = False Then: strNP09 = textTM22
      'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
      If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
         strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
      Else
      '2014/10/6 END
         strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
      End If
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ

      If rsA.State <> adStateClosed Then rsA.Close
      StrSQLa = "select * from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='102' and cp27 is null and cp57 is null "
      Set rsA = New ADODB.Recordset
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount <> 0 Then
         strSql = "update caseprogress set cp06=" & strNP08 & ",cp07=" & strNP09 & " where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='102' and cp27 is null and cp57 is null "
      Else
         If rsA.State <> adStateClosed Then rsA.Close
         StrSQLa = "select * from nextprogress where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07='102' and np06 is null "
         Set rsA = New ADODB.Recordset
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount <> 0 Then
            strSql = "update nextprogress set np08=" & strNP08 & ",np09=" & strNP09 & " where np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "' and np07='102' and np06 is null "
         Else
            strNP22 = GetNextProgressNo()
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
                             "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         End If
      End If
      cnnConnection.Execute strSql
      '2011/9/22 END
      
'2011/9/22 modify by sonia «e¤w§ì¥À®×¬O§_ºÞ¨î²Ä¤G´Á,¬G§ï¥Hm_blnReceiveSecond§PÂ_
'      '¥À®×¦³¦¬ 717 ®É¡A¤£ºÞ¡A­Y¦³ 716 ªº¤]¤£ºÞ¡A¥u¦³ 715 ªº ¤l®×­n±¾²Ä¤G´Áµù¥U¶O ¡A¦ý¶È­­´Á°_¤é+3¦~-1¤Ñ ¤j©ó ¨t²Î¤éªº¤~°µ
'      If rsA.State <> adStateClosed Then rsA.Close
'      m_MonTM11 = ""
'      m_MonTM14 = ""
'      m_MonTM21 = ""
'      StrSQLa = "select * from trademark where tm01='" & m_MonTM01 & "' and tm02='" & m_MonTM02 & "' and tm03='" & m_MonTM03 & "' and tm04='" & m_MonTM04 & "' "
'      Set rsA = New ADODB.Recordset
'      rsA.CursorLocation = adUseClient
'      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'          m_MonTM11 = CheckStr(rsA.Fields("tm11"))
'          m_MonTM14 = CheckStr(rsA.Fields("tm14"))
'          m_MonTM21 = CheckStr(rsA.Fields("tm21"))
'      End If
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      If (m_MonTM21 >= 20031128) Or (m_MonTM11 < 20031128 And m_MonTM14 >= 20030901 And m_MonTM14 <> "") Then
'        If ChangeWDateStringToWString(DateAdd("d", -1, DateAdd("yyyy", 3, ChangeWStringToWDateString(m_MonTM21)))) <= strSrvDate(1) Then
'            If rsA.State <> adStateClosed Then rsA.Close
'            StrSQLa = "select * from caseprogress where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10 in ('716','717') "
'            Set rsA = New ADODB.Recordset
'            rsA.CursorLocation = adUseClient
'            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount = 0 Then
'                If rsA.State <> adStateClosed Then rsA.Close
'                StrSQLa = "select * from caseprogress where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='715' "
'                Set rsA = New ADODB.Recordset
'                rsA.CursorLocation = adUseClient
'                rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'                If rsA.RecordCount <> 0 Then
                'Modify By Sindy 2012/12/19 101¦~7¤ë°Ó¼Ð·s­×ªk¼o°£¤G´Áµù¥U¶OÃº¶O¨î«× +And Val(m_TM13) < 20120701
                If m_blnReceiveSecond = False And Val(m_TM13) < 20120701 Then
                    '­n±¾²Ä¤G´Áªº´Á­­µ¹¤l®×
                    'ªk©w´Á­­
                    strCP07 = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 3, ChangeWStringToWDateString(m_MonTM21))))
                    '¥»©Ò´Á­­
                    'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
                    If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                       strCP06 = PUB_GetOurDeadline(DBDATE(strCP07))
                    Else
                    '2014/10/6 END
                       strCP06 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strCP07))))
                    End If
                    strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
                    
                    strNP07 = "716"
                    strNP22 = GetNextProgressNo()
                    strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                                    "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                                    DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
                    cnnConnection.Execute strSql
                End If
'                If rsA.State <> adStateClosed Then rsA.Close
                'add by nickc 2007/03/06 ¥Ó½Ð°ê®a¬O¥xÆW®É¡A±N715©Î717µo¤åªº¡A¤Wcp24='1'¡Acp25=¨Ó¨ç¦¬¤å¤é¡A¨Ã±N npªº 305 np06¤W Y
'            Else
'                If m_TM10 = "000" Then
                strSql = "update caseprogress set cp24='1' ,cp25=" & strCP05 & " where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10 in ('715','717')  and cp27 is not null "
                cnnConnection.Execute strSql
                strSql = "update nextprogress set np06='Y' where np06 is null and np07=305 and np01 in (select cp09 from caseprogress where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10 in ('715','717')  and cp27 is not null ) "
                cnnConnection.Execute strSql
'                End If
'            End If
'        End If
'      End If
      
      '¦³´Á­­®É
      If textNP08.Enabled = True And textNP09.Enabled = True Then
             '­Yµe­±¦³¿é¤J·s´Á­­¥H·s´Á­­¬°¥D¡A¨S¦³ªº¸Ü±NÄ~©Ó¥À®×´Á­­
             If Trim(textNP08) <> "" And Trim(textNP09) <> "" Then
                If UpForm.IsHaveNp202 Then
                      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                          "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',202," & _
                          DBDATE(textNP08) & "," & DBDATE(textNP09) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo & ")"
                      cnnConnection.Execute strSql
                ElseIf UpForm.IsHaveCp202 Then
                     If Trim(textNP08) <> "" Then
                         strSql = "update caseprogress set cp06=" & DBDATE(textNP08) & ",cp07=" & DBDATE(textNP09) & ",cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     Else
                         strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     End If
                     cnnConnection.Execute strSql
                     strSql = "update caseprogress set cp43='" & m_CP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     cnnConnection.Execute strSql
                     strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     cnnConnection.Execute strSql
                End If
             Else
                If UpForm.IsHaveNp202 Then
                      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                          "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',202," & _
                          m_MonNP08 & "," & m_MonNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo & ")"
                      cnnConnection.Execute strSql
                ElseIf UpForm.IsHaveCp202 Then
                     If Trim(textNP08) <> "" Then
                         strSql = "update caseprogress set cp06=" & DBDATE(textNP08) & ",cp07=" & DBDATE(textNP09) & ",cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     Else
                         strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     End If
                     cnnConnection.Execute strSql
                     strSql = "update caseprogress set cp43='" & m_CP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     cnnConnection.Execute strSql
                     strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                     cnnConnection.Execute strSql
                End If
             End If
             If UpForm.IsHaveNp202 Then
                  strSql = "update nextprogress set np06='N',np15=np15||'Âà¤J¤l®×¡A¤l®×®×¸¹¡G" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' where np02='" & m_MonTM01 & "' and np03='" & m_MonTM02 & "' and np04='" & m_MonTM03 & "' and np05='" & m_MonTM04 & "' and np06 is null and np07=202 "
                  cnnConnection.Execute strSql
             ElseIf UpForm.IsHaveCp202 Then
                  strSql = "update caseprogress set cp57=to_number(to_char(sysdate,'YYYYMMDD')),cp64=cp64||'Âà¤J¤l®×¡A¤l®×®×¸¹¡G" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202' and cp27 is null "
                  cnnConnection.Execute strSql
             End If
             '¥À®×¤À³Îµo¤å«áªº¦¬¤å¤Îµo¤å®×¥ó¬ÒÂà¤J¦³´Á­­ªº¤l®×
             Dim m_MonCP27 As String
             strSql = "select cp27 from caseprogress where cp09='" & m_MonCP09 & "' "
             m_MonCP27 = ""
             Set rsTmp = New ADODB.Recordset
             If rsTmp.State = 1 Then rsTmp.Close
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
             If rsTmp.RecordCount > 0 Then
                 m_MonCP27 = CheckStr(rsTmp.Fields("cp27"))
             End If
             If m_MonCP27 <> "" Then
                 strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & "' where cp05>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                 cnnConnection.Execute strSql
                 strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & "' where cp27>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                 cnnConnection.Execute strSql
                 strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp05>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                 cnnConnection.Execute strSql
                 strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                 cnnConnection.Execute strSql
             End If
      End If
      '2008/10/24 ADD BY SONIA ¤À³Î¥À®×³¬¨÷
      Set rsA = New ADODB.Recordset
      If rsA.State = 1 Then rsA.Close
      strSql = "select * from divisioncase,trademark where dc05='" & m_MonTM01 & "' and dc06='" & m_MonTM02 & "' and dc07='" & m_MonTM03 & "' and dc08='" & m_MonTM04 & "' and dc01=tm01(+) and dc02=tm02(+) and dc03=tm03(+) and dc04=tm04(+) and (tm16 is null or tm16='') "
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount = 0 Then
         strSql = "update trademark set tm29='Y',tm30=to_number(to_char(sysdate,'YYYYMMDD')),tm31='87' where tm01='" & m_MonTM01 & "' and tm02='" & m_MonTM02 & "' and tm03='" & m_MonTM03 & "' and tm04='" & m_MonTM04 & "' and (tm29 is null or tm29='') "
         cnnConnection.Execute strSql
      End If
      If rsA.State = 1 Then rsA.Close
      '2008/10/24 END
      
    'Added by Morgan 2023/1/17 ¹q¤l¤½¤å
    ElseIf m_DocNo <> "" Then
       PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
    'end 2023/1/17
    End If

'add by nickc 2006/08/14
If Me.Visible = True Then
 '911107 nick transation
  cnnConnection.CommitTrans
End If
   'add by nickc 2006/08/14
   If UpForm Is Nothing Or Me.Visible = False Then
        If Me.textCreFee.Text = "Y" Then
            '6:¦C¦L·s¼Wªº½Ð´Ú¸ê®Æ
            ProcessPrint
            'Added by Lydia 2016/11/17 ¥H½Ð´Ú¹ï¶HÀË¬d¬O§_¦s¦b©ó°ê¥~©T©w±H¶Ê´Ú³æ¥N²z¤HÀÉ(ACC225)¥B¤U¦¸±Hµo¤é´Á¡Ö¨t²Î¤é¡A­Y¦s¦b«hÅã¥Ü°T®§´£¿ô¾Þ§@¤H­û
            If m_strSerialNo <> "" And strA1K28 <> "" Then
               If PUB_ChkAcc225MsgList(m_strSerialNo, strA1K28, m_TM01, m_TM02, m_TM03, m_TM04) Then
               End If
            End If
            'end 2016/11/17
        End If
       ' ¦C¦L©w½Z
       If textPrint <> "N" Then
          'Modified by Lydia 2023/02/23 ³qª¾¨ç¤ÎÄ¶¤å¡BÃÒ®ÑPDF¡A¦P®É¦s¦ÜFCT_WORKFLOW\(¬Û¹ïÀ³®×¸¹ªº¸ê®Æ§¨)
          'PrintLetter
          PrintLetterNew
          m_blnPrintAddress = True
          'add by nick 2004/09/24
    '    '·s¼W¦a§}±ø¦Cªí¸ê®Æ
        'Modify By Sindy 2025/10/2 ¨ú®ø¦a§}±ø
'        pub_AddressListSN = pub_AddressListSN + 1
'        PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
       End If
    End If
     Exit Function
CheckingErr:
    'add by nickc 2006/08/14
    If Me.Visible = True Then
        cnnConnection.RollbackTrans
        MsgBox (Err.Description)
    End If
    'edit by nick 2004/11/03
    OnSaveData = False
End Function

'Add By Cheng 2002/06/06
Private Sub ProcessPrint()
Screen.MousePointer = vbHourglass
'Modify By Cheng 2003/01/16
'¦C¦L½Ð´Ú³æ®É¨Ï¥Î¦@¦Pªºªí³æ(Frmacc2480)
'For Each prnPrint In Printers
'   If prnPrint.DeviceName = Combo2 Then
'      Set Printer = prnPrint
'   End If
'Next
'PrintData
'For Each prnPrint In Printers
'   If prnPrint.DeviceName = strPrint Then
'      Set Printer = prnPrint
'   End If
'Next
Load Frmacc2480: DoEvents
Frmacc2480.Text1.Text = m_strSerialNo
Frmacc2480.Text2.Text = m_strSerialNo
Frmacc2480.Combo1.Text = Me.Combo2.Text
Frmacc2480.Command2_Click: DoEvents
Unload Frmacc2480
Screen.MousePointer = vbDefault
End Sub

'Add By Cheng 2002/06/06
Private Function GetUSRate() As Double
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetUSRate = 0
'Modify By Cheng 2002/12/13
'À³¥H¥Á°ê¦~§ì³Ì±µªñ¨t²Î¤éªº¸ê®Æ
'strSQLA = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & ServerDate & " AND ROWNUM = 1 ORDER BY USXR01 "
StrSQLa = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & (ServerDate - 19110000) & " ORDER BY USXR01 DESC "
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    'Modify By Cheng 2002/12/13
'   GetUSRate = rsA.Fields(0).Value
   GetUSRate = CDbl(rsA.Fields(0).Value)
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub Form_Unload(Cancel As Integer)
    '­Y¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
    If Me.Combo2.Text <> Me.Combo2.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
    End If
    'Add By Cheng 2002/07/19
    Set frm03020404_03 = Nothing
End Sub

Private Sub Text1_GotFocus()
    'Add By Cheng 2003/01/28
    TextInverse Me.Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
    
    Cancel = False
    'Add By Cheng 2003/01/28
    '­Y¦³¿é¤JÃÒ®Ñ¤é´Á
    If IsEmptyText(Text1) = False Then
       ' ÀË¬d¤é´Á®æ¦¡
       'edit by nickc 2006/09/08
       'If CheckIsTaiwanDate(Text1, False) = False Then
       If CheckIsDate(Text1, False) = False Then
          Cancel = True
          strTit = "¸ê®ÆÀË®Ö"
          strMsg = "ÃÒ®Ñ¤é´Á®æ¦¡¿é¤J¿ù»~"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          Text1_GotFocus
       End If
     End If
End Sub

Private Sub Text2_GotFocus()
    TextInverse Me.Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 89 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    '93.7.7 add by sonia §ó§ïÃÒ®Ñ¤£¦L©w½Z¤£¦L¦a§}±ø
    'edit by nick 2004/09/24 ¤é¤å°£¥~
    'If KeyAscii = 89 Then
    If KeyAscii = 89 And GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) <> "3" Then
       textPrint = "N"
    End If
    '93.7.7 end
End Sub

Private Sub textCreFee_Change()
    'Marked By Cheng 2004/05/11
    '¨ú®ø«ü©w½Ð´Ú³æ¦Lªí¾÷, ¥Î¦C¦Lµe­±¤Wªº½Ð´Ú³æ¦Lªí¾÷
'    'Add By Cheng 2002/12/13
'    If Me.textCreFee.Text = "Y" Then
'        Label18.Visible = True
'        Me.Combo2.Visible = True
'    Else
'        Label18.Visible = False
'        Me.Combo2.Visible = False
'    End If
    'End
End Sub

Private Sub textCreFee_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    'Add By Cheng 2003/09/23
    'Begin
    If KeyAscii <> 8 And KeyAscii <> 89 Then
        KeyAscii = 0
    End If
    'End
End Sub

' ¬O§_²£¥Í½Ð´Ú¸ê®Æ
Private Sub textCreFee_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
    'Modify By Cheng 2003/09/23
    'Begin
'   If IsEmptyText(textCreFee) = False Then
'      Select Case textCreFee
'         Case " ", "Y":
'         Case Else:
'            Cancel = True
'            strTit = "¸ê®ÆÀË®Ö"
'            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎY"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textCreFee_GotFocus
'      End Select
'   End If
    'End
    'Marked By Cheng 2004/05/11
'   'Add By Cheng 2002/06/05
'   If Me.textCreFee.Text = "Y" Then
'      Label18.Visible = True
'      Me.Combo2.Visible = True
'   Else
'      Label18.Visible = False
'      Me.Combo2.Visible = False
'   End If
    'End
End Sub

Private Sub textNP08_GotFocus()
InverseTextBox textNP08
End Sub

Private Sub textNP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   
   Cancel = False
   If IsEmptyText(textNP08) = False Then
      If CheckIsTaiwanDate(textNP08, False) = False Then
         Cancel = True
         strMsg = "¤é´Á¤£¥¿½T"
         strTit = "¤l®×·s¥»©Ò´Á­­"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP08_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
      Else
          textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08, True), 1)
      'end 2020/07/07
      End If
   End If
EXITSUB:
End Sub

Private Sub textNP09_GotFocus()
    InverseTextBox textNP09
End Sub

Private Sub textNP09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   
   Cancel = False
   If IsEmptyText(textNP09) = False Then
      If CheckIsTaiwanDate(textNP09, False) = False Then
         Cancel = True
         strMsg = "¤é´Á¤£¥¿½T"
         strTit = "¤l®×·sªk©w´Á­­"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP09_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub
Private Sub textPrint_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    'Add By Cheng 2003/09/23
    'Begin
    If KeyAscii <> 8 And KeyAscii <> 78 Then
        KeyAscii = 0
    End If
    'End
End Sub

' ¦C¦L©w½Z
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
    'Modify By Cheng 2003/09/23
    'Begin
'   If IsEmptyText(textPrint) = False Then
'      Select Case textPrint
'         Case " ", "N":
'         Case Else:
'            Cancel = True
'            strTit = "¸ê®ÆÀË®Ö"
'            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎN"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textPrint_GotFocus
'      End Select
'   End If
    'End
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim Cancel As Boolean
   
   CheckDataValid = False
   'add by nickc 2007/03/06 ¥Ó½Ð¤é¦b 92/11/28 «eªº¡A­Y§Ñ°O½Ð´Ú¡A«h­n¸ß°Ý¤@¤U
   If DBDATE(Val(m_TM11)) < 20031128 And UCase(Trim(textCreFee)) <> "Y" And textCreFee.Locked = False Then
       If MsgBox("¦¹®×¥Ó½Ð¤é¦b 92/11/28 «e¡A½Ð°Ý¬O§_­n½Ð´Ú¡H", vbYesNo) = vbYes Then
           textCreFee = "Y"
       End If
   End If
   
   ' µù¥U¸¹¤Îµù¥U¤½§i¤é¤£¥iªÅ¥Õ
   If Me.textTM14.Text = "" Then
       strTit = "¸ê®ÆÀË®Ö"
       strMsg = "½Ð¿é¤Jµù¥U¤½§i¤é"
       nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       textTM14.SetFocus
       GoTo EXITSUB
   End If
   If Me.textTM15.Text = "" Then
       strTit = "¸ê®ÆÀË®Ö"
       strMsg = "½Ð¿é¤J¼f©w¸¹"
       nResponse = MsgBox(strMsg, vbOKOnly, strTit)
       textTM15.SetFocus
       GoTo EXITSUB
   End If
   ' ±M¥Î´Á­­°_¤é¤£¥iªÅ¥Õ
   If IsEmptyText(textTM21) = True Then
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "½Ð¿é¤J±M¥Î´Á­­°_¤é"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM21.SetFocus
      GoTo EXITSUB
   End If
   ' ±M¥Î´Á­­¤î¤é¤£¥iªÅ¥Õ
   If IsEmptyText(textTM22) = True Then
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "½Ð¿é¤J±M¥Î´Á­­¤î¤é"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM22.SetFocus
      GoTo EXITSUB
   End If
   
   'Add By Sindy 2012/5/18
   If LabNP07.Caption <> "" Then
      'ÀË¬d¨Ó¨ç´Á­­--¤é´Á
      If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
         If Me.Option4(2).Value = True Then
            If Me.Text12.Text = "" Then
               MsgBox "½Ð¿é¤J¨Ó¨ç´Á­­!!!", vbExclamation + vbOKOnly
               Me.Text12.SetFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
   
   'Add By Sindy 2012/7/9 ¥H¨¾­×§ï´Á­­¤Ñ¼Æ©Î¤ë¼Æ,­«·s­pºâ´Á­­
   If Me.Text10.Enabled = True Then
      Cancel = False
      Text10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text11.Enabled = True Then
      Cancel = False
      Text11_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2012/7/9 End
      
   CheckDataValid = True
EXITSUB:
End Function

' ¬O§_¦C¦LÂ½Ä¶¨ç
Private Sub textPrtTrans_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' ¬O§_¦C¦LÂ½Ä¶¨ç
Private Sub textPrtTrans_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrtTrans) = False Then
      Select Case textPrtTrans
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎN"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrtTrans_GotFocus
      End Select
   End If
End Sub

Private Sub textTM14_GotFocus()
    TextInverse Me.textTM14
End Sub

Private Sub textTM14_Validate(Cancel As Boolean)
   If IsEmptyText(textTM14) = False Then
      'edit by nick 2004/10/06
      'If CheckIsTaiwanDate(textTM14, False) = False Then
      If CheckIsDate(textTM14, False) = False Then
         Cancel = True
         'MsgBox "½Ð¿é¤J¥Á°ê¦~", vbOKOnly, "¸ê®ÆÀË®Ö"
         MsgBox "½Ð¿é¤J¦è¤¸¦~", vbOKOnly, "¸ê®ÆÀË®Ö"
         textTM14.SetFocus
         Exit Sub
      End If
      'Added by Lydia 2023/03/29 ¨ó§U±±ºÞ°w¹ï¥xÆWµù¥UÃÒ¿é¤J¡A¤½§i¤é´Á¥u¯à¿é¤J1¸¹©Î16¸¹
      If m_TM01 = "FCT" And m_TM10 = "000" And InStr("01,16,", Format(PUB_DBDAY(textTM14), "00")) = 0 Then
         Cancel = True
         MsgBox "¤½§i¤é´Á¥u¯à¿é¤J1¸¹©Î16¸¹", vbOKOnly, "¸ê®ÆÀË®Ö"
      End If
      'end 2023/03/29
      '2010/4/7 ADD BY SONIA
      'If Text1 = "" Then Text1 = textTM14
      'Modify By Sindy 2012/1/6 ªü½¬:µù¥UÃÒ¿é¤J¸ê®Æ¤¤¤§¡¨ÃÒ®Ñ¤é´Á¡¨«Y³]©w¦Û°Ê±a¤½§i¤é¡A¦ý¤½§i¤é¿é¤J¿ù»~­«·s¿é¤J®ÉÃÒ®Ñ¤é´Á¤£·|¸òµÛ§ó¥¿¡A½Ð­×§ï¡AÁÂÁÂ!
      If m_CP10 <> "308" Then Text1 = textTM14   '2013/3/19 MODIFY BY SONIA ¤À³Î®×¤£¥i±a¤½§i¤é,§_«h·|±a¨ì¥À®×ªº¤½§i¤é, FCT-034085
      '2010/4/7 END
      
      'Add By Sindy 2014/4/1 ¶ñ¤J¹w³]­È
      If IsEmptyText(textTM21) = True Then textTM21 = GetTM2122Date(1)
      If IsEmptyText(textTM22) = True Then textTM22 = GetTM2122Date(2)
      '2014/4/1 END
   End If
End Sub

Private Sub textTM15_GotFocus()
    TextInverse Me.textTM15
End Sub

'Add By Sindy 2010/9/1
Private Sub textTM15_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If IsEmptyText(textTM15) = False Then
      'ÀË¬d¼f©w¸¹©Ò¿é¤Jªºªø«×¬O§_¥¿½T
      'Add By Sindy 2017/5/17 + strRetrunText
      If PUB_ChkTm12Tm15Length("2", textTM15, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, , , strRetrunText) = False Then
         Cancel = True
         textTM15_GotFocus
         Exit Sub
      'Add By Sindy 2017/5/17
      Else
         textTM15 = strRetrunText
      '2017/5/17 END
      End If
   End If
End Sub

' ±M¥Î´Á­­°_¤é
Private Sub textTM21_Validate(Cancel As Boolean)
   Dim strDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textTM21) = False Then
      ' ÀË¬d¤é´Á®æ¦¡
      'edit by nick 2004/10/06
      'If CheckIsTaiwanDate(textTM21, False) = False Then
      If CheckIsDate(textTM21, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªº±M¥Î´Á­­°_¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21_GotFocus
      End If
      strDate = GetTM2122Date(1) 'Modify By Sindy 2014/3/31 ²¾¦ÜGetTM2122Date¨ç¼Æ
      If textTM21 <> strDate Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
        'Modify By Cheng 2002/12/13
'         strMsg = "±M¥Î´Á­­°_¤éÀ³¬°<" & strDate & ">"
         strMsg = "±M¥Î´Á­­°_¤éÀ³¬°<" & strDate & ">¡A¬O§_Ä~Äò§@·~¡H"
         nResponse = MsgBox(strMsg, vbYesNo, strTit)
         If nResponse = vbNo Then
            textTM21_GotFocus
         Else
            Cancel = False
         End If
      End If
   End If
End Sub

' ±M¥Î´Á­­¤î¤é
Private Sub textTM22_Validate(Cancel As Boolean)
   Dim strDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bolHaveData As Boolean
   Cancel = False
   
   If IsEmptyText(textTM22) = False Then
      ' ÀË¬d¤é´Á®æ¦¡
      'edit by nick 2004/10/06
      'If CheckIsTaiwanDate(textTM22, False) = False Then
      If CheckIsDate(textTM22, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªº±M¥Î´Á­­¤î¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22_GotFocus
      End If
      strDate = GetTM2122Date(2, bolHaveData) 'Modify By Sindy 2014/3/31 ²¾¦ÜGetTM2122Date¨ç¼Æ
      Select Case m_TM08
         Case "1", "4", "7", "8", "9":
            If textTM22 <> strDate Then
               Cancel = True
               strTit = "¸ê®ÆÀË®Ö"
               strMsg = "±M¥Î´Á­­¤î¤éÀ³¬°<" & strDate & ">¡A¬O§_Ä~Äò§@·~¡H"
               nResponse = MsgBox(strMsg, vbYesNo, strTit)
               If nResponse = vbNo Then
                  textTM22_GotFocus
               Else
                  Cancel = False
               End If
            End If
         Case Else
            If bolHaveData = False Then
               Cancel = True
               strTit = "¸ê®ÆÀË®Ö"
               strMsg = "µL¦¹®×¥¿°Ó¼Ðªº®×¥ó¸ê®Æ¡A¬O§_Ä~Äò§@·~¡H"
               nResponse = MsgBox(strMsg, vbYesNo, strTit)
               If nResponse = vbNo Then
                  textTM22_GotFocus
               Else
                  Cancel = False
               End If
            Else
               If strDate <> "" Then
                  If Val(DBDATE(textTM22)) <> Val(strDate) Then
                     Cancel = True
                     strTit = "¸ê®ÆÀË®Ö"
                     strMsg = "±M¥Î´Á­­¤î¤éÀ³¬°<" & DBDATE(rsTmp.Fields("TM22")) & ">¡A¬O§_Ä~Äò§@·~¡H"
                     nResponse = MsgBox(strMsg, vbYesNo, strTit)
                     If nResponse = vbNo Then
                        textTM22_GotFocus
                     Else
                        Cancel = False
                     End If
                  End If
               End If
            End If
      End Select
   End If
End Sub

'Add By Sindy 2014/3/31 ±N­pºâ±M¥Î´Á­­°_¨´¤éªº¤½¦¡©ñ¦b¤@°_
'strType : 1.TM21
'          2.TM22
Private Function GetTM2122Date(strType As Integer, Optional bolHaveData As Boolean) As String
Dim rsTmp As ADODB.Recordset
Dim strSql As String
   
   GetTM2122Date = ""
   
   '¦³¤½§i¤é
   If Val(textTM14) > 0 Then
      Select Case strType
         Case 1 'TM21
            '±M¥Î´Á¶¡°_¤é¬°¤½§i¤é+¤T­Ó¤ë
            'Modify By Cheng 2003/09/02
      '     strDate = TAIWANDATE(DateSerial(Val(DBYEAR(textTM14)), Val(DBMONTH(textTM14)) + 3, Val(DBDAY(textTM14))))
            '93.6.21 MODIFY BY SONIA ¥þ³¡¨Ì·sªk, ±M¥Î´Á°_¤é¬°¤½§i¤éIf Val(DBDATE(m_TM11)) < 20031128 Then
            'If Me.textTM14.Text <> "" Then
            '    strDate = TAIWANDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(textTM14))))
            'Else
            '    strDate = ""
            'End If
            'edit by nick 2004/10/06
            'Modified Lydia 2019/12/09 ¥þ³¡§ï¥Î·sªk, ¥xÆW®×=±M¥Î´Á°_¤é¬°¤½§i¤é
            'If Val(DBDATE(textTM14)) < 20030816 Then
            '   'strDate = TAIWANDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(textTM14))))
            '   GetTM2122Date = DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(textTM14))))
            'Else
            '   'strDate = TAIWANDATE(textTM14)
               GetTM2122Date = DBDATE(textTM14)
            'End If
            ''93.6.21 END
            
         Case 2 'TM22
            'Modified Lydia 2019/12/09 ¥þ³¡§ï¥Î·sªk, ¥xÆW®×=±M¥Î´Á¤î¤é¬°¤½§i¤é+10¦~-1¤Ñ
            'Select Case m_TM08
            '   'modify by sonia 2013/11/27 ¥[9¹ÎÅé°Ó¼Ð
            '   Case "1", "4", "7", "8", "9":
            '      '±M¥Î´Á¶¡¤î¤é¬°¤½§i¤é+¤T­Ó¤ë°_¤Q¦~´î¤@¤Ñ
            '      'Modify By Cheng 2003/09/02
            '      'strDate = TAIWANDATE(DateSerial(Val(DBYEAR(textTM14)) + 10, Val(DBMONTH(textTM14)) + 3, Val(DBDAY(textTM14)) - 1))
             '     '93.6.21 MODIFY BY SONIA ·sªk:±M¥Î´Á¤î¤é¬°¤½§i¤é°_¤Q¦~´î¤@¤Ñ
             '     'strDate = TAIWANDATE(DateAdd("d", -1, DateAdd("yyyy", 10, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(textTM14))))))
             '     'edit by nick 2004/10/06
             '     If Val(DBDATE(textTM14)) < 20030816 Then
             '        'strDate = TAIWANDATE(DateAdd("d", -1, DateAdd("yyyy", 10, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(textTM14))))))
            '         'Modified by Lydia 2019/11/13 §ï¥Î¦@¥Î¼Ò²Õ
            '         'GetTM2122Date = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 10, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(textTM14))))))
            '         'Modified by Lydia 2019/12/05 +´î¤@¤Ñ=Y
            '         GetTM2122Date = PUB_GetEndDate(CompDate(1, 3, DBDATE(textTM14)), 10, "Y")
            '      Else
            '         'strDate = TAIWANDATE(DateAdd("d", -1, DateAdd("yyyy", 10, ChangeWStringToWDateString(DBDATE(textTM14)))))
            '         '±M¥Î´Á¶¡¤î¤é¬°¤½§i¤é¥[¤Q¦~´î¤@¤Ñ
            '         'Modified by Lydia 2019/11/13 §ï¥Î¦@¥Î¼Ò²Õ
            '         'GetTM2122Date = DBDATE(DateAdd("d", -1, DateAdd("yyyy", 10, ChangeWStringToWDateString(DBDATE(textTM14)))))
            '         GetTM2122Date = PUB_GetEndDate(DBDATE(textTM14), 10, m_NA85)
            '      End If
            '      '93.6.21 END
            '   Case Else
            '      '91.12.20 modify by sonia
            '      'strSQL = "SELECT * FROM TRADEMARK " & _
            '      '         "WHERE TM15 = '" & m_TM27 & "' "
            '      '­Y°Ó¼ÐºØÃþ¬°2,3«h§ì1; ­Y¬°5,6«h§ì4
            '      If m_TM08 = "2" Or m_TM08 = "3" Then
            '          strSql = "Select * From TradeMark Where TM15 = '" & m_TM27 & "' And TM08 = '1' "
            '      ElseIf m_TM08 = "5" Or m_TM08 = "6" Then
            '          strSql = "Select * From TradeMark Where TM15 = '" & m_TM27 & "' And TM08 = '4' "
            '      Else
            '          strSql = "Select * From TradeMark Where TM15 = '" & m_TM27 & "' "
            '      End If
            '      '91.12.22 end
            '      Set rsTmp = New ADODB.Recordset
            '      rsTmp.CursorLocation = adUseClient
            '      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            '      If rsTmp.RecordCount > 0 Then
            '         bolHaveData = True
            '         If IsNull(rsTmp.Fields("TM22")) = False Then
            '            GetTM2122Date = DBDATE(rsTmp.Fields("TM22"))
            '         End If
            '      '91.12.22 ADD BY SONIA
            '      Else
            '         bolHaveData = False
            '      End If
            '      rsTmp.Close
            '      Set rsTmp = Nothing
            'End Select
                   'Modify By Sindy 2022/3/7 + m_TM10 : ©µ®i«á¤§±M¥Î´Á­­¦~«×­Õ¦³2¤ë29¤é®É¡A±M¥Î´Á­­¤î¤éÀ³¬°2¤ë29¤é¡A¦Ó«D¥H¥[10¦~¤§¤è¦¡­pºâ¬°2¤ë28¤é
                   GetTM2122Date = PUB_GetEndDate(DBDATE(textTM14), 10, m_NA85, m_TM10)
            'end 2019/12/09
      End Select
   End If
End Function

Private Sub textCreFee_GotFocus()
   InverseTextBox textCreFee
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPrtTrans_GotFocus()
   InverseTextBox textPrtTrans
End Sub

Private Sub textTM21_GotFocus()
   InverseTextBox textTM21
End Sub

Private Sub textTM22_GotFocus()
   InverseTextBox textTM22
End Sub

' ¦C¦L©w½Z«e±N¨Ò¥~Äæ¦ì¥[¤J¨ì¦C¦L©w½Z¨Ò¥~Äæ¦ìÀÉ®×¤¤
Private Sub InsExpField()
   Dim strSql As String
   Dim strTemp As String
   Dim strET03 As String
   
   ' ®×¥ó©Ê½è
   'Select Case m_CP10
      ' ¥Ó½Ð
   '   Case "101":
         ' ©w½Z»y¤å
         Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
            ' ¤¤¤å
            Case "1":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               '2005/8/26 MODIFY BY SONIA
               'EndLetter "05", strCP09, "01", strUserNum
'               If Query716717_cp Then
                  'edit by nickc 2005/09/30 ÄA­Ë
                  'EndLetter "05", strCP09, "01", strUserNum
                  EndLetter "05", strCP09, "21", strUserNum
'               Else
'                  'edit by nickc 2005/09/30 ÄA­Ë
'                  'EndLetter "05", strCP09, "21", strUserNum
'                  EndLetter "05", strCP09, "01", strUserNum
'               End If
               '2005/8/26 END
               
            ' ­^¤å
            Case "2":
'                '­Y¥Ó½Ð¤é¤p©ó920901
                'edit by nick 2004/09/24 ­ì¥ý¦³µù°O
'                If DBDATE(Val(m_TM11)) < 20031128 Then
'                    'Modify By Cheng 2004/03/18
''                    '­Y¤½§i¤é¤p©ó920901
''                    If DBDATE(Val(m_TM14)) < 20030901 Then
'                    '­Y±M¥Î´Á°_¤é¤p©ó921201(¥ÎÂÂ©w½Z)
'                    If Val(DBDATE(Me.textTM21.Text)) < 20031201 Then
'                       Select Case m_TM08
'                          ' Áp¦X°Ó¼Ð, Áp¦XªA°È¼Ð³¹
'                            'Modify By Cheng 2003/03/12
'        '                  Case "2", "5":
'                          Case "2":
'                             ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                             EndLetter "05", strCP09, "02", strUserNum
'                             ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                             If textPrtTrans <> "N" Then
'                                'Modify By Cheng 2003/12/26
'                                '¨Ï¥Î·sªºÄ¶¤å©w½Z
''                                ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
''                                EndLetter "05", strCP09, "03", strUserNum
''                                'Add By Cheng 2003/01/28
''                                '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
''                                If Me.Text1.Text <> "" Then
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "03" & "','" & strUserNum & _
''                                             "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
''                                    cnnConnection.Execute strSQL
''                                End If
''                                'Add By Cheng 2003/02/19
''                                '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
''                                If m_TM67 <> "" Then
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "03" & "','" & strUserNum & _
''                                             "','©ñ±ó±M¥ÎÅv','The following part disclaimed¡G" & m_TM67 & "')"
''                                    cnnConnection.Execute strSQL
''                                End If
'                                ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                                EndLetter "05", strCP09, "13", strUserNum
'                                'Add By Cheng 2003/01/28
'                                '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
'                                If Me.Text1.Text <> "" Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                'Add By Cheng 2003/02/19
'                                '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
'                                If m_TM67 <> "" Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed¡G" & ChgSQL(m_TM67) & "')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                '¨Ò¥~Äæ¦ì--ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù
'                                If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Or InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹") > 0 Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù','(Service Mark of prior Trademark Law)')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                'add by nickc 2007/03/08 ¥[¤J¦P·N®Ñ°Ó¼Ð¸¹¼Æ
'                                If m_TM118 <> "" Then
'                                    'Modify By Sindy 2012/11/06 23-I-13=>30-I-10
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & vbCrLf & "*In accordance with the proviso of Article 30-I-10 of the Trademark Law, this mark is granted registration with consent from the proprietor(s) of Reg. No(s). " & ChgSQL(m_TM118) & ".') "
'                                    cnnConnection.Execute strSql
'                                End If
'
'                             End If
'                          'Áp¦XªA°È¼Ð³¹
'                          Case "5":
'                             ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                             EndLetter "05", strCP09, "10", strUserNum
'                             ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                             If textPrtTrans <> "N" Then
'                                'Modify By Cheng 2003/12/26
'                                '¨Ï¥Î·sªºÄ¶¤å©w½Z
''                                ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
''                                EndLetter "05", strCP09, "11", strUserNum
''                                'Add By Cheng 2003/01/28
''                                '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
''                                If Me.Text1.Text <> "" Then
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "11" & "','" & strUserNum & _
''                                             "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
''                                    cnnConnection.Execute strSQL
''                                End If
''                                'Add By Cheng 2003/02/19
''                                '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
''                                If m_TM67 <> "" Then
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "11" & "','" & strUserNum & _
''                                             "','©ñ±ó±M¥ÎÅv','The following part disclaimed¡G" & m_TM67 & "')"
''                                    cnnConnection.Execute strSQL
''                                End If
'                                ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                                EndLetter "05", strCP09, "13", strUserNum
'                                'Add By Cheng 2003/01/28
'                                '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
'                                If Me.Text1.Text <> "" Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                'Add By Cheng 2003/02/19
'                                '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
'                                If m_TM67 <> "" Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed¡G" & ChgSQL(m_TM67) & "')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                '¨Ò¥~Äæ¦ì--ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù
'                                If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Or InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹") > 0 Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù','(Service Mark of prior Trademark Law)')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                'add by nickc 2007/03/08 ¥[¤J¦P·N®Ñ°Ó¼Ð¸¹¼Æ
'                                If m_TM118 <> "" Then
'                                    'Modify By Sindy 2012/11/06 23-I-13=>30-I-10
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & vbCrLf & "*In accordance with the proviso of Article 30-I-10 of the Trademark Law, this mark is granted registration with consent from the proprietor(s) of Reg. No(s). " & ChgSQL(m_TM118) & "') "
'                                    cnnConnection.Execute strSql
'                                End If
'                             End If
'                          'Add By Cheng 2003/01/17
'                          'ªA°È¼Ð³¹
'                          Case "4"
'                             ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                             EndLetter "05", strCP09, "06", strUserNum
'                             ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                             If textPrtTrans <> "N" Then
'                                'Modify By Cheng 2003/12/26
'                                '¨Ï¥Î·sªºÄ¶¤å©w½Z
''                                ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
''                                EndLetter "05", strCP09, "07", strUserNum
''                                'Add By Cheng 2003/01/28
''                                '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
''                                If Me.Text1.Text <> "" Then
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & _
''                                             "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
''                                    cnnConnection.Execute strSQL
''                                End If
''                                'Add By Cheng 2003/02/19
''                                '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
''                                If m_TM67 <> "" Then
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "07" & "','" & strUserNum & _
''                                             "','©ñ±ó±M¥ÎÅv','The following part disclaimed¡G" & m_TM67 & "')"
''                                    cnnConnection.Execute strSQL
''                                End If
'                                ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                                EndLetter "05", strCP09, "13", strUserNum
'                                'Add By Cheng 2003/01/28
'                                '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
'                                If Me.Text1.Text <> "" Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                'Add By Cheng 2003/02/19
'                                '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
'                                If m_TM67 <> "" Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed¡G" & ChgSQL(m_TM67) & "')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                '¨Ò¥~Äæ¦ì--ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù
'                                If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Or InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹") > 0 Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù','(Service Mark of prior Trademark Law)')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                'add by nickc 2007/03/08 ¥[¤J¦P·N®Ñ°Ó¼Ð¸¹¼Æ
'                                If m_TM118 <> "" Then
'                                    'Modify By Sindy 2012/11/06 23-I-13=>30-I-10
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & vbCrLf & "*In accordance with the proviso of Article 30-I-10 of the Trademark Law, this mark is granted registration with consent from the proprietor(s) of Reg. No(s). " & ChgSQL(m_TM118) & "') "
'                                    cnnConnection.Execute strSql
'                                End If
'                             End If
'                          ' ¨ä¥¦
'                          Case Else:
'                            '­Y®×¥ó³Æµù¦³°O¿ý­ì¬°Áp¦X¼Ð³¹©Î­ì¬°Áp¦XªA°È¼Ð³¹,  «h¨Ï¥ÎÁp¦X¼Ð³¹©w½Z
'                            If InStr(m_TM58, "­ì¬°Áp¦X°Ó¼Ð") > 0 Or InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹") > 0 Then
'                                ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                                EndLetter "05", strCP09, "02", strUserNum
'                            '¨ä¥L
'                            Else
'                                ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                                EndLetter "05", strCP09, "04", strUserNum
'                            End If
'                             ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                             If textPrtTrans <> "N" Then
'                                'Modify By Cheng 2003/12/26
'                                '¨Ï¥Î·sªºÄ¶¤å©w½Z
''                                ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
''                                EndLetter "05", strCP09, "05", strUserNum
''                                'Add By Cheng 2003/01/28
''                                '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
''                                If Me.Text1.Text <> "" Then
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "05" & "','" & strUserNum & _
''                                             "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
''                                    cnnConnection.Execute strSQL
''                                End If
''                                'Add By Cheng 2003/02/19
''                                '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
''                                If m_TM67 <> "" Then
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "05" & "','" & strUserNum & _
''                                             "','©ñ±ó±M¥ÎÅv','The following part disclaimed¡G" & m_TM67 & "')"
''                                    cnnConnection.Execute strSQL
''                                End If
'                                ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                                EndLetter "05", strCP09, "13", strUserNum
'                                'Add By Cheng 2003/01/28
'                                '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
'                                If Me.Text1.Text <> "" Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                'Add By Cheng 2003/02/19
'                                '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
'                                If m_TM67 <> "" Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed¡G" & ChgSQL(m_TM67) & "')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                '¨Ò¥~Äæ¦ì--ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù
'                                If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Or InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹") > 0 Then
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù','(Service Mark of prior Trademark Law)')"
'                                    cnnConnection.Execute strSql
'                                End If
'                                'add by nickc 2007/03/08 ¥[¤J¦P·N®Ñ°Ó¼Ð¸¹¼Æ
'                                If m_TM118 <> "" Then
'                                    'Modify By Sindy 2012/11/06 23-I-13=>30-I-10
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                             "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & vbCrLf & "*In accordance with the proviso of Article 30-I-10 of the Trademark Law, this mark is granted registration with consent from the proprietor(s) of Reg. No(s). " & ChgSQL(m_TM118) & "') "
'                                    cnnConnection.Execute strSql
'                                End If
'                             End If
'                       End Select
''                    '­Y¤½§i¤é¤j©óµ¥©ó920901
'                    '­Y±M¥Î´Á°_¤é¤j©óµ¥©ó921201(¥Î·s©w½Z)
'                    Else
'                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                        EndLetter "05", strCP09, "12", strUserNum
'                        ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                        If textPrtTrans <> "N" Then
'                           ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                           EndLetter "05", strCP09, "13", strUserNum
'                           'Add By Sindy 2015/6/23
'                           If m_TM08 = "7" Then 'ÃÒ©ú¼Ð³¹
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','°Ó¼ÐºØÃþ','CERTIFICATION MARK')"
'                              cnnConnection.Execute strSql
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','Class','')"
'                              cnnConnection.Execute strSql
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','ªA°È¶µ¥Ø','Contents of Certification : ')"
'                              cnnConnection.Execute strSql
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','Trademark','')"
'                              cnnConnection.Execute strSql
'                           Else
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','°Ó¼ÐºØÃþ','TRADEMARK')"
'                              cnnConnection.Execute strSql
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','Class','Class(es) : " & textTM09 & "')"
'                              cnnConnection.Execute strSql
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','ªA°È¶µ¥Ø','Specification of Goods/Services :')"
'                              cnnConnection.Execute strSql
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                       "','Trademark','Trademark ')"
'                              cnnConnection.Execute strSql
'                           End If
'                           '2015/6/23 END
'                           'Add By Cheng 2003/01/28
'                           '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
'                           If Me.Text1.Text <> "" Then
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                        "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
'                               cnnConnection.Execute strSql
'                           End If
'                           'Add By Cheng 2003/02/19
'                           '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
'                           If m_TM67 <> "" Then
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                        "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed¡G" & ChgSQL(m_TM67) & "')"
'                               cnnConnection.Execute strSql
'                           End If
'                           '¨Ò¥~Äæ¦ì--ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù
''                           If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Then
'                           If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Or InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹") > 0 Then
'                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                        "','ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù','(Service Mark of prior Trademark Law)')"
'                               cnnConnection.Execute strSql
'                           End If
'                            'add by nickc 2007/03/08 ¥[¤J¦P·N®Ñ°Ó¼Ð¸¹¼Æ
'                            If m_TM118 <> "" Then
'                                'Modify By Sindy 2012/11/06 23-I-13=>30-I-10
'                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
'                                         "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & vbCrLf & "*In accordance with the proviso of Article 30-I-10 of the Trademark Law, this mark is granted registration with consent from the proprietor(s) of Reg. No(s). " & ChgSQL(m_TM118) & "') "
'                                cnnConnection.Execute strSql
'                            End If
'                        End If
'                    End If
''                '­Y¥Ó½Ð¤é¤j©óµ¥©ó921128
'                Else
                     'Modify By Sindy 2022/8/25
                     If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, "1701", strET03, , "05") = True Then
                        EndLetter "05", strCP09, strET03, strUserNum
                     Else
                     '2022/8/25 END
                        'Add by Sindy 2020/4/24 ¬O§_°±¤î¶l°È
                        If m_NA86 = "Y" Then
                           strET03 = "23"
                           EndLetter "05", strCP09, strET03, strUserNum
                        Else
                        '2020/4/24 END
                          'edit by nick 2004/09/24
      '                    If Query716717_cp Then
                              'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
                              If Val(strSrvDate(1)) >= 20120701 Then
                                 strET03 = "22"
                                 EndLetter "05", strCP09, strET03, strUserNum
'                              Else
'                              '2012/6/27 End
'                                 strET03 = "19"
'                                 EndLetter "05", strCP09, strET03, strUserNum
                              End If
      '                    Else
      '                        EndLetter "05", strCP09, "18", strUserNum
      '                    End If
                        End If
                     End If
                     'Add By Sindy 2015/6/23
                     If m_TM08 = "7" Then 'ÃÒ©ú¼Ð³¹
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','°Ó¼ÐºØÃþ','Certification Mark')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','Class','')"
                        cnnConnection.Execute strSql
                     Else
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','°Ó¼ÐºØÃþ','Trademark')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "05" & "','" & strCP09 & "','" & strET03 & "','" & strUserNum & _
                                 "','Class','Class(es) : " & textTM09 & "')"
                        cnnConnection.Execute strSql
                     End If
                     '2015/6/23 ENd
                    'edit by nick 2004/10/07
                    If textPrtTrans <> "N" Then
                       ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                       EndLetter "05", strCP09, "13", strUserNum
                        'Add By Sindy 2015/6/23
                        If m_TM08 = "7" Then 'ÃÒ©ú¼Ð³¹
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','°Ó¼ÐºØÃþ','CERTIFICATION MARK')"
                           cnnConnection.Execute strSql
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','Class','')"
                           cnnConnection.Execute strSql
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','ªA°È¶µ¥Ø','Contents of Certification : ')"
                           cnnConnection.Execute strSql
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','Trademark','')"
                           cnnConnection.Execute strSql
                        Else
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','°Ó¼ÐºØÃþ','TRADEMARK')"
                           cnnConnection.Execute strSql
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','Class','Class(es) : " & textTM09 & "')"
                           cnnConnection.Execute strSql
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','ªA°È¶µ¥Ø','Specification of Goods/Services :')"
                           cnnConnection.Execute strSql
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','Trademark','Trademark ')"
                           cnnConnection.Execute strSql
                        End If
                        '2015/6/23 END
                       '¨Ò¥~Äæ¦ì--ÃÒ®Ñ¤é´Á
                       If Me.Text1.Text <> "" Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','ÃÒ®Ñ¤é´Á','" & DBDATE(Me.Text1.Text) & "')"
                           cnnConnection.Execute strSql
                       End If
                       '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
                       If m_TM67 <> "" Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed¡G" & ChgSQL(m_TM67) & "')"
                           cnnConnection.Execute strSql
                       End If
                       '¨Ò¥~Äæ¦ì--ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù
'                           If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Then
                       If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹") > 0 Or InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹") > 0 Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                    "','ÂÂªkµù¥U¤§ªA°È¼Ð³¹¥[µù','(Service Mark of prior Trademark Law)')"
                           cnnConnection.Execute strSql
                       End If
                        'add by nickc 2007/03/08 ¥[¤J¦P·N®Ñ°Ó¼Ð¸¹¼Æ
                        If m_TM118 <> "" Then
                            'Modify By Sindy 2012/11/06 23-I-13=>30-I-10
                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & "05" & "','" & strCP09 & "','" & "13" & "','" & strUserNum & _
                                     "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & vbCrLf & "*In accordance with the proviso of Article 30-I-10 of the Trademark Law, this mark is granted registration with consent from the proprietor(s) of Reg. No(s). " & ChgSQL(m_TM118) & "') "
                            cnnConnection.Execute strSql
                        End If
                    End If
'                End If

            ' ¤é¤å
            Case "3":
                'edit by nick 2004/09/24
                If Trim(Text2.Text) <> "Y" Then
                   '­Y¥Ó½Ð¤é¤p©ó921128(¥ÎÂÂ©w½Z)
'                   If Val(DBDATE(m_TM11)) < 20031128 Then
'                       ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                       EndLetter "05", strCP09, "14", strUserNum
'                       ' Áp¦X°Ó¼Ð
'                       If IsEmptyText(m_TM27) = False Then
'                          ' Áp¦X°Ó¼Ð
'                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "05" & "','" & strCP09 & "','" & "14" & "','" & strUserNum & _
'                                   "','Áp¦X°Ó¼Ð','" & "¨Ì¦s ¥¿°Ó¼Ð µn¿ýµf¸¹ : (" & m_TM27 & ")" & "')"
'                          cnnConnection.Execute strSql
'                       End If
'                       ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                       If textPrtTrans <> "N" Then
'                          ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                          EndLetter "05", strCP09, "15", strUserNum
'                          ' Áp¦X°Ó¼Ð
'                          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "05" & "','" & strCP09 & "','" & "15" & "','" & strUserNum & _
'                                   "','Áp¦X°Ó¼Ð','" & "¨Ì¦s ¥¿°Ó¼Ð µn¿ýµf¸¹ : (" & m_TM27 & ")" & "')"
'                          cnnConnection.Execute strSql
'                          ' °Ó«~°Ï¤À
'                          If m_TM08 = "4" Then
'                             ' °Ó«~°Ï¤À
'                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                      "VALUES ('" & "05" & "','" & strCP09 & "','" & "15" & "','" & strUserNum & _
'                                      "','°Ó«~°Ï¤À','" & "ªA°È°Ï¤À" & "')"
'                             cnnConnection.Execute strSql
'                          Else
'                             ' °Ó«~°Ï¤À
'                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                      "VALUES ('" & "05" & "','" & strCP09 & "','" & "15" & "','" & strUserNum & _
'                                      "','°Ó«~°Ï¤À','" & "°Ó«~°Ï¤À" & "')"
'                             cnnConnection.Execute strSql
'                          End If
'                          ' «ü©w°Ó«~
'                          If m_TM08 = "4" Then
'                             ' «ü©w°Ó«~
'                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                      "VALUES ('" & "05" & "','" & strCP09 & "','" & "15" & "','" & strUserNum & _
'                                      "','«ü©w°Ó«~','" & "«ü©w§Ð°È" & "')"
'                             cnnConnection.Execute strSql
'                          Else
'                             ' «ü©w°Ó«~
'                             strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                      "VALUES ('" & "05" & "','" & strCP09 & "','" & "15" & "','" & strUserNum & _
'                                      "','«ü©w°Ó«~','" & "«ü©w°Ó«~" & "')"
'                             cnnConnection.Execute strSql
'                          End If
'                       End If
'                   '­Y¥Ó½Ð¤é¤j©óµ¥©ó921128(¥Î·s©w½Z)
'                   Else
                       ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                       If Is716Have = False Then
                           EndLetter "05", strCP09, "17", strUserNum
                       Else
                           EndLetter "05", strCP09, "16", strUserNum
                       End If
                       ' Áp¦X°Ó¼Ð
                       If IsEmptyText(m_TM27) = False Then
                          ' Áp¦X°Ó¼Ð
                          'Removed by Morgan 2023/3/15 ©w½Z¨S¥Î¨ì
                          'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          '         "VALUES ('" & "05" & "','" & strCP09 & "','" & "16" & "','" & strUserNum & _
                          '         "','Áp¦X°Ó¼Ð','" & "¨Ì¦s ¥¿°Ó¼Ð µn¿ýµf¸¹ : (" & m_TM27 & ")" & "')"
                          'cnnConnection.Execute strSql
                          'end 2023/3/15
                       End If
                       ' ¬O§_¦C¦LÂ½Ä¶¨ç
                       If textPrtTrans <> "N" Then
                          ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                          'edit by nick 2004/08/17 ¦]¬°¸­©ö¶³»¡­×ªk«e«áªºÄ¶¤å¬Ò¬Û¦P
                          'EndLetter "05", strCP09, "17", strUserNum
                          EndLetter "05", strCP09, "15", strUserNum
                           'Add By Cheng 2003/02/19
                           '¨Ò¥~Äæ¦ì--©ñ±ó±M¥ÎÅv
                           If m_TM67 <> "" Then
                               'edit by nick 2004/08/17 ¦]¬°¸­©ö¶³»¡­×ªk«e«áªºÄ¶¤å¬Ò¬Û¦P
                               'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "17" & "','" & strUserNum & _
                                        "','©ñ±ó±M¥ÎÅv','°Ó¼Ð¨£¥»ÇRÆèÇr¡u" & ChgSQL(m_TM67) & "¡vÇUˆü¥e“¸Çy¦³þêÇQÆê¡C')"
                               'Modify By Sindy 2022/10/12 ˆü¥e“¸Çy¦³ §ï¬° °Ó¼Ð“¸Çy¥D±i
                               'Modified by Morgan 2023/3/15
                               'strExc(1) = "°Ó¼Ð¨£¥»ÇRÆèÇr¡u" & ChgSQL(m_TM67) & "¡vÇU°Ó¼Ð“¸Çy¥D±iþêÇQÆê¡C"
                               strExc(1) = PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv1") & ChgSQL(m_TM67) & PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv2")
                               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & "05" & "','" & strCP09 & "','" & "15" & "','" & strUserNum & _
                                        "','©ñ±ó±M¥ÎÅv','" & strExc(1) & "')"
                               cnnConnection.Execute strSql
                           End If
                           'Add By Sindy 2010/11/17
                           If m_TM118 <> "" Then
                              'Modified by Morgan 2023/3/15
                              'strExc(1) = "°Ó¼Ðªk²Ä30’f²Ä1¶µ²Ä10†AÇU³W©wÇR°òþøþà¡Bµn“÷°Ó¼Ð²Ä" & ChgSQL(m_TM118) & "†AÇU°Ó¼Ð“¸ªÌÇU¦P·NÇRÇoÇqµn“÷Çy³\¥iþìÇr¡C"
                              'Modified by Lydia 2023/04/12 debug: m_TM67=> m_TM118 ; ex. FCT49319¡B49320µù¥UÃÒ¥¼±a¥X¦P·N®Ñ°Ó¼Ð¸¹¼Æ
                              strExc(1) = PUB_GetUniText(Me.Name, "¦P·N®Ñ°Ó¼Ð¸¹¼Æ1") & ChgSQL(m_TM118) & PUB_GetUniText(Me.Name, "¦P·N®Ñ°Ó¼Ð¸¹¼Æ2")
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "05" & "','" & strCP09 & "','" & "15" & "','" & strUserNum & _
                                       "','¦P·N®Ñ°Ó¼Ð¸¹¼Æ','" & strExc(1) & "')"
                              cnnConnection.Execute strSql
                           End If
                           '2010/11/17 End
                       End If
'                   End If
               'Modify By Sindy 2010/4/14 ¦¹©w½Z§ï¦Ü§ó§ïµo¤å®É¤~°µ
'               'add by nick 2004/09/24 ¥[¤J·s¤é¤å©w½Z
'               Else
'                   EndLetter "05", strCP09, "20", strUserNum
               End If
         End Select
      'Case Else:
   'End Select
End Sub

Private Sub PrintLetter()
Dim ET03 As String 'Add By Sindy 2022/8/25

   ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
   InsExpField
   
   ' ®×¥ó©Ê½è
   'Select Case m_CP10
      ' ¥Ó½Ð
      'Case "101":
         ' ©w½Z»y¤å
         Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
            ' ¤¤¤å
            Case "1":
               '2005/8/26 MODIFY BY SONIA
                ' ¦C¦L©w½Z
                'NowPrint strCP09, "05", "01", False, strUserNum, 0
'                If Query716717_cp Then
                    'edit by nickc 2005/09/30 ÄA­Ë
                    'NowPrint strCP09, "05", "01", False, strUserNum, 0
                    NowPrint strCP09, "05", "21", False, strUserNum, 0
'                Else
'                    'edit by nickc 2005/09/30 ÄA­Ë
'                    'NowPrint strCP09, "05", "21", False, strUserNum, 0
'                    NowPrint strCP09, "05", "01", False, strUserNum, 0
'                End If
                '2005/8/26 END
                'Add By Cheng 2003/02/27
                '³]©w­n¦C¦L¦a§}±ø
                m_blnPrintAddress = True
            ' ­^¤å
            Case "2":
'                '­Y¥Ó½Ð¤é¤p©ó921128
                'edit by nick 2004/09/24 ­ì¥ý¬O³Q¤Wµù°O
'                If DBDATE(Val(m_TM11)) < 20031128 Then
''                    '­Y¤½§i¤é¤p©ó920901
''                    If DBDATE(Val(m_TM14)) < 20030901 Then
'                    '­Y±M¥Î´Á°_¤é¤p©ó921201(¥ÎÂÂ©w½Z)
'                    If Val(DBDATE(Me.textTM21.Text)) < 20031201 Then
'                       Select Case m_TM08
'                          ' Áp¦X°Ó¼Ð, Áp¦XªA°È¼Ð³¹
'                            'Modify By Cheng 2003/03/12
'        '                  Case "2", "5":
'                            'Áp¦X°Ó¼Ð
'                          Case "2":
'                             ' ¦C¦L©w½Z
'                             NowPrint strCP09, "05", "02", False, strUserNum, 0
'                            'Add By Cheng 2003/02/27
'                            '³]©w­n¦C¦L¦a§}±ø
'                            m_blnPrintAddress = True
'                             ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                             If textPrtTrans <> "N" Then
'                                ' ¦C¦L©w½Z
'                                'Modify By Cheng 2003/12/26
'                                '¨Ï¥Î·sªºÄ¶¤å©w½Z
''                                NowPrint strCP09, "05", "03", False, strUserNum, 0
'                                NowPrint strCP09, "05", "13", False, strUserNum, 0
'                             End If
'                            'Áp¦XªA°È°Ó¼Ð
'                          Case "5":
'                             ' ¦C¦L©w½Z
'                             NowPrint strCP09, "05", "10", False, strUserNum, 0
'                            'Add By Cheng 2003/02/27
'                            '³]©w­n¦C¦L¦a§}±ø
'                            m_blnPrintAddress = True
'                             ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                             If textPrtTrans <> "N" Then
'                                ' ¦C¦L©w½Z
'                                'Modify By Cheng 2003/12/26
'                                '¨Ï¥Î·sªºÄ¶¤å©w½Z
''                                NowPrint strCP09, "05", "11", False, strUserNum, 0
'                                NowPrint strCP09, "05", "13", False, strUserNum, 0
'                             End If
'                          'Add By Cheng 2003/01/16
'                          'ªA°È¼Ð³¹
'                          Case "4"
'                             ' ¦C¦L©w½Z
'                             NowPrint strCP09, "05", "06", False, strUserNum, 0
'                            'Add By Cheng 2003/02/27
'                            '³]©w­n¦C¦L¦a§}±ø
'                            m_blnPrintAddress = True
'                             ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                             If textPrtTrans <> "N" Then
'                                ' ¦C¦L©w½Z
'                                'Modify By Cheng 2003/12/26
'                                '¨Ï¥Î·sªºÄ¶¤å©w½Z
''                                NowPrint strCP09, "05", "07", False, strUserNum, 0
'                                NowPrint strCP09, "05", "13", False, strUserNum, 0
'                             End If
'                          ' ¨ä¥¦
'                          Case Else:
'                            '­Y®×¥ó³Æµù¦³°O¿ý­ì¬°Áp¦X¼Ð³¹©Î­ì¬°Áp¦XªA°È¼Ð³¹,  «h¨Ï¥ÎÁp¦X¼Ð³¹©w½Z
'                            If InStr(m_TM58, "­ì¬°Áp¦X°Ó¼Ð") > 0 Or InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹") > 0 Then
'                                ' ¦C¦L©w½Z
'                                NowPrint strCP09, "05", "02", False, strUserNum, 0
'                            '¨ä¥L
'                            Else
'                                ' ¦C¦L©w½Z
'                                NowPrint strCP09, "05", "04", False, strUserNum, 0
'                            End If
'                            'Add By Cheng 2003/02/27
'                            '³]©w­n¦C¦L¦a§}±ø
'                            m_blnPrintAddress = True
'                             ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                             If textPrtTrans <> "N" Then
'                                ' ¦C¦L©w½Z
'                                'Modify By Cheng 2003/12/26
'                                '¨Ï¥Î·sªºÄ¶¤å©w½Z
''                                NowPrint strCP09, "05", "05", False, strUserNum, 0
'                                NowPrint strCP09, "05", "13", False, strUserNum, 0
'                             End If
'                       End Select
''                    '­Y¤½§i¤é¤j©óµ¥©ó920901
'                    '­Y±M¥Î´Á°_¤é¤j©óµ¥©ó921201(¥Î·s©w½Z)
'                    Else
'                         ' ¦C¦L©w½Z
'                         NowPrint strCP09, "05", "12", False, strUserNum, 0
'                        'Add By Cheng 2003/02/27
'                        '³]©w­n¦C¦L¦a§}±ø
'                        m_blnPrintAddress = True
'                         ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                         If textPrtTrans <> "N" Then
'                            ' ¦C¦L©w½Z
'                            NowPrint strCP09, "05", "13", False, strUserNum, 0
'                         End If
'                    End If
''               '­Y¥Ó½Ð¤é¤j©óµ¥©ó921128
'                Else
                     'Modify By Sindy 2022/8/25
                     If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, "1701", ET03, , "05") = True Then
                        NowPrint strCP09, "05", ET03, False, strUserNum, 0
                     Else
                     '2022/8/25 END
                        'Add by Sindy 2020/4/24 ¬O§_°±¤î¶l°È
                        If m_NA86 = "Y" Then
                           NowPrint strCP09, "05", "23", False, strUserNum, 0
                        Else
                        '2020/4/24 END
                        'add by nick 2004/09/24
                        'If Query716717_cp Then
                           'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
                           If Val(strSrvDate(1)) >= 20120701 Then
                              NowPrint strCP09, "05", "22", False, strUserNum, 0
'                           Else
'                           '2012/6/27 End
'                              NowPrint strCP09, "05", "19", False, strUserNum, 0
                           End If
                        'Else
                        '   NowPrint strCP09, "05", "18", False, strUserNum, 0
                        'End If
                        End If
                     End If
                     '³]©w­n¦C¦L¦a§}±ø
                     m_blnPrintAddress = True
                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
                     If textPrtTrans <> "N" Then
                        ' ¦C¦L©w½Z
                        NowPrint strCP09, "05", "13", False, strUserNum, 0
                     End If
'                End If
            ' ¤é¤å
            Case "3":
                'add by nick 2004/09/24
                If Trim(Text2.Text) <> "Y" Then
                     '­Y¥Ó½Ð¤é¤p©ó921128(¥ÎÂÂ©w½Z)
'                     If Val(DBDATE(m_TM11)) < 20031128 Then
'                         ' ¦C¦L©w½Z
'                         NowPrint strCP09, "05", "14", False, strUserNum, 0
'                     '­Y¥Ó½Ð¤é¤j©óµ¥©ó921128(¥Î·s©w½Z)
'                     Else
                         ' ¦C¦L©w½Z
                         'edit by nick 2004/08/17
                         If Is716Have = False Then
                             NowPrint strCP09, "05", "17", False, strUserNum, 0
                         Else
                             NowPrint strCP09, "05", "16", False, strUserNum, 0
                         End If
'                     End If
                     'Add By Cheng 2003/02/27
                     '³]©w­n¦C¦L¦a§}±ø
                     m_blnPrintAddress = True
                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
                     If textPrtTrans <> "N" Then
                        ' ¦C¦L©w½Z
                        'edit by nick 2004/08/17 ¦]¬°¸­©ö¶³»¡­×ªk«e«áªºÄ¶¤å¬Ò¬Û¦P
                        'NowPrint strCP09, "05", "17", False, strUserNum, 0
                        NowPrint strCP09, "05", "15", False, strUserNum, 0
                     End If
                'Modify By Sindy 2010/4/14 ¦¹©w½Z§ï¦Ü§ó§ïµo¤å®É¤~°µ
'                'add by nick 2004/09/24 ¥[¤J·s¤é¤å©w½Z
'                Else
'                    m_blnPrintAddress = True
'                    NowPrint strCP09, "05", "20", False, strUserNum, 0
                End If
         End Select
      'Case Else:
   'End Select
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse

TxtValidate = False

'Add By Sindy 2010/12/24
If Me.textTM15.Enabled = True Then
   Cancel = False
   textTM15_Validate Cancel
   If Cancel = True Then
      textTM15.SetFocus
      Exit Function
   End If
End If

If Me.textTM14.Enabled = True Then
   Cancel = False
   textTM14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textTM21.Enabled = True Then
   Cancel = False
   textTM21_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textTM22.Enabled = True Then
   Cancel = False
   textTM22_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.Text1.Enabled = True Then
   Cancel = False
   Text1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

'add by nick 2004/09/24 §PÂ_¦³µL²Ä¤G´Á©Î¬O¥þ´Áªº
' Åª¨ú®×¥ó¶i«×ÀÉ
Private Function Query716717_cp() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' ¨ú±o®×¥ó¶i«×ÀÉÀÉ®×¤¤Äæ¦ì
   strSql = "SELECT count(*) FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' and cp10 in ('716','717') and cp27 is not null "
            
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04
   If rsTmp.Fields(0).Value > 0 Then
        Query716717_cp = True
   Else
        Query716717_cp = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' Åª¨ú°Ó¼Ð°ò¥»ÀÉ
Private Sub QueryMonTradeMark()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   
   m_blnReceiveSecond = False '2011/9/22 add by sonia
   ' ¨ú±o°Ó¼Ð°ò¥»ÀÉªº¬ÛÃö¶µ¥Ø
   strSql = "SELECT * FROM TradeMark,divisioncase " & _
            "WHERE dc01 = '" & m_TM01 & "' AND " & _
                  "dc02 = '" & m_TM02 & "' AND " & _
                  "dc03 = '" & m_TM03 & "' AND " & _
                  "dc04 = '" & m_TM04 & "' and dc05=tm01(+) and dc06=tm02(+) and dc07=tm03(+) and dc08=tm04(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTM12 = CheckStr(rsTmp.Fields("TM12"))       '2008/10/24 ADD BY SONIA ¤À³Î¤l®×¥Ó½Ð®×¸¹¹w³]¥À®×¥Ó½Ð®×¸¹
      textTM14 = (CheckStr(rsTmp.Fields("TM14")))
      textTM21 = (CheckStr(rsTmp.Fields("TM21")))
      textTM22 = (CheckStr(rsTmp.Fields("TM22")))
      '2011/9/22 ADD BY SONIA ¥À®×­Y¤£ºÞ¨î²Ä¤G´Á,¤À³Î®×¤]¤£ºÞ¨î
      If InStr("" & rsTmp.Fields("TM58"), "²Ä¤G´Á") > 0 Then
         m_blnReceiveSecond = True
      End If
      '2011/9/19 end
      m_MonTM01 = CheckStr(rsTmp.Fields("tm01"))
      m_MonTM02 = CheckStr(rsTmp.Fields("tm02"))
      m_MonTM03 = CheckStr(rsTmp.Fields("tm03"))
      m_MonTM04 = CheckStr(rsTmp.Fields("tm04"))
      If textNP08.Enabled = True And textNP09.Enabled = True Then
           strSql = "SELECT * FROM nextprogress " & _
                    "WHERE np02 = '" & m_MonTM01 & "' AND " & _
                         " np03 = '" & m_MonTM02 & "' AND " & _
                         " np04 = '" & m_MonTM03 & "' AND " & _
                         " np05 = '" & m_MonTM04 & "' and np06 is null and np07=202 "
          rsTmp.Close
          rsTmp.CursorLocation = adUseClient
          rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
          If rsTmp.RecordCount > 0 Then
              m_MonNP08 = CheckStr(rsTmp.Fields("np08"))
              m_MonNP09 = CheckStr(rsTmp.Fields("np09"))
          End If
      End If
   End If
   '2011/9/22 add by sonia ¥À®×¬O§_¤w¦¬²Ä¤G´Á
   If m_blnReceiveSecond = False Then
      strSql = "Select * From Caseprogress Where " & ChgCaseprogress(m_MonTM01 & m_MonTM02 & m_MonTM03 & m_MonTM04) & " And (CP10='716' OR CP10='717')"
      rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then m_blnReceiveSecond = True
   End If
   '2011/9/22 end
   
   '2011/9/22 add by sonia §ì»P¥À®×ÂI¿ï¦¬¤å¸¹¤§¬Û¦P®×¥ó©Ê½èªº¤l®×¦¬¤å¸¹T-175229(§_«h¤l®×T-175230·|§ì¨ì²§Ä³µªÅG602)
   strSql = "SELECT c1.cp09,c1.cp10,c2.cp09 FROM CaseProgress c1,caseprogress c2 WHERE c1.CP09= '" & frm02010401_6.oKey & "' " & _
            "and c2.cp01='" & m_TM01 & "' and c2.cp02='" & m_TM02 & "' and c2.cp03='" & m_TM03 & "' and c2.cp04='" & m_TM04 & "' and c1.cp10=c2.cp10 "
   rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(2)) = False Then
         m_CP09 = rsTmp.Fields(2)
      End If
      If IsNull(rsTmp.Fields(1)) = False Then
         m_CP10 = rsTmp.Fields(1)
      End If
   End If
   '2011/9/22 END
   
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2012/5/18
Private Sub Option1_Click(Index As Integer)
   If Me.Option4(0).Value Then
      Text10_Validate False
   ElseIf Me.Option4(1).Value Then
      Text11_Validate False
   ElseIf Me.Option4(2).Value Then
      Text12_Validate False
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_LostFocus()
   '«D¥xÆW"¤Ñ"¸õÂ÷®É¨ì"¥»©Ò´Á­­"Äæ¦ì
   If m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then
      If textNP08.Enabled = True Then textNP08.SetFocus
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
   CloseIme
End Sub

Private Sub Text11_LostFocus()
   '«D¥xÆW"¤ë"¸õÂ÷®É¨ì"¥»©Ò´Á­­"Äæ¦ì
   'If m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then
   '   If textNP08.Enabled = True Then textNP08.SetFocus
   'End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_LostFocus()
   '«D¥xÆW"¤é"¸õÂ÷®É¨ì"¥»©Ò´Á­­"Äæ¦ì
   If m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then
      If textNP08.Enabled = True Then textNP08.SetFocus
   End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
   Else
      If ChkDate(Text12) Then
         If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
            If Val(Text12) < Val(strSrvDate(2)) Then
               MsgBox "¨Ó¨ç´Á­­¤£¥i¤p©ó¨t²Î¤é !", vbCritical
               Cancel = True
            Else
               textNP09 = Text12
               'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
               If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                  textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
               Else
               '2014/10/6 END
                  textNP08 = TransDate(CompDate(2, -2, TransDate(textNP09, 2)), 1)
               End If
               '¥»©Ò´Á­­­Y«D¤u§@¤Ñ«h§ì³Ìªñ¤u§@¤Ñ
'               Me.textNP08.Text = TransDate(PUB_GetWorkDay1(Me.textNP08.Text, True), 1)
            End If
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub GetTime()
   Dim i As Integer
   Dim strFromDate As String '´Á­­°_ºâ¤é
   
   'Add By Sindy 2012/8/30
   If Option4(0).Value = False And Option4(1).Value = False Then Exit Sub
   '2012/8/30 End
   
   strFromDate = DBDATE(textCP05S)
   
   If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
      '¤å¨ì¤Ñ¼Æ
      If Option4(0).Value = True Then
         textNP09 = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
         If Option1(0).Value = True Then textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
         If Val(Text10) >= 60 Then
            i = -4
         Else
            i = -2
         End If
      '¤å¨ì¤ë¼Æ
      ElseIf Option4(1).Value = True Then
         textNP09 = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
         If Option1(0).Value = True Then textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
      If textNP09 <> "" Then
         'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
         If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
            textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
         Else
         '2014/10/6 END
            textNP08 = TransDate(CompDate(2, i, TransDate(textNP09, 2)), 1)
         End If
      End If
      '¥»©Ò´Á­­­Y«D¤u§@¤Ñ«h§ì³Ìªñ¤u§@¤Ñ
'      Me.textNP08.Text = TransDate(PUB_GetWorkDay1(Me.textNP08.Text, True), 1)
   End If
End Sub

'Åª¨ú¨Ó¨ç´Á­­
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
Dim strFromDate As String '´Á­­°_ºâ¤é
   
   strFromDate = DBDATE(textCP05S)
   
   ChgType = False
   If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   
   ' ®×¥ó©Ê½è
   strRvType = LabNP07.Caption '202.¥Ó½Ð·N¨£®Ñ
   If strRvType = "" Then Exit Function
   
   If ClsPDGetCaseProperty(m_TM01, strRvType, strTempName, bolTmp) Then
      textNP08 = ""
      textNP09 = ""
      
      If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
         strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & strRvType & "'"
         If strExc(0) <> "" Then
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            With RsTemp
               If intI = 1 Then
                  If Not IsNull(.Fields(1)) Then
                     '¤å¨ì¤Ñ¼Æ
                     Option4(0).Value = True
                     Text10 = .Fields(1)
                     textNP09 = TransDate(CompDate(2, Text10, TransDate(strFromDate, 2)), 1)
                  ElseIf Not IsNull(.Fields(2)) Then
                     '¤å¨ì¤ë¼Æ
                     Option4(1).Value = True
                     Text11 = .Fields(2)
                     textNP09 = TransDate(CompDate(1, .Fields(2), TransDate(strFromDate, 2)), 1)
                  Else
                     '¤å¨ì¤Ñ¼Æ
                     Option4(0).Value = True
                     Text10 = ""
                     Text11 = ""
                  End If
                  If textNP09 <> "" And Not IsNull(.Fields(0)) Then
                     '¤å¨ì·í¤é
                     If .Fields(0) = "1" Then
                        Option1(0).Value = True
                        textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
                     '¤å¨ì¦¸¤é
                     Else
                        Option1(1).Value = True
                     End If
                  End If
                  '¤å¨ì¤Ñ¼Æ
                  If Text10 <> "" Then
                     If Val(Text10) >= 60 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  '¤å¨ì¤ë¼Æ
                  ElseIf Not IsNull(.Fields(2)) Then
                     If Val(.Fields(2)) >= 2 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  End If
                  If textNP09 <> "" Then
                     'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
                     If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                        textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
                     Else
                     '2014/10/6 END
                        textNP08 = TransDate(CompDate(2, i, TransDate(textNP09, 2)), 1)
                     End If
                  End If
                  '¥»©Ò´Á­­­Y«D¤u§@¤Ñ«h§ì³Ìªñ¤u§@¤Ñ
'                  Me.textNP08.Text = TransDate(PUB_GetWorkDay1(Me.textNP08.Text, True), 1)
               End If
            End With
         End If
      End If
      ChgType = True
   End If
End Function

'Added by Lydia 2023/02/24 ¦sÀÉ¨ìFCT_WorkFlow
Private Sub PrintLetterNew()
Dim ET03 As String '©w½Z
Dim ET03_1 As String 'Ä¶¤å
Dim stLang As String '©w½Z»y¤å
Dim strFilePath As String, strFN01 As String, strFN02 As String 'Memo by Lydia 2023/06/05 strFN03§ï¦b¤W¤è«Å§i
   
   ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
   InsExpField
   stLang = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
   'Modified by Lydia 2023/05/03 §ï¦¨¦@¥Î¼Ò²Õ¡G³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW
   strFilePath = Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04)
   If Pub_GetFCTeFileName(strFilePath, m_TM01, m_TM02, m_TM03, m_TM04, "1701", , strFN01, strFN02, strFN03) = False Then
      Exit Sub
   End If
   'end 2023/05/03
   
         ' ©w½Z»y¤å
         Select Case stLang
            ' ¤¤¤å
            Case "1":
                    NowPrint strCP09, "05", "21", True, strUserNum, 0
                    m_blnPrintAddress = True '³]©w­n¦C¦L¦a§}±ø
                    'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                    If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                        Sleep 100
                    End If
                    'end 2023/05/03
            ' ­^¤å
            Case "2":
''               '­Y¥Ó½Ð¤é¤j©óµ¥©ó921128
                     'Modify By Sindy 2022/8/25
                     If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, "1701", ET03, , "05") = True Then
                        NowPrint strCP09, "05", ET03, True, strUserNum, 0
                        'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                        If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                            Sleep 100
                        End If
                        'end 2023/05/03
                     Else
                     '2022/8/25 END
                        'Add by Sindy 2020/4/24 ¬O§_°±¤î¶l°È
                        If m_NA86 = "Y" Then
                           NowPrint strCP09, "05", "23", True, strUserNum, 0
                           'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                           If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                               Sleep 100
                           End If
                           'end 2023/05/03
                        Else
                           'Modify By Sindy 2012/6/27 °Ó¼Ð­×ªk
                           If Val(strSrvDate(1)) >= 20120701 Then
                              NowPrint strCP09, "05", "22", True, strUserNum, 0
                              'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                              If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                                  Sleep 100
                              End If
                              'end 2023/05/03
'                           Else
'                           '2012/6/27 End
'                              NowPrint strCP09, "05", "19", True, strUserNum, 0
'                              'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
'                              If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
'                                  Sleep 100
'                              End If
'                              'end 2023/05/03
                           End If
                        End If
                     End If
                     '³]©w­n¦C¦L¦a§}±ø
                     m_blnPrintAddress = True
                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
                     If textPrtTrans <> "N" Then
                        ' ¦C¦L©w½Z
                        NowPrint strCP09, "05", "13", True, strUserNum, 0
                        'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                         If PUB_PrintWord2File(g_WordAp, strFilePath, strFN02) = True Then
                            Sleep 100
                        End If
                        'end 2023/05/03
                     End If
            ' ¤é¤å
            Case "3":
                If Trim(Text2.Text) <> "Y" Then
                         If Is716Have = False Then
                             NowPrint strCP09, "05", "17", True, strUserNum, 0
                             'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                             If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                                 Sleep 100
                             End If
                             'end 2023/05/03
                         Else
                             NowPrint strCP09, "05", "16", True, strUserNum, 0
                             'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                             If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                                 Sleep 100
                             End If
                             'end 2023/05/03
                         End If
                     '³]©w­n¦C¦L¦a§}±ø
                     m_blnPrintAddress = True
                     ' ¬O§_¦C¦LÂ½Ä¶¨ç
                     If textPrtTrans <> "N" Then
                        ' ¦C¦L©w½Z
                        NowPrint strCP09, "05", "15", True, strUserNum, 0
                        'Modified by Lydia 2023/05/03 §ï¦@¥Î¼Ò²Õ
                        If PUB_PrintWord2File(g_WordAp, strFilePath, strFN02) = True Then
                            Sleep 100
                        End If
                        'end 2023/05/03
                     End If
                End If
         End Select
   '¤U¸ü¨÷©v°ÏªºÃÒ®ÑPDF
   'Mark by Lydia 2023/06/05 ¹q¤l©Î¯È¥»ÃÒ®Ñ²Î¤@¦b³Ì«á¤U¸ü¨÷©v°ÏªºÃÒ®ÑPDF
   'strSql = "select cpp14 From casepaperpdf where cpp01='" & NowCP09 & "' and instr(upper(cpp02),upper('." & IIf(m_TM136 = "1", "CERT", "1701") & ".PDF'))>0"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   'If intI = 1 Then
   '   If PUB_GetFtpFile("" & RsTemp.Fields("cpp14"), Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04) & "\" & strFN03, "Casepaperpdf") = True Then
   '   End If
   'End If
   'end 2023/06/05
End Sub

