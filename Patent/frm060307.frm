VERSION 5.00
Begin VB.Form frm060307 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "¦~ÃÒ¶O½Ð´Ú¨ç"
   ClientHeight    =   3804
   ClientLeft      =   1248
   ClientTop       =   2316
   ClientWidth     =   5028
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3804
   ScaleWidth      =   5028
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Caption         =   "¾ã§å"
      Height          =   2325
      Left            =   90
      TabIndex        =   17
      Top             =   540
      Width           =   4845
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   6
         Left            =   2850
         MaxLength       =   7
         TabIndex        =   3
         Top             =   585
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   4
         Left            =   3090
         MaxLength       =   2
         TabIndex        =   7
         Top             =   930
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   3
         Left            =   2850
         MaxLength       =   1
         TabIndex        =   6
         Top             =   930
         Width           =   255
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   2
         Left            =   2010
         MaxLength       =   6
         TabIndex        =   5
         Top             =   930
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   264
         Index           =   1
         Left            =   1530
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "FCP"
         Top             =   930
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   264
         Index           =   5
         Left            =   1530
         MaxLength       =   7
         TabIndex        =   2
         Top             =   585
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "·s¼W"
         Default         =   -1  'True
         Height          =   400
         Index           =   0
         Left            =   3510
         TabIndex        =   8
         Top             =   915
         Width           =   600
      End
      Begin VB.CommandButton Command1 
         Caption         =   "§R°£"
         Height          =   400
         Index           =   1
         Left            =   4110
         TabIndex        =   9
         Top             =   915
         Width           =   600
      End
      Begin VB.ListBox List1 
         Height          =   948
         ItemData        =   "frm060307.frx":0000
         Left            =   1545
         List            =   "frm060307.frx":0002
         TabIndex        =   10
         Top             =   1245
         Width           =   1935
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2370
         X2              =   2490
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   1
         Left            =   2850
         TabIndex        =   22
         Top             =   330
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   0
         Left            =   1530
         TabIndex        =   21
         Top             =   330
         Width           =   480
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2610
         X2              =   2730
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¤W¦¸¦C¦Lµo¤å¤é¡G"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   20
         Top             =   330
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥»¦¸¦C¦Lµo¤å¤é¡G"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   645
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥»¦¸¤£¾A¥Î®×¸¹¡G"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   18
         Top             =   945
         Width           =   1440
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "¦~¶O"
      Height          =   315
      Index           =   1
      Left            =   1710
      TabIndex        =   1
      Top             =   150
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "»âÃÒ¤ÎÃº¦~¶O"
      Height          =   315
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1515
      TabIndex        =   12
      Top             =   3360
      Width           =   3315
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1515
      TabIndex        =   11
      Top             =   3030
      Width           =   3315
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Height          =   400
      Index           =   0
      Left            =   3330
      TabIndex        =   13
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4095
      TabIndex        =   14
      Top             =   90
      Width           =   756
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "¦a§}±ø¦Lªí¾÷¡G"
      Height          =   180
      Index           =   0
      Left            =   195
      TabIndex        =   16
      Top             =   3450
      Width           =   1260
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "½Ð´Ú³æ¤Î©w½Z¡G"
      Height          =   180
      Index           =   11
      Left            =   195
      TabIndex        =   15
      Top             =   3090
      Width           =   1260
   End
End
Attribute VB_Name = "frm060307"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/10/26 ¤é¤å¤w§ï©ñ©w½Z¤º
'Memo By Morgan 2012/12/10 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo by Morgan2010/12/27 ¥Ó½Ð®×¸¹Äæ¤w­×§ï
'2010/12/6 memo by sonia ­û¤u½s¸¹Äæ¤w­×§ï
'Memo by Morgan2010/8/13 ¤é´ÁÄæ¤w­×§ï
Option Explicit

Dim intWhere As Integer, strReceiveNo As String
Dim m_LetterLanguage As String
Dim m_LetterKind As Integer
Dim m_strCP01 As String '¥»©Ò®×¸¹
Dim m_strCP02 As String '¥»©Ò®×¸¹
Dim m_strCP03 As String '¥»©Ò®×¸¹
Dim m_strCP04 As String '¥»©Ò®×¸¹
Dim m_strCP09 As String 'Á`¦¬¤å¸¹
Dim m_strCP10 As String '®×¥ó©Ê½è
Dim m_strCP16 As String '¶O¥Î
Dim m_strCP17 As String '³W¶O
Dim m_strCP18 As String 'ÂI¼Æ
Dim m_strCP27 As String 'µo¤å¤é
Dim m_strPA72 As String '¦~¶O¤wÃº¦~«×
Dim m_strPA73 As String '¦~¶OÃº¶O¦~«×

Dim m_strNP09 As String '¤U¦¸Ãº¶O¤é(ªk©w)
Dim m_strNP08 As String '¤U¦¸Ãº¶O¤é(¥»©Ò)
Dim m_strNP23 As String '¤U¦¸Ãº¶O¤é(¬ù©w) 'Add By Sindy 2021/4/27

Dim m_strPA09 As String '¥Ó½Ð°ê®a
Dim m_strPA08 As String '±M§QºØÃþ
'Added by Morgan 2019/8/6
Dim m_strPA10 As String '¥Ó½Ð¤é´Á
Dim m_strPA25 As String '±M¥Î´Á¤î¤é

Dim m_strStartDBNo As String  '°_©lD/N No.
Dim m_strEndDBNo As String    'ºI¤îD/N No.

'Dim m_CustList() As String    ' ¼È¦s¥Ó½Ð¤H
'Dim m_CustListCount As Integer
'Dim m_CP() As String          '¼È¦s¥»©Ò®×¸¹

Dim m_AddWorkFCount As Integer     '¼È¦s¦~¶O³æµ§¤£¶]µ§¼Æ

Dim strPrinter As String
Dim PLeft(0 To 11) As Integer, Page As Integer, strTemp(0 To 11) As String, iPrint As Integer
Dim m_PayBefore As String  '¦¬´Ú«á¿ì®×

Const ET01 As String = "10"
'Add By Cheng 2003/01/30
Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer, j As Integer, i As Integer
'Add by Morgan 2004/6/25
Dim m_bolNew As Boolean '¬O§_¥Î·sªk
Dim m_bol412 As Boolean '¬O§_¦³µo¤å©µ½w¤½§i
Dim pa(1 To 4) As String
Dim stPA14 As String
Dim m_stPS As String 'Add by Morgan 2004/10/15 ½Ð´Ú¨ç³Æµù
'Add by Morgan 2008/4/3
Dim m_bolEmail As Boolean '¬O§_¥HEMail³qª¾(²£¥Í¹q¤lÀÉ)
Dim m_PA26 As String '¥Ó½Ð¤H
Dim m_PA75 As String '¥N²z¤H Added by Morgan 2016/10/13
Dim m_bolPlusPaper As Boolean, m_iCopy As Integer
'Added by Morgan 2014/6/3
Dim m_bolDNEmail As Boolean, m_bolDNPlusPaper As Boolean
Dim strPrinter2 As String 'Add By Sindy 2015/9/24
Dim m_NationID As String 'Add By Sindy 2017/2/15 °êÄy¥N½X
'Added by Lydia 2019/10/01 ¦~¶Oµo¤å§t¾ã§å¦~¶O¡G¶Ç¤J­Ó®×¤§¦¬¤å¸¹,®×¥ó©Ê½è(±N­ì¥»«ü©w¤H­û©w´Á¶]µ{¦¡, §ï¦b­Ó®×µo¤å¦Û¦æ³B²z)
Public m_KeyCP09 As String
Public m_KeyCP10 As String
Dim m_OsPrinter  As String 'Added by Lydia 2019/12/27 §@·~¨t²Î¹w³]¦Lªí¾÷
'Added by Lydia 2020/08/17 ¾ã¦Xµo¤å©M³qª¾¨ç¡A¥þ¥H¾ã§åµ{¦¡°õ¦æ
Public m_bTransOK As Boolean '¥~³¡©I¥s:¬O§_§¹¦¨( Transaction±Nµo¤å©M¦~ÃÒ¶O½Ð´Ú¨ç¤@¨Ö¥]¤J)
Dim m_OptKind As String 'Ãþ§O¡G ªÅ¥Õ¡G«D¥~³¡©I¥s, 1-­Ó®×µo¤åfrm060104_7/frm060104_a, 2-¾ã§åµo¤åfrm060104_j, 3-³qª¾¨çfrm060306_2
Dim m_OptPAID As String '¿é¤J¡G¤w¦¬´Ú(1-¤£±HD/N, 2-±HD/N)
Dim m_OptRecDate As String '¿é¤J¡G·í¤Ñ½Ð´ÚY
Dim m_OptMemo As String '¿é¤J¡G¹O´Á¸ÉÃº(¨Ó·½ªí³æªº³]©w¤§´y­z)
Dim m_eFlag As String '¬O§_e/E¤Æ
Dim m_OptMCrec As String '¿é¤J¡G¤H¤uEmailºûÅ@


'Added by Lydia 2020/08/17 ¥~³¡©I¥s³]©w
Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' ²M°£·j´MªºKey
   If bClear = True Then
      m_OptKind = ""
      m_OptPAID = ""
      m_OptRecDate = ""
      m_OptMemo = ""
      m_OptMCrec = ""
   End If
   
   Select Case nType
      Case 0: 'Ãþ§O 'ªÅ¥Õ¡G«D¥~³¡©I¥s, 1-­Ó®×µo¤åfrm060104_7/frm060104_a, 2-¾ã§åµo¤åfrm060104_j, 3-³qª¾¨çfrm060306_2
          m_OptKind = strData
      Case 1: '¤w¦¬´Ú(1-¤£±HD/N, 2-±HD/N)
          m_OptPAID = strData
      Case 2: '·í¤Ñ½Ð´ÚY
          m_OptRecDate = strData
      Case 3: '¹O´Á¸ÉÃº(¨Ó·½ªí³æªº³]©w¤§´y­z)
          m_OptMemo = strData
      Case 4: '¤H¤uEmailºûÅ@
          m_OptMCrec = strData
   End Select
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
'Modified by Lydia 2019/08/28 ¥[¤j¯Á¤Þ strTxt(1 To 6)=> strTxt(1 To 20)
Dim strTxt(1 To 20) As String, i As Integer, s As Integer, strTmp As String
Dim strAnnuity As String
Dim strTemp1 As Variant
Dim strTemp2 As Variant
Dim strSendDate As String
Dim intYear As Integer
Dim ii As Integer

   EndLetter ET01, m_strCP09, ET03, strUserNum
   
   If m_strCP10 = »âÃÒ¤ÎÃº¦~¶O Or m_strCP10 = ¦~¶O Then
      '­Y¦³Ãº¦~¶O¤é´Á
      If m_strPA73 <> "" Then
         strTemp1 = Split(UCase(m_strPA72), ",")
         strTemp2 = Split(UCase(m_strPA73), ",")
         For i = 0 To UBound(strTemp2)
            If Val(strTemp2(i)) = m_strCP27 Then
               If Val(strSendDate) <> m_strCP27 Then
               
                  'Modify by Morgan 2004/10/11 ¥[¤é¤å©w½Z
                  'Modify by Morgan 2016/8/18 +¤¤¤å
                  'If m_LetterLanguage = 3 Then
                  If m_LetterLanguage <> 2 Then
                     strAnnuity = strTemp1(i)
                  Else
                  
                     Select Case Val(strTemp1(i))
                        Case 1
                           strAnnuity = strTemp1(i) & "st to "
                        Case 2
                           strAnnuity = strTemp1(i) & "nd to "
                        Case 3
                           strAnnuity = strTemp1(i) & "rd to "
                        Case Else
                           strAnnuity = strTemp1(i) & "th to "
                     End Select
                     
                  End If
                  
                  strSendDate = strTemp2(i)
               End If
               s = i
               intYear = intYear + 1
            End If
         Next i
      End If
        '­Y¦P®ÉÃº¦h¦~
      If intYear > 1 Then
         'Modify by Morgan 2004/10/11 ¦Ò¼{¤é¤å©w½Z
         'Modify by Morgan 2016/8/18 +¤¤¤å
         'If m_LetterLanguage = 3 Then
         If m_LetterLanguage <> 2 Then
            strAnnuity = strAnnuity & " ~ " & strTemp1(s)
         Else
            Select Case Val(strTemp1(s))
               Case 1
                  strAnnuity = strAnnuity & strTemp1(s) & "st"
               Case 2
                  strAnnuity = strAnnuity & strTemp1(s) & "nd"
               Case 3
                  strAnnuity = strAnnuity & strTemp1(s) & "rd"
               Case Else
                  strAnnuity = strAnnuity & strTemp1(s) & "th"
            End Select
         End If
        '­Y¥uÃº¤@¦~
      Else
        '­YµLÃº¦~¶O¤é´Á
         If IsEmpty(strTemp2) Then
            'Modify by Morgan 2004/10/11 ¦Ò¼{¤é¤å©w½Z
            'Modify by Morgan 2016/8/18 +¤¤¤å
            'If m_LetterLanguage = 3 Then
            If m_LetterLanguage <> 2 Then
               strAnnuity = "1"
            Else
               strAnnuity = "1st"
            End If
        '­Y¦³Ãº¦~¶O¤é´Á
         Else
            'Modify by Morgan 2004/10/11 ¦Ò¼{¤é¤å©w½Z
            'Modify by Morgan 2016/8/18 +¤¤¤å
            'If m_LetterLanguage <> 3 Then
            If m_LetterLanguage = 2 Then
               
               Select Case Val(strTemp1(UBound(strTemp2)))
                  Case 1
                     strAnnuity = strTemp1(UBound(strTemp2)) & "st"
                  Case 2
                     strAnnuity = strTemp1(UBound(strTemp2)) & "nd"
                  Case 3
                     strAnnuity = strTemp1(UBound(strTemp2)) & "rd"
                  Case Else
                     strAnnuity = strTemp1(UBound(strTemp2)) & "th"
               End Select
            End If
         End If
      End If
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','²Ä´X¦~¦Ü´X¦~¶O'," & CNULL(strAnnuity) & ")"
      If Val(m_strNP09) <> 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','¦~¶Oªk©w´Á­­'," & DBDATE(m_strNP09) & ")"
         ii = ii + 1
         'Modify By Sindy 2021/4/27
         If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','¦~¶O¬ù©w´Á­­'," & DBDATE(m_strNP23) & ")"
         Else
         '2021/4/27 END
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','¦~¶O¥»©Ò´Á­­'," & DBDATE(m_strNP08) & ")"
         End If
      End If
      
      'Add by Morgan 2004/6/25
      '»âÃÒ¥B¦³µo¤å©µ½w¤½§i
      If m_bol412 = True Then
         ii = ii + 1
         'Modify by Morgan 2005/1/13 ¥[¤é¤å©µ½w¤½§i
         'Modified by Morgan 2022/10/26
         'If m_LetterLanguage = 3 Then
         '   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         '      "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','©µ½w¤½§i','¡@¤S¡B®Æª÷¯Ç¥IÇU¤â“dþàÇO¦P®ÉÇR¡Bµn“÷¤½³øÇU´¦¸üÇy…üÇpþîÇr¦®Çy¦P§½ÇR¤W¥Ó­PþêÇeþêþò¡C " & "')"
         'Else
         '   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         '      "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','©µ½w¤½§i','    We have also filed a request for delaying publication of allowed claim(s)." & Chr(13) & "')"
         'End If
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','©µ½w¤½§i­n¦L','¡ð')"
         'end 2022/10/26
      End If
    
   End If
   
   'Add By Cheng 2003/05/20
   '½Ð´Ú¨ç¤é´Á
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç¤é´Á'," & strSrvDate(1) & ")"
   
   'Add by Morgan 2004/10/15 ½Ð´Ú¨ç³Æµù
   If m_stPS <> "" Then
      'Add By Sindy 2015/9/24
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','¦³½Ð´Ú¨ç³Æµù®É¤£¦L','¡ð')"
      '2015/9/24 END
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç³Æµù','" & ChgSQL(m_stPS) & "')"
   End If
   '2004/10/15 end
   
   'Added by Morgan 2016/10/13
   If m_PA75 = "Y54391000" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','µo²¼','µo²¼')"
   
   End If
   'end 2016/10/13
  
'Modified by Morgan 2019/11/5 ³]­p±M¥Î´Á¤w§ó·s,­^¤å¦~¶O½Ð´Ú¨ç¤w¦X¨Ö
'   'Added by Morgan 2019/8/6
'   '108.11.1·sªk³]­p®×±M¥Î´Á¥Ñ12¦~©µªø¬°15¦~
'   If m_strPA08 = "3" Then
'      '±M¥Î´Á§ó·s«e¯S®í±±¨î,§ó·s«á¥i§ï³£§ì±M¥Î´Á¤î¤é
'      If strSrvDate(1) < 20191101 Then
'         'strExc(1) = CompDate(2, -1, CompDate(0, 15, m_strPA10)) 'Removed by Morgan 2019/8/16 §ï¼g¦@¥Î¨Ò¥~Äæ¦ì
'
'         '±M¥Î´Á§ó·s«e¯S®í±±¨î,§ó·s«á¥i³s¦P©w½Z¤º¨Ò¥~Äæ¦ì¤@¨Ö²¾°£
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'            "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','108/10/31«e³]­p®×¤£¦L','¡ð')"
'
'      Else
'         'strExc(1) = m_strPA25 'Removed by Morgan 2019/8/16 §ï¼g¦@¥Î¨Ò¥~Äæ¦ì
'      End If
'
'      'Removed by Morgan 2019/8/16 §ï¼g¦@¥Î¨Ò¥~Äæ¦ì
'      'ii = ii + 1
'      'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','³]­p®×15¦~©¡º¡¤é','" & strExc(1) & "')"
'
'   End If
'   'end2019/8/6

   '¦Û°Ê¥NÃº
   'Modified by Morgan 2020/5/26
   'If m_LetterKind = "3" Then
   If m_LetterKind = "2" Then
   'end 2020/5/26
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','¦Û°Ê¥NÃº¤£¦L','¡ð')"
   End If
   
   'Added by Lydia 2020/08/17 ¤w¦¬´Ú¤£±HD/N=¦³½Ð´Ú¨ç³Æµù®É¤£¦L
   If m_OptPAID <> "" Then
      If m_stPS = "" And m_OptPAID = "1" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
             "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','¦³½Ð´Ú¨ç³Æµù®É¤£¦L','¡ð')"
'         '¤é¤å©w½Z
'         If m_LetterLanguage = "3" Then
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','¤£±HD/N','¡ð')"
'         Else '¤¤¤å,­^¤å©w½Z =¦³½Ð´Ú¨ç³Æµù®É¤£¦L
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'                "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','¦³½Ð´Ú¨ç³Æµù®É¤£¦L','¡ð')"
'         End If
      End If
   End If
   'end 2020/08/17
   
   '³Ì«á¤@¦~
   If m_LetterKind = "4" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','³Ì«á¤@¦~¤£¦L','¡ð')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','³Ì«á¤@¦~¤~¦L','¡ð')"
   End If
'end 2019/11/5
   
'Removed by Morgan 2015/10/12 ©w½Z³æ¤@¤Æ,¨ú®ø--¦¿¦p¥É
'   'Add by Morgan 2008/4/28 ­^¤å¦~¶O½Ð´Ú¨ç X49199001 ¤£¥Îªþ©x¤è¦¬¾Ú
'   If m_strCP10 = ¦~¶O Then
'      If m_PA26 = "X49199001" Then
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'            "('" & ET01 & "','" & m_strCP09 & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó','our debit note')"
'      End If
'   End If
'   'end 2008/4/28
   
   'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
   'If Not objLawDll.ExecSQL(ii, strTxt) Then
   If Not ClsLawExecSQL(ii, strTxt, True) Then
      MsgBox "Àx¦s¨Ò¥~Äæ¦ì¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
   End If

End Sub

'Modified by Lydia 2019/10/01 §ï¦¨¦@¥Î¼Ò²Õ
'Private Sub cmdok_Click(Index As Integer)
Public Sub cmdok_Click(Index As Integer)
Dim strTmp As String, strTmp1 As String, strTmp2 As String, rsTemp1 As ADODB.Recordset
Dim rsMain As ADODB.Recordset
Dim strTxt(1 To 2) As String, i As Integer
Dim blnTransaction As Boolean
Dim stNo As String 'Add by Morgan 2004/12/8 ¥N²z¤H
Dim stNo1 As String 'Added by Morgan 2013/3/25
Dim stCon As String, opt1idx As Integer
Dim stLstFNo As String, stCtMan As String
Dim strCaseNo As String, strFileName As String, strFullFileName As String 'Add By Sindy 2017/2/15
Dim m_CP148 As String, m_Case2cp148 As String 'Added by Lydia 2019/10/08 ¦~¶Oµo¤å(¶Ç¤J¦¬¤å¸¹)-¯S®í½Ð´Ú³æ
Dim m_CP60 As String, iCopy As Integer  'Added by Lydia 2020/08/17
Dim stNo2 As String, stNo3 As String, stNo4 As String, stNo5 As String    'Added by Lydia 2023/08/01 ¥Ó½Ð¤H2~5

On Error GoTo ErrorHandler
   
   blnTransaction = False
   Select Case Index
      Case 0 '½T©w
         If m_KeyCP09 = "" Then 'Added by Lydia 2019/10/01 §PÂ_«D¦~¶Oµo¤å
            If (Option1(0).Value = True Or Option1(1).Value = True) Then
                'Modify by Morgan 2011/3/15 ±qValidate ²¾¨Ó
                If Text1(5) = "" Or Text1(6) = "" Then
                   MsgBox "¥»¦¸¦C¦Lµo¤å¤é¤£¥iªÅ¥Õ¡A½Ð­«·s¿é¤J !", vbCritical
                   If Text1(5) = "" Then
                      Text1(5).SetFocus
                   Else
                      Text1(6).SetFocus
                   End If
                   Exit Sub
                End If
                'end 2011/3/15
                
                If PUB_CheckKeyInDate(Me.Text1(5)) = -1 Then
                   Me.Text1(5).SetFocus
                   Text1_GotFocus 5
                   Exit Sub
                End If
                If PUB_CheckKeyInDate(Me.Text1(6)) = -1 Then
                   Me.Text1(6).SetFocus
                   Text1_GotFocus 6
                   Exit Sub
                End If
            End If
         'Added by Lydia 2019/10/14 ¦~¶Oµo¤å:¨t²Î¤é=µo¤å¤é(¤£°µ°O¿ý)
         Else
             Text1(5) = strSrvDate(2)
             Text1(6) = strSrvDate(2)
         End If '2019/10/01

         m_bTransOK = False 'Added by Lydia 2020/08/17
         Screen.MousePointer = vbHourglass
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 ²M°£¬d¸ß¦Lªí°O¿ýÀÉÄæ¦ì
         strTmp = ""
         m_strNP09 = 0
         For i = 0 To List1.ListCount - 1
            strTmp = strTmp & List1.List(i) & ","
         Next
         
         ' ²M°£
         'ClearCustList
         
         ' ²M°£¦~¶O³æµ§¤£¶]¤u§@ÀÉ
         strExc(0) = "DELETE R060307 WHERE ID='" & strUserNum & "'"
         cnnConnection.Execute strExc(0)
         
         pub_QL05 = pub_QL05 & ";" & Label1(1) & Text1(5) & "-" & Text1(6) 'Add By Sindy 2010/12/7
         
         intI = 1
         'Modified by Morgan 2013/4/18
         'stCon = " AND CP10 IN ('" & »âÃÒ¤ÎÃº¦~¶O & "','" & ¥[µù°l¥[ & "','" & ¥[µùÁp¦X & "','" & ¦~¶O & "')"
         If m_KeyCP09 = "" Then 'Added by Lydia 2019/10/01 §PÂ_«D¦~¶Oµo¤å
            If Option1(0).Value = True Then
               'Modify By Sindy 2015/9/23 ¤wµL ¥[µù°l¥[,¥[µùÁp¦X
               'stCon = " AND CP10 IN ('" & »âÃÒ¤ÎÃº¦~¶O & "','" & ¥[µù°l¥[ & "','" & ¥[µùÁp¦X & "')"
               stCon = " AND CP10='" & »âÃÒ¤ÎÃº¦~¶O & "'"
               opt1idx = 0
            ElseIf Option1(1).Value = True Then
               stCon = " AND CP10='" & ¦~¶O & "'"
               opt1idx = 1
            End If
            'end 2013/4/18
         'Added by Lydia 2019/10/01
            If (Option1(0).Value = True Or Option1(1).Value = True) Then
                stCon = stCon & " and CP27 BETWEEN " & TransDate(Text1(5), 2) & " AND " & TransDate(Text1(6), 2)
            End If
         Else  '¦~¶Oµo¤å(¶Ç¤J¦¬¤å¸¹)
            If m_KeyCP10 = »âÃÒ¤ÎÃº¦~¶O Then
                opt1idx = 0
                Option1(0).Value = True
            Else
                opt1idx = 1
                Option1(1).Value = True
            End If
            If Len(m_KeyCP09) > 9 Then '¦~¶O¾ã§åµo¤å
                stCon = " and cp09 in (" & GetAddStr(m_KeyCP09) & ") and cp10='" & m_KeyCP10 & "' and nvl(cp27,0) > 0 "
            Else
                stCon = " and cp09='" & m_KeyCP09 & "' and cp10='" & m_KeyCP10 & "' and nvl(cp27,0) > 0 "
            End If
            stCon = stCon & " and cp158>0 " 'Added by Lydia 2020/08/17
         End If
         'end 2019/10/01
         
         'Modified by Morgan 2014/5/29 §ï±Æ§Ç:ºÞ¨î¤H,¦¬«H¤H,¥»©Ò®×¸¹
         'strExc(0) = "SELECT CP01,CP02,CP03,CP04,CP09,CP10,CP16,CP17,CP27,PA14,CP18 FROM CASEPROGRESS,PATENT WHERE CP01='FCP' " & stCon & _
            " and CP27 BETWEEN " & TransDate(Text1(5), 2) & " AND " & TransDate(Text1(6), 2) & " AND (CP20<>'N' OR CP20 IS NULL) AND CP60 IS NULL AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL ORDER BY 1,2,3,4"
         'Modified by Morgan 2015/6/8 +PA26
         'Modify By Sindy 2017/2/15 + ,nvl(fa10,cu10) NaName
         'Modify By Sindy 2017/3/7 + nvl(n1.na51,n2.na51) FCPEmp
         'Modified by Lydia 2019/10/01 §ï¦¨¶Ç¤J±ø¥ó; 2019/10/08 +CP148
         'strExc(0) = "select CP01,CP02,CP03,CP04,CP09,CP10,CP16,CP17,CP27,PA14,CP18,FNo,nvl(n1.na16,n2.NA16) CtMan,PA26,PA75,nvl(fa10,cu10) NaID,nvl(n1.na51,n2.na51) FCPEmp" & _
            " from (SELECT CP01,CP02,CP03,CP04,CP09,CP10,CP16,CP17,CP27,PA14,CP18,Nvl(PA76, Nvl(CU96, Nvl(PA75, PA26))) FNo,PA26,PA75" & _
            " FROM CASEPROGRESS,PATENT,CUSTOMER WHERE CP01='FCP' " & stCon & " and CP27 BETWEEN " & TransDate(Text1(5), 2) & " AND " & TransDate(Text1(6), 2) & _
            " AND (CP20<>'N' OR CP20 IS NULL) AND CP60 IS NULL AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL" & _
            " and cu01(+)=substr(PA26,1,8) and cu02(+)=substr(PA26,9)),fagent,customer,nation n1,nation n2" & _
            " where fa01(+)=substr(fno,1,8) and fa02(+)=substr(fno,9) and cu01(+)=substr(fno,1,8) and cu02(+)=substr(fno,9) and n1.na01(+)=fa10 and n2.na01(+)=cu10"
         'Modified by Lydia 2020/08/17 ®³±¼AND CP60 IS NULL¡Aª½±µ§PÂ_CP60
         'Modified by Lydia 2020/08/17 +¬O§_e/E¤Æ ;  GetEmailFlag(CP01||CP02||CP03||CP04) as eFlag
         strExc(0) = "select CP01,CP02,CP03,CP04,CP09,CP10,CP16,CP17,CP27,PA14,CP18,FNo,nvl(n1.na16,n2.NA16) CtMan,PA26,PA75,nvl(fa10,cu10) NaID,nvl(n1.na51,n2.na51) FCPEmp,CP148, CP60, GetEmailFlag(CP01||CP02||CP03||CP04) as eFlag " & _
            " from (SELECT CP01,CP02,CP03,CP04,CP09,CP10,CP16,CP17,CP27,PA14,CP18,Nvl(PA76, Nvl(CU96, Nvl(PA75, PA26))) FNo,PA26,PA75,CP148,CP60" & _
            " FROM CASEPROGRESS,PATENT,CUSTOMER WHERE CP01='FCP' " & stCon & _
            " AND (CP20<>'N' OR CP20 IS NULL)  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL" & _
            " and cu01(+)=substr(PA26,1,8) and cu02(+)=substr(PA26,9)),fagent,customer,nation n1,nation n2" & _
            " where fa01(+)=substr(fno,1,8) and fa02(+)=substr(fno,9) and cu01(+)=substr(fno,1,8) and cu02(+)=substr(fno,9) and n1.na01(+)=fa10 and n2.na01(+)=cu10"
            
         If Option1(0).Value = True Then
            strExc(0) = strExc(0) & " order by FCPEmp,NaID,cp01,cp02,cp03,cp04"
         Else
            strExc(0) = strExc(0) & " order by CtMan,Fno,cp01,cp02,cp03,cp04"
         End If
         
         Set rsMain = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            blnTransaction = True
            m_strStartDBNo = "": m_strEndDBNo = ""
            With rsMain
               InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
               'Add by Sindy 2015/9/24
               'Modified by Lydia 2019/12/27
               'pub_OsPrinter = PUB_GetOsDefaultPrinter
               m_OsPrinter = PUB_GetOsDefaultPrinter
               PUB_SetOsDefaultPrinter Combo2.Text
               PUB_SetWordActivePrinter
               PUB_RestorePrinter Combo2.Text
               '2015/9/24 END
               
               Do While Not .EOF
                  'Added by Lydia 2020/08/17
                  If m_OptKind = "1" Then
                       'Transaction±Nµo¤å©M¦~ÃÒ¶O½Ð´Ú¨ç¤@¨Ö¥]¤J; ¦]¬°5/20ªºFCP762756µo¤åµ{¦¡¥¢±Ñ,¸g¹q¸£¤¤¤ß´ú¸Õm51¦sdn.pdf©ótyping2,³y¦¨¦³µo¤å¤é«oµL¹ê»Ú½Ð´Ú³æ
                  Else
                  'end 2020/08/17
                       cnnConnection.BeginTrans
                  End If 'Added by Lydia 2020/08/17
                  m_strCP01 = .Fields(0): m_strCP02 = .Fields(1): m_strCP03 = .Fields(2): m_strCP04 = .Fields(3)
                  m_strCP09 = .Fields(4): m_strCP10 = .Fields(5): m_strCP27 = .Fields(8)
                  stCtMan = .Fields("CtMan") 'Added by Morgan 2014/5/29
                  m_PA26 = "" & .Fields("PA26") 'Added by Morgan 2015/6/8
                  m_PA75 = "" & .Fields("PA75") 'Added by Morgan 2016/10/13 FC¥N²z¤H
                  m_NationID = "" & .Fields("NaID") 'Add By Sindy 2017/2/15 °êÄy¥N½X
                  m_CP148 = "" & .Fields("cp148") 'Added by Lydia 2019/10/08 (­Ó®×)¬O§_¯S®í½Ð´Ú³æ
                  'Added by Lydia 2020/08/17
                  m_CP60 = "" & .Fields("cp60") '½Ð´Ú³æ
                  m_eFlag = "" & .Fields("eFlag") '¬O§_e/E¤Æ
                  'end 2020/08/17
                  'Add by Morgan 2008/4/3
                  m_bolEmail = PUB_GetEMailFlag(m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04, IIf(m_strCP10 = "605", True, False), , m_bolPlusPaper)
                  'Added by Morgan 2014/6/3
                  If m_bolEmail = False Then
                     m_bolDNEmail = PUB_GetEMailFlag(m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04, IIf(m_strCP10 = "605", True, False), , m_bolDNPlusPaper, , True)
                  Else
                     m_bolDNEmail = m_bolEmail
                     m_bolDNPlusPaper = m_bolPlusPaper
                  End If
                  'end 2014/6/3
                  
                  'Add by Morgan 2009/10/20 +§PÂ_¬O§_EMail¦P®É±H¯È¥»
                  'Remove by Lydia 2020/08/17 «DE¤Æ¥X¯È¥»¤@¥÷; ­Y¦³«ü©w¥÷¼Æ>1«h´î¤@¥÷,«ü©w¥÷¼Æ=1«h¤£¥X¯È¥»
                  'If m_bolPlusPaper Then
                  '   m_iCopy = 0
                  'Else
                  '   m_iCopy = 1
                  'End If
                  ''end 2009/10/20
                  'end 2020/08/17
                  m_iCopy = 0 'E¤Æ¤£¥X¯È¥»
                  iCopy = 0 'Added by Lydia 2021/05/25 «ü©w¥÷¼ÆÂk¹s; ex. ¾ã§å¦~¶Oµo¤åFCP-040386ªº¥N²z¤HY23699«ü©w¦L4¥÷,¶Ç¤JPUB_GetCopySetting­n=0,¤~·|§ì¸ê®Æ
                  If m_bolEmail = False Or m_bolPlusPaper = True Then
                      m_iCopy = 1
                      '«DE¤Æ¥X¯È¥»¤@¥÷; ­Y¦³«ü©w¥÷¼Æ>1«h´î¤@¥÷,«ü©w¥÷¼Æ=1«h¤£¥X¯È¥»
                      PUB_GetCopySetting iCopy, m_strCP01, m_strCP02, m_strCP03, m_strCP04, m_strCP10, ET01
                      'Added by Lydia 2020/12/25 ­Y¬OE+±H¦L¤@¥÷,©T©w³£­n¦L; ex.Y53715000(E+±H)¥u¦L¤@¥÷
                      If m_bolEmail = True Or m_bolPlusPaper = True Then
                          If iCopy = 0 Then
                             m_iCopy = 1
                          ElseIf iCopy > 1 Then
                             m_iCopy = iCopy - 1
                          End If
                      Else
                      'end 2020/12/25
                          If iCopy = 1 Then
                              m_iCopy = 0
                          ElseIf iCopy > 1 Then
                              m_iCopy = iCopy - 1
                          End If
                      End If 'Added by Lydia 2020/12/23
                  End If
                  'end 2020/08/17
                  'Added by Lydia 2020/08/26 ³]©wDN¥X¯È¥»¡A¤~¯à¦P®É²£¥ÍPDFÀÉ©M¯È¥»
                  If m_bolEmail = False And m_iCopy > 0 Then
                      m_bolDNEmail = True
                      m_bolDNPlusPaper = True
                  End If
                  'end 2020/08/26
                  
                  'Add by Morgan 2004/6/25
                  Erase pa: m_bolNew = False: m_bol412 = False: stPA14 = "" & .Fields("PA14")
                  If Val(stPA14) = 0 Or Val(stPA14) >= 20040701 Then
                     m_bolNew = True
                     pa(1) = m_strCP01: pa(2) = m_strCP02: pa(3) = m_strCP03: pa(4) = m_strCP04
                     m_bol412 = PUB_Check412(pa)
                  End If
                  
                  If IsNull(.Fields(6)) Then
                     m_strCP16 = 0
                  Else
                     m_strCP16 = .Fields(6)
                  End If
                  If IsNull(.Fields(7)) Then
                     m_strCP17 = 0
                  Else
                     m_strCP17 = .Fields(7)
                  End If
                  'Add by Morgan 2007/9/5
                  If IsNull(.Fields("cp18")) Then
                     m_strCP18 = 0
                  Else
                     m_strCP18 = .Fields("cp18")
                  End If
                  
                  strTmp2 = m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04
                  
                  'Add by Morgan 2004/10/15 ¥N²z¤H Y20412 ªº©Ò¦³½Ð´Ú¨ç¹w³]P.S. Our debit note will be seprated from the rest of this letter and dealt with separately.
                  stNo = GetPrjPeopleNum6(m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04)
                  
                  'Modified by Morgan 2013/3/25 +X28186ªº½Ð´Ú¨ç¥[§PÂ_¥N²z¤HY53495®×¥ó¤~±a³Æµù
                  'm_stPS = PUB_GetDNPS(stNo, m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04, m_strCP10)
                  stNo1 = GetPrjPeopleNum1(m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04)
                  'Added by Lydia 2023/08/01
                  stNo2 = GetPrjPeopleNum2(m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04)
                  stNo3 = GetPrjPeopleNum3(m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04)
                  stNo4 = GetPrjPeopleNum4(m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04)
                  stNo5 = GetPrjPeopleNum5(m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04)
                  'end 2023/08/01
                  'Add by Lydia 2015/02/06 §ï¦¨ªí³æºûÅ@, ¦@¥Î¼Ò²ÕPUB_GetDebitNotePS
                  
'                  m_stPS = PUB_GetDNPS(stNo, m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04, m_strCP10, stNo1)
'                  'end 2013/3/25
'                  'Add by Morgan 2008/11/13 ¥Ó½Ð¤H X28186 ªº©Ò¦³½Ð´Ú¨ç¤]­n¹w³]
'                  If m_stPS = "" Then
'                     'Modified by Morgan 2013/3/25 ¥Ó½Ð¤H§ï¥Î¤W­±¤w³]©wªº stNo1 ÅÜ¼Æ¥BµL»Ý¦A¦Ò¼{¬O§_¦h¤H¥Ó½Ð(­ìX28186»Ý¨D¤w¤£¦s¦b)
'                     'stNo = GetPrjPeopleNum1(m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04)
'                     'If GetPrjPeopleNum2(m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04) = "" Then
'                     '   m_stPS = PUB_GetDNPS(stNo, m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04, m_strCP10)
'                     'End If
'                     m_stPS = PUB_GetDNPS(stNo1, m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04, m_strCP10)
'                     'end 2013/3/25
'                  End If
                  'End 2008/11/13
                  'Modified by Lydia 2023/08/01 ¼W¥[¥Ó½Ð¤H2~5
                  'm_stPS = PUB_GetDebitNotePS(m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04, m_strCP10, stNo, stNo1)
                  m_stPS = PUB_GetDebitNotePS(m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04, m_strCP10, stNo, stNo1 & "," & stNo2 & "," & stNo3 & "," & stNo4 & "," & stNo5)
                  'end 2015/02/06
                  If m_stPS <> "" Then
                     m_stPS = "P.S. " & m_stPS
                  End If
                  '2004/10/15 end
                  
                  'Modified by Lydia 2019/10/15 Sharon: ¥~±M°Q½×¨M©w,µo¤å¯S®í½Ð´Ú³æªº©w½Z§ï¨ì°h©Ó¿ì§@·~§¹¦¨«á,µ{§Ç¤H­û¦A­Ó®×¶]©w½Z; ¬°¤F©w½Z¤é´Á¯à°÷¤@­P
                  'If InStr(strTmp, strTmp2) = 0 Then
                  'Moeified by Lydia 2020/08/17 ¯S®í½Ð´Ú³æ¥i¥H¥X©w½Z
                  'If m_KeyCP09 <> "" And m_CP148 = "Y" Then
                  '    m_Case2cp148 = m_Case2cp148 & "," & .Fields("cp01") & "-" & .Fields("cp02") & IIf("" & .Fields("cp03") & .Fields("cp04") <> "000", "-" & .Fields("cp03") & "-" & .Fields("cp04"), "")
                  'ElseIf InStr(strTmp, strTmp2) = 0 Then
                  ''end 2019/10/15
                  If InStr(strTmp, strTmp2) = 0 Then
                  'end 2020/08/17
                     '¨ú±o©w½Z»y¤å
                     'Modify by Morgan 2006/5/25
                     'm_LetterLanguage = GetLetterLanguage(m_strCP01, m_strCP02, m_strCP03, m_strCP04)
                     m_LetterLanguage = PUB_GetLanguage(m_strCP01, m_strCP02, m_strCP03, m_strCP04, m_strCP10, "1")

                     '¨ú±o©w½Z»y¤å­^¤å¤§¤ÀÃþ
                     'Modify by Morgan 2004/10/12 ¥[¤é¤å
                     'If m_LetterLanguage = 2 Then
                     'Modified by Morgan 2016/10/13 ¥[¤¤¤å
                     'If m_LetterLanguage = 2 Or m_LetterLanguage = 3 Then
                        m_LetterKind = CetLetterKind(m_strCP01, m_strCP02, m_strCP03, m_strCP04, m_strCP10)
                     'End If
                     'end 2016/10/13
                     
                     Select Case m_strCP10
'                     Case ¥[µù°l¥[, ¥[µùÁp¦X
'                        Select Case m_LetterLanguage
'                           Case "1":    ' ¤¤¤å
'                              strTmp1 = "01"
'                           Case "2":    ' ­^¤å
'                              strTmp1 = "05"
'                           Case "3":    ' ¤é¤å
'                              strTmp1 = "06"
'                        End Select
'                        StartLetter ET01, strTmp1
'
'                        '·s¼W¦a§}±ø¦Cªí¸ê®Æ
'                        pub_AddressListSN = pub_AddressListSN + 1 '½Ð´Ú³æ²M³æ·|¥Î
'                        'Add by Morgan 2008/4/3 §PÂ_¬O§_²£¥Í¹q¤lÀÉ
'                        If m_bolEmail Then
'                           NowPrint m_strCP09, ET01, strTmp1, False, strUserNum, , , , , m_iCopy, , True, True
'                        Else
'                        'end 2008/4/3
'                           NowPrint m_strCP09, ET01, strTmp1, False, strUserNum
'                        End If
'
'                        If Not m_bolEmail Or m_bolPlusPaper Then
'                           If stLstFNo <> .Fields("FNo") Then 'Added by Morgan 2014/5/29 ¦P¦¬¥ó¤H¥u­n¦L¤@±i
'                              PUB_AddNewAddressList strUserNum, m_strCP01, m_strCP02, m_strCP03, m_strCP04, "" & pub_AddressListSN, "0", m_strCP10
'                           End If
'                        End If
'
'                        PUB_AddNewLetterList "¦~ÃÒ¶O½Ð´Ú¨ç", Me.Text1(5).Text & "-" & Me.Text1(6).Text, m_strCP01, m_strCP02, m_strCP03, m_strCP04, IIf(m_bolEmail, IIf(m_bolPlusPaper, "¢Ó", "¢í"), ""), stCtMan, .Fields("FNo") 'Modified by Morgan 2014/5/29 +²M³æe¤Æ¥[µù°O,ºÞ¨î¤H,¦¬¥ó¤H
'                        If AddACC1K0 = False Then
'                           cnnConnection.RollbackTrans
'                           MsgBox "·s¼W½Ð´Ú³æ¸ê®Æ¿ù»~ !", vbCritical
'                           Screen.MousePointer = vbDefault
'                           Exit Sub
'                        End If
                     Case »âÃÒ¤ÎÃº¦~¶O
                        'Add By Sindy 2015/9/23 ¦C¦LFCP©Ó¿ì³æ
                        'Modified by Lydia 2019/03/04 §ó´«Ãþ§O¥N¸¹;
                        'Call PUB_PrintFCPEmpBill(m_strCP01, m_strCP02, m_strCP03, m_strCP04, ET01, m_strCP09, , , "1")
                        'Added by Lydia 2019/10/08 ¦~¶Oµo¤å(¶Ç¤J¦¬¤å¸¹)-¯S®í½Ð´Ú³æ¡A¥u¦C¦L©w½Z
                        'Modified by Lydia 2020/08/17 ¤£¦L©Ó¿ì³æ§ïµoemail
                        'If m_KeyCP09 <> "" And m_CP148 = "Y" Then
                        '     m_Case2cp148 = m_Case2cp148 & "," & .Fields("cp01") & "-" & .Fields("cp02") & IIf("" & .Fields("cp03") & .Fields("cp04") <> "000", "-" & .Fields("cp03") & "-" & .Fields("cp04"), "")
                        'Else
                        ''end 2019/10/08
                        '     Call PUB_PrintFCPEmpBill(m_strCP01, m_strCP02, m_strCP03, m_strCP04, "06", m_strCP09, , , "1")
                        'End If 'end 2019/10/08
                        'If m_OptKind <> "3" Then '½Ð´Ú³qª¾¨ç-¦~ÃÒ¶O½Ð´Ú¨ç(³æµ§¥X½Ð´Ú³æ)¡G¥»§@·~¤£µoemail³qª¾¶i¦æ½Ð´Ú 'Remove by Lydia 2020/08/24
                            If m_OptMCrec <> "" Then  '¤H¤uEmailºûÅ@
                                 cnnConnection.Execute m_OptMCrec, intI
                            ElseIf Not (m_OptKind = "3" And m_CP60 <> "") Then 'Modified by Lydia 2020/08/24 ½Ð´Ú³qª¾¨ç-¦~ÃÒ¶O½Ð´Ú¨ç(³æµ§¥X½Ð´Ú³æ):­Y¦b³o¦¸²£¥Í½Ð´Ú³æ,¤~¥i¥Hµoemail³qª¾¶i¦æ½Ð´Ú
                                 'Modified by Lydia 2022/04/12  °Ï§O¾ã§å¦~¶Oµo¤å
                                 'Call PUB_GetFCPEmpMail("1", m_strCP09, m_eFlag, m_CP148, m_OptPAID, m_OptRecDate, m_OptMemo, "")
                                 Call PUB_GetFCPEmpMail(IIf(m_OptKind = "2", "3", "1"), m_strCP09, m_eFlag, m_CP148, m_OptPAID, m_OptRecDate, m_OptMemo, "")
                            End If
                        'End If
                        'end 2020/08/17
                        
                        'Modified by Morgan 2019/12/18 ²¾¨ì©w½Z¦C¦L«e¥ý·s¼W¡A¦]¬°©w½Z¤º®e­n§PÂ_¬O§_¦³¦P®É½Ð¤G®Ö926
                        
                        'Modified by Lydia 2020/08/17 °£¯S®í½Ð´Ú³æ,¬Ò­n¦L½Ð´Ú³æ
                        'If Not (m_KeyCP09 <> "" And m_CP148 = "Y") Then 'Added by Lydia 2019/10/08 ¦~¶Oµo¤å(¶Ç¤J¦¬¤å¸¹)-¯S®í½Ð´Ú³æ¡A¥u¦C¦L©w½Z
                        If m_CP148 <> "Y" Then
                            'Added by Lydia 2020/08/17 ¤w¦³½Ð´Ú³æ=>­«¦L
                            If m_CP60 <> "" Then
                                 Call AddPrintDN(m_CP60)
                            Else
                            'end 2020/08/17
                                If AddACC1K0 = False Then
                                   cnnConnection.RollbackTrans
                                   MsgBox "·s¼W½Ð´Ú³æ¸ê®Æ¿ù»~ !", vbCritical
                                   Screen.MousePointer = vbDefault
                                   Exit Sub
                                End If
                            End If 'Added by Lydia 2020/08/17
                        End If
                        
                        Select Case m_LetterLanguage
                           Case "1":    ' ¤¤¤å
                              strTmp1 = "01"
                           Case "2":    ' ­^¤å
                              'Modified by Morgan 2019/11/12 ­^¤å¦~¶O½Ð´Ú¨ç¤w¦X¨Ö
                              'Select Case m_LetterKind
                              '   Case "1":    '¤@¯ë
                              '      strTmp1 = "02"
                              '   Case "2":    '¦~¶O¦Û°Ê¥NÃº
                              '      strTmp1 = "03"
                              '   Case "4":    '³Ì«á¤@¦~
                              '      strTmp1 = "04"
                              'End Select
                              strTmp1 = "15"
                              'end2019/11/12
                              
                           Case "3":    ' ¤é¤å
                              strTmp1 = "06"
                              'Added by Lydia 2020/08/17 ¤w¦¬´Ú
                              'Mark by Lydia 2020/08/17 »âÃÒ(¤é¤å)¤w¦¬´Ú,­ì¥»³B²zª¬ªp¬°10,¨Ö¤J¤@¯ë06(by frm060306_2)
                              'If m_OptPAID <> "" Then
                              '    strTmp1 = "10"
                              'End If
                              ''end 2020/08/17
                              
                        End Select
                                                
                        'Add by Morgan 2004/6/25   ·sªk©w½Z
                        'Removed by Morgan 2019/11/12 ©w½Z¤w§R°£
                        'If m_bolNew = True And InStr("02,03,04", strTmp1) > 0 Then
                        '   strTmp1 = Format(Val(strTmp1) + 10, "00")
                        '   'Modify by Morgan 2004/7/20 ­YµL¤½§i¤é¥t¥X©w½Z
                        '   If Val(stPA14) = 0 Then strTmp1 = "15"
                        'End If
                        'end 2019/11/12
                        
                        StartLetter ET01, strTmp1
                        '·s¼W¦a§}±ø¦Cªí¸ê®Æ
                        pub_AddressListSN = pub_AddressListSN + 1 '½Ð´Ú³æ²M³æ·|¥Î
                        'Add by Morgan 2008/4/3 §PÂ_¬O§_²£¥Í¹q¤lÀÉ(Word)
                        'Modified by Lydia 2020/08/17 ¥X¯È¥»
                        'If m_bolEmail Then
                        If m_iCopy > 0 Then
                           NowPrint m_strCP09, ET01, strTmp1, False, strUserNum, , , , , m_iCopy, , True, True
                        Else
                        'end 2008/4/3
                           'Modified by Lydia 2020/08/16 «DE¤Æ¤]­n²£¥Í¹q¤lÀÉ
                           'NowPrint m_strCP09, ET01, strTmp1, False, strUserNum
                           NowPrint m_strCP09, ET01, strTmp1, False, strUserNum, , , , , , , , True
                        End If
                        
                        'Add By Sindy 2017/2/15 ¤é¥»°Ï­n²£¥ÍPDF¹q¤lÀÉ¦Ü­Ó®×¸ê®Æ§¨
                        'If m_bolEmail And Left(m_NationID, 3) = "011" Then   'Mark by Lydia 2020/08/17 ¢Ó¤Æ©M«D¢Ó¤Æ¬Ò²£¥Í¹q¤lÀÉ(Typing2)
                           strCaseNo = .Fields("CP01") & .Fields("CP02") & IIf(.Fields("CP03") & .Fields("CP04") <> "000", .Fields("CP03") & .Fields("CP04"), "")
                           'Modified by Lydia 2020/08/17 §ï¥Î®×¥ó©Ê½è
                           'strFileName = PUB_GetEFilePath(.Fields("CP01")) & "\" & .Fields("CP01") & "\" & Left(.Fields("CP02"), 3) & "\" & strCaseNo & "\" & strCaseNo & "_" & strSrvDate(1) & "." & IIf(Option1(0).Value = True, »âÃÒ¤ÎÃº¦~¶O, ¦~¶O) & ".CUS.PDF"
                           strFileName = PUB_GetEFilePath(.Fields("CP01")) & "\" & .Fields("CP01") & "\" & Left(.Fields("CP02"), 3) & "\" & strCaseNo & "\" & strCaseNo & "_" & strSrvDate(1) & "." & m_strCP10 & ".CUS.PDF"
                           'Modified by Sindy 2020/4/28 FCP-42850 ¥u²£¥Í«È¤á¨ç¹q¤lÀÉ,¤£­n³s¦P¥Ó½Ð®Ñ¤]©ñ¥X¨Ó
                           'Call PUB_PrintLetter(m_strCP09, , , True, strFullFileName)
                           'Modified by Lydia 2020/08/17 ¦³¥÷¼Æ,¤@¨Ö°õ¦æ¦C¦L
                           'Call PUB_PrintLetter(m_strCP09, , , True, strFullFileName, , True)
                           Call PUB_PrintLetter(m_strCP09, , , True, strFullFileName, IIf(m_iCopy = 0, False, True), True)
                           '2020/4/28 END
                           
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
                          
                        'Mark by Lydia 2020/08/17 ¢Ó¤Æ©M«D¢Ó¤Æ¬Ò²£¥Í¹q¤lÀÉ(Typing2)
                        'Else
                        ''2017/2/15 END
                        '   PUB_PrintLetter m_strCP09 'ª½±µ¦L¥X©w½Z Add By Sindy 2015/9/23
                        'End If
                        'end 2020/08/17
                        
                        If Not m_bolEmail Or m_bolPlusPaper Then
                           If stLstFNo <> .Fields("FNo") Then 'Added by Morgan 2014/5/29 ¦P¦¬¥ó¤H¥u­n¦L¤@±i
'                              'Add By Sindy 2015/9/21 ¤é¤å©w½Z¤~­n¦L¦a§}±ø
'                              If m_LetterLanguage = "3" Or Val(¥~±M¶}µ¡«H¨ç±Ò¥Î¤é) >= Val(strSrvDate(1)) Then
'                              '2015/9/21 END
                                 PUB_AddNewAddressList strUserNum, m_strCP01, m_strCP02, m_strCP03, m_strCP04, "" & pub_AddressListSN, "0", m_strCP10
'                              End If
                           End If
                        End If
                        
                        '·s¼W¾ã§å©w½Z¦C¦L²M³æ¸ê®ÆNvl(PA105, Nvl(PA76, Nvl(PA88
                        'Modified by Lydia 2020/08/17 ±Æ°£­Ó®×
                        'If m_KeyCP09 = "" Or Len(m_KeyCP09) > 9 Then  'Added by Lydia 2019/10/01 §PÂ_¡G¾ã§å½Ð´Ú¨ç ©Î ¦~¶O¾ã§åµo¤å
                        If (m_OptKind = "0" And opt1idx < 2) Or Len(m_KeyCP09) > 9 Then
                             PUB_AddNewLetterList "¦~ÃÒ¶O½Ð´Ú¨ç", Me.Text1(5).Text & "-" & Me.Text1(6).Text, m_strCP01, m_strCP02, m_strCP03, m_strCP04, IIf(m_bolEmail, IIf(m_bolPlusPaper, "¢Ó", "¢í"), ""), stCtMan, .Fields("FNo") 'Modified by Morgan 2014/5/29 +²M³æe¤Æ¥[µù°O,ºÞ¨î¤H,¦¬¥ó¤H
                        End If
                        

                     Case ¦~¶O
                        '93.10.4 MODIFY BY SONIA §PÂ_¬O§_¦~¶O³æµ§¤£¶], ¤£ºÞ½Ð´Ú¹ï¶H
                        'strExc(0) = "SELECT PA26,PA75,PA107 FROM PATENT WHERE " & ChgPatent(strTmp2)
                        'Memo by Lydia 2019/10/01 ¬O§_¦~¶O³æµ§¤£¶](PA107),¥Ø«e¼È¤£¨Ï¥Î,©Ò¥H¤£·|°õ¦æAddWorkF
                        strExc(0) = "SELECT PA26,Nvl(PA76, PA75),PA107 FROM PATENT WHERE " & ChgPatent(strTmp2)
                        '93.10.4 END
                        intI = 1
                        Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If IsNull(rsTemp1.Fields(2)) Then
                              'Modified by Morgan 2022/7/5 ¤¤¶¡¦¬¤å¦~¶O¥i¯à¨S¦³¥Ó½Ð¤H
                              'm_PA26 = rsTemp1.Fields(0)
                              m_PA26 = "" & rsTemp1.Fields(0)
                              'end 2022/7/5
                              strTxt(1) = ""
                              strTxt(2) = ""
                              If Not IsNull(rsTemp1.Fields(0)) Then strTxt(1) = rsTemp1.Fields(0)
                              If Not IsNull(rsTemp1.Fields(1)) Then strTxt(2) = rsTemp1.Fields(1)
                              If Not CU73FA40(strTxt(1), strTxt(2)) Then
                                 'Add By Sindy 2015/9/23 ¦C¦LFCP©Ó¿ì³æ
                                 'Modified by Lydia 2019/03/04 §ó´«Ãþ§O¥N¸¹;
                                 'Call PUB_PrintFCPEmpBill(m_strCP01, m_strCP02, m_strCP03, m_strCP04, ET01, m_strCP09, , , "2")
                                 'Added by Lydia 2019/10/08 ¦~¶Oµo¤å(¶Ç¤J¦¬¤å¸¹)-¯S®í½Ð´Ú³æ¡A¥u¦C¦L©w½Z
                                 'Modified by Lydia 2020/08/17 ¤£¦L©Ó¿ì³æ§ïµoemail
                                 'If m_KeyCP09 <> "" And m_CP148 = "Y" Then
                                 '     m_Case2cp148 = m_Case2cp148 & "," & .Fields("cp01") & "-" & .Fields("cp02") & IIf("" & .Fields("cp03") & .Fields("cp04") <> "000", "-" & .Fields("cp03") & "-" & .Fields("cp04"), "")
                                 'Else
                                 ''end 2019/10/08
                                 '     Call PUB_PrintFCPEmpBill(m_strCP01, m_strCP02, m_strCP03, m_strCP04, "06", m_strCP09, , , "2")
                                 'End If 'end 2019/10/08
                                 'If m_OptKind <> "3" Then  '½Ð´Ú³qª¾¨ç-¦~ÃÒ¶O½Ð´Ú¨ç(³æµ§¥X½Ð´Ú³æ)¡G¥»§@·~¤£µoemail³qª¾¶i¦æ½Ð´Ú  'Remove by Lydia 2020/08/24
                                    If m_OptMCrec <> "" Then  '¤H¤uEmailºûÅ@
                                         cnnConnection.Execute m_OptMCrec, intI
                                    ElseIf Not (m_OptKind = "3" And m_CP60 <> "") Then 'Modified by Lydia 2020/08/24 ½Ð´Ú³qª¾¨ç-¦~ÃÒ¶O½Ð´Ú¨ç(³æµ§¥X½Ð´Ú³æ):­Y¦b³o¦¸²£¥Í½Ð´Ú³æ,¤~¥i¥Hµoemail³qª¾¶i¦æ½Ð´Ú
                                         'Modified by Lydia 2022/04/12  °Ï§O¾ã§å¦~¶Oµo¤å
                                         'Call PUB_GetFCPEmpMail("1", m_strCP09, m_eFlag, m_CP148, m_OptPAID, m_OptRecDate, m_OptMemo, "")
                                         Call PUB_GetFCPEmpMail(IIf(m_OptKind = "2", "3", "1"), m_strCP09, m_eFlag, m_CP148, m_OptPAID, m_OptRecDate, m_OptMemo, "")
                                    End If
                                 'End If
                                 'end 2020/08/17
                                 
                                 Select Case m_LetterLanguage
                                    Case "1":    ' ¤¤¤å
                                       strTmp1 = "01"
                                    Case "2":    ' ­^¤å
                                       'Modified by Morgan 2019/11/5 ­^¤å¦~¶O½Ð´Ú¨ç¤w¦X¨Ö
                                       'Select Case m_LetterKind
                                       '   Case "1":    '¤@¯ë
                                       '      strTmp1 = "02"
                                       '   Case "2":    '¦~¶O¦Û°Ê¥NÃº
                                       '      strTmp1 = "03"
                                       '   Case "4":    '³Ì«á¤@¦~
                                       '      strTmp1 = "04"
                                       'End Select
                                       strTmp1 = "02"
                                       'end 2019/11/5
                                    Case "3":    ' ¤é¤å
                                       strTmp1 = "06"
                                       'Add by Morgan 2004/10/12 ¦~¶O¦Û°Ê¥NÃº
                                       If m_LetterKind = 2 Then
                                          strTmp1 = "09"
                                          'Mark by Lydia 2020/08/17 ¦~¶O(¤é¤å)¦Û°Ê¥NÃº+¤w¦¬´Ú,­ì¥»³B²zª¬ªp¬°11,¨Ö¤J10(by frm060306_2)
                                          'If m_OptPAID <> "" Then strTmp1 = "11" 'Added by Lydia 2020/08/17 ¤w¦¬´Ú
                                       'Added by Lydia 2020/08/17
                                       Else
                                          If m_OptPAID <> "" Then '¤w¦¬´Ú
                                              strTmp1 = "10"
                                          End If
                                       'end 2020/08/17
                                       End If
                                 End Select
                                 
                                 StartLetter ET01, strTmp1
                                 '·s¼W¦a§}±ø¦Cªí¸ê®Æ
                                 pub_AddressListSN = pub_AddressListSN + 1 '½Ð´Ú³æ²M³æ·|¥Î
                                 'Add by Morgan 2008/4/3 §PÂ_¬O§_²£¥Í¹q¤lÀÉ(Word)
                                 'Modified by Lydia 2020/08/17 ¥X¯È¥»
                                 'If m_bolEmail Then
                                 If m_iCopy > 0 Then
                                    NowPrint m_strCP09, ET01, strTmp1, False, strUserNum, , , , , m_iCopy, , True, True
                                 Else
                                 'end 2008/4/3
                                    'Modified by Lydia 2020/08/17 «DE¤Æ¤]­n²£¥Í¹q¤lÀÉ
                                    'NowPrint m_strCP09, ET01, strTmp1, False, strUserNum
                                    NowPrint m_strCP09, ET01, strTmp1, False, strUserNum, , , , , , , , True
                                 End If
                                 
                                 'Add By Sindy 2017/2/15 ¤é¥»°Ï­n²£¥ÍPDF¹q¤lÀÉ¦Ü­Ó®×¸ê®Æ§¨
                                 'If m_bolEmail And Left(m_NationID, 3) = "011" Then 'Mark by Lydia 2020/08/17 ¢Ó¤Æ©M«D¢Ó¤Æ¬Ò²£¥Í¹q¤lÀÉ(Typing2)
                                    strCaseNo = .Fields("CP01") & .Fields("CP02") & IIf(.Fields("CP03") & .Fields("CP04") <> "000", .Fields("CP03") & .Fields("CP04"), "")
                                    'Modified by Lydia 2020/08/17 §ï¥Î®×¥ó©Ê½è
                                    'strFileName = PUB_GetEFilePath(.Fields("CP01")) & "\" & .Fields("CP01") & "\" & Left(.Fields("CP02"), 3) & "\" & strCaseNo & "\" & strCaseNo & "_" & strSrvDate(1) & "." & IIf(Option1(0).Value = True, »âÃÒ¤ÎÃº¦~¶O, ¦~¶O) & ".CUS.PDF"
                                    strFileName = PUB_GetEFilePath(.Fields("CP01")) & "\" & .Fields("CP01") & "\" & Left(.Fields("CP02"), 3) & "\" & strCaseNo & "\" & strCaseNo & "_" & strSrvDate(1) & "." & m_strCP10 & ".CUS.PDF"
                                    'Modified by Sindy 2020/4/28 FCP-42850 ¥u²£¥Í«È¤á¨ç¹q¤lÀÉ,¤£­n³s¦P¥Ó½Ð®Ñ¤]©ñ¥X¨Ó
                                    'Call PUB_PrintLetter(m_strCP09, , , True, strFullFileName)
                                    'Modified by Lydia 2020/08/17 ¦³¥÷¼Æ,¤@¨Ö°õ¦æ¦C¦L
                                    'Call PUB_PrintLetter(m_strCP09, , , True, strFullFileName, , True)
                                    Call PUB_PrintLetter(m_strCP09, , , True, strFullFileName, IIf(m_iCopy = 0, False, True), True)
                                    '2020/4/28 END
                                    
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

                                 'Mark by Lydia 2020/08/17 ¢Ó¤Æ©M«D¢Ó¤Æ¬Ò²£¥Í¹q¤lÀÉ(Typing2)
                                 'Else
                                 ''2017/2/15 END
                                  '  PUB_PrintLetter m_strCP09 'ª½±µ¦L¥X©w½Z Add By Sindy 2015/9/23
                                 'End If
                                 'end 2020/08/17
                                 
                                 If Not m_bolEmail Or m_bolPlusPaper Then
                                    If stLstFNo <> .Fields("FNo") Then 'Added by Morgan 2014/5/29 ¦P¦¬¥ó¤H¥u­n¦L¤@±i
'                                       'Add By Sindy 2015/9/21 ¤é¤å©w½Z¤~­n¦L¦a§}±ø
'                                       If m_LetterLanguage = "3" Or Val(¥~±M¶}µ¡«H¨ç±Ò¥Î¤é) >= Val(strSrvDate(1)) Then
'                                       '2015/9/21 END
                                          PUB_AddNewAddressList strUserNum, m_strCP01, m_strCP02, m_strCP03, m_strCP04, "" & pub_AddressListSN, "0", m_strCP10
'                                       End If
                                    End If
                                 End If
                                 
                                 '·s¼W¾ã§å©w½Z¦C¦L²M³æ¸ê®Æ
                                 'Add by Morgan 2010/5/3 ¦~¶O¥N²z¤H¬° Y51774ªÌ²M³æ¥»©Ò®×¸¹«e¥[¡µ(ªþ±H¥ó²M³æ)
                                 strExc(2) = ""
                                 If m_strCP10 = "605" Then
                                    strExc(1) = PUB_GetReceiver(m_strCP01, m_strCP02, m_strCP03, m_strCP04, "605")
                                    If strExc(1) = "Y51774000" Then
                                       strExc(2) = "¡µ"
                                    End If
                                 Else
                                    strExc(2) = ""
                                 End If
                                 'end 2010/5/3
                                 
                                '·s¼W¾ã§å©w½Z¦C¦L²M³æ¸ê®Æ
                                 'Modified by Morgan 2013/4/18
                                 'PUB_AddNewLetterList "¦~ÃÒ¶O½Ð´Ú¨ç", Me.Text1(5).Text & "-" & Me.Text1(6).Text, m_strCP01, m_strCP02, m_strCP03, m_strCP04, strExc(2)
                                 'Modified by Lydia 2020/08/17 ±Æ°£­Ó®×
                                 'If m_KeyCP09 = "" Or Len(m_KeyCP09) > 9 Then  'Added by Lydia 2019/10/01 §PÂ_¡G¾ã§å½Ð´Ú¨ç ©Î ¦~¶O¾ã§åµo¤å
                                 If (m_OptKind = "0" And opt1idx < 2) Or Len(m_KeyCP09) > 9 Then
                                     PUB_AddNewLetterList "¦~ÃÒ¶O½Ð´Ú¨ç", Me.Text1(5).Text & "-" & Me.Text1(6).Text, m_strCP01, m_strCP02, m_strCP03, m_strCP04, strExc(2) & IIf(m_bolEmail, IIf(m_bolPlusPaper, "¢Ó", "¢í"), ""), stCtMan, .Fields("FNo") 'Modified by Morgan 2014/5/29 +²M³æe¤Æ¥[µù°O,ºÞ¨î¤H,¦¬¥ó¤H
                                 End If
                                 
                                 'Modified by Lydia 2020/06/26 °£¯S®í½Ð´Ú³æ,¬Ò­n¦L½Ð´Ú³æ
                                 'If Not (m_KeyCP09 <> "" And m_CP148 = "Y") Then 'Added by Lydia 2019/10/08 ¦~¶Oµo¤å(¶Ç¤J¦¬¤å¸¹)-¯S®í½Ð´Ú³æ¡A¥u¦C¦L©w½Z
                                 If m_CP148 <> "Y" Then
                                        'Added by Lydia 2020/08/17 ¤w¦³½Ð´Ú³æ=>­«¦L
                                        If m_CP60 <> "" Then
                                             Call AddPrintDN(m_CP60)
                                        Else
                                        'end 2020/08/17
                                            If AddACC1K0 = False Then
                                               cnnConnection.RollbackTrans
                                               MsgBox "·s¼W½Ð´Ú³æ¸ê®Æ¿ù»~ !", vbCritical
                                               Screen.MousePointer = vbDefault
                                               Exit Sub
                                            End If
                                        End If 'Added by Lydia 2020/08/17
                                 End If
                              Else
                                 '¦~¶O³æµ§¤£¶],¦L²M³æ¤ÎD/N²M³æ
                                 If AddWorkF = False Then
                                    cnnConnection.RollbackTrans
                                    MsgBox "·s¼W¦~¶O¾ã§å¤u§@ÀÉ¸ê®Æ¿ù»~ !", vbCritical
                                    Screen.MousePointer = vbDefault
                                    Exit Sub
                                 End If
                              End If
                           Else
                              '¦~¶O³æµ§¤£¶],¦L²M³æ¤ÎD/N²M³æ
                              If AddWorkF = False Then
                                cnnConnection.RollbackTrans
                                 MsgBox "·s¼W¦~¶O¾ã§å¤u§@ÀÉ¸ê®Æ¿ù»~ !", vbCritical
                                 Screen.MousePointer = vbDefault
                                 Exit Sub
                              End If
                           End If
                        End If
                     End Select
                  End If
                  stLstFNo = .Fields("FNo") 'Added by Morgan 2014/5/29
                  
                  'Added by Lydia 2020/08/17 ¥~³¡©I¥s:¬O§_§¹¦¨
                  If m_OptKind = "1" Then
                      m_bTransOK = True
                  Else
                  'end 2020/08/17
                      cnnConnection.CommitTrans
                  End If 'Added by Lydia 2020/08/17
                  
                  'Added by Lydia 2019/10/18 Sharon: §ï¦¨¤@­Ó®×¤l¦L§¹©Ó¿ì³æ->½Ð´Ú©w½Z->±b³æ¡A±µµÛ¦A¦L¤U¤@­Ó®×¸¹; ¦]¬°¥H«eªº±b³æ»Ý­n®M¦³«HÀYªº¯È±i¡A©Ò¥H¶°¤¤¦b³Ì«á¤@°_¦C¦L¡A²{¦b±b³æ³£¬Oª½±µ¦L«HÀY¡C
                  '¦C¦L½Ð´Ú³æ
                  PUB_PrintDebitNote strUserNum, Me.Combo2.Text, IIf(m_KeyCP09 <> "", Me.Name, "")
                  '§R°£½Ð´Ú³æ¦Cªí¸ê®Æ
                  PUB_DeleteDebitNoteList strUserNum
                  'end 2019/10/18
                  .MoveNext
               Loop
               
               'Add by Sindy 2015/9/24
               'Modified by Lydia 2019/12/27
               'PUB_SetOsDefaultPrinter pub_OsPrinter
               PUB_SetOsDefaultPrinter m_OsPrinter
               PUB_RestorePrinter strPrinter2
               '2015/9/24 END
            End With
            blnTransaction = False
            
            'Modified by Morgan 2013/4/17 §ï¬ö¿ý¦b¸ê®Æ®w,§_«h´«¹q¸£©Î¨Ï¥ÎªÌ·|Åª¤£¨ì
            'SaveSetting "TAIE", "FCP", "DATE71", Text1(5).Text
            'SaveSetting "TAIE", "FCP", "DATE72", Text1(6).Text
            'Modified by Lydia 2020/08/17
            'If m_KeyCP09 = "" Then 'Added by Lydia 2019/10/01 §PÂ_«D¦~¶Oµo¤å
            If m_OptKind = "0" And opt1idx < 2 Then
                PUB_SaveLastDate Me.Name, opt1idx & "DATE71", Text1(5).Text
                PUB_SaveLastDate Me.Name, opt1idx & "DATE72", Text1(6).Text
                'end 2013/4/17
            End If
            
            ProcessPrint   '¦C¦L½Ð´Ú³æ
            Screen.MousePointer = vbDefault
            'Added by Lydia 2019/10/08
            If m_KeyCP09 <> "" Then
                If m_Case2cp148 <> "" Then
                     m_Case2cp148 = Replace(m_Case2cp148, ",", vbCrLf)
                     MsgBox "¯S®í½Ð´Ú³æ¡G" & m_Case2cp148 & vbCrLf & "»Ý°h©Ó¿ì½Ð´Ú¡I", vbInformation, "µo¤å¦Û°Ê²£¥Í±b³æ"
                End If
           
            Else  '§PÂ_«D¦~¶Oµo¤å
            'end 2019/10/08
                MsgBox "¦C¦Lµ²§ô !", vbInformation
                If (Option1(0).Value = True Or Option1(1).Value = True) Then   'Added by Lydia 2020/08/17  ±Æ°£­Ó®×
                'Modified by Morgan 2013/4/17 §ï¬ö¿ý¦b¸ê®Æ®w,§_«h´«¹q¸£©Î¨Ï¥ÎªÌ·|Åª¤£¨ì
                'Label2(0).Caption = GetSetting("TAIE", "FCP", "DATE71", "")
                'Label2(1).Caption = GetSetting("TAIE", "FCP", "DATE72", "")
                    Label2(0).Caption = PUB_GetLastDate(Me.Name, opt1idx & "DATE71")
                    Label2(1).Caption = PUB_GetLastDate(Me.Name, opt1idx & "DATE72")
                'end 2013/4/17
                End If 'Added by Lydia 2020/08/17
            End If

            'Modified by Lydia 2020/08/17 ±Æ°£­Ó®× And opt1idx < 2
            If m_KeyCP09 = "" And opt1idx < 2 Then   'Added by Lydia 2019/10/01 §PÂ_«D¦~¶Oµo¤å
                Text1(5).Text = ""
                Text1(6).Text = ""
                Text1(5).SetFocus
                
                For i = 0 To List1.ListCount - 1
                   List1.RemoveItem 0
                Next
            'Added by Lydia 2019/10/01
            'Else
            '    Call cmdok_Click(1)  '¸õ¥X
            End If
            'end 2019/10/01
         Else
            InsertQueryLog (0) 'Add By Sindy 2010/12/7
            MsgBox "µL²Å¦X±ø¥ó¤§¸ê®Æ¥i¦C¦L !", vbInformation
            Screen.MousePointer = vbDefault
         End If
      Case 1 'µ²§ô
         Me.Enabled = False
         Unload Me
   End Select

Exit Sub
ErrorHandler:
    If blnTransaction = True Then cnnConnection.RollbackTrans
    If Err.Number <> 0 Then MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub ProcessPrint()

   Dim rsTmp As ADODB.Recordset
   '¥ý¦C¦C¦L¦~¶O¾ã§å©ú²Ó,¦A¦C¦L½Ð´Ú³æ
   If m_AddWorkFCount > 0 Then
      MsgBox "«ö½T©w«á¶}©l¦C¦L¦~¶O¾ã§å©ú²Ó¡A½Ð§ó´«¯È±i!", vbOKOnly + vbInformation, "¦C¦L¦~¶O¾ã§å©ú²Ó"
      PrintWorkF
   End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ¨ú±o­n¦LªíªººØÃþ
' Input : strPA01 ==> ¥»©Ò®×¸¹¨t²ÎÃþ§O
'         strPA02 ==> ¥»©Ò®×¸¹¬y¤ô¸¹
'         strPA03 ==> ¥»©Ò®×¸¹
'         strPA04 ==> ¥»©Ò®×¸¹
' Output : 1 = ªí¤@¯ë
'          2 = ªí¦~¶O¦Û°Ê¥NÃº
'          3 = ªí°l¥[Áp¦X
'          4 = ªí³Ì«á¤@¦~
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CetLetterKind(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String, ByVal strCP10 As String) As Integer
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsTmp As ADODB.Recordset
   Dim rsSubTmp As ADODB.Recordset
   Dim nKind As Integer
   Dim strFA01 As String
   Dim strFA02 As String
   Dim strCU01 As String
   Dim strCU02 As String
   nKind = 0
   
   ' PA03 <> 0 ®É¬°°l¥[Áp¦X
   If m_strCP03 <> "0" Then
      nKind = 3
      GoTo EXITSUB
   End If
   ' PA03 <> 0 ®É¬°°l¥[Áp¦X
   If m_strCP10 = "602" Or m_strCP10 = "603" Then
      nKind = 3
      GoTo EXITSUB
   End If
   '¬O§_³Ì«á¤@¦~
   'Modify by Morgan 2005/9/29 ¥[¥»©Ò´Á­­
   'Modify By Sindy 2021/4/27 + ,MAX(NP23)
   strSql = "SELECT MAX(NP09),MAX(NP08),MAX(NP23) FROM NEXTPROGRESS " & _
            "WHERE NP02 = '" & m_strCP01 & "' AND " & _
                  "NP03 = '" & m_strCP02 & "' AND " & _
                  "NP04 = '" & m_strCP03 & "' AND " & _
                  "NP05 = '" & m_strCP04 & "' AND NP07 = '605' AND NP06 IS NULL"
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount = 1 And IsNull(rsTmp.Fields(0)) Then
      nKind = 4
   End If
    'Modify By Cheng 2004/03/15
    'Á×§KNullªº¿ù»~
'   m_strNP09 = rsTmp.Fields(0)
   m_strNP09 = "" & rsTmp.Fields(0)
   m_strNP08 = "" & rsTmp.Fields(1)
   m_strNP23 = "" & rsTmp.Fields(2) 'Add By Sindy 2021/4/27
    'End
   Set rsTmp = Nothing
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & m_strCP01 & "' AND " & _
                  "PA02 = '" & m_strCP02 & "' AND " & _
                  "PA03 = '" & m_strCP03 & "' AND " & _
                  "PA04 = '" & m_strCP04 & "' "
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      m_strPA09 = "": m_strPA08 = ""
      m_strPA72 = "": m_strPA73 = ""
      If Not IsNull(rsTmp.Fields("PA08")) Then m_strPA08 = rsTmp.Fields("PA08")
      If Not IsNull(rsTmp.Fields("PA09")) Then m_strPA09 = rsTmp.Fields("PA09")
      If Not IsNull(rsTmp.Fields("PA72")) Then m_strPA72 = rsTmp.Fields("PA72")
      If Not IsNull(rsTmp.Fields("PA73")) Then m_strPA73 = rsTmp.Fields("PA73")
      m_strPA10 = "" & rsTmp.Fields("PA10") 'Added by Morgan 2019/8/6
      m_strPA25 = "" & rsTmp.Fields("PA25") 'Added by Morgan 2019/8/6
      If nKind = 4 Then
         rsTmp.Close
         GoTo EXITSUB
      Else
         ' PA70 = Y ®É¬°¦Û°Ê¥NÃº   '¦¹³B¥u³qª¾«áÄò¦~¶O¬O§_¦Û°Ê¥NÃº, ©Ò¥H¤£ºÞ»âÃÒ¬O§_¦Û°Ê¥NÃº
         If IsNull(rsTmp.Fields("PA70")) = False Then
            If rsTmp.Fields("PA70") = "Y" Then
               nKind = 2
               rsTmp.Close
               GoTo EXITSUB
            End If
         End If
         ' PA75 ¨ú¥N²z¤HÀÉ
         If IsNull(rsTmp.Fields("PA75")) = False Then
            If IsEmptyText(rsTmp.Fields("PA75")) = False Then
               If Len(rsTmp.Fields("PA75")) > 8 Then
                  strFA01 = Mid(rsTmp.Fields("PA75"), 1, 8)
                  strFA02 = Mid(rsTmp.Fields("PA75"), 9, 1)
               Else
                  strFA01 = rsTmp.Fields("PA75") & String(8 - Len(rsTmp.Fields("PA75")), "0")
                  strFA02 = "0"
               End If
               strSubSQL = "SELECT * FROM FAGENT " & _
                           "WHERE FA01 = '" & strFA01 & "' AND " & _
                                 "FA02 = '" & strFA02 & "' "
               Set rsSubTmp = New ADODB.Recordset
               rsSubTmp.CursorLocation = adUseClient
               rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
               If rsSubTmp.RecordCount > 0 Then
                  If IsNull(rsSubTmp.Fields("FA41")) = False Then
                     If rsSubTmp.Fields("FA41") = "Y" Then
                        nKind = 2
                        rsSubTmp.Close
                        rsTmp.Close
                        GoTo EXITSUB
                     End If
                  End If
               End If
               rsSubTmp.Close
               
               nKind = 1
               rsTmp.Close
               GoTo EXITSUB
            End If
         End If
         ' ¨ú«È¤áÀÉ
         If IsNull(rsTmp.Fields("PA26")) = False Then
            If IsEmptyText(rsTmp.Fields("PA26")) = False Then
               If Len(rsTmp.Fields("PA26")) > 8 Then
                  strCU01 = Mid(rsTmp.Fields("PA26"), 1, 8)
                  strCU02 = Mid(rsTmp.Fields("PA26"), 9, 1)
               Else
                  strCU01 = rsTmp.Fields("PA26") & String(8 - Len(rsTmp.Fields("PA26")), "0")
                  strCU02 = "0"
               End If
               strSubSQL = "SELECT * FROM CUSTOMER " & _
                           "WHERE CU01 = '" & strCU01 & "' AND " & _
                                 "CU02 = '" & strCU02 & "' "
               Set rsSubTmp = New ADODB.Recordset
               rsSubTmp.CursorLocation = adUseClient
               rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
               If rsSubTmp.RecordCount > 0 Then
                  If IsNull(rsSubTmp.Fields("CU74")) = False Then
                     If rsSubTmp.Fields("CU74") = "Y" Then
                        nKind = 2
                        rsSubTmp.Close
                        rsTmp.Close
                        GoTo EXITSUB
                     End If
                  End If
               End If
               rsSubTmp.Close
               
               nKind = 1
               rsTmp.Close
               GoTo EXITSUB
            End If
         End If
      End If
   End If
   
EXITSUB:
   If nKind = 0 Then nKind = 1
   Set rsSubTmp = Nothing
   Set rsTmp = Nothing
   CetLetterKind = nKind
End Function

Private Function AddACC1K0() As Boolean

   Dim m_strSerialNo As String '½Ð´Ú³æ¸¹
   Dim strAgentNo As String '¥N²z¤H½s¸¹
   Dim strPrintCust  As String '¬O§_¦C¦L¥Ó½Ð¤H
   Dim dblUSRate As Double '¬üª÷¶×²v
   Dim rsA As New ADODB.Recordset
   Dim StrSQLa As String
   Dim strA1K27 As String '¦C¦L¹ï¶H
   Dim strA1K28 As String '½Ð´Ú¹ï¶H
   Dim strA1K05 As String '½Ð´Ú³æ³Æµù Add by Morgan 2011/4/18
   Dim strDisc As String '§é¦©
   Dim str926CP09 As String, str926CP16 As String, str926Disc As String 'Added by Morgan 2019/12/16
   Dim bol926In605 As Boolean 'Added by Morgan 2025/8/11 'Added by Morgan 2025/8/11 »âÃÒ½Ð´Úª÷ÃB¬O§_§t¤G®Ö
   Dim strA1K02 As String  'Added by Morgan 2025/9/11

   AddACC1K0 = False
   'µL¶O¥Î¤Î³W¶O®É, ¥H¤wÃº¦~«×¤§³Ì«á¤@¦~­pºâªA°È¶O¤Î³W¶O,¨Ã§ó·sµo¤å¤é¦^PA73
   If (m_strCP16 = 0 Or m_strCP17 = 0) And (m_strCP10 <> "602" And m_strCP10 <> "603") Then
      GetPatentYearFee
   End If
   'ÀË¬d¬O§_¬°¦¬´Ú«á¿ì®×
   'Modified by Morgan 2017/8/2 ¥ý¨ú®ø§ï¤H¤u±±ºÞ,¬yµ{­n­×§ï--David
   'm_PayBefore = CheckPayBefore
   m_PayBefore = "N"
   'end 2017/8/2
   If m_PayBefore = "N" Then
      '¶}©l·s¼W°ê¥~½Ð´Ú¸ê®Æ
      '1:¥ý¥H"X"§ìACC1R0¤§°ê¥~½Ð´Ú³æªº¦Û°Ê½s¸¹, ¨Ã§ó·s¨ä¬y¤ô¸¹
      adoTaie.Execute "update acc1r0 set a1r04 = a1r04 where a1r01 = 'X'" 'Added by Morgan 2018/10/24 ­n¥ýÂê¦í¡A§_«h¥i¯à·|»P¨ä¥L¦P®É·s¼W½Ð´Ú³æªº¨Ï¥ÎªÌ¨ú¨ì¬Û¦P³æ¸¹ Ex:X10715838
      m_strSerialNo = AccAutoNo(MsgText(815), 5)
      AccSaveAutoNo MsgText(815), Right(m_strSerialNo, 5)
      '2:·s¼WACC1K0
      '¥N²z¤H½s¸¹
      strAgentNo = PUB_GetA1K03(m_strCP01, m_strCP02, m_strCP03, m_strCP04)
      '¦C¦L¹ï¶H
      strA1K27 = PUB_GetA1K27(m_strCP01, m_strCP02, m_strCP03, m_strCP04, m_strCP10)
      If strA1K27 = "" Then strA1K27 = strAgentNo
      '½Ð´Ú¹ï¶H
      strA1K28 = PUB_GetA1K28(m_strCP01, m_strCP02, m_strCP03, m_strCP04, m_strCP10)
      If strA1K28 = "" Then strA1K28 = strAgentNo
      'Modify by Morgan 2004/12/16 §ï³W«h
      'D/N¬O§_¦C¦L¥Ó½Ð¤H
      'strPrintCust = PUB_GetA1K04(m_strCP01, m_strCP02, m_strCP03, m_strCP04)
      strPrintCust = PUB_GetA1K04(m_strCP01, m_strCP02, m_strCP03, m_strCP04, strA1K28, m_strCP10)
      '2004/12/16 end
      
   '   dblUSRate = GetUSRate '¬üª÷¶×²v

        'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
         Dim strA1K33 As String, strA1K18 As String
         'Modify By Sindy 2016/11/30
         'strA1K33 = PUB_GetInitCurrPrintType(m_strCP01, strA1K28, strA1K18, dblUSRate)
         'Modified by Morgan 2018/4/27 +strA1K27
         strA1K33 = PUB_GetInitCurrPrintType(m_strCP01, strA1K28, strA1K18, dblUSRate, m_strCP02, m_strCP03, m_strCP04, strA1K27)
         '2016/11/30 END
        
      'Add by Morgan 2011/4/18
      'Modified by Morgan 2014/12/3
      'strA1K05 = PUB_GetDNRemark(strA1K28)
      strA1K05 = PUB_GetDNRemark(strA1K28, m_strCP01, m_strCP02, m_strCP03, m_strCP04)
      'end 2014/12/3
      
      'Modified by Morgan 2014/3/10 §PÂ_¬O§_¦³¯S®í½Ð´Úª÷ÃB
      'Modified by Morgan 2015/6/8
      'If PUB_GetSpecInvoiceFee(strA1K28, "FCP", "601", strSrvDate(2), dblUSRate, strExc(1)) = True Then
      'Modified by Morgan 2017/9/5 +Y54179000
      'Modified by Morgan 2024/11/14 ©T©w³ø»ù¥x¹ô³£§ïµL±ø¥ó¶i¦ì,§_«h·í±Ë¥h«á¦AÂà¦^¬üª÷¤S±Ë¥h´N·|®t1¬ü¤¸. Ex: 2¬ü¤¸*29.24=58.48¥x¹ô-> 58¥x¹ô/29.24=1.98¬ü¤¸->1¬ü¤¸
      'Modified by Morgan 2024/11/14 +Y27696 Bobst Mex SA»âÃÒ¤Î¤G®Ö±b³æ¤§ªA°È¶O³]©w¬°©T©wª÷ÃB(»âÃÒ¤ÎÃº¦~¶OªA°È¶O¡G170¬üª÷,¤G¦¸®Ö¹ïªA°È¶O¡G135¬üª÷,ªA°È¶OÁ`ª÷ÃB¡G305¬üª÷)
      'Modified by Morgan 2025/8/8 §ï¼g¦¨¨ç¼Æ¤½¥Î
      'If strA1K28 = "Y54179000" Or (m_strCP10 = »âÃÒ¤ÎÃº¦~¶O And (strA1K28 = "Y45814010" Or Left(m_PA26, 6) = "X48310" Or strA1K28 = "Y27696000")) Then
      '   strExc(2) = PUB_GetUSXRate_1(strSrvDate(2), "USD")
      '   'BASF »âÃÒ©T©w½Ð´Úª÷ÃB USD$180(§t®Ö¹ï¤w­ã±M§Q)
      '   If strA1K28 = "Y45814010" Then
      '       'strExc(1) = Round(180 * Val(strExc(2)))
      '       strExc(1) = UInt(180 * Val(strExc(2)))
      '   'SYNGENTA »âÃÒ©T©w½Ð´Úª÷ÃB USD$95
      '   ElseIf Left(m_PA26, 6) = "X48310" Then
      '       'strExc(1) = Round(95 * Val(strExc(2)))
      '       strExc(1) = UInt(95 * Val(strExc(2)))
      '   'Added by Morgan 2017/9/5
      '   'Y54179000 Longitude Licensing Ltd.
      '   ElseIf strA1K28 = "Y54179000" Then
      '      '»âÃÒ¤Î²Ä¤@¦~¦~¶O½Ð´Úª÷ÃB¬°USD72
      '      If m_strCP10 = »âÃÒ¤ÎÃº¦~¶O Then
      '         'strExc(1) = Round(72 * Val(strExc(2)))
      '         strExc(1) = UInt(72 * Val(strExc(2)))
      '      Else
      '         intI = Val(Mid(m_strPA72, InStrRev(m_strPA72, ",") + 1))
      '         '10-20¦~ USD120
      '         If intI >= 10 Then
      '            'strExc(1) = Round(120 * Val(strExc(2)))
      '            strExc(1) = UInt(120 * Val(strExc(2)))
      '         '7-9¦~ USD84
      '         ElseIf intI >= 7 Then
      '            'strExc(1) = Round(84 * Val(strExc(2)))
      '            strExc(1) = UInt(84 * Val(strExc(2)))
      '         '4-6¦~ USD65
      '         ElseIf intI >= 4 Then
      '            'strExc(1) = Round(65 * Val(strExc(2)))
      '            strExc(1) = UInt(65 * Val(strExc(2)))
      '         '1-3¦~ USD60
      '         Else
      '            'strExc(1) = Round(60 * Val(strExc(2)))
      '            strExc(1) = UInt(60 * Val(strExc(2)))
      '         End If
      '      End If
      '   'end 2017/9/5
      '   'Added by Morgan 2024/11/14 +Y27696 Bobst Mex SA»âÃÒ¤Î¤G®Ö±b³æ¤§ªA°È¶O³]©w¬°©T©wª÷ÃB(»âÃÒ¤ÎÃº¦~¶OªA°È¶O¡G170¬üª÷,¤G¦¸®Ö¹ïªA°È¶O¡G135¬üª÷,ªA°È¶OÁ`ª÷ÃB¡G305¬üª÷)
      '   ElseIf strA1K28 = "Y27696000" Then
      '      strExc(1) = UInt(170 * Val(strExc(2)))
      '   'end 2024/11/14
      '   End If
      Call PUB_Get601605SpecDN(m_strCP10, m_strCP01, m_strCP02, m_strCP03, m_strCP04, m_strPA08, m_strPA72, m_PA26, m_PA75, strA1K28, strExc(1), str926CP09, str926CP16, str926Disc, bol926In605)
      strA1K02 = strSrvDate(2)
      'Added by Morgan 2025/9/11
      '½Ð´Ú¹ï¶H¬°Y56199 Coupang Corp.ªº©Ò¦³±b³æ¡A½Ð´Ú¤é¬Ò¬°·í¤ë16¸¹--Kahn
      If strA1K28 = "Y56199000" And Right(strA1K02, 2) > "16" Then
         strA1K02 = (Val(strA1K02) \ 100) & "16"
         strA1K02 = TransDate(CompDate(1, 1, strA1K02), 1)
      End If
      'end 2025/9/11
      
      If Val(strExc(1)) > 0 Then
      'end 2025/88/8
         m_strCP16 = Val(m_strCP17) + Val(strExc(1))
         m_strCP18 = Round(Val(strExc(1)) / 1000, 1)
      'end 2015/6/8
      
        'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
'         strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K05,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21) " & _
                  "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",'" & strA1K05 & "',0,NULL," & Val(m_strCP17) & "," & dblUSRate & "," & (Val(strExc(1)) + Val(m_strCP17)) & ",NULL,'" & m_strCP01 & "','" & m_strCP02 & "','" & m_strCP03 & "','" & m_strCP04 & "','USD',0, " & Fix((Val(strExc(1)) + Val(m_strCP17)) / dblUSRate) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & strSrvDate(2) & ",to_char(sysdate,'HH24miss'),'" & strUserNum & "')"
          strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K05,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21,A1K33) " & _
                  "VALUES  ('" & m_strSerialNo & "'," & strA1K02 & ",'" & strA1K05 & "',0,NULL," & Val(m_strCP17) & "," & dblUSRate & "," & (Val(strExc(1)) + Val(m_strCP17)) & ",NULL,'" & m_strCP01 & "','" & m_strCP02 & "','" & m_strCP03 & "','" & m_strCP04 & "','" & strA1K18 & "',0, " & Fix((Val(strExc(1)) + Val(m_strCP17)) / dblUSRate) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & strSrvDate(2) & ",to_char(sysdate,'HH24miss'),'" & strUserNum & "','" & strA1K33 & "')"
        
         cnnConnection.Execute strSql
         '3:·s¼W¤Gµ§ACC1L0
         strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
                  "VALUES  ('" & m_strSerialNo & "','FCP','',0,'001','" & m_strCP10 & "'," & Val(strExc(1)) & "," & strSrvDate(2) & ",to_char(sysdate,'HH24miss'),'" & strUserNum & "')"
         cnnConnection.Execute strSql
      Else
      'end 2014/3/10
         
         strDisc = 1 - (PUB_GetA1L07Disc(m_strCP01, m_strCP02, m_strCP03, m_strCP04, m_strCP10, strSrvDate(2)) / 100)
         'A1K11­n¥ý¦©°£§é¦©¤~¦sÀÉ
         '¬üª÷¨ú¾ã¼Æ¦ì(µL±ø¥ó±Ë¥h)
          'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
'         strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K05,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21) " & _
                  "VALUES  ('" & m_strSerialNo & "'," & strSrvDate(2) & ",'" & strA1K05 & "',0,NULL," & Val(m_strCP17) & "," & dblUSRate & "," & Val(m_strCP16) - Val((Val(m_strCP16) - Val(m_strCP17)) * Val(strDisc)) & ",NULL,'" & m_strCP01 & "','" & m_strCP02 & "','" & m_strCP03 & "','" & m_strCP04 & "','USD',0, " & Fix(Val("" & IIf(dblUSRate = 0, (Val(m_strCP16) - Val((Val(m_strCP16) - Val(m_strCP17)) * Val(strDisc))), ((Val(m_strCP16) - Val((Val(m_strCP16) - Val(m_strCP17)) * Val(strDisc))) / dblUSRate)))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & strSrvDate(2) & ",to_char(sysdate,'HH24miss'),'" & strUserNum & "')"
                  strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K05,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K19,A1K20,A1K21,A1K33) " & _
                  "VALUES  ('" & m_strSerialNo & "'," & strA1K02 & ",'" & strA1K05 & "',0,NULL," & Val(m_strCP17) & "," & dblUSRate & "," & Val(m_strCP16) - Val((Val(m_strCP16) - Val(m_strCP17)) * Val(strDisc)) & ",NULL,'" & m_strCP01 & "','" & m_strCP02 & "','" & m_strCP03 & "','" & m_strCP04 & "','" & strA1K18 & "',0, " & Fix(Val("" & IIf(dblUSRate = 0, (Val(m_strCP16) - Val((Val(m_strCP16) - Val(m_strCP17)) * Val(strDisc))), ((Val(m_strCP16) - Val((Val(m_strCP16) - Val(m_strCP17)) * Val(strDisc))) / dblUSRate)))) & ",'" & strAgentNo & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "'," & strSrvDate(2) & ",to_char(sysdate,'HH24miss'),'" & strUserNum & "','" & strA1K33 & "')"

         cnnConnection.Execute strSql
         '3:·s¼W¤Gµ§ACC1L0
         strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
                  "VALUES  ('" & m_strSerialNo & "','FCP',''," & (Val(m_strCP16) - Val(m_strCP17)) * Val(strDisc) & ",'001','" & m_strCP10 & "'," & (Val(m_strCP16) - Val(m_strCP17)) & "," & strSrvDate(2) & ",to_char(sysdate,'HH24miss'),'" & strUserNum & "')"
         cnnConnection.Execute strSql
      End If
      
      '92.4.18 MODIFY BY SONIA
      If Val(m_strCP17) <> 0 Then
      '92.4.18 END
         strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
                  "VALUES  ('" & m_strSerialNo & "','FCP','',0 ,'002','" & m_strCP10 & "99'," & Val(m_strCP17) & "," & strSrvDate(2) & ",to_char(sysdate,'HH24miss'),'" & strUserNum & "')"
         cnnConnection.Execute strSql
      End If
      
      'Added by Morgan 2019/12/16
      '»âÃÒ¦P®É½Ð®Ö¹ï¤w­ã±M§Q
      'Removed by Morgan 2025/8/11 §ï¦b¤W­±©I¥s¨ç¼Æ¤@°_³]©w
      'If m_strCP10 = »âÃÒ¤ÎÃº¦~¶O Then
         'If PUB_If601Plus926(m_PA75, m_strCP01, m_strCP02, m_strCP03, m_strCP04, str926CP09) = True Then
      'end 2025/8/11
      
            If str926CP09 <> "" Then
            
         'Removed by Morgan 2025/8/11 §ï¦b¤W­±©I¥s¨ç¼Æ¤@°_³]©w
         '      str926Disc = 1 - (PUB_GetA1L07Disc(m_strCP01, m_strCP02, m_strCP03, m_strCP04, "926", strSrvDate(2)) / 100)
         '      'Added by Morgan 2024/11/14 +Y27696 Bobst Mex SA»âÃÒ¤Î¤G®Ö±b³æ¤§ªA°È¶O³]©w¬°©T©wª÷ÃB(»âÃÒ¤ÎÃº¦~¶OªA°È¶O¡G170¬üª÷,¤G¦¸®Ö¹ïªA°È¶O¡G135¬üª÷,ªA°È¶OÁ`ª÷ÃB¡G305¬üª÷)
         '      If strA1K28 = "Y27696000" Then
         '         strExc(2) = PUB_GetUSXRate_1(strSrvDate(2), "USD")
         '         str926CP16 = UInt(135 * Val(strExc(2)))
         '         str926Disc = 0 'Added by Morgan 2024/12/12
         '      'end 2024/11/14
         '      ElseIf m_strPA08 = "3" Then
         '         str926CP16 = "3000"
         '      Else
         '         str926CP16 = "4500"
         '      End If
         'end 2025/8/11
         
               If bol926In605 = False Then 'Added by Morgan 2025/8/11
               
                  strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L06,A1L07,A1L02,A1L04,A1L05,A1L08,A1L09,A1L10) " & _
                           "VALUES  ('" & m_strSerialNo & "','FCP',''," & (Val(str926CP16) * Val(str926Disc)) & ",'003','926'," & Val(str926CP16) & "," & strSrvDate(2) & ",to_char(sysdate,'HH24miss'),'" & strUserNum & "')"
                  cnnConnection.Execute strSql, intI
               
                  '+¥x¹ôª÷ÃB(¥~¹ô¤U­±¨ç¼Æ·|§ó·s)
                  strSql = "update acc1k0 set a1k11=a1k11+" & (Val(str926CP16) - (Val(str926CP16) * Val(str926Disc))) & " where a1k01='" & m_strSerialNo & "'"
                  cnnConnection.Execute strSql, intI
               End If
               
               strSql = "UPDATE CASEPROGRESS SET CP16=" & str926CP16 & ",CP17=0,CP18=" & (Val(str926CP16) / 1000) & ",CP60='" & m_strSerialNo & "' WHERE CP09='" & str926CP09 & "'"
               cnnConnection.Execute strSql, intI
            End If
            
      'Removed by Morgan 2025/8/11 §ï¦b¤W­±©I¥s¨ç¼Æ¤@°_³]©w
         'End If
      'End If
      'end 2025/8/11
      'end 2019/12/16
      
      PUB_UpdateA1k08 m_strSerialNo 'Added by Morgan 2012/11/2 §ó·s½Ð´Ú³æ¥~¹ôª÷ÃB
      
   Else
      'Modified by Lydia 2017/07/18 debug (ex.FCP38125,¤é´Á1060621~1060705)
      'StrSQLa = "SELECT MAX(A1K01) FROM ACC1K0,ACC1W0 WHERE WHERE A1K13='" & m_strCP01 & "' AND A1K14='" & m_strCP02 & "' AND A1K15='" & m_strCP03 & "' AND A1K16='" & m_strCP04 & "' AND A1K01=A1W01(+) ADD A1W01 IS NULL "
      StrSQLa = "SELECT MAX(A1K01) FROM ACC1K0,ACC1W0 WHERE A1K13='" & m_strCP01 & "' AND A1K14='" & m_strCP02 & "' AND A1K15='" & m_strCP03 & "' AND A1K16='" & m_strCP04 & "' AND A1K01=A1W01(+) "
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      'Modified by Lydia 2017/07/18 debug
      'If rsA.RecordCount > 0 Then
      If Not rsA.EOF Then
      'end 2017/07/18
         m_strSerialNo = rsA.Fields(0).Value
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   
   '4:·s¼WACC1W0
   strSql = "INSERT INTO ACC1W0 " & _
            "VALUES  ('" & m_strSerialNo & "','" & m_strCP09 & "')"
   cnnConnection.Execute strSql
   '5:§ó·s½Ð´Ú³æ¸¹
   strSql = "UPDATE CASEPROGRESS SET CP16='" & m_strCP16 & "',CP17='" & m_strCP17 & "',CP18='" & m_strCP18 & "',CP60='" & m_strSerialNo & "' WHERE CP09='" & m_strCP09 & "'"
   cnnConnection.Execute strSql
   
   PUB_PointAutoassign m_strSerialNo, True 'Add by Morgan 2010/4/21 ¦Û°Ê¤À°tÂI¼Æ
   
   '½Ð´Ú³æ
   'Modify by Morgan 2008/4/3
   'If m_strStartDBNo = "" Then m_strStartDBNo = m_strSerialNo
   'If m_strSerialNo < m_strStartDBNo Then m_strStartDBNo = m_strSerialNo
   'If m_strSerialNo > m_strEndDBNo Then m_strEndDBNo = m_strSerialNo
   'Modified by Morgan 2014/6/3
   'PUB_AddNewDebitNoteList strUserNum, m_strSerialNo, "" & pub_AddressListSN, IIf(m_bolEmail, "Y", ""), IIf(m_bolPlusPaper, "Y", "")
   'Modified by Lydia 2020/08/17 +¤w¦¬´ÚIIf(m_OptPAID <> "", "Y", "")
   PUB_AddNewDebitNoteList strUserNum, m_strSerialNo, "" & pub_AddressListSN, IIf(m_bolDNEmail, "Y", ""), IIf(m_bolDNPlusPaper, "Y", ""), IIf(m_OptPAID <> "", "Y", "")
   'end 2014/6/3
   'end 2008/4/3
   
    'Added by Lydia 2016/11/21 ¾ã§å¦C¦L:¥H½Ð´Ú¹ï¶HÀË¬d¬O§_¦s¦b©ó°ê¥~©T©w±H¶Ê´Ú³æ¥N²z¤HÀÉ(ACC225)¥B¤U¦¸±Hµo¤é´Á¡Ö¨t²Î¤é¡A­Y¦s¦b«h·s¼W¦C¦L²M³æ
   If m_strSerialNo <> "" And strA1K28 <> "" Then
      If m_KeyCP09 = "" Or Len(m_KeyCP09) > 9 Then  'Added by Lydia 2019/11/04 §PÂ_¡G¾ã§å½Ð´Ú¨ç ©Î ¦~¶O¾ã§åµo¤å
            'Modified by Lydia 2020/06/23 §ï¥Î®×¥ó©Ê½è
            'If PUB_ChkAcc225MsgList(m_strSerialNo, strA1K28, m_strCP01, m_strCP02, m_strCP03, m_strCP04, IIf(Option1(0).Value = True, "¦~ÃÒ¶O½Ð´Ú¨ç", "¦~¶O½Ð´Ú¨ç")) Then
            If PUB_ChkAcc225MsgList(m_strSerialNo, strA1K28, m_strCP01, m_strCP02, m_strCP03, m_strCP04, IIf(m_strCP10 = "601", "¦~ÃÒ¶O½Ð´Ú¨ç", "¦~¶O½Ð´Ú¨ç")) Then
            End If
      'Added by Lydia 2019/11/04 ¦~¶O³æµ§,ª½±µµoemail
      Else
            If PUB_ChkAcc225MsgList(m_strSerialNo, strA1K28, m_strCP01, m_strCP02, m_strCP03, m_strCP04, "") Then
            End If
      End If
      'end 2019/11/04
   End If
    'end 2016/11/21
   
    'Added by Lydia 2020/08/17 µo¤å->(·s¼W½Ð´Ú³æ)¤w¦¬´Ú,¦Û°ÊEmailµ¹°]°È³B
    If m_strSerialNo <> "" And m_strCP01 = "FCP" And m_OptPAID <> "" And (m_OptKind = "1" Or m_OptKind = "2") Then
       'Modifie by Lydia 2024/01/29 §ï¦¨¯S®í³]©w
       'strExc(0) = "A2008" '°]°È¹w³]¦¬¥ó¤H, CCµ¹¾Þ§@ªÌ
       strExc(0) = Pub_GetSpecMan("¥~±M½Ð´Ú³æ¤w¦¬´Ú³qª¾¤H­û")
       If strExc(0) <> "" Then
       'end 2024/01/29
          Call ClsPDGetCaseProperty(m_strCP01, m_strCP10, strExc(2))
          strExc(1) = m_strCP01 & "-" & m_strCP02 & IIf(m_strCP03 & m_strCP04 = "000", "", "-" & m_strCP03 & "-" & m_strCP04) & strExc(2) & "¤w¦¬´Ú¡A½Ð´Ú³æ½s¸¹¬°" & m_strSerialNo
          strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
             "values('" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
             ",'" & ChgSQL(strExc(1)) & "','¦P¥D¦®','" & strUserNum & "' )"
          cnnConnection.Execute strSql, intI
       End If
    End If
    'end 2020/08/17
   
   AddACC1K0 = True
   
   'Remove by Morgan 2008/4/3 ²Î¤@§ï©I¥s¤½¥Îµ{¦¡¦C¦L
   ''¦a§}±ø
   'ReDim Preserve m_CustList(m_CustListCount + 1)
   'ReDim Preserve m_CP(m_CustListCount + 1)
   'm_CustList(m_CustListCount) = strAgentNo
   'm_CP(m_CustListCount) = CheckStr(m_strCP01) & "-" & CheckStr(m_strCP02) & "-" & CheckStr(m_strCP03) & "-" & CheckStr(m_strCP04)
   'm_CustListCount = m_CustListCount + 1
   
End Function

Private Function GetUSRate() As Double
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetUSRate = 0
'strSQLA = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & (ServerDate - 19110000) & " AND ROWNUM = 1 ORDER BY USXR01 "
StrSQLa = "SELECT USXR02 FROM USXRATE WHERE USXR01<=" & (ServerDate - 19110000) & " ORDER BY USXR01 DESC "
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
'
'Private Function GetAgentNO() As String
'Dim rsA As New ADODB.Recordset
'Dim strSQLA As String
'
'GetAgentNO = ""
''¨ú±o±M§Q°ò¥»ÀÉªº"FC¥N²z¤H"
'strSQLA = "Select PA75 From PATENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' AND PA75 IS NOT NULL "
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'rsA.CursorLocation = adUseClient
'rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   GetAgentNO = rsA.Fields(0).Value
'Else
'   '¨ú±o±M§Q°ò¥»ÀÉªº"¥Ó½Ð¤H1"
'   strSQLA = "SELECT PA26 FROM PATENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' AND PA26 IS NOT NULL "
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   rsA.CursorLocation = adUseClient
'   rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      GetAgentNO = rsA.Fields(0).Value
'   End If
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'End Function

'Private Function GetPrintNO() As String
'Dim rsA As New ADODB.Recordset
'Dim strSQLA As String
'
'GetPrintNO = ""
''¨ú±o±M§Q°ò¥»ÀÉªº"FC¥N²z¤H"
'strSQLA = "Select PA75 From PATENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' AND PA75 IS NOT NULL "
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'rsA.CursorLocation = adUseClient
'rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   GetPrintNO = rsA.Fields(0).Value
'   '¨ú±o°ê¥~¥N²z¤HÀÉªº"©T©w½Ð´Ú¹ï¶H"
'   strSQLA = "Select FA30 From FAGENT WHERE SUBSTR('" & GetAgentNO & "',1,8)=FA01 AND SUBSTR('" & GetAgentNO & "',9,1)=FA02 AND FA30 IS NOT NULL "
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   rsA.CursorLocation = adUseClient
'   rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      GetPrintNO = rsA.Fields(0).Value
'   End If
'Else
'   '¨ú±o±M§Q°ò¥»ÀÉªº"¥Ó½Ð¤H1"
'   strSQLA = "SELECT PA26 FROM PATENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' AND PA26 IS NOT NULL "
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   rsA.CursorLocation = adUseClient
'   rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      GetPrintNO = rsA.Fields(0).Value
'      '¨ú±o«È¤á°ò¥»ÀÉªº"©T©w½Ð´Ú¹ï¶H"
'      strSQLA = "SELECT CU57 FROM CUSTOMER WHERE CU01=SUBSTR('" & GetAgentNO & "',1,8) AND CU02=SUBSTR('" & GetAgentNO & "',9,1) AND CU57 IS NOT NULL "
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      rsA.CursorLocation = adUseClient
'      rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         GetPrintNO = rsA.Fields(0).Value
'      End If
'   End If
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'End Function

'Private Function GetAccountNO() As String
'Dim rsA As New ADODB.Recordset
'Dim strSQLA As String
'
'GetAccountNO = ""
''¨ú±o±M§Q°ò¥»ÀÉªº"¦~¶O½Ð´Ú¹ï¶H"
'strSQLA = "SELECT PA105 FROM PATENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' AND PA105 IS NOT NULL "
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'rsA.CursorLocation = adUseClient
'rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 And m_strCP10 = ¦~¶O Then
'   GetAccountNO = rsA.Fields(0).Value
'Else
'   '¨ú±o±M§Q°ò¥»ÀÉªº"©T©w½Ð´Ú¹ï¶H"
'   strSQLA = "SELECT PA88 FROM PATENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' AND PA88 IS NOT NULL "
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   rsA.CursorLocation = adUseClient
'   rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      GetAccountNO = rsA.Fields(0).Value
'   Else
'      '¨ú±o±M§Q°ò¥»ÀÉªº"FC¥N²z¤H"
'      strSQLA = "Select PA75 From PATENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' AND PA75 IS NOT NULL "
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      rsA.CursorLocation = adUseClient
'      rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         GetAccountNO = rsA.Fields(0).Value
'         '¨ú±o°ê¥~¥N²z¤HÀÉªº"¦~¶O½Ð´Ú¹ï¶H"
'         strSQLA = "Select FA62 From FAGENT WHERE SUBSTR('" & GetAgentNO & "',1,8)=FA01 AND SUBSTR('" & GetAgentNO & "',9,1)=FA02 AND FA62 IS NOT NULL "
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'         rsA.CursorLocation = adUseClient
'         rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount > 0 And m_strCP10 = ¦~¶O Then
'            GetAccountNO = rsA.Fields(0).Value
'         Else
'            '¨ú±o°ê¥~¥N²z¤HÀÉªº"©T©w½Ð´Ú¹ï¶H"
'            strSQLA = "Select FA30 From FAGENT WHERE SUBSTR('" & GetAgentNO & "',1,8)=FA01 AND SUBSTR('" & GetAgentNO & "',9,1)=FA02 AND FA30 IS NOT NULL "
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            rsA.CursorLocation = adUseClient
'            rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount > 0 Then
'               GetAccountNO = rsA.Fields(0).Value
'            End If
'         End If
'      Else
'         '¨ú±o±M§Q°ò¥»ÀÉªº"¥Ó½Ð¤H1"
'         strSQLA = "SELECT PA26 FROM PATENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' AND PA26 IS NOT NULL "
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'         rsA.CursorLocation = adUseClient
'         rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount > 0 Then
'            GetAccountNO = rsA.Fields(0).Value
'            '¨ú±o«È¤á°ò¥»ÀÉªº"¦~¶O½Ð´Ú¹ï¶H"
'            strSQLA = "SELECT CU97 FROM CUSTOMER WHERE CU01=SUBSTR('" & GetAgentNO & "',1,8) AND CU02=SUBSTR('" & GetAgentNO & "',9,1) AND CU97 IS NOT NULL "
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            rsA.CursorLocation = adUseClient
'            rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount > 0 And m_strCP10 = ¦~¶O Then
'               GetAccountNO = rsA.Fields(0).Value
'            Else
'               '¨ú±o«È¤á°ò¥»ÀÉªº"©T©w½Ð´Ú¹ï¶H"
'               strSQLA = "SELECT CU57 FROM CUSTOMER WHERE CU01=SUBSTR('" & GetAgentNO & "',1,8) AND CU02=SUBSTR('" & GetAgentNO & "',9,1) AND CU57 IS NOT NULL "
'               If rsA.State <> adStateClosed Then rsA.Close
'               Set rsA = Nothing
'               rsA.CursorLocation = adUseClient
'               rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'               If rsA.RecordCount > 0 Then
'                  GetAccountNO = rsA.Fields(0).Value
'               End If
'            End If
'         End If
'      End If
'   End If
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'End Function

'Private Function GetPrintCust() As String
'Dim rsA As New ADODB.Recordset
'Dim strSQLA As String
'
'GetPrintCust = ""
''¨ú±o±M§Q°ò¥»ÀÉªº"D/N¬O§_¦C¦L¥Ó½Ð¤H"
'strSQLA = "SELECT PA78 FROM PATENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' AND PA78 IS NOT NULL "
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'rsA.CursorLocation = adUseClient
'rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   GetPrintCust = rsA.Fields(0).Value
'Else
'   '¨ú±o°ê¥~¥N²z¤HÀÉªº"D/N¬O§_¦C¦L¥Ó½Ð¤H"
'   '92.04.02 nick add left join
'   'strSQLA = "SELECT FA44 FROM PATENT, FAGENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' " & _
'            " AND SUBSTR(PA75,1,8)=FA01 AND SUBSTR(PA75,9,1)=FA02 AND FA44 IS NOT NULL "
'    strSQLA = "SELECT FA44 FROM PATENT, FAGENT WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' " & _
'            " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA44 IS NOT NULL "
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   rsA.CursorLocation = adUseClient
'   rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'      GetPrintCust = rsA.Fields(0).Value
'   Else
'      '¨ú±o«È¤á°ò¥»ÀÉªº"D/N¬O§_¦C¦L¥Ó½Ð¤H"
'      '92.04.02 nick add left join
'      'strSQLA = "SELECT CU77 FROM PATENT, CUSTOMER WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' " & _
'               " AND SUBSTR(PA26,1,8)=CU01 AND SUBSTR(PA26,9,1)=CU02 AND CU77 IS NOT NULL "
'      strSQLA = "SELECT CU77 FROM PATENT, CUSTOMER WHERE PA01='" & m_strCP01 & "' AND PA02='" & m_strCP02 & "' AND PA03='" & m_strCP03 & "' AND PA04='" & m_strCP04 & "' " & _
'               " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND CU77 IS NOT NULL "
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

Private Sub Command1_Click(Index As Integer)
 Dim strTmp As String
   If Index = 0 And Text1(2).Text <> "" Then
      strTmp = Text1(1) & Text1(2)
      If Text1(3).Text = "" Then
         strTmp = strTmp & "0"
      Else
         strTmp = strTmp & Text1(3).Text
      End If
      If Text1(4).Text = "" Then
         strTmp = strTmp & "00"
      Else
         strTmp = strTmp & Text1(4).Text
      End If
      intI = 1
      strExc(0) = "SELECT PA57 FROM PATENT WHERE " & ChgPatent(strTmp)
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) = "Y" Then
            MsgBox "¥²¶·¬°¥¼³¬¨÷¤§®×¸¹¡A½Ð­«·s¿é¤J !", vbCritical
         Else
            List1.AddItem strTmp
            Text1(2).Text = ""
         End If
      Else
         MsgBox "®×¸¹¤£¦s¦b¡A½Ð­«·s¿é¤J !", vbCritical
      End If
      Text1(2).SetFocus
   Else
      If List1.ListIndex > -1 Then List1.RemoveItem List1.ListIndex
   End If
End Sub

Private Sub Form_Load()
Dim ii As Integer
'Add By Cheng 2003/02/05
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   MoveFormToCenter Me
   intWhere = °ê¥~_FC
   
   'Modified by Morgan 2013/4/17 §ï¬ö¿ý¦b¸ê®Æ®w,§_«h´«¹q¸£©Î¨Ï¥ÎªÌ·|Åª¤£¨ì
   'Label2(0).Caption = GetSetting("TAIE", "FCP", "DATE71", "")
   'Label2(1).Caption = GetSetting("TAIE", "FCP", "DATE72", "")
   Option1(0).Value = True
   'end 2013/4/17
   
   'm_CustListCount = 0
   
'Modify by Morgan 2011/3/15 §ï¦@¥Î¥B¤£­n±Æ°£¹w³]¦Lªí¾÷
   PUB_SetPrinter Me.Name, Combo1, strPrinter
   'Modify By Sindy 2015/9/24 +strPrinter2
   PUB_SetPrinter Me.Name, Combo2, strPrinter2
'end 2011/3/15
   
   If m_OptKind = "" Then m_OptKind = "0" 'Added by Lydia 2020/08/17 «D¥~³¡©I¥s,¹w³]Ãþ§O

End Sub

Private Sub Form_Unload(Cancel As Integer)
   '¦C¦L½Ð´Ú³æ
   'Modified by Lydia 2019/10/01 ¶Ç¤Jµ{¦¡¦WºÙ
   'PUB_PrintDebitNote strUserNum, Me.Combo2.Text
   'Remove by Lydia 2019/10/18 Sharon: §ï¦¨¤@­Ó®×¤l¦L§¹©Ó¿ì³æ->½Ð´Ú©w½Z->±b³æ
   'PUB_PrintDebitNote strUserNum, Me.Combo2.Text, IIf(m_KeyCP09 <> "", Me.Name, "")
   ''§R°£½Ð´Ú³æ¦Cªí¸ê®Æ
   'PUB_DeleteDebitNoteList strUserNum
   
   'Modified by Lydia 2020/08/17
   'If m_KeyCP09 = "" Or Len(m_KeyCP09) > 9 Then   'Added by Lydia 2019/10/01 §PÂ_¡G¾ã§å½Ð´Ú¨ç ©Î ¦~¶O¾ã§åµo¤å
   'Modified by Lydia 2020/09/23 debug: ¦~¶O¾ã§åµo¤å,¨S¦³¦C¦L²M³æ
   'If m_OptKind < "2" And (m_KeyCP09 = "" Or Len(m_KeyCP09) > 9) Then
   If m_OptKind <> "3" And (m_KeyCP09 = "" Or Len(m_KeyCP09) > 9) Then
        '¦C¦L©w½Z¾ã§å¦C¦L²M³æ
        'Modify By Sindy 2017/2/13
        'PUB_PrintLetterList strUserNum, "7", Me.Combo2.Text, strPrinter2
        'Modified by Lydia 2020/08/17
        'PUB_PrintLetterList strUserNum, "7", Me.Combo2.Text, strPrinter2, IIf(Option1(0).Value = True, Option1(0).Caption, "")
        'Modified by Lydia 2020/09/23 ­ì¥»¾ã§å-»âÃÒ¤ÎÃº¦~¶O,²M³æ¶Ç¤JOption1(0).Caption·|¥H©Ó¿ì¤H¬°¥D; ²{¦b¤w¸g¨S¦³¾ã§å-»âÃÒ¤ÎÃº¦~¶O
        'PUB_PrintLetterList strUserNum, "7", Me.Combo2.Text, strPrinter2, "¦~¶O"
        PUB_PrintLetterList strUserNum, "7", Me.Combo2.Text, strPrinter2
        '2017/2/13 END
        '§R°£©w½Z¾ã§å¦C¦L¸ê®Æ
        'Modified by Lydia +¶Ç¤J§R°£±ø¥ó
        'PUB_DeleteLetterList strUserNum
        PUB_DeleteLetterList strUserNum, "and LL02='¦~ÃÒ¶O½Ð´Ú¨ç' "
   End If
   
   'Added by Lydia 2016/11/21
   '¦C¦L:°ê¥~©T©w±H¶Ê´Ú³æ²M³æ
   PUB_PrintAcc225List strUserNum, Me.Combo2.Text
   '§R°£:°ê¥~©T©w±H¶Ê´Ú³æ²M³æ
   PUB_DeleteAcc225List strUserNum
   'end 2016/11/21
        
   '¦C¦L¦a§}±ø
   PUB_PrintAddressList strUserNum, Me.Combo1.Text
   '§R°£¦a§}±ø¦Cªí¸ê®Æ
   PUB_DeleteAddressList strUserNum
   'ªì©l¤Æ§Ç¸¹
   pub_AddressListSN = 0
   
   '­Y¦a§}±ø¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '­Y½Ð´Ú³æ¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   '2004/10/26 end
   
   'Added by Lydia 2020/08/17 ­Ó®×µo¤å­n¦^¨ì«eµe­±,¤~·|commit
   If m_OptKind <> "1" Then
       PUB_SendMailCache
   End If
   'end 2020/08/17
   
   Set frm060307 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
   If m_KeyCP09 = "" Then 'Added by Lydia 2019/10/01 §PÂ_«D¦~¶Oµo¤å
        Label2(0).Caption = PUB_GetLastDate(Me.Name, Index & "DATE71")
        Label2(1).Caption = PUB_GetLastDate(Me.Name, Index & "DATE72")
        QueryOtherData
   End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   If Index = 6 Then
      If Text1(6) <> "" Then
         If Not ChkRange(Text1(5), Text1(6), "¥»¦¸¦C¦Lµo¤å¤é") Then
            Text1(5).SetFocus
            TextInverse Text1(5)
         End If
      End If
   End If
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1
         If Text1(Index).Text <> "FCP" Then
            MsgBox "¨t²Î§O¥²»Ý¬° FCP¡A½Ð­«·s¿é¤J !", vbCritical
            Cancel = True
            GoTo RunEnd
         End If
      Case 5, 6
         If Text1(Index) <> "" Then
            If ChkDate(Text1(Index)) Then
               If Index = 5 Then
                  If Val(Label2(1).Caption) >= Val(Text1(Index).Text) Then
                     MsgBox "¥»¦¸¦C¦Lµo¤å¤é¤£¥i¤p©ó©Îµ¥©ó¤W¦¸¦C¦Lµo¤å¤é¡A½Ð­«·s¿é¤J !", vbCritical
                     Cancel = True
                     GoTo RunEnd
                  End If
               End If
            Else
               Cancel = True
               GoTo RunEnd
            End If
            QueryOtherData
         End If
   End Select
RunEnd:
   If Cancel Then TextInverse Text1(Index)
End Sub

'Add By Sindy 2015/9/23
Private Sub QueryOtherData()
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim bFind As Boolean
Dim strData As String
Dim stCon As String
   
   If Text1(5) = "" Or Text1(6) = "" Then Exit Sub 'µL¥»¦¸¦C¦Lµo¤å¤é
   
   Screen.MousePointer = vbHourglass
   
   If Option1(0).Value = True Then
      stCon = " AND CP10='" & »âÃÒ¤ÎÃº¦~¶O & "'"
   ElseIf Option1(1).Value = True Then
      stCon = " AND CP10='" & ¦~¶O & "'"
   End If
            
   '¦³¯S®í½Ð´Ú
   strSql = "select CP01,CP02,CP03,CP04,CP09,CP10,CP16,CP17,CP27,PA14,CP18,FNo,nvl(n1.na16,n2.NA16) CtMan,PA26" & _
            " from (SELECT CP01,CP02,CP03,CP04,CP09,CP10,CP16,CP17,CP27,PA14,CP18,Nvl(PA76, Nvl(CU96, Nvl(PA75, PA26))) FNo,PA26" & _
            " FROM CASEPROGRESS,PATENT,CUSTOMER WHERE CP01='FCP' " & stCon & " and CP27 BETWEEN " & TransDate(Text1(5), 2) & " AND " & TransDate(Text1(6), 2) & _
            " AND (CP20<>'N' OR CP20 IS NULL) AND CP60 IS NULL AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND PA57 IS NULL and cp148='Y'" & _
            " and cu01(+)=substr(PA26,1,8) and cu02(+)=substr(PA26,9)),fagent,customer,nation n1,nation n2" & _
            " where fa01(+)=substr(fno,1,8) and fa02(+)=substr(fno,9) and cu01(+)=substr(fno,1,8) and cu02(+)=substr(fno,9) and n1.na01(+)=fa10 and n2.na01(+)=cu10" & _
            " order by cp01,cp02,cp03,cp04"
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         bFind = False
         strData = rsTmp.Fields("CP01") & rsTmp.Fields("CP02") & rsTmp.Fields("CP03") & rsTmp.Fields("CP04")
         For i = 0 To List1.ListCount - 1
            If List1.List(i) = strData Then
               bFind = True: Exit For
            End If
         Next i
         If bFind = False Then
            List1.AddItem strData
         End If
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
End Sub

'92.1.19 COPY BY SONIA FROM frm060104_7
'­pºâ¬ÛÃö¶O¥Î
Private Sub GetPatentYearFee()

Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim ArrYear
Dim ii As Integer, intEnd As Integer

   ArrYear = Split(m_strPA72, ",")
   intEnd = Val(ArrYear(UBound(ArrYear)))
   '¨ú±o®×¥ó©Ê½è¬°¦~¶Oªº¬ÛÃö¶O¥Î
   StrSQLa = "Select * From PatentYearFee Where YF01='" & m_strPA09 & "' AND YF02='" & m_strPA08 & "' AND YF03='Y00000000' AND YF04='" & ¦~¶O & "' AND YF05=" & Val(intEnd)
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If Not rsA.EOF Then
      m_strCP17 = Val("" & rsA.Fields("YF07").Value)
      m_strCP16 = Val("" & rsA.Fields("YF06").Value) + Val("" & rsA.Fields("YF07").Value)
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   '¨ú±o®×¥ó©Ê½è¬°»âÃÒªº¬ÛÃö¶O¥Î
   If m_strCP10 <> ¦~¶O Then
      StrSQLa = "Select * From PatentYearFee Where YF01='" & m_strPA09 & "' AND YF02='" & m_strPA08 & "' AND YF03='Y00000000' AND YF04='" & m_strCP10 & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If Not rsA.EOF Then
         m_strCP17 = Val(m_strCP17) + Val("" & rsA.Fields("YF07").Value)
         m_strCP16 = Val(m_strCP16) + Val("" & rsA.Fields("YF06").Value) + Val("" & rsA.Fields("YF07").Value)
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   '93.3.5 modify by sonia
   'm_strCP18 = Round(((Val(m_strCP16) - Val(m_strCP17)) / 1000), 1)
   'Remove by Morgan 2007/9/5 ÁÙ­ì
   'm_strCP18 = Val(m_strCP16) / 1000
   m_strCP18 = Round(((Val(m_strCP16) - Val(m_strCP17)) / 1000), 1)
   'end 2007/9/5
   '93.3.5 end
End Sub

Private Function CheckPayBefore() As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String

   CheckPayBefore = "N"
   ChgCaseNo m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04, pa
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 ¤£¥Î dll ¤F If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
     
      '­Y¦³FC¥N²z¤H§PÂ_FA39="Y" ,­YµLFC¥N²z¤H«h¥Î¥Ó½Ð¤H§PÂ_ CU72="Y" , ¤~¥i·s¼W½Ð´Ú³æ¸ê®Æ
        '­Y¦³¥N²z¤H
      If pa(75) <> "" Then
        pa(75) = Left(pa(75) & "000000000", 9)
        StrSQLa = "Select FA39 From Fagent Where FA01='" & Left(pa(75), 8) & "' And FA02='" & Mid(pa(75), 9, 1) & "' " & _
                            " And FA39 IS NOT NULL AND FA39 ='Y' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            CheckPayBefore = "Y"
            Exit Function
        End If
        '­Y¦³¥Ó½Ð¤H
      ElseIf pa(26) <> "" Then
        pa(26) = Left(pa(26) & "000000000", 9)
        StrSQLa = "Select CU72 From Customer Where CU01='" & Left(pa(26), 8) & "' And CU02='" & Mid(pa(26), 9, 1) & "' " & _
                            " And CU72 IS NOT NULL AND CU72 ='Y' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            CheckPayBefore = "Y"
            Exit Function
        End If
      '­YµL¥N²z¤H»P¥Ó½Ð¤H
      Else
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            cnnConnection.RollbackTrans
            Exit Function
      End If
   End If
End Function

'Remove by Morgan 2008/4/3 ²Î¤@§ï©I¥s¤½¥Îµ{¦¡¦C¦L
'Private Sub PrintAddress()
'   Dim nPageNo As String
'   Dim strCust As String
'   Dim nPos As Integer
'
'   ' ¬y¤ô¸¹
'   nPageNo = 1
'   For nPos = 0 To m_CustListCount - 1
'      strCust = m_CustList(nPos)
'      If Len(strCust) > 8 Then
'         strCust = Left(strCust, 8)
'      Else
'         strCust = strCust & String(8 - Len(strCust), "0")
'      End If
'
'      ' ¦C¦L¦a§}±ø
'      Load frm083014
'      '****** 90.11.29  nick
'      frm083014.Hide
'      '**********************
'      '****** 91.08.07   nick  ¥[¤J¥»©Ò®×¸¹
'      frm083014.Text1(6).Text = m_CP(nPos)
'      '************************************
'      'Add By Cheng 2002/12/20
'      '¶Ç¥»©Ò®×¸¹
'      frm083014.SetCaseNo m_CP(nPos)
'      '§t¤£±HÂø»xªº«È¤á
'      frm083014.Text1(5).Text = "Y"
'
'      frm083014.Text1(1).Text = strCust
'      ' ¥u¦L¤@¥÷
'      frm083014.Text1(3).Text = "1"
'      ' ¦L­^¤å
'      frm083014.Text1(4).Text = "2"
'      ' ¦a§}±ø¬y¤ô¸¹
'      frm083014.SetPageNo nPageNo
'      ' ³]©w¦Lªí¾÷
'      frm083014.SetPrinter Combo1.List(Combo1.ListIndex)
'      frm083014.cmdPrint_Click
'      frm083014.cmdBack_Click
'      ' ¬y¤ô¸¹»¼¼W
'      nPageNo = nPageNo + 1
'   Next nPos
'
'   ' ²M°£¥Ó½Ð¤H¦ê¦C
'   ClearCustList
'End Sub

' ²M°£¥Ó½Ð¤H¥N½X¼È¦s°Ï
'Private Sub ClearCustList()
'   If m_CustListCount > 0 Then
'      Erase m_CustList
'      Erase m_CP
'   End If
'   m_CustListCount = 0
'End Sub

Private Sub PrintWorkF()
Dim i As Integer, SavDay1 As String

'2008/9/12 ADD BY SONIA
Printer.FontSize = 12
Printer.Font.Name = "Times New Roman"
'2008/9/12 END
strSql = "SELECT * FROM R060307 WHERE ID='" & strUserNum & "' ORDER BY R06030701,R06030705,R06030707,R06030708,R06030709,R06030710 "
CheckOC
SavDay1 = ""
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        'PrintTitle
        Do While .EOF = False
            For i = 1 To 11
                strTemp(i) = CheckStr(.Fields(i - 1))
                If i = 9 Then
                   If .Fields(8) = "0" And .Fields(9) = "00" Then
                      strTemp(9) = "": strTemp(10) = ""
                      i = 10
                   End If
                End If
            Next i
            If SavDay1 <> "" And strTemp(1) <> SavDay1 Then
               PrintEnd
                Printer.NewPage
            End If
            If strTemp(1) <> SavDay1 Then
                Page = 1
                PrintTitle
                SavDay1 = strTemp(1)
            End If
            If iPrint > 14500 Then
                Printer.CurrentX = 100
                Printer.CurrentY = iPrint
                Printer.Print String(138, "-")
                Printer.NewPage
                Page = Page + 1
                PrintTitle False
            End If
            PrintDatil
            If iPrint >= 14500 Then
                Printer.NewPage
                Page = Page + 1
                PrintTitle False
            End If
            .MoveNext
        Loop
        PrintEnd
    End With
End If
Printer.EndDoc
CheckOC

End Sub

Sub PrintDatil()
Dim strTH As String

For i = 2 To 10
   If i > 8 And strTemp(i) <> "" Then
      Printer.CurrentX = PLeft(i) - 80
      Printer.CurrentY = iPrint
      Printer.Print "-"
   End If
   If i = 5 Then
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print ChgEngDate(strTemp(i))
   Else
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   End If
   If i = 6 Then
      If Val(strTemp(i)) > 0 Then
         Printer.CurrentX = PLeft(i) + 300
         Printer.CurrentY = iPrint
         Select Case Val(strTemp(i))
            Case 1
               strTH = "st"
            Case 2
               strTH = "nd"
            Case 3
               strTH = "rd"
            Case Else
               strTH = "th"
         End Select
         Printer.Print strTH
      End If
   End If
   If i = 7 Then
      Printer.CurrentX = PLeft(i) + 400
      Printer.CurrentY = iPrint
      Printer.Print "-"
   End If
Next i
iPrint = iPrint + 300
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print strTemp(11)
iPrint = iPrint + 600
End Sub

Sub GetPleft()
Erase PLeft
PLeft(2) = 500
PLeft(3) = 1700
PLeft(4) = 2800
'92.9.2 MODIFY BY SONIA
'PLeft(5) = 6700
PLeft(5) = 6400
'92.9.2 END
PLeft(6) = 8500
PLeft(7) = 9100
PLeft(8) = 9600
PLeft(9) = 10350
PLeft(10) = 10550
PLeft(11) = 2800
End Sub

Private Sub PrintTitle(Optional bolPrint As Boolean = True)
Dim rsTmp As New ADODB.Recordset
Dim StrSQLa As String

GetPleft
iPrint = 2900
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 8000
Printer.CurrentY = iPrint
Printer.Print ChgEngDate(GetTodayDate)
iPrint = iPrint + 300
Printer.CurrentX = 8000
Printer.CurrentY = iPrint
Printer.Print "Page : " & str(Page)
iPrint = iPrint + 300
If bolPrint Then
   '«HÀY¦a§}strTemp(i)
   Set rsTmp = Nothing
   'Modify by Morgan 2011/5/26 +CU102,FA70
   If Left(strTemp(1), 1) = "X" Then
      StrSQLa = "SELECT CU05,CU88,CU89,CU90,CU24,CU25,CU26,CU27,CU28," & _
         "CU65,CU66,CU67,CU68,CU69,CU102 FROM CUSTOMER WHERE " & ChgCustomer("" & strTemp(1))
   Else
      StrSQLa = "SELECT FA05,FA63,FA64,FA65,FA18,FA19,FA20,FA21,FA22," & _
         "FA32,FA33,FA34,FA35,FA36,FA70 FROM FAGENT WHERE " & ChgFagent("" & strTemp(1))
   End If
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      With rsTmp
         For i = 0 To 3
            If Not IsNull(.Fields(i)) Then
               Printer.CurrentX = 500
               Printer.CurrentY = iPrint
               Printer.Print .Fields(i)
               iPrint = iPrint + 300
            End If
         Next
         If .Fields(9) <> "" Or .Fields(10) <> "" Or .Fields(11) <> "" Or .Fields(12) <> "" Or .Fields(13) <> "" Then
            For i = 9 To 13
               If Not IsNull(.Fields(i)) Then
                  Printer.CurrentX = 500
                  Printer.CurrentY = iPrint
                  Printer.Print .Fields(i)
                  iPrint = iPrint + 300
               End If
            Next
         Else
            For i = 4 To 8
               If Not IsNull(.Fields(i)) Then
                  Printer.CurrentX = 500
                  Printer.CurrentY = iPrint
                  Printer.Print .Fields(i)
                  iPrint = iPrint + 300
               End If
            Next
            'Add by Morgan 2011/5/26
            '¦a§}6
            If Not IsNull(.Fields(14)) Then
               Printer.CurrentX = 500
               Printer.CurrentY = iPrint
               Printer.Print .Fields(14)
               iPrint = iPrint + 300
            End If
         End If
      End With
   End If
   
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "Re¡G Payment of the Annuities of Taiwan Patents"
   iPrint = iPrint + 600
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   'Modified by Morgan 2024/4/10 ¹ï¥~²Î¤@¥Î Dear Colleagues --ªLÁ`
   'Printer.Print "Dear Sirs,"
   Printer.Print "Dear Colleagues,"
   'end 2024/4/10
   iPrint = iPrint + 600
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "We are pleased to advise that the annuity of each of the following Taiwan patents has been paid before the "
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "due date of this year and the payment schedule for the next year is also listed for your reference :"
End If

iPrint = iPrint + 500
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(138, "-")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "PAT NO.     APPLN NO.    YOUR REF                                        DUE  DATE         YEAR    OUR REF"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "                                          CASE NO."
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(138, "-")
iPrint = iPrint + 300
If iPrint >= 14500 Then
   Page = Page + 1
   Printer.NewPage
   PrintTitle False
   Exit Sub
End If
End Sub

Sub PrintEnd()
iPrint = iPrint + 600
If iPrint + 2100 >= 14900 Then
   Page = Page + 1
   Printer.NewPage
   PrintTitle False
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "With best regards,"
iPrint = iPrint + 600
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "                                                                                               Best regards,"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "                                                                                               Tai E International"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "                                                                                               Patent & Law Office"
iPrint = iPrint + 600
Printer.CurrentX = 500
Printer.CurrentY = iPrint
'Modify by Morgan 2006/7/6
'Printer.Print "ICL/dy"
Printer.Print "CTY/dy"
End Sub

'¦~¶O³æµ§¤£¶], ¥ý¼g¤J¤u§@ÀÉ, ³Ì«á¦A¦C¦L
Private Function AddWorkF() As Boolean
Dim strAgentNo As String '¥N²z¤H½s¸¹
Dim strPA72 As String
Dim strKey(5) As String
Dim strCaseFee(1 To 2) As String
Dim varPA72 As Variant
Dim varRef As Variant
Dim ii As Integer
Dim m_NextPA72 As String
Dim rsTmp As New ADODB.Recordset
Dim StrSQLa As String

'   On Error GoTo ErrorHandler
   AddWorkF = False
   
   '³Ì«á¤@¦~¤£·s¼W
   'Modify by Morgan 2004/12/7
   'If m_LetterLanguage = "4" Then
   If m_LetterKind = "4" Then
      'Remove by Morgan 2007/9/27 Ãº³Ì«á¤@¦~ªº®×¥ó¤]­n¦L,¦ý¤£±¾´Á­­
      'AddWorkF = True
      'Exit Function
      'end 2007/9/27
   End If
'   cnnConnection.BeginTrans
'   strAgentNo = GetAccountNO '½Ð´Ú¹ï¶H
   '93.10.4 CANCEL BY SONIA ³æµ§¤£¶]«H¨ç¤§©ïÀY¶¶§Ç¬° ¦~¶O¥N²z¤H->¥N²z¤H->¥Ó½Ð¤H
   'strAgentNo = PUB_GetA1K28(m_strCP01, m_strCP02, m_strCP03, m_strCP04, m_strCP10)
   '93.10.4 END
   Set rsTmp = Nothing
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & m_strCP01 & "' AND " & _
                  "PA02 = '" & m_strCP02 & "' AND " & _
                  "PA03 = '" & m_strCP03 & "' AND " & _
                  "PA04 = '" & m_strCP04 & "' "
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("PA72")) = False Then strPA72 = rsTmp.Fields("PA72")
      strKey(0) = ""
      strKey(1) = m_strCP01
      strKey(2) = m_strCP02
      strKey(3) = m_strCP03
      strKey(4) = m_strCP04
      'Add by Morgan 2007/9/27 ³Ì«á¤@¦~
      If m_LetterKind = "4" Then
         m_strNP09 = "NULL"
         m_NextPA72 = "NULL"
      Else
      'end 2007/9/27
         If GetMoneyDate(rsTmp.Fields("PA08"), rsTmp.Fields("PA09"), strKey, strCaseFee(1), strCaseFee(2)) Then
            ' ´M§ä¤U¦¸Ãº¦~¶Oªº¦ì¸m
            '­Y¤w¦³Ãº¶O°O¿ý
            If IsEmptyText(strPA72) = False Then
                 varPA72 = Split(strPA72, ",")
                 If IsEmptyText(strCaseFee(2)) = False Then
                    varRef = Split(strCaseFee(2), ",")
                    For ii = LBound(varRef) To UBound(varRef)
                          If Val(varPA72(UBound(varPA72))) < Val(varRef(ii)) Then
                               m_NextPA72 = varRef(ii)
                               Exit For
                          End If
                    Next ii
                 End If
             '­YµLÃº¶O°O¿ý
             Else
                 If IsEmptyText(strCaseFee(2)) = False Then
                     varRef = Split(strCaseFee(2), ",")
                     m_NextPA72 = varRef(LBound(varRef))
                 End If
            End If
         End If
      End If
      '93.10.4 ADD BY SONIA ³æµ§¤£¶]«H¨ç¤§©ïÀY¶¶§Ç¬° ¦~¶O¥N²z¤H->¥N²z¤H->¥Ó½Ð¤H
      strAgentNo = ""
      If IsNull(rsTmp.Fields("PA76")) = False Then strAgentNo = rsTmp.Fields("PA76")
      If strAgentNo = "" And IsNull(rsTmp.Fields("PA75")) = False Then
         strAgentNo = rsTmp.Fields("PA75")
      End If
      If strAgentNo = "" And IsNull(rsTmp.Fields("PA26")) = False Then
         strAgentNo = rsTmp.Fields("PA26")
      End If
      '93.10.4 END
      
      strSql = "INSERT INTO R060307 " & _
               "VALUES  ('" & strAgentNo & "','" & rsTmp.Fields("PA22") & "','" & rsTmp.Fields("PA11") & "','" & rsTmp.Fields("PA77") & "'," & DBDATE(m_strNP09) & "," & m_NextPA72 & ",'" & m_strCP01 & "','" & m_strCP02 & "','" & m_strCP03 & "','" & m_strCP04 & "','" & rsTmp.Fields("PA48") & "','" & strUserNum & "')"
      cnnConnection.Execute strSql
'      cnnConnection.CommitTrans
      AddWorkF = True
      '¦~¶O³æµ§¤£¶]µ§¼Æ
      m_AddWorkFCount = m_AddWorkFCount + 1
      Exit Function
   End If
   
'   Exit Function
'ErrorHandler:
      
'   cnnConnection.RollbackTrans
'   AddWorkF = False

End Function

'Added by Lydia 2020/08/17 ¼W¥[¦C¦LDebitNote
Private Function AddPrintDN(ByVal pSerialNo As String) As String
Dim strA1K28 As String '½Ð´Ú¹ï¶H
Dim strAgentNo As String '¥N²z¤H½s¸¹

    '¥N²z¤H½s¸¹
    strAgentNo = PUB_GetA1K03(m_strCP01, m_strCP02, m_strCP03, m_strCP04)
    '½Ð´Ú¹ï¶H
    strA1K28 = PUB_GetA1K28(m_strCP01, m_strCP02, m_strCP03, m_strCP04, m_strCP10)
    If strA1K28 = "" Then strA1K28 = strAgentNo

    PUB_AddNewDebitNoteList strUserNum, pSerialNo, "" & pub_AddressListSN, IIf(m_bolDNEmail, "Y", ""), IIf(m_bolDNPlusPaper, "Y", ""), IIf(m_OptPAID <> "", "Y", "")
   
    '¾ã§å¦C¦L:¥H½Ð´Ú¹ï¶HÀË¬d¬O§_¦s¦b©ó°ê¥~©T©w±H¶Ê´Ú³æ¥N²z¤HÀÉ(ACC225)¥B¤U¦¸±Hµo¤é´Á¡Ö¨t²Î¤é¡A­Y¦s¦b«h·s¼W¦C¦L²M³æ
    If pSerialNo <> "" And strA1K28 <> "" Then
      If m_KeyCP09 = "" Or Len(m_KeyCP09) > 9 Then  '§PÂ_¡G¾ã§å½Ð´Ú¨ç ©Î ¦~¶O¾ã§åµo¤å
            If PUB_ChkAcc225MsgList(pSerialNo, strA1K28, m_strCP01, m_strCP02, m_strCP03, m_strCP04, IIf(m_strCP10 = "601", "¦~ÃÒ¶O½Ð´Ú¨ç", "¦~¶O½Ð´Ú¨ç")) Then
            End If
      '¦~¶O³æµ§,ª½±µµoemail
      Else
            If PUB_ChkAcc225MsgList(pSerialNo, strA1K28, m_strCP01, m_strCP02, m_strCP03, m_strCP04, "") Then
            End If
      End If
    End If
    
End Function

'Added by Lydia 2020/08/17 ¥~±M©Ó¿ì³æ=>µoemail³qª¾
'Mark by Lydia 2020/08/17 §ï¦¨¦@¥Î¡A©ñ¦bbasLetter
'Private Sub GetFCPEmpMail(ByVal pKeyNo As String, ByVal pStrType As String, ByVal pCP148 As String, ByVal pPaid As String, ByVal pRecDate As String)
''pKeyNo : ¦¬¤å¸¹
''pStrType : ¬O§_e/E¤Æ
''pCP148 : ¯S®í½Ð´Ú³æ
''pPaid : ¤w¦¬´Ú(1-¤£±HD/N, 2-±HD/N)
''pRecDate : ·í¤Ñ½Ð´ÚY
'Dim strB1 As String, intB As Integer, strB2 As String
'Dim rsBD As New ADODB.Recordset
'Dim intR As Integer
'Dim strTo As String, strCC As String, strSubject As String, strContent As String, strSpeed As String
'Dim tmpArr1 As Variant
'
'    strB1 = "select cp01,cp02,cp03,cp04,cp10,decode(pa09,'000',cpm03,cpm04) cp10n," & _
'                "cp60,cp152,cp53,cp54,Substr(Nvl(Fa10,Cu10),1,3) Fcna01 ,pa75,pa26,pa27,pa28,pa29,pa30 " & _
'                "from caseprogress, patent,casepropertymap,fagent,customer " & _
'                "where cp09='" & pKeyNo & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) " & _
'                "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) "
'    intB = 1
'    Set rsBD = ClsLawReadRstMsg(intB, strB1)
'    If intB = 1 Then
'        '§ì©Ó¿ì³æ¯S®í³]©w(FcpEmpBill): ³Æµù¡B³t§O¡BEmail¦¬¥ó¤H+CC
'        Call PUB_GetFcpEMPBillSpec(rsBD.Fields("cp01") & rsBD.Fields("cp02") & rsBD.Fields("cp03") & rsBD.Fields("cp04"), "06", "" & rsBD.Fields("pa75"), _
'                                   rsBD.Fields("pa26") & "," & rsBD.Fields("pa27") & "," & rsBD.Fields("pa28") & "," & rsBD.Fields("pa29") & "," & rsBD.Fields("pa30"), , strB2, strSpeed, , , , strTo, strCC)
'        '---------
'        If pCP148 = "Y" Then
'            'ex: ¡i¤w°e¥ó-»âÃÒ¤ÎÃº¦~¶O/¦~¶O¡j¥i½Ð´Úµo¤å(email) Our Ref: FCP-0XXXXX [INCOM.601/605]
'            strSubject = "¡i¤w°e¥ó-" & rsBD.Fields("cp10n") & " X1¡j¥i½Ð´Úµo¤å( X2) Our Ref: " & rsBD.Fields("cp01") & "-" & rsBD.Fields("cp02") & IIf(rsBD.Fields("cp03") <> "0", "-" & rsBD.Fields("cp03"), "") & IIf(rsBD.Fields("cp04") <> "00", "-" & rsBD.Fields("cp04"), "") & " [INCOM." & rsBD.Fields("cp10") & "]"
'        Else
'
'        End If
'        'ex: ¡i¤w°e¥ó-»âÃÒ¤ÎÃº¦~¶O/¦~¶O¡j¥i½Ð´Úµo¤å(email) Our Ref: FCP-0XXXXX [INCOM.601/605]
'        'ex: ¡i¤w°e¥ó-»âÃÒ¤ÎÃº¦~¶O/¦~¶O-¯S®í½Ð´Ú¡j½Ð¶i¦æ½Ð´Ú Our Ref: FCP-0XXXXX [INCOM.601/605]
'        strSubject = "¡i¤w°e¥ó-" & rsBD.Fields("cp10n") & " X1¡j" & IIf(pCP148 = "Y", "½Ð¶i¦æ½Ð´Ú ( X2)", "¥i½Ð´Úµo¤å( X2)") & " Our Ref: " & rsBD.Fields("cp01") & "-" & rsBD.Fields("cp02") & IIf(rsBD.Fields("cp03") <> "0", "-" & rsBD.Fields("cp03"), "") & IIf(rsBD.Fields("cp04") <> "00", "-" & rsBD.Fields("cp04"), "") & " [INCOM." & rsBD.Fields("cp10") & "]"
'
'        'Àu¥ý¥D¦®: ¯S®í½Ð´Ú>·í¤Ñ½Ð´Ú>¤w¦¬´Ú>¤@¯ë
'        strB1 = ""
'        If pCP148 = "Y" Then strB1 = strB1 & ",¯S®í½Ð´Ú"
'        If pRecDate = "Y" Then strB1 = strB1 & ",·í¤Ñ½Ð´Ú"
'        If pPaid = "1" Then
'             strB1 = strB1 & ",¤w¦¬´Ú¤£±HD/N"
'        ElseIf pPaid = "2" Then
'            strB1 = strB1 & ",¤w¦¬´Ú±HD/N"
'        End If
'        If strB1 <> "" Then
'           strSubject = Replace(strSubject, " X1", "-" & Mid(strB1, 2))
'        Else
'           strSubject = Replace(strSubject, " X1", "")
'        End If
'        '³t§O: e/E¤Æ
'        strB1 = ""
'        If strSpeed <> "" Then
'             strB1 = strB1 & "," & strSpeed '§ì©Ó¿ì³æ³]©wFcpEmpBill
'        Else
'             '°Ñ¦Ò©Ó¿ì³æªº¹w³] (PUB_PrintFCPEmpBill)
'            If pStrType = "E" Then 'E+±H
'               If "" & rsBD.Fields("fcna01") = "011" Then '¤é¥»
'                  strSpeed = "Email+¥­«H"
'               ElseIf "" & rsBD.Fields("fcna01") = "101" Then '¬ü°ê
'                  strSpeed = "Email+±¾¸¹"
'               Else
'                  strSpeed = "Email+¥­«H"
'               End If
'            ElseIf pStrType = "e" Then 'E¤Æ
'               strSpeed = "Email"
'            Else '«DE¤Æ
'               If "" & rsBD.Fields("fcna01") = "011" Then '¤é¥»
'                  strSpeed = "±¾¸¹"
'               ElseIf "" & rsBD.Fields("fcna01") = "101" Then '¬ü°ê
'                  strSpeed = "±¾¸¹"
'               Else
'                  strSpeed = "¥­«H"
'               End If
'            End If
'            If "" & rsBD.Fields("fcna01") = "231" Then '¼w°ê
'               If pStrType = "E" Then 'E+±H
'                  strSpeed = "Email+±¾¸¹"
'               ElseIf pStrType = "e" Then 'E¤Æ
'                  strSpeed = "Email"
'               Else '«DE¤Æ
'                  strSpeed = "±¾¸¹"
'               End If
'            End If
'            '¥­«H§ï±¾¸¹ --¦]¬°2020/02/21¶l§½·~°È½Õ¾ã
'            If strSpeed <> "" Then strSpeed = Replace(strSpeed, "¥­«H", "±¾¸¹")
'            strB1 = strB1 & "," & strSpeed
'        End If
'        If strB1 = "" Or (pCP148 = "Y" And strSpeed = "") Then '¯S®í½Ð´Ú­n§PÂ_¦³µL©Ó¿ì³æ¯S®í³]©w
'           strSubject = Replace(strSubject, "( X2)", "")
'        Else
'           strSubject = Replace(strSubject, " X2", Mid(strB1, 2))
'        End If
'
'        '¤º¤å
'        intR = 1
'        '¹O´Á¸ÉÃº(¨Ó·½ªí³æªº³]©w¤§´y­z)
'        If m_OptMemo <> "" Then
'            strContent = strContent & intR & ". " & m_OptMemo & vbCrLf
'            intR = intR + 1
'        End If
'        If strB2 <> "" Then '§ì©Ó¿ì³æ³]©wFcpEmpBill
'           '­Y¦³´«¦æ¡A«e­±¥[3ªÅ¥Õ
'           strContent = strContent & intR & ". " & Replace(strB2, Chr(13) & Chr(10), Chr(13) & Chr(10) & "¡@") & vbCrLf
'           intR = intR + 1
'        ElseIf Not (pCP148 = "Y" Or pRecDate = "Y") Then
'           If pStrType = "e" Then '¤pe,Email
'           Else '¤jE,E+±H¡@¡F¯È¥»
'               strContent = strContent & intR & ". ¤w¥X¯È¥»¡Aµ{§Ç½Ð¾ã²z½Ð´Ú¯È¥»·|©Ó¿ì±H¥X¡C" & vbCrLf
'               intR = intR + 1
'           End If
'        End If
'
'        If pCP148 = "Y" Then
'            strContent = strContent & intR & ". ¯S®í½Ð´Ú¡A½Ð¶i¦æ¤H¤u½Ð´Ú§@·~¡C" & vbCrLf
'            intR = intR + 1
'        End If
'        If pRecDate = "Y" Then
'            strContent = strContent & intR & ". µ{§Ç¡G«Ý¦¬¾Ú¦Ü½Ð¾ã²z½Ð´Úªþ¥óemail©Ó¿ì±H¥X¤Îbackup¡C" & vbCrLf
'            intR = intR + 1
'        End If
'        If pStrType = "" Then '¯È¥»
'            strContent = strContent & intR & ". ½Ð¦L¦¬¾Ú¯È¥»¡A¦¬¾Ú¤U¸ü¤é¡G" & ChangeWStringToTDateString("" & rsBD.Fields("cp152")) & "¡C" & vbCrLf
'        Else
'            strContent = strContent & intR & ". ¦¬¾Ú¤U¸ü¤é¡G" & ChangeWStringToTDateString("" & rsBD.Fields("cp152")) & "¡C" & vbCrLf
'        End If
'        intR = intR + 1
'        strContent = strContent & intR & ". Ãº¶O¦~«×¡G²Ä" & IIf(Val("" & rsBD.Fields("cp53")) = 0, "  ", rsBD.Fields("cp53")) & "¦~¡ã²Ä" & IIf(Val("" & rsBD.Fields("cp54")) = 0, "  ", rsBD.Fields("cp54")) & "¦~" & vbCrLf
'        intR = intR + 1
'        If pStrType = "" And pCP148 <> "Y" Then  '«DE¤Æ¡A¥u¦s¯È¥»  'Modified by Lydia 2020/08/07 ¯S®í½Ð´Ú©ñ¦bTyping2
'           '½Ð´Ú©w½Z¡B±b³æBackup¦^¦s¨÷©v°Ï[paper.601/605] (¹ï©Ó¿ì¤H­û)
'            strContent = strContent & intR & ". ½Ð´Ú©w½Z¡B±b³æBackup¦^¦s¨÷©v°Ï[paper." & rsBD.Fields("cp10") & "]" & vbCrLf
'        ElseIf pCP148 = "Y" Then '¯S®í½Ð´Ú¤£¦sDN
'            strContent = strContent & intR & ". ½Ð´Ú©w½Z¤w¦sTyping2 ¡C" & vbCrLf
'        Else
'            strContent = strContent & intR & ". ½Ð´Ú©w½Z¡B±b³æ¤w¦sTyping2 ¡C" & vbCrLf
'        End If
'        intR = intR + 1
'        'Email¦¬¥ó¤H: «ü©w¤H­û¥N¸¹F001-FCP©Ó¿ìºÞ¨î¤HNA51¡BF002-FCP©Ó¿ìºÞ¨î¤H¤§¥DºÞ¡BF011-FCPµ{§ÇºÞ¨î¤HNA16¡BF012-FCPµ{§ÇºÞ¨î¤H¤§¥DºÞ¡C
'        If strTo = "" Then 'µL¯S®í³]©w¡A¤@¯ë
'           'Memo by Lydia 2020/08/13 Sharon: ·í¤Ñ½Ð´Ú>¯S®í½Ð´Ú>¤@¯ë(e/E/¯È¥»)
'           If pRecDate = "Y" Then '·í¤Ñ½Ð´ÚY
'                strTo = "F011;"
'           ElseIf pCP148 = "Y" Then '¯S®í½Ð´Ú³æ
'                strTo = "F001;"
'           ElseIf pStrType = "e" Then
'                strTo = "F001;"
'           ElseIf pStrType = "E" Then
'                strTo = "F001;F011;"
'           Else    '«DE¤Æ,¥u±H¯È¥»
'                strTo = "F011;"
'           End If
'        Else
'           strTo = Replace(strTo, ",", ";") & ";"
'           strTo = Replace(Replace(Replace(Replace(strTo, "01;", "F001;"), "02;", "F002;"), "11;", "F011;"), "12;", "F012;")
'        End If
'        'Email°Æ¥»¦¬¥ó¤H
'        If strCC = "" Then 'µL¯S®í³]©w¡A¤@¯ë
'           'Memo by Lydia 2020/08/13 Sharon: ·í¤Ñ½Ð´Ú>¯S®í½Ð´Ú>¤@¯ë(e/E/¯È¥»)
'           If pRecDate = "Y" Then '·í¤Ñ½Ð´ÚY
'                'strCC = "F012;" '8/13  µ{§Ç¥DºÞ¤£¥ÎCC
'           ElseIf pCP148 = "Y" Then '¯S®í½Ð´Ú³æ
'                strCC = "F002;F011;"
'           ElseIf pStrType = "e" Then
'                strCC = "F002;F011;"
'           ElseIf pStrType = "E" Then
'                strCC = "F002;"
'           Else    '«DE¤Æ,¥u±H¯È¥»
'                'strCC = "F012;"  '8/13  µ{§Ç¥DºÞ¤£¥ÎCC
'           End If
'        Else
'           strCC = Replace(strCC, ",", ";") & ";"
'           strCC = Replace(Replace(Replace(Replace(strCC, "01;", "F001;"), "02;", "F002;"), "11;", "F011;"), "12;", "F012;")
'        End If
'        '¦¬¥ó¤H·í¤¤¦³©Ó¿ì,¤~¹w³]CC¦^backup, ¦Û°ÊÂk¤J¨÷©v°Ï
'        If InStr(strTo & ";", "F001;") > 0 Or InStr(strTo & ";", "F002;") > 0 Then
'            strCC = strCC & "backup;"
'        End If
'
'        tmpArr1 = Split(strTo & ";" & strCC, ";")
'        For intB = 0 To UBound(tmpArr1)
'            If Trim(tmpArr1(intB)) <> "" And InStr("F001;F002;F011;F012;", Trim(tmpArr1(intB))) > 0 Then
'                If InStr(strTo & ";" & strCC, Trim(tmpArr1(intB))) > 0 Then '¥¼¸m´«
'                   strB1 = ""
'                   Select Case Trim(tmpArr1(intB))
'                       Case "F001"  'FCP©Ó¿ì
'                            strB1 = PUB_GetFCPSalesNo(rsBD.Fields("cp01"), rsBD.Fields("cp02"), rsBD.Fields("cp03"), rsBD.Fields("cp04"), rsBD.Fields("cp10"))
'                            strTo = Replace(strTo, "F001;", strB1 & ";")
'                            strCC = Replace(strCC, "F001;", strB1 & ";")
'                       Case "F002"  'FCP©Ó¿ì¥DºÞ
'                            strB1 = PUB_GetFCPProSup(PUB_GetFCPSalesNo(rsBD.Fields("cp01"), rsBD.Fields("cp02"), rsBD.Fields("cp03"), rsBD.Fields("cp04"), rsBD.Fields("cp10")))
'                            strTo = Replace(strTo, "F002;", strB1 & ";")
'                            strCC = Replace(strCC, "F002;", strB1 & ";")
'                       Case "F011"  'FCPµ{§Ç
'                            strB1 = PUB_GetFCPHandler(rsBD.Fields("cp01"), rsBD.Fields("cp02"), rsBD.Fields("cp03"), rsBD.Fields("cp04"), rsBD.Fields("cp10"))
'                            strTo = Replace(strTo, "F011;", strB1 & ";")
'                            strCC = Replace(strCC, "F011;", strB1 & ";")
'                       Case "F012"  'FCPµ{§Ç¥DºÞ
'                            strB1 = PUB_GetFCPProSup(PUB_GetFCPHandler(rsBD.Fields("cp01"), rsBD.Fields("cp02"), rsBD.Fields("cp03"), rsBD.Fields("cp04"), rsBD.Fields("cp10")))
'                            strTo = Replace(strTo, "F012;", strB1 & ";")
'                            strCC = Replace(strCC, "F012;", strB1 & ";")
'                   End Select
'                End If
'            End If
'        Next intB
'
'        strB1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'                    "values('" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'                    ",'" & ChgSQL(strSubject) & "','" & ChgSQL(strContent) & "'," & CNULL(strCC) & ")"
'        cnnConnection.Execute strB1, intI
'    End If
'
'    Set rsBD = Nothing
'End Sub

