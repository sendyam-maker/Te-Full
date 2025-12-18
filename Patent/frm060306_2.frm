VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060306_2 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "½Ð´Ú³qª¾¨ç-¦~ÃÒ¶O½Ð´Ú¨ç"
   ClientHeight    =   5430
   ClientLeft      =   1410
   ClientTop       =   1080
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7080
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   3
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3870
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   " ¦^«eµe­±(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4830
      TabIndex        =   3
      Top             =   30
      Width           =   1155
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   2
      Left            =   1296
      MaxLength       =   1
      TabIndex        =   0
      Top             =   4170
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   1
      Left            =   1296
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3870
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   2436
      Left            =   48
      TabIndex        =   5
      Top             =   480
      Width           =   6975
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   7
         Left            =   1260
         TabIndex        =   34
         Top             =   2205
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   6
         Left            =   1260
         TabIndex        =   33
         Top             =   1965
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   5
         Left            =   1260
         TabIndex        =   32
         Top             =   1740
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   4
         Left            =   1260
         TabIndex        =   31
         Top             =   1500
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   3
         Left            =   1260
         TabIndex        =   30
         Top             =   1245
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   2
         Left            =   1260
         TabIndex        =   29
         Top             =   1005
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   1
         Left            =   1260
         TabIndex        =   28
         Top             =   405
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label2 
         Height          =   180
         Index           =   0
         Left            =   1260
         TabIndex        =   27
         Top             =   150
         Width           =   5595
         VariousPropertyBits=   268435483
         Caption         =   "Label2"
         Size            =   "9869;317"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   285
         Left            =   1260
         TabIndex        =   25
         Top             =   630
         Width           =   5655
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "9975;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥»©Ò®×¸¹"
         Height          =   180
         Index           =   0
         Left            =   84
         TabIndex        =   16
         Top             =   168
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "½Ð´Ú¨ç¤é´Á"
         Height          =   180
         Index           =   1
         Left            =   84
         TabIndex        =   15
         Top             =   408
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "±M§Q¦WºÙ"
         Height          =   180
         Index           =   2
         Left            =   84
         TabIndex        =   14
         Top             =   648
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(¤¤)"
         Height          =   180
         Index           =   3
         Left            =   84
         TabIndex        =   13
         Top             =   1008
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(¤é)"
         Height          =   180
         Index           =   4
         Left            =   84
         TabIndex        =   12
         Top             =   1488
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(­^)"
         Height          =   180
         Index           =   5
         Left            =   84
         TabIndex        =   11
         Top             =   1248
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(¤¤)"
         Height          =   180
         Index           =   6
         Left            =   84
         TabIndex        =   10
         Top             =   1728
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(¤é)"
         Height          =   180
         Index           =   7
         Left            =   84
         TabIndex        =   9
         Top             =   2208
         Width           =   936
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(­^)"
         Height          =   180
         Index           =   8
         Left            =   84
         TabIndex        =   8
         Top             =   1968
         Width           =   936
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4050
      TabIndex        =   2
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   6012
      TabIndex        =   4
      Top             =   36
      Width           =   800
   End
   Begin VB.TextBox Text2 
      Height          =   264
      Index           =   0
      Left            =   1416
      MaxLength       =   1
      TabIndex        =   1
      Top             =   5112
      Width           =   255
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   1290
      TabIndex        =   38
      Top             =   3645
      Width           =   5745
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10134;317"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   11
      Left            =   1290
      TabIndex        =   37
      Top             =   3405
      Width           =   5745
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10134;317"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   1290
      TabIndex        =   36
      Top             =   3165
      Width           =   5745
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "10134;317"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   8
      Left            =   90
      TabIndex        =   35
      Top             =   2970
      Width           =   4005
      VariousPropertyBits=   268435483
      Caption         =   "Label2"
      Size            =   "7064;317"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   570
      Index           =   0
      Left            =   1290
      TabIndex        =   26
      Top             =   4470
      Width           =   5715
      VariousPropertyBits=   -1466939365
      MaxLength       =   2000
      ScrollBars      =   3
      Size            =   "10081;1005"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_¯S®í½Ð´Ú         (Y:¬O)"
      Height          =   180
      Index           =   10
      Left            =   4260
      TabIndex        =   24
      Top             =   3912
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¤w¦¬´Ú                      (1-¤£±HD/N, 2-±HD/N)"
      Height          =   180
      Index           =   16
      Left            =   90
      TabIndex        =   23
      Top             =   4212
      Width           =   3150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó©Ê½è"
      Height          =   180
      Index           =   15
      Left            =   96
      TabIndex        =   22
      Top             =   3672
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¦~¶O¦Û°Ê¥NÃº          (Y:¦Û°Ê¥NÃº)"
      Height          =   180
      Index           =   9
      Left            =   96
      TabIndex        =   21
      Top             =   3912
      Width           =   2532
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "P. S."
      Height          =   180
      Index           =   11
      Left            =   96
      TabIndex        =   20
      Top             =   4470
      Width           =   312
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "°Æ¥»¦¬¨ü¤H"
      Height          =   180
      Index           =   12
      Left            =   96
      TabIndex        =   19
      Top             =   3192
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "°Æ¥»Ápµ¸¤H"
      Height          =   180
      Index           =   13
      Left            =   96
      TabIndex        =   18
      Top             =   3432
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_­×§ï½Ð´Ú¨ç        (Y)"
      Height          =   180
      Index           =   14
      Left            =   96
      TabIndex        =   17
      Top             =   5112
      Width           =   1860
   End
End
Attribute VB_Name = "frm060306_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/10/26 ¤é¤å¤w§ï©ñ©w½Z¤º
'Memo By Sindy 2021/7/16 Form2.0¤w­×§ï
'Memo By Morgan 2012/12/10 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo by Morgan2010/12/27 ¥Ó½Ð®×¸¹Äæ¤w­×§ï
'2010/12/6 memo by sonia ­û¤u½s¸¹Äæ¤w­×§ï
'Memo by Morgan2010/8/13 ¤é´ÁÄæ¤w­×§ï
Option Explicit

Dim intWhere As Integer, strReceiveNo As String
Dim m_LetterLanguage As String
Dim m_LetterKind As Integer
'edit by nickc 2007/02/02
'Dim m_CP10 As String, pA(1 To T_PA) As String
Dim m_CP10 As String, pa() As String
Dim m_strNP09 As String '¤U¦¸Ãº¶O¤é(ªk©w)
Dim m_strNP08 As String '¤U¦¸Ãº¶O¤é(¥»©Ò)
Dim m_strNP23 As String '¤U¦¸Ãº¶O¤é(¬ù©w) 'Add By Sindy 2021/4/27
Const ET01 As String = "10"
'Add by Morgan 2004/6/25
Dim m_bolNew As Boolean '¬O§_¥Î·sªk
Dim m_bol412 As Boolean '¬O§_¦³µo¤å©µ½w¤½§i
Dim m_NationID As String 'Add By Sindy 2017/2/17


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
'Modified by Lydia 2019/08/15 strTxt(1 To 6)=>strTxt(1 To 20) ex.FCP-59547 ¶W¥X¯Á¤Þ
Dim strTxt(1 To 20) As String, i As Integer, j As Integer, strTmp As String
'Add By Cheng 2003/02/13
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
Dim strTemp1
Dim strTemp2
Dim strSendDate As String
Dim intYear As Integer
Dim strAnnuity As String
Dim s As Integer
    
    ii = 0
    EndLetter ET01, strReceiveNo, ET03, strUserNum
    'Add By Cheng 2003/02/24
   '½Ð´Ú¨ç¤é´Á
   If frm060306.Text5.Text <> "" Then
      ii = ii + 1
      strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç¤é´Á','" & DBDATE(frm060306.Text5.Text) & "')"
   End If
    '½Ð´Ú¨ç³Æµù
    If Text1(0).Text <> "" Then
       'Added by Lydia 2019/10/16 »P¾ã§å(frm060307)¤@­P¡A¦³P.S´N§R±¼¡¨along with our debit note¡¨
       ii = ii + 1
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦³½Ð´Ú¨ç³Æµù®É¤£¦L','¡ð')"
       'end 2019/10/16
       ii = ii + 1
       strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
          "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','½Ð´Ú¨ç³Æµù','P.S. " & ChgSQL(Text1(0).Text) & "')"
    End If
    'Add By Cheng 2003/02/13
    '¨ú±oÃº¦~¶O¬ÛÃö¸ê®Æ
    StrSQLa = "Select * From Patent,Caseprogress Where CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And CP09='" & strReceiveNo & "' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        '­Y¦³Ãº¦~¶O¤é´Á
        If "" & rsA("PA73").Value <> "" Then
           strTemp1 = Split(UCase(rsA("PA72").Value), ",")
           strTemp2 = Split(UCase(rsA("PA73").Value), ",")
           For i = 0 To UBound(strTemp2)
              If Val(strTemp2(i)) = "" & rsA("CP27").Value Then
                 If Val(strSendDate) <> "" & rsA("CP27").Value Then
                 
                     'Modify by Morgan 2004/10/12 ¥[¤é¤å©w½Z
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
        
            'Modify by Morgan 2004/10/12 ¦Ò¼{¤é¤å©w½Z
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
               'Modify by Morgan 2004/10/12 ¦Ò¼{¤é¤å©w½Z
               'Modify by Morgan 2016/8/18 +¤¤¤å
               'If m_LetterLanguage = 3 Then
               If m_LetterLanguage <> 2 Then
                  strAnnuity = "1"
               Else
                  strAnnuity = "1st"
               End If
          '­Y¦³Ãº¦~¶O¤é´Á
           Else
               'Modify by Morgan 2004/10/12 ¦Ò¼{¤é¤å©w½Z
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
           "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','²Ä´X¦~¦Ü´X¦~¶O'," & CNULL(strAnnuity) & ")"
    End If
    
   If m_strNP09 <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦~¶Oªk©w´Á­­'," & DBDATE(m_strNP09) & ")"
      ii = ii + 1
      'Modify By Sindy 2021/4/27
      If strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦~¶O¬ù©w´Á­­'," & DBDATE(m_strNP23) & ")"
      Else
      '2021/4/27 END
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦~¶O¥»©Ò´Á­­'," & DBDATE(m_strNP08) & ")"
      End If
   End If
    
    'Add by Morgan 2004/6/25
    '»âÃÒ¥B¦³µo¤å©µ½w¤½§i
    If m_bol412 = True Then
         ii = ii + 1
         'Modified by Morgan 2022/10/26
         'If m_LetterLanguage = 3 Then
         '   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','©µ½w¤½§i','¡@¤S¡B®Æª÷¯Ç¥IÇU¤â“dþàÇO¦P®ÉÇR¡Bµn“÷¤½³øÇU´¦¸üÇy…üÇpþîÇr¦®Çy¦P§½ÇR¤W¥Ó­PþêÇeþêþò¡C " & "')"
         'Else
         '   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','©µ½w¤½§i','    We have also filed a request for delaying publication of allowed claim(s)." & Chr(13) & "')"
         'End If
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','©µ½w¤½§i­n¦L','¡ð')"
         'end 2022/10/26
    End If
    
   'Added by Morgan 2016/10/13
   If ChangeCustomerL(pa(75)) = "Y54391000" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','µo²¼','µo²¼')"
   
   End If
   'end 2016/10/13
    
'Modified by Morgan 2019/11/5 ³]­p±M¥Î´Á¤w§ó·s,­^¤å¦~¶O½Ð´Ú¨ç¦X¨Ö
'   'Added by Morgan 2019/8/6
'   '108.11.1·sªk³]­p®×±M¥Î´Á¥Ñ12¦~©µªø¬°15¦~
'   If pa(8) = "3" Then
'      '±M¥Î´Á§ó·s«e¯S®í±±¨î,§ó·s«á¥i§ï³£§ì±M¥Î´Á¤î¤é
'      If strSrvDate(1) < 20191101 Then
'         'strExc(1) = CompDate(2, -1, CompDate(0, 15, pa(10))) 'Removed by Morgan 2019/8/16 §ï¼g¦@¥Î¨Ò¥~Äæ¦ì
'
'         '±M¥Î´Á§ó·s«e¯S®í±±¨î,§ó·s«á¥i³s¦P©w½Z¤º¨Ò¥~Äæ¦ì¤@¨Ö²¾°£
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'            "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','108/10/31«e³]­p®×¤£¦L','¡ð')"
'
'      Else
'         'strExc(1) = DBDATE(pa(25)) 'Removed by Morgan 2019/8/16 §ï¼g¦@¥Î¨Ò¥~Äæ¦ì
'      End If
'
'      'Removed by Morgan 2019/8/16 §ï¼g¦@¥Î¨Ò¥~Äæ¦ì
'      'ii = ii + 1
'      'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','³]­p®×15¦~©¡º¡¤é','" & strExc(1) & "')"
'
'   End If
'   'end2019/8/6
   
   '¦Û°Ê¥NÃº
   If Text2(1) = "Y" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¦Û°Ê¥NÃº¤£¦L','¡ð')"
   End If
   
   '¤w¦¬´Ú
   If Text2(2) = "Y" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','¤w¦¬´Ú¤£¦L','¡ð')"
   End If
   
   '³Ì«á¤@¦~
   If m_LetterKind = "4" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','³Ì«á¤@¦~¤£¦L','¡ð')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','³Ì«á¤@¦~¤~¦L','¡ð')"
   End If
'end 2019/11/5
          
'Removed by Morgan 2015/10/12 ©w½Z³æ¤@¤Æ,¨ú®ø--¦¿¦p¥É
'   'Add by Morgan 2008/4/28 ­^¤å¦~¶O½Ð´Ú¨ç X49199001 ¤£¥Îªþ©x¤è¦¬¾Ú
'   If pa(26) = "X49199001" Then
'      ii = ii + 1
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
'         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','ªþ¥ó','our debit note')"
'   End If
'   'end 2008/4/28
   
    If ii > 0 Then
       'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
       'If Not objLawDll.ExecSQL(ii, strTxt) Then
       If Not ClsLawExecSQL(ii, strTxt) Then
          MsgBox "Àx¦s¨Ò¥~Äæ¦ì¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
       End If
    End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim bolChk As Boolean
   Dim strTmp As String
   Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
   Dim strCaseNo As String, strFileName As String, strFullFileName As String 'Add By Sindy 2017/2/17
   Dim nFrm As Form 'Added by Lydia 2020/08/17
   
   Select Case Index
      Case 2
         Unload frm060306
         Unload Me
      Case 0
        Screen.MousePointer = vbHourglass
          'Added by Lydia 2020/08/17 ²Î¤@¨ì¦~ÃÒ¶O½Ð´Ú¨ç°õ¦æ
          If pa(1) = "FCP" And (m_CP10 = "601" Or m_CP10 = "605") Then
                'ÀË¬dªí³æ¬O§_¤w¶}±Ò¡A­Y¬O¡A«hÃö³¬
                For Each nFrm In Forms
                   If StrComp(nFrm.Name, "frm060307", vbTextCompare) = 0 Then
                      Unload frm060307
                      Exit For
                   End If
                Next
                frm060307.m_KeyCP09 = strReceiveNo
                frm060307.m_KeyCP10 = m_CP10
                Call frm060307.SetData(0, "3", True) '¥~³¡©I¥s,¹w³]Ãþ§O
                Call frm060307.SetData(1, Text2(2)) '¤w¦¬´Ú
                frm060307.Show
                Call frm060307.cmdOK_Click(0)
                Unload frm060307
          Else
          'end 2020/08/17
                 If Text2(0).Text = "Y" Then bolChk = True
                 '¨ú±o©w½Z»y¤å
                 'Modify by Morgan 2006/5/25
                 'm_LetterLanguage = GetLetterLanguage(pA(1), pA(2), pA(3), pA(4))
                 m_LetterLanguage = PUB_GetLanguage(pa(1), pa(2), pa(3), pa(4), m_CP10, "1")
                 '¨ú±o©w½Z»y¤å­^¤å¤§¤ÀÃþ
                 'Modify by Morgan 2004/10/12 ¥[¤é¤å
                 'If m_LetterLanguage = 2 Then
                 If m_LetterLanguage = 2 Or m_LetterLanguage = 3 Then
                    m_LetterKind = CetLetterKind(pa(1), pa(2), pa(3), pa(4), m_CP10)
                 End If
                 Select Case m_CP10
                    Case "602", "603"
                       Select Case m_LetterLanguage
                          Case "1":    ' ¤¤¤å
                             strTmp = "01"
                          Case "2":    ' ­^¤å
                             strTmp = "05"
                          Case "3":    ' ¤é¤å
                             strTmp = "06"
                       End Select
                    Case "601"
                       '½Ð¦A°Ï¤À¤¤¤å01, ­^¤å(¤@¯ë02©Î¦~¶O¦Û°Ê¥NÃº03©Î³Ì«á¤@¦~04), ¤é¤å06
                       Select Case m_LetterLanguage
                          Case "1":    ' ¤¤¤å
                             strTmp = "01"
                             
                          Case "2":    ' ­^¤å
                             'Modified by Morgan 2019/11/5 ­^¤å½Ð´Ú¨ç¤w¦X¨Ö
                             'Select Case m_LetterKind
                             '   Case "1":    '¤@¯ë
                             '      strTmp = "02"
                             '   Case "2":    '¦~¶O¦Û°Ê¥NÃº
                             '      strTmp = "03"
                             '   Case "4":    '³Ì«á¤@¦~
                             '      strTmp = "04"
                             'End Select
                             strTmp = "15"
                             'end 2019/11/5
                             
                          Case "3":    ' ¤é¤å
                             'Add by Morgan 2004/10/12 ¦Ò¼{¤é¤å©w½Z
                             strTmp = "06"
                             
                             'Add by Morgan 2005/7/20
                             '­Y¬O§_¤w¦¬´Ú¦³¤W"Y"
                             If Me.Text2(2).Text = "Y" Then
                                strTmp = "10"
                             End If
                             '2005/7/20 end
                             
                       End Select
                       
                    Case "605"
                       
                       Select Case m_LetterLanguage
                          Case "1":    ' ¤¤¤å
                             strTmp = "01"
                             
                          Case "2":    ' ­^¤å
                             'Modified by Morgan 2019/11/5 ­^¤å½Ð´Ú¨ç¤w¦X¨Ö
                             'Select Case m_LetterKind
                             '   Case "1":    '¤@¯ë
                             '      strTmp = "02"
                             '   Case "2":    '¦~¶O¦Û°Ê¥NÃº
                             '      strTmp = "03"
                             '   Case "4":    '³Ì«á¤@¦~
                             '      strTmp = "04"
                             'End Select
                             strTmp = "02"
                             'end 2019/11/5
                             
                          Case "3":    ' ¤é¤å
                             'Add by Morgan 2004/10/12 ¦~¶O¦Û°Ê¥NÃº
                             If m_LetterKind = 2 Or Text2(1) = "Y" Then
                                strTmp = "09"
                                'Add by Morgan 2005/7/20
                                '­Y¬O§_¤w¦¬´Ú¦³¤W"Y"
                                If Me.Text2(2).Text = "Y" Then
                                   strTmp = "11"
                                End If
                                '2005/7/20 end
                             Else
                                strTmp = "06"
                                'Add by Morgan 2005/7/20
                                '­Y¬O§_¤w¦¬´Ú¦³¤W"Y"
                                If Me.Text2(2).Text = "Y" Then
                                   strTmp = "10"
                                End If
                                '2005/7/20 end
                             End If
                       End Select
                 End Select
                 
                 'Removed by Morgan 2019/11/5 ­^¤å½Ð´Ú¨ç¤w¦X¨Ö
                 ''·í©w½Z»y¤å¬°­^¤å¥B¦~¶O¬°¤@¯ë©Î¦Û°Ê¥NÃº®É
                 'If m_LetterLanguage = "2" Then
                 '   If (m_LetterKind = "1" Or m_LetterKind = "2") Then
                 '      '­Y¦~¶O¦Û°Ê¥NÃº¦³¤W"Y"
                 '      If Text2(1) = "Y" Then
                 '          'Modify By Cheng 2003/02/18
                 '          '­Y¬O§_¤w¦¬´Ú¦³¤W"Y"
                 '          If Me.Text2(2).Text = "Y" Then
                 '              strTmp = "08"
                 '          Else
                 '              strTmp = "03"
                 '          End If
                 '      '­Y¦~¶O¦Û°Ê¥NÃº¨S¤W"Y"
                 '      Else
                 '          'Modify By Cheng 2003/02/18
                 '          '­Y¬O§_¤w¦¬´Ú¦³¤W"Y"
                 '          If Me.Text2(2).Text = "Y" Then
                 '              strTmp = "07"
                 '          Else
                 '              strTmp = "02"
                 '          End If
                 '      End If
                 '   End If
                 'End If
                 
                 ''Add by Morgan 2004/6/25   ·sªk©w½Z
                 'If m_bolNew = True Then
                 '   If m_CP10 = "601" And InStr("02,03,04,07,08", strTmp) > 0 Then
                 '      strTmp = Format(Val(strTmp) + 10, "00")
                 '      'Modify by Morgan 2004/7/20 ­YµL¤½§i¤é¥t¥X©w½Z
                 '      If Val(pa(14)) = 0 Then
                 '         '¤@¯ë
                 '         If InStr("12,13,14", strTmp) > 0 Then
                 '            strTmp = "15"
                 '         '¤w¦¬´Ú
                 '         Else
                 '            strTmp = "16"
                 '         End If
                 '      End If
                 '   End If
                 'End If
                 
                 ''Add by Morgan 2005/9/22 ³Ì«á¤@¦~¤]­n°Ï¤À¬O§_¤w¦¬´Ú
                 'If (m_CP10 = "605" And strTmp = "04") Then
                 '   '­Y¬O§_¤w¦¬´Ú¦³¤W"Y"
                 '   If Me.Text2(2).Text = "Y" Then
                 '      strTmp = "12"
                 '   End If
                 'End If
                 'end 2019/11/5
                                   
                 StartLetter ET01, strTmp
                 
                 'Modify by Morgan 2008/3/24 §PÂ_¬O§_²£¥Í¹q¤lÀÉ
                 'NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum
                 bolEmail = PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4), IIf(m_CP10 = "605", True, False), , bolPlusPaper)
                 'Add by Morgan 2009/10/20 +§PÂ_¬O§_EMail¦P®É±H¯È¥»
                 If bolPlusPaper Then
                    iCopy = 0
                 Else
                    iCopy = 1
                 End If
                 'end 2009/10/20
                 If bolEmail Then
                    NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum, , , , , iCopy, , True, True
                    'Add By Sindy 2017/2/17 ¤é¥»°Ï­n²£¥ÍPDF¹q¤lÀÉ¦Ü­Ó®×¸ê®Æ§¨
                    Call PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4), m_CP10, m_NationID)
                    If Left(m_NationID, 3) = "011" Then
                       strCaseNo = pa(1) & pa(2) & IIf(pa(3) & pa(4) <> "000", pa(3) & pa(4), "")
                       strFileName = PUB_GetEFilePath(pa(1)) & "\" & pa(1) & "\" & Left(pa(2), 3) & "\" & strCaseNo & "\" & strCaseNo & "_" & strSrvDate(1) & "." & m_CP10 & ".CUS.PDF"
                       'Modified by Lydia 2019/10/22 FCP-59598°µ¯S®í½Ð´Ú³æ,¥h±¼³Æµù¤º®e©M­n­×§ï½Ð´Ú¨çText2(0)=Y=>LD16=* ; ¦]¬°ld16¤w¤W*, µLªk²£¥Íapp.path\*.pdf
                       'Call PUB_PrintLetter(strReceiveNo, , , True, strFullFileName, False)
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
                    '2017/2/17 END
                    If bolChk = False Then
                       MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(pa(1)) & " ]¡I"
                    End If
                 Else
                    NowPrint strReceiveNo, ET01, strTmp, bolChk, strUserNum
                 End If
                 'end 2008/3/31
                 
                 '·s¼W¦a§}±ø¦Cªí¸ê®Æ
                 If Not bolEmail Or bolPlusPaper Then
        '            'Add By Sindy 2015/9/21 ¤é¤å©w½Z¤~­n¦L¦a§}±ø
        '            If m_LetterLanguage = "3" Or Val(¥~±M¶}µ¡«H¨ç±Ò¥Î¤é) >= Val(strSrvDate(1)) Then
        '            '2015/9/21 END
                    'Add By Sindy 2017/3/20 ¤é¤å©w½Z¤~­n¦L¦a§}±ø
                    If frm060306.m_FCna01 = "101" Or m_LetterLanguage = "3" Then '¬ü°ê ©Î ¤é¤å©w½Z¤~­n¦L¦a§}±ø
                    '2017/3/20 END
                       pub_AddressListSN = pub_AddressListSN + 1
                       PUB_AddNewAddressList strUserNum, frm060306.Text1.Text, frm060306.Text2.Text, frm060306.Text3.Text, frm060306.Text4.Text, "" & pub_AddressListSN, "0", m_CP10
                    End If
                 End If
          End If 'Added by Lydia 2020/08/17
          
         frm060306.Show
         frm060306.Clear
        Screen.MousePointer = vbDefault
         Unload Me
      Case 1
         frm060306.Show
         Unload Me
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060306_2 = Nothing
End Sub

Private Sub Form_Load()
   
   MoveFormToCenter Me
   intWhere = °ê¥~_FC
   ReadPatent
   m_bolNew = False: m_bol412 = False
   If Val(pa(14)) = 0 Or Val(pa(14)) >= 930701 Then
      m_bolNew = True
      m_bol412 = PUB_Check412(pa)
   End If

End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, strTmp As String
 Dim varTmp As Variant, i As Integer, iStart As Integer, iEnd As Integer
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
            
            'Modified by Lydia 2020/08/13 +CP148
            strExc(0) = "SELECT CP27,CP148 FROM CASEPROGRESS WHERE CP09='" & strReceiveNo & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Text2(3).Text = "" & RsTemp.Fields("CP148") 'Added by Lydia 2020/08/13
               varTmp = Split(pa(73), ",")
               iStart = 0
               iEnd = 0
               For i = 0 To UBound(varTmp)
                    'Modify By Cheng 2003/02/13
                    '¥Î¦r¦ê¤ñ¸û
'                  If rsTemp.Fields(0) = Format(varTmp(i)) Then
                  If "" & RsTemp.Fields(0) = Format(varTmp(i)) Then
                     If iStart = 0 Then iStart = i + 1
                     iEnd = i + 1
                  End If
               Next
               '92.3.17 ADD BY SONIA
               If iStart = iEnd And iStart = 0 Then
                  iStart = i: iEnd = i
               End If
               '92.3.17 END
                'Modify By Cheng 2003/02/13
                'Ãº¶O¦~«×Åã¥Ü®æ¦¡
                '­Y¥uÃº¤@¦~
                If iStart = iEnd Then
                    Label2(8) = "Ãº¶O¦~«×²Ä " & iEnd & " ¦~¦~¶O"
                '­Y¦P®ÉÃº¦h¦~
                Else
                    Label2(8) = "Ãº¶O¦~«×²Ä " & iStart & " ¦Ü " & iEnd & " ¦~¦~¶O"
                End If
            End If
            'Modify By Cheng 2003/02/10
            '­Y¦³°Æ¥»¦¬¨ü¤H
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
            'Modify By Cheng 2003/02/10
            '­Y¦³°Æ¥»¦¬¨ü¤H
            If pa(35) <> "" Then
                'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
                'If objLawDll.LawGetName(pa(35), strTmp) Then Label2(10) = strTmp
                If ClsLawLawGetName(pa(35), strTmp) Then Label2(10) = strTmp
            End If
            Label2(11) = pa(36)
         End If
   End Select
   
   m_CP10 = ""
   strExc(0) = "SELECT CPM03,CP10 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE CP09='" & strReceiveNo & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Label2(12) = RsTemp.Fields(0)
      m_CP10 = RsTemp.Fields(1)
   End If
   
   m_LetterKind = CetLetterKind(pa(1), pa(2), pa(3), pa(4), m_CP10)
   If m_LetterKind = "2" Then Text2(1) = "Y"

End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Added by Lydia 2020/08/17 ¤w¦¬´Ú  (1-¤£±HD/N, 2-±HD/N)
   '¦~¶O¦Û°Ê¥NÃºText2(1)Äæ¦ìÂê©w,§ï¦¨¼Ò²Õ§PÂ_
   If Index = 2 Then
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   Else
   'end 2020/08/17
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If 'Added by Lydia 2020/08/17
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
   
   
'Removed by Morgan 2016/5/13 ¥Ø«e¥u¦³­l¥Í³]­p¶W¹L9½X¥Î¤@¯ë©w½Z--David
'   '­Y¥Ó½Ð®×¸¹½X¼Æ¶W¹L8½X«h¬°°l¥[Áp¦X®×
'   If IsNull(pa(11)) = False Then
'      'Modify by Morgan 2010/12/28 ¥Ó½Ð®×¸¹§ï½X¼Æ
'      'If Len(pa(11)) > 8 Then
'      If Len(pa(11)) > 9 Then
'         nKind = 3
'         GoTo EXITSUB
'      End If
'   End If
'end 2016/5/13
   
   '¬O§_³Ì«á¤@¦~
   'Modify by Morgan 2005/9/29 ¥[¥»©Ò´Á­­
   'Modified by Morgan 2015/5/29 +¤£Äò¿ì¥¼¹O´Á±ø¥ó --·¶ªÚ Ex.FCP-44244
   'Modify By Sindy 2021/4/27 + ,MAX(NP23)
   strSql = "SELECT MAX(NP09),MAX(NP08),MAX(NP23) FROM NEXTPROGRESS " & _
            "WHERE NP02 = '" & pa(1) & "' AND " & _
                  "NP03 = '" & pa(2) & "' AND " & _
                  "NP04 = '" & pa(3) & "' AND " & _
                  "NP05 = '" & pa(4) & "' AND NP07 = '605' AND (NP06 IS NULL or (NP06='N' and np09>" & strSrvDate(1) & "))"
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount = 1 And IsNull(rsTmp.Fields(0)) Then
      nKind = 4
   End If
   m_strNP09 = "" & rsTmp.Fields(0)
   m_strNP08 = "" & rsTmp.Fields(1)
   m_strNP23 = "" & rsTmp.Fields(2) 'Add By Sindy 2021/4/27
   
   Set rsTmp = Nothing
      
   If nKind = 4 Then
        'Modify By Cheng 2003/02/10
'      rsTmp.Close
        Set rsTmp = Nothing
      GoTo EXITSUB
   Else
      ' PA70 = Y ®É¬°¦Û°Ê¥NÃº   '¦¹³B¥u³qª¾«áÄò¦~¶O¬O§_¦Û°Ê¥NÃº, ©Ò¥H¤£ºÞ»âÃÒ¬O§_¦Û°Ê¥NÃº
      If IsNull(pa(70)) = False Then
         If pa(70) = "Y" Then
            nKind = 2
            'Modify By Cheng 2003/02/10
'            rsTmp.Close
            Set rsTmp = Nothing
            GoTo EXITSUB
         End If
      End If
      ' PA75 ¨ú¥N²z¤HÀÉ
      If IsNull(pa(75)) = False Then
         If IsEmptyText(pa(75)) = False Then
            If Len(pa(75)) > 8 Then
               strFA01 = Mid(pa(75), 1, 8)
               strFA02 = Mid(pa(75), 9, 1)
            Else
               strFA01 = pa(75) & String(8 - Len(pa(75)), "0")
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
                    'Modify By Cheng 2003/02/10
'                     rsTmp.Close
                    Set rsTmp = Nothing
                     GoTo EXITSUB
                  End If
               End If
            End If
            rsSubTmp.Close
            
            nKind = 1
            'Modify By Cheng 2003/02/10
'            rsTmp.Close
            Set rsTmp = Nothing
            GoTo EXITSUB
         End If
      End If
      ' ¨ú«È¤áÀÉ
      If IsNull(pa(26)) = False Then
         If IsEmptyText(pa(26)) = False Then
            If Len(pa(26)) > 8 Then
               strCU01 = Mid(pa(26), 1, 8)
               strCU02 = Mid(pa(26), 9, 1)
            Else
               strCU01 = pa(26) & String(8 - Len(pa(26)), "0")
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
                    'Modify By Cheng 2003/02/10
'                     rsTmp.Close
                    Set rsTmp = Nothing
                     GoTo EXITSUB
                  End If
               End If
            End If
            rsSubTmp.Close
            
            nKind = 1
            'Modify By Cheng 2003/02/10
'            rsTmp.Close
            Set rsTmp = Nothing
            GoTo EXITSUB
         End If
      End If
   End If
   
EXITSUB:
   If nKind = 0 Then nKind = 1
   Set rsSubTmp = Nothing
   Set rsTmp = Nothing
   CetLetterKind = nKind
End Function
