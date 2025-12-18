VERSION 5.00
Begin VB.Form frm050316 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "¥N²z¤H®×¥óÁ`Ã¯"
   ClientHeight    =   4275
   ClientLeft      =   3690
   ClientTop       =   1830
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4470
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   10
      Top             =   2700
      Width           =   255
   End
   Begin VB.OptionButton opt 
      Caption         =   "µo¤å¤é"
      Height          =   300
      Index           =   1
      Left            =   1470
      TabIndex        =   6
      Top             =   1650
      Width           =   1044
   End
   Begin VB.OptionButton opt 
      Caption         =   "¦¬¤å¤é"
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1650
      Value           =   -1  'True
      Width           =   1044
   End
   Begin VB.Frame Frame1 
      Caption         =   "³]©w"
      Height          =   645
      Left            =   240
      TabIndex        =   21
      Top             =   3450
      Width           =   3825
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   12
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "¦Lªí¾÷"
         Height          =   180
         Index           =   2
         Left            =   105
         TabIndex        =   22
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   7
      Top             =   2010
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   3120
      MaxLength       =   7
      TabIndex        =   8
      Top             =   2010
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "µ²§ô(&X)"
      Height          =   400
      Index           =   1
      Left            =   3330
      TabIndex        =   14
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2550
      TabIndex        =   13
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   11
      Top             =   3060
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   3120
      MaxLength       =   4
      TabIndex        =   2
      Top             =   930
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   1
      Top             =   930
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2340
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   3120
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1275
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1440
      MaxLength       =   9
      TabIndex        =   3
      Top             =   1275
      Width           =   1200
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2595
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "¬O§_§t®Ö»é¡G"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   25
      Top             =   2745
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "(N:¤£§t)"
      Height          =   180
      Index           =   0
      Left            =   1785
      TabIndex        =   24
      Top             =   2745
      Width           =   645
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2760
      X2              =   3000
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð°ê®a¡G"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   23
      Top             =   930
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   2760
      X2              =   3000
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "(1.¤¤¤å 2. ­^¤å 3.¤é¤å)"
      Height          =   180
      Index           =   1
      Left            =   1800
      TabIndex        =   20
      Top             =   3120
      Width           =   1740
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "³øªí»y¤å¡G"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   19
      Top             =   3120
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(1.¥Ø«eµ{§Ç 2.©Ò¦³µ{§Ç)"
      Height          =   180
      Left            =   1785
      TabIndex        =   18
      Top             =   2400
      Width           =   1875
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   2760
      X2              =   3000
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "¤º®e¡G"
      Height          =   180
      Left            =   360
      TabIndex        =   17
      Top             =   2370
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¥N²z¤H½s¸¹¡G"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   1275
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¨t²ÎÃþ§O¡G"
      Height          =   180
      Left            =   360
      TabIndex        =   15
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "frm050316"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/17 SQLDate¤wÀË¬d
'Memo by Morgan2010/12/28 ¥Ó½Ð®×¸¹Äæ¤w­×§ï
'Memo By Sindy 2010/11/26 ­û¤u½s¸¹Äæ¤w­×§ï
'Memo by Morgan2010/8/12 ¤é´ÁÄæ¤w­×§ï
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, i As Integer, j As Integer, k As Integer, s As Integer
Dim strTemp(0 To 20) As String, StrTemp5(0 To 14) As String, iPrint As Integer, Page As Integer, strSQL11 As String, strSQL22 As String, strSQL33 As String, strSQL44 As String, strSQL55 As String
Dim PLeft(0 To 17) As Integer, strTemp1 As Variant, strTemp2 As Variant, SeekPrint As Integer, SeekPrintL As Integer, SeekTempPrint As String, IntTot As Integer, StrSQL221 As String, StrSQL222 As String, StrSQL223 As String

'Add By Cheng 2002/02/27
Dim m_ColCustName As String '¥Ó½Ð¤H¦WºÙÄæ¦ì
Dim m_ColCustAdd As String '¥Ó½Ð¤H¦a§}Äæ¦ì
Dim m_ColAgName As String '¥N²z¤H¦WºÙÄæ¦ì
Dim m_ColAgAdd As String '¥N²z¤H¦a§}Äæ¦ì
'add by nickc 2007/05/25 §PÂ_¬O§_¦³¦L¹LP
Dim IsHavePrintP As Boolean

'*******************
'   ¥N²z¤H®×¥óÁ`Ã¯
'   ªô¤p©j»¡¤À¦¨±M§Q©M°Ó¼Ð¨âºØ®æ¦¡¦U¤À¤¤¡A­^¡A¤é
'   °£¤F°Ó¼Ð¥~¨ä¥L³£¬O±M§Q®æ¦¡
'   ±M§Q     lawcase , servicepractice  , patent
'   °Ó¼Ð     trademark
'   90/07/17   nick keyin »P­ì¥»±ø´Ú¥þ³¡³£¤£¤@¼Ë
'*******************
Private Sub cmdok_Click(Index As Integer)
'Add By Cheng 2002/11/15
On Error GoTo ErrorHandler

Select Case Index
Case 0 '½T©w
     PUB_RestorePrinter Combo1 'Modified by Morgan 2017/11/21 ³]©w¦Lªí¾÷§ï©I¥s¤½¥Î¨ç¼Æ,­ìµ{¦¡²¾°£
     'MODIFY BY SONIA 2016/1/4
     'Printer.PaperSize = 39
     Printer.PaperSize = PUB_GetPaperSize(15)
     'END 2016/1/4
     Printer.Orientation = 1
     DoEvents
     'Printer.Orientation = 2
     'DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("¨t²ÎÃþ§O¤£¥iªÅ¥Õ!!", , "USER ¿é¤J¿ù»~")
         txt1(0).SetFocus
         Exit Sub
     Else
         If Len(txt1(4)) = 0 Then
             s = MsgBox("¥N²z¤H½s¸¹°Ï¶¡¤£¥iªÅ¥Õ!!", , "USER ¿é¤J¿ù»~")
             txt1(3).SetFocus
             txt1_GotFocus (3)
             Exit Sub
         Else
             'Add By Cheng 2002/03/19
            If PUB_CheckKeyInDate(Me.txt1(8)) = -1 Then
               Me.txt1(8).SetFocus
               txt1_GotFocus 8
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(7)) = -1 Then
               Me.txt1(7).SetFocus
               txt1_GotFocus 7
               Exit Sub
            End If
             
             If Len(txt1(5)) = 0 Then
                 s = MsgBox("¤º®e¤£¥iªÅ¥Õ!!", , "USER ¿é¤J¿ù»~")
                 txt1(5).SetFocus
                 Exit Sub
             Else
                 If Len(txt1(6)) = 0 Then
                     s = MsgBox("³øªí»y¤å¤£¥iªÅ¥Õ!!", , "USER ¿é¤J¿ù»~")
                     txt1(6).SetFocus
                     Exit Sub
                 Else
                     If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Then
                         If Mid(Trim(txt1(3)), 1, 6) <> Mid(Trim(txt1(4)), 1, 6) Then
                             s = MsgBox("¥N²z¤H½s¸¹«e¤»½X¥²¶·¬Û¦P!!", , "USER ¿é¤J¿ù»~")
                             txt1(3).SetFocus
                             txt1_GotFocus (3)
                             Exit Sub
                         End If
                        Screen.MousePointer = vbHourglass
                        Me.Enabled = False
                        'Modify By Cheng 2003/11/04
                        '­Y¥Ñ¹q¸£¤¤¤ß¶i¤J
                        If UCase(App.EXEName) = "TECOMPUTER" Or UCase(App.EXEName) = "COMPUTER" Then
                            'add by nickc 2007/05/25
                            IsHavePrintP = False
                            '±M§Q
                            ProcessP
                            If IsHavePrintP = True Then
                                'MODIFY BY SONIA 2016/1/4
                                'Printer.PaperSize = 39
                                Printer.PaperSize = PUB_GetPaperSize(15)
                                'END 2016/1/4
                                Printer.Orientation = 1
                            End If
                            '°Ó¼Ð
                            ProcessT
                        '­Y¥Ñ¨ä¥L¨t²Î¶i¤J
                        Else
                            If intPCaseKind = 2 Then
                               '°Ó¼Ð
                               ProcessT
                            Else
                               '±M§Q
                               ProcessP
                            End If
                        End If
                        Me.Enabled = True
                        Screen.MousePointer = vbDefault
                    End If
                 End If
             End If
         End If
     End If
Case 1 'µ²§ô
    'Add By Cheng 2003/02/05
    '­Y¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    Unload Me
Case Else
End Select
'Add By Cheng 2002/11/15
Exit Sub
ErrorHandler:
    Select Case Err.Number
    Case 380
        MsgBox "¦Lªí¾÷¿ï¾Ü¿ù»~!!!"
    Case Else
        MsgBox "(" & Err.Number & ")" & Err.Description
    End Select
    Screen.MousePointer = vbDefault 'Added by Lydia 2018/03/16
End Sub

Private Sub ProcessP()
'Modify By Cheng 2002/02/27
'****************************************************
'¥N²z¤H¦WºÙ, ¥Ó½Ð¤H¦WºÙ, ¦a§}, ®×¥ó©Ê½è¦WºÙ, ®×¥ó¦WºÙ
'­Y¤¤¤å³øªí  ¤¤-->­^-->¤é
'­Y­^¤å³øªí  ­^-->¤é-->NULL
'­Y¤é¤å³øªí  ¤é-->­^-->NULL
'****************************************************
ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 ²M°£¬d¸ß¦Lªí°O¿ýÀÉÄæ¦ì
Screen.MousePointer = vbHourglass
cnnConnection.Execute "delete from R050316_C WHERE ID='" & strUserNum & "' "     '»y¤å¬°¤¤¤å      ¹ï¤º
cnnConnection.Execute "DELETE FROM R050316_E WHERE ID='" & strUserNum & "' "     '»y¤å¬°­^¤å    ¹ï¥~
cnnConnection.Execute "DELETE FROM R050316_J WHERE ID='" & strUserNum & "' "     '»y¤å¬°¤é¤å   ¹ï¥~
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
strSQL11 = " AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+)  "
strSQL22 = ""
strSQL33 = " AND SP01=CP01(+) AND SP02=CP02(+) AND SP03=CP03(+) AND SP04=CP04(+) "
strSQL44 = " AND LC01=CP01(+) AND LC02=CP02(+) AND LC03=CP03(+) AND LC04=CP04(+) "
strSQL55 = ""
StrSQL221 = " AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) "
StrSQL222 = " AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) "
StrSQL223 = " AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) "
If Len(txt1(0)) <> 0 Then
    strSQL11 = strSQL11 & " and pa01 in (" & SQLGrpStr(txt1(0), 1) & ") "
    StrSQL221 = StrSQL221 & " and cp01 in (" & SQLGrpStr(txt1(0), 1) & ") "
    strSQL33 = strSQL33 & " and sp01 in (" & SQLGrpStr(txt1(0), 5) & ") "
    StrSQL222 = StrSQL222 & " and cp01 in (" & SQLGrpStr(txt1(0), 5) & ") "
    strSQL44 = strSQL44 & " and lc01 in (" & SQLGrpStr(txt1(0), 3) & ") "
    StrSQL223 = StrSQL223 & " and cp01 in (" & SQLGrpStr(txt1(0), 3) & ") "
    pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/10/4
End If
If Len(txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & txt1(1) & "' "
    StrSQL3 = StrSQL3 + " AND SP09>='" & txt1(1) & "' "
    StrSQL4 = StrSQL4 + " AND LC15>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & txt1(2) & "' "
    StrSQL3 = StrSQL3 + " AND SP09<='" & txt1(2) & "' "
    StrSQL4 = StrSQL4 + " AND LC15<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2(3) & txt1(1) & "-" & txt1(2)  'Add By Sindy 2010/10/4
End If
If Len(txt1(4)) <> 0 Then
    strSQL11 = strSQL11 + " AND (PA75>='" & GetNewFagent(txt1(3)) & "' AND PA75<='" & GetNewFagent(txt1(4)) & "') "
    strSQL22 = strSQL22 + " AND (CP44>='" & GetNewFagent(txt1(3)) & "' AND CP44<='" & GetNewFagent(txt1(4)) & "') "
    strSQL33 = strSQL33 + " AND (SP26>='" & GetNewFagent(txt1(3)) & "' AND SP26<='" & GetNewFagent(txt1(4)) & "') "
    strSQL44 = strSQL44 + " AND (LC22>='" & GetNewFagent(txt1(3)) & "' AND LC22<='" & GetNewFagent(txt1(4)) & "') "
    strSQL55 = strSQL55 + " AND (CP44>='" & GetNewFagent(txt1(3)) & "' AND CP44<='" & GetNewFagent(txt1(4)) & "') "
    pub_QL05 = pub_QL05 & ";" & Label2(0) & txt1(3) & "-" & txt1(4)  'Add By Sindy 2010/10/4
End If
'¦¬µo¤å°_¤é
If Len(txt1(8)) <> 0 Then
    'Modify By Cheng 2002/11/15
    If Me.opt(0).Value Then
        strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(8)))
        StrSQL3 = StrSQL3 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(8)))
        StrSQL4 = StrSQL4 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(8)))
        strSQL5 = strSQL5 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(8)))
    Else
        strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(8)))
        StrSQL3 = StrSQL3 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(8)))
        StrSQL4 = StrSQL4 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(8)))
        strSQL5 = strSQL5 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(8)))
    End If
End If
'¦¬µo¤å¨´¤é
If Len(txt1(7)) <> 0 Then
    'Modify By Cheng 2002/11/15
    If Me.opt(0).Value Then
        strSQL1 = strSQL1 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(7)))
        StrSQL3 = StrSQL3 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(7)))
        StrSQL4 = StrSQL4 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(7)))
        strSQL5 = strSQL5 + " AND CP05<=" & Val(ChangeTStringToWString(txt1(7)))
    Else
        strSQL1 = strSQL1 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(7)))
        StrSQL3 = StrSQL3 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(7)))
        StrSQL4 = StrSQL4 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(7)))
        strSQL5 = strSQL5 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(7)))
    End If
'Add By Cheng 2002/03/19
Else
    'Modify By Cheng 2002/11/15
    '­Y¦³¿é¤J°_¤é¥¼¿é¤J¨´¤é§Y¥H¨t²Î¤é¨ú¥N¤§
    If Len(txt1(8)) <> 0 Then
        If Me.opt(0).Value Then
            strSQL1 = strSQL1 + " AND CP05<=" & Val(ServerDate)
            StrSQL3 = StrSQL3 + " AND CP05<=" & Val(ServerDate)
            StrSQL4 = StrSQL4 + " AND CP05<=" & Val(ServerDate)
            strSQL5 = strSQL5 + " AND CP05<=" & Val(ServerDate)
        Else
            strSQL1 = strSQL1 + " AND CP27<=" & Val(ServerDate)
            StrSQL3 = StrSQL3 + " AND CP27<=" & Val(ServerDate)
            StrSQL4 = StrSQL4 + " AND CP27<=" & Val(ServerDate)
            strSQL5 = strSQL5 + " AND CP27<=" & Val(ServerDate)
        End If
    End If
End If
If Len(txt1(8)) <> 0 Or Len(txt1(7)) <> 0 Then
   If Me.opt(0).Value Then
      pub_QL05 = pub_QL05 & ";¦¬¤å¤é¡G" & txt1(8) & "-" & txt1(7)   'Add By Sindy 2010/10/4
   Else
      pub_QL05 = pub_QL05 & ";µo¤å¤é¡G" & txt1(8) & "-" & txt1(7)   'Add By Sindy 2010/10/4
   End If
End If

'Add By Cheng 2003/07/17
'¬O§_§t®Ö»é
If Me.txt1(9).Text <> "" Then
    strSQL1 = strSQL1 + " AND ( PA16<>'2' OR PA16 IS NULL ) "
    pub_QL05 = pub_QL05 & ";" & Label6(0) & "¤£§t"  'Add By Sindy 2010/10/4
End If

CheckOC
Select Case txt1(6)
Case 1 '¤¤¤å³øªí
      pub_QL05 = pub_QL05 & ";" & Label6(1) & "¤¤¤å"   'Add By Sindy 2010/10/4
         '                   StrSQL = " SELECT PA75,pa77,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,PA26),pa11,NVL(na04,PA09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,PA06,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL11
         '                                                                                                                                                                                                                                                             'FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
         'StrSQL = StrSQL + " union all select sp26,sp27,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,SP08),sp11,NVL(na04,SP09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp06,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
         'StrSQL = StrSQL + " union all select lc22,lc23,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,LC11),'',NVL(na04,LC15),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc06,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL44
         'StrSQL = StrSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,PA26),pa11,NVL(na04,PA09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,PA06,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
         'StrSQL = StrSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,SP08),sp11,NVL(na04,SP09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp06,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
         'StrSQL = StrSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,LC11),'',NVL(na04,LC15),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc06,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223

         '               StrSQL = " SELECT PA75,pa77,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,PA26),pa11,NVL(na04,PA09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,PA06,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL11
         'StrSQL = StrSQL + " union all select sp26,sp27,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,SP08),sp11,NVL(na04,SP09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp06,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
         'StrSQL = StrSQL + " union all select lc22,lc23,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,LC11),'',NVL(na04,LC15),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc06,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL44
         'StrSQL = StrSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,PA26),pa11,NVL(na04,PA09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,PA06,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
         'StrSQL = StrSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,SP08),sp11,NVL(na04,SP09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp06,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
         'StrSQL = StrSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,LC11),'',NVL(na04,LC15),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc06,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
    '¥Ø«eµ{§Ç
    If txt1(5) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label3 & "¥Ø«eµ{§Ç"   'Add By Sindy 2010/10/4
'                        strSQL = " SELECT PA75,pa77,NVL(cu04,PA26),pa11,NVL(na03,PA09),NVL(decode(pa09,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),PA05,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(cu04,SP08),sp11,NVL(na03,SP09),NVL(decode(sp09,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp05,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(cu04,LC11),'',NVL(na03,LC15),NVL(decode(lc15,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc05,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu04,PA26),pa11,NVL(na03,PA09),NVL(decode(pa09,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),PA05,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu04,SP08),sp11,NVL(na03,SP09),NVL(decode(sp09,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp05,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu04,LC11),'',NVL(na03,LC15),NVL(decode(lc15,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc05,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE  cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
      
      '910819 Sieg 307
      If intPCaseKind = ±M§Q And intPWhere = °ê¥~_CF Then
        'Modify By Cheng 2002/11/15
        '­×§ï¥Ó½Ð¤H¦WºÙ
'         strSQL = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE  cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
'92.03.27 nick §ï¥»©Ò®×¸¹
'         strSQL = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE  cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
         'Modified by Lydia 2018/06/05 ­×§ïÅã¥Ü®×¥ó©Ê½è '020',CPM04,CPM03 => '000',CPM03,CPM04
         strSql = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & strSQL11
         strSql = strSql + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL33
         strSql = strSql + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL44
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & strSQL22 & StrSQL221
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL22 & StrSQL222
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE  cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL22 & StrSQL223
         'end 2018/06/05
      Else
        'Modify By Cheng 2002/11/15
        '­×§ï¥Ó½Ð¤H¦WºÙ
'         strSQL = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),Nvl(PA05,nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE  cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
'92.03.27 nick §ï¥»©Ò®×¸¹
'         strSQL = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),Nvl(PA05,nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE  cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
         'Modified by Lydia 2018/06/05 ­×§ïÅã¥Ü®×¥ó©Ê½è '020',CPM04,CPM03 => '000',CPM03,CPM04
         strSql = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'020',ptm04,ptm03),Nvl(PA05,nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & strSQL11
         strSql = strSql + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL33
         strSql = strSql + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL44
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'020',ptm04,ptm03),Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & strSQL1 & strSQL22 & StrSQL221
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL22 & StrSQL222
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE  cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL22 & StrSQL223
         'end 2018/06/05
      End If
      
    '©Ò¦³µ{§Ç
    Else
      pub_QL05 = pub_QL05 & ";" & Label3 & "©Ò¦³µ{§Ç"   'Add By Sindy 2010/10/4
'                        strSQL = " SELECT PA75,pa77,NVL(cu04,PA26),pa11,NVL(na03,PA09),NVL(decode(pa09,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),PA05,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=pTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(cu04,SP08),sp11,NVL(na03,SP09),NVL(decode(sp09,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp05,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(cu04,LC11),'',NVL(na03,LC15),NVL(decode(lc15,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc05,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu04,PA26),pa11,NVL(na03,PA09),NVL(decode(pa09,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),PA05,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu04,SP08),sp11,NVL(na03,SP09),NVL(decode(sp09,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp05,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu04,LC11),'',NVL(na03,LC15),NVL(decode(lc15,'000',cpm03,cpm04),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc05,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
      '910819 Sieg 307
      If intPCaseKind = ±M§Q And intPWhere = °ê¥~_CF Then
        'Modify By Cheng 2002/11/14
        '­×§ï¥Ó½Ð¤H¦WºÙ
'         strSQL = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=pTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
'92.03.27 nick §ï¥»©Ò®×¸¹
'         strSQL = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=pTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
         'Modified by Lydia 2018/06/05 ­×§ïÅã¥Ü®×¥ó©Ê½è '020',CPM04,CPM03 => '000',CPM03,CPM04
         strSql = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=pTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL33
         strSql = strSql + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL44
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),ptm03,Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL22 & StrSQL222
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL22 & StrSQL223
         'end 2018/06/05
      Else
        'Modify By Cheng 2002/11/15
        '­×§ï¥Ó½Ð¤H¦WºÙ
'         strSQL = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=pTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),Null),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
'92.03.27 nick §ï¥»©Ò®×¸¹
'         strSQL = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=pTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'000',ptm03,ptm04),Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C')  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
         'Modified by Lydia 2018/06/05 ­×§ïÅã¥Ü®×¥ó©Ê½è '020',CPM04,CPM03 => '000',CPM03,CPM04
         strSql = " SELECT PA75,pa77,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'020',ptm04,ptm03),Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=pTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select sp26,sp27,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL33
         strSql = strSql + " union all select lc22,lc23,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL44
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),PA26),pa11,NVL(na03,PA09),NVL(Nvl(decode(pa09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s¦b','N','®ø·À',''),decode(pa09,'020',ptm04,ptm03),Nvl(PA05,Nvl(PA06,PA07)),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),SP08),sp11,NVL(na03,SP09),NVL(Nvl(decode(sp09,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp05,Nvl(SP06,SP07)),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP08,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL22 & StrSQL222
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),LC11),'',NVL(na03,LC15),NVL(Nvl(decode(lc15,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc05,Nvl(LC06,LC07)),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D'))  AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL22 & StrSQL223
         'end 2018/06/05
      End If
    End If
    adoRecordset.CursorLocation = adUseClient
    k = 0
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        With adoRecordset
            .MoveFirst
            DoEvents
            Do While .EOF = False
                For i = 0 To 17
                    strTemp(i) = CheckStr(.Fields(i))
                Next i
                If strTemp(16) = "Y" Then
                    strTemp(9) = "*" + strTemp(9)
                    strTemp(16) = "³¬¨÷"
                Else
                    strTemp(16) = ""
                End If
                Select Case CheckSys(CheckStr(.Fields(18)))
                Case "1", "5"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>411 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (411,1204,997,998,995,996,999) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "3", "4", "7", "8"
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & " and np07<>6001 AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & strNpSqlOfNoSalesDuty & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "2", "6"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>305 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (305,997,998,995,996) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2

                End Select
                strSql = "INSERT INTO R050316_C values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
                k = k + 1
                DoEvents
                .MoveNext
            Loop
        End With
        '¥Ø«eµ{§Ç
        If Val(txt1(5)) = 1 Then
            'Add By Cheng 2002/03/28
            '¥ý±NµLµo¤å¤éªº¸ê®Æ¶ñ¤J¤@­È(MaxDate), ¥H¨Ï¦¹Äæ¦ì­È³Ì¤j
            cnnConnection.Execute "Update R050316_C Set R013007='MaxDate' Where R013007 is null Or length(R013007)=0"
            '¥ý§ìµLµo¤å¤é¥B¦¬¤å¤é, ¦¬¤å¸¹³Ì¤jªÌ, ­Y³£¦³µo¤å¤é«h§ì³Ì¤jªÌ
            strSql = "SELECT R013010,MAX(R013007) FROM R050316_C WHERE ((R013018<'B' or R013018>'C')) AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_C WHERE ((R013018<'B' or R013018>'C')) AND (R013010 NOT IN " & strSQL1 & " OR R013007 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            '¦A±Nµo¤å¤éªº¸ê®Æ¬°MaxDateªÌ, §ó·s¬° NULL
            cnnConnection.Execute "Update R050316_C Set R013007=NULL Where R013007='MaxDate'"
            
            strSql = "SELECT R013010,MAX(R013016) FROM R050316_C WHERE ((R013018<'B' or R013018>'C')) AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_C WHERE ((R013018<'B' or R013018>'C')) AND (R013010 NOT IN " & strSQL1 & " OR R013016 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MIN(R013008) FROM R050316_C WHERE ((R013018<'B' or R013018>'C')) AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_C WHERE ((R013018<'B' or R013018>'C')) AND (R013010 NOT IN " & strSQL1 & " OR R013008 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_C WHERE ((R013018<'B' or R013018>'C')) AND ID='" & strUserNum & "'  GROUP BY R013010 "
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_c WHERE "
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSQL1 = " ("
                strSQL2 = " ("
                Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
                Loop
                strSQL1 = strSQL1 + ") "
                strSQL2 = strSQL2 + ") "
                cnnConnection.Execute "DELETE FROM R050316_C WHERE ((R013018<'B' or R013018>'C')) AND (R013010 NOT IN " & strSQL1 & " OR R013018 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
        End If
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/4
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    CheckOC
Case 2 '­^¤å³øªí
    pub_QL05 = pub_QL05 & ";" & Label6(1) & "­^¤å"   'Add By Sindy 2010/10/4
    'strSQL = " SELECT PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22," & SQLDate("PA10") & ",PA11,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",PA57,PA77,PA06,CP09,PA75 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL1
    'strSQL = strSQL + " union all select TM01||'-'||TM02||'-'||TM03||'-'||TM04,TM15," & SQLDate("TM11") & ",TM12,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",TM29,TM45,TM06,CP09,TM44 FROM TRADEMARK,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cp01=CPM01(+) AND CP10=CPM02(+) AND '2'=ptm02(+) AND TM08=PTM02(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL2
    'strSQL = strSQL + " union all select SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14," & SQLDate("SP10") & ",SP11,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",SP15,SP27,SP06,CP09,SP26 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cP01=CPM01(+) AND CP10=CPM02(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL3
    'strSQL = strSQL + " union all select LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'',CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",LC08,LC23,LC06,CP09,LC22 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cp01=CPM01(+) AND CP10=CPM02(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL4
    'strSQL = strSQL + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,PA22," & SQLDate("PA10") & ",PA11,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",PA57,CP45,PA06,CP09,CP44 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL5
    'strSQL = strSQL + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,TM15," & SQLDate("TM11") & ",TM12,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",TM29,CP45,TM06,CP09,CP44 FROM TRADEMARK,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cp01=CPM01(+) AND CP10=CPM02(+) AND '2'=ptm02(+) AND TM08=PTM02(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL5
    'strSQL = strSQL + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,SP14," & SQLDate("SP10") & ",SP11,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",SP15,CP45,SP06,CP09,CP44 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cP01=CPM01(+) AND CP10=CPM02(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL5
    'strSQL = strSQL + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,'',0,'',CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",LC08,CP45,LC06,CP09,CP44 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND Cp01=CPM01(+) AND CP10=CPM02(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL5
    '89/11/23   §ï nick
    '                    strSQL = " SELECT PA75,pa77,NVL(cu06,PA26),pa11,NVL(na04,PA09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,PA07,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL11
    ''     strSQL = strSQL + " union all select sp26,sp27,NVL(cu06,SP08),sp11,NVL(na04,SP09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp07,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
    '     strSQL = strSQL + " union all select lc22,lc23,NVL(cu06,LC11),'',NVL(na04,LC15),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc07,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL44
    '     strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,PA26),pa11,NVL(na04,PA09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,PA07,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
    '     strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,SP08),sp11,NVL(na04,SP09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp07,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
    '     strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,LC11),'',NVL(na04,LC15),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc07,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
    'Else
    '                    strSQL = " SELECT PA75,pa77,NVL(cu06,PA26),pa11,NVL(na04,PA09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,PA07,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL11
    '     strSQL = strSQL + " union all select sp26,sp27,NVL(cu06,SP08),sp11,NVL(na04,SP09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp07,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
    '     strSQL = strSQL + " union all select lc22,lc23,NVL(cu06,LC11),'',NVL(na04,LC15),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc07,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL44
    '     strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,PA26),pa11,NVL(na04,PA09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,PA07,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
    '     strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,SP08),sp11,NVL(na04,SP09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp07,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
    '     strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,LC11),'',NVL(na04,LC15),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc07,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
    
    If txt1(5) = "1" Then
         pub_QL05 = pub_QL05 & ";" & Label3 & "¥Ø«eµ{§Ç"   'Add By Sindy 2010/10/4
    'FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL11
'                        strSQL = " SELECT PA75,pa77,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,PA26),pa11,NVL(na04,PA09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,PA06,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND CP09<'B' AND CP09>='C' AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'                                                                                                                                                                                                                                                                      'FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select sp26,sp27,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,SP08),sp11,NVL(na04,SP09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp06,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND CP09<'B' AND CP09>='C' AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,LC11),'',NVL(na04,LC15),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc06,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND CP09<'B' AND CP09>='C' AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,PA26),pa11,NVL(na04,PA09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,PA06,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND CP09<'B' AND CP09>='C' AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,SP08),sp11,NVL(na04,SP09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp06,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND CP09<'B' AND CP09>='C' AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,LC11),'',NVL(na04,LC15),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc06,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND CP09<'B' AND CP09>='C' AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
'92.03.27 nick §ï¥»©Ò®×¸¹
'                        strSQL = " SELECT PA75,pa77,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),PA26),pa11,NVL(na04,PA09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,Nvl(PA06,PA07),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or CP09>='C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),SP08),sp11,NVL(na04,SP09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp06,SP07),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or CP09>='C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),LC11),'',NVL(na04,LC15),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc06,LC07),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or CP09>='C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),PA26),pa11,NVL(na04,PA09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,Nvl(PA06,PA07),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or CP09>='C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),SP08),sp11,NVL(na04,SP09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp06,SP07),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or CP09>='C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),LC11),'',NVL(na04,LC15),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc06,LC07),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or CP09>='C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
                        strSql = " SELECT PA75,pa77,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),PA26),pa11,NVL(na04,PA09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,Nvl(PA06,PA07),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=PTM01(+) AND PA08=PTM02(+)  AND (CP09<'B' or CP09>='C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select sp26,sp27,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),SP08),sp11,NVL(na04,SP09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp06,SP07),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or CP09>='C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL33
         strSql = strSql + " union all select lc22,lc23,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),LC11),'',NVL(na04,LC15),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc06,LC07),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or CP09>='C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL44
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),PA26),pa11,NVL(na04,PA09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,Nvl(PA06,PA07),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or CP09>='C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),SP08),sp11,NVL(na04,SP09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp06,SP07),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or CP09>='C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL22 & StrSQL222
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),LC11),'',NVL(na04,LC15),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc06,LC07),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or CP09>='C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL22 & StrSQL223
    Else
         pub_QL05 = pub_QL05 & ";" & Label3 & "©Ò¦³µ{§Ç"   'Add By Sindy 2010/10/4
'                        strSQL = " SELECT PA75,pa77,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,PA26),pa11,NVL(na04,PA09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,PA06,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND CP09<'B' AND CP09>='C' AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,SP08),sp11,NVL(na04,SP09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp06,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'B' AND CP09>='C' AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,LC11),'',NVL(na04,LC15),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc06,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND CP09<'B' AND CP09>='C' AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,PA26),pa11,NVL(na04,PA09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,PA06,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND CP09<'B' AND CP09>='C' AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,SP08),sp11,NVL(na04,SP09),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp06,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'B' AND CP09>='C' AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,LC11),'',NVL(na04,LC15),NVL(cpm10,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc06,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND CP09<'B' AND CP09>='C' AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
'92.03.27 nick §ï¥»©Ò®×¸¹
'                        strSQL = " SELECT PA75,pa77,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),PA26),pa11,NVL(na04,PA09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,Nvl(PA06,PA07),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or CP09>='C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),SP08),sp11,NVL(na04,SP09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp06,SP07),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or CP09>='C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),LC11),'',NVL(na04,LC15),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc06,LC07),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or CP09>='C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),PA26),pa11,NVL(na04,PA09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,Nvl(PA06,PA07),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' Or CP09>='C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),SP08),sp11,NVL(na04,SP09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp06,SP07),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' Or CP09>='C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),LC11),'',NVL(na04,LC15),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc06,LC07),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' Or CP09>='C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
                        strSql = " SELECT PA75,pa77,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),PA26),pa11,NVL(na04,PA09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,Nvl(PA06,PA07),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or CP09>='C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select sp26,sp27,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),SP08),sp11,NVL(na04,SP09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp06,SP07),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or CP09>='C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL33
         strSql = strSql + " union all select lc22,lc23,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),LC11),'',NVL(na04,LC15),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc06,LC07),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or CP09>='C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL44
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),PA26),pa11,NVL(na04,PA09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,Nvl(PA06,PA07),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' Or CP09>='C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),SP08),sp11,NVL(na04,SP09),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp06,SP07),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' Or CP09>='C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL22 & StrSQL222
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),LC11),'',NVL(na04,LC15),NVL(Nvl(cpm10,CPM13),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc06,LC07),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' Or CP09>='C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL22 & StrSQL223
    End If
    adoRecordset.CursorLocation = adUseClient
    k = 0
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        With adoRecordset
            .MoveFirst
            DoEvents
            Do While .EOF = False
                For i = 0 To 17
                    strTemp(i) = CheckStr(.Fields(i))
                Next i
                If strTemp(16) = "Y" Then
                    strTemp(9) = "*" + strTemp(9)
                    strTemp(16) = "Closed"
                Else
                    strTemp(16) = ""
                End If
                Select Case CheckSys(CheckStr(.Fields(18)))
                Case "1", "5"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>411 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (411,997,998,995,996,999,1204) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "3", "4", "7", "8"
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & " and np07<>6001 AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & strNpSqlOfNoSalesDuty & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "2", "6"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>305 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (305,997,998,995,996) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                
                
                End Select
                strSql = "INSERT INTO R050316_E values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & strUserNum & "') "
                'strSQL = "INSERT INTO R050316_C values ('" & chgsql(strTemp(0)) & "','" & chgsql(strTemp(1)) & "','" & chgsql(strTemp(2)) & "','" & chgsql(strTemp(3)) & "','" & chgsql(strTemp(4)) & "','" & chgsql(strTemp(5)) & "','" & chgsql(strTemp(6)) & "','" & chgsql(strTemp(7)) & "','" & chgsql(strTemp(8)) & "','" & chgsql(strTemp(9)) & "','" & chgsql(strTemp(10)) & "','" & chgsql(strTemp(11)) & "','" & chgsql(strTemp(12)) & "','" & chgsql(strTemp(13)) & "','" & chgsql(strTemp(14)) & "','" & chgsql(strTemp(15)) & "','" & chgsql(strTemp(16)) & "','" & chgsql(strTemp(17)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
                k = k + 1
                DoEvents
                .MoveNext
            Loop
        End With
        '¥Ø«eµ{§Ç
        If Val(txt1(5)) = 1 Then
            'Add By Cheng 2002/03/28
            '¥ý±NµLµo¤å¤éªº¸ê®Æ¶ñ¤J¤@­È(MaxDate), ¥H¨Ï¦¹Äæ¦ì­È³Ì¤j
            cnnConnection.Execute "Update R050316_E Set R013007='MaxDate' Where R013007 is null Or length(R013007)=0"
            '¥ý§ìµLµo¤å¤é¥B¦¬¤å¤é, ¦¬¤å¸¹³Ì¤jªÌ, ­Y³£¦³µo¤å¤é«h§ì³Ì¤jªÌ
            strSql = "SELECT R013010,MAX(R013007) FROM R050316_E WHERE ((R013018<'B' or R013018>'C')) AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_E WHERE ((R013018<'B' or R013018>'C')) AND (R013010 NOT IN " & strSQL1 & " OR R013007 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            '¦A±Nµo¤å¤éªº¸ê®Æ¬°MaxDateªÌ, §ó·s¬° NULL
            cnnConnection.Execute "Update R050316_E Set R013007=NULL Where R013007='MaxDate'"
            
            strSql = "SELECT R013010,MAX(R013016) FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            '
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013016 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MIN(R013008) FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013008 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "'  GROUP BY R013010 "
            'CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_E WHERE "
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSQL1 = " ("
                strSQL2 = " ("
                Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
                Loop
                strSQL1 = strSQL1 + ") "
                strSQL2 = strSQL2 + ") "
                cnnConnection.Execute "DELETE FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013018 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
                'cnnConnection.Execute "DELETE FROM R050316_E WHERE SUBSTR(R013012,1,1)='C' AND R013012<>'" & chgsql(CheckStr(adoRecordset1.Fields(0))) & "' AND ID='" & strUserNum & "' "
            End If
            CheckOC2
        End If
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/4
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    CheckOC
Case 3 '¤é¤å³øªí
    pub_QL05 = pub_QL05 & ";" & Label6(1) & "¤é¤å"   'Add By Sindy 2010/10/4
    If txt1(5) = "1" Then
         pub_QL05 = pub_QL05 & ";" & Label3 & "¥Ø«eµ{§Ç"   'Add By Sindy 2010/10/4
'                        strSQL = " SELECT PA75,pa77,NVL(cu06,PA26),pa11,NVL(na04,PA09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,PA07,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(cu06,SP08),sp11,NVL(na04,SP09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp07,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(cu06,LC11),'',NVL(na04,LC15),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc07,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,PA26),pa11,NVL(na04,PA09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,PA07,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,SP08),sp11,NVL(na04,SP09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp07,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,LC11),'',NVL(na04,LC15),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc07,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
'92.03.27 nick §ï¥»©Ò®×¸¹
'                        strSQL = " SELECT PA75,pa77,NVL(cu06,Decode(CU05,Null,PA26,CU05||' '||CU88||' '||CU89||' '||CU90)),pa11,NVL(na04,PA09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,Nvl(PA07,PA06),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(cu06,Decode(CU05,Null,SP08,CU05||' '||CU88||' '||CU89||' '||CU90)),sp11,NVL(na04,SP09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp07,SP06),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(cu06,Decode(CU05,Null,LC11,CU05||' '||CU88||' '||CU89||' '||CU90)),'',NVL(na04,LC15),NVL(nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc07,LC06),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,PA26,CU05||' '||CU88||' '||CU89||' '||CU90)),pa11,NVL(na04,PA09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,Nvl(PA07,PA06),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,SP08,CU05||' '||CU88||' '||CU89||' '||CU90)),sp11,NVL(na04,SP09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp07,SP06),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,LC11,CU05||' '||CU88||' '||CU89||' '||CU90)),'',NVL(na04,LC15),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc07,LC06),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
          'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
                        strSql = " SELECT PA75,pa77,NVL(cu06,Decode(CU05,Null,PA26,CU05||' '||CU88||' '||CU89||' '||CU90)),pa11,NVL(na04,PA09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,Nvl(PA07,PA06),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select sp26,sp27,NVL(cu06,Decode(CU05,Null,SP08,CU05||' '||CU88||' '||CU89||' '||CU90)),sp11,NVL(na04,SP09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp07,SP06),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL33
         strSql = strSql + " union all select lc22,lc23,NVL(cu06,Decode(CU05,Null,LC11,CU05||' '||CU88||' '||CU89||' '||CU90)),'',NVL(na04,LC15),NVL(nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc07,LC06),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL44
         strSql = strSql + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,PA26,CU05||' '||CU88||' '||CU89||' '||CU90)),pa11,NVL(na04,PA09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,Nvl(PA07,PA06),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
         strSql = strSql + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,SP08,CU05||' '||CU88||' '||CU89||' '||CU90)),sp11,NVL(na04,SP09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp07,SP06),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL22 & StrSQL222
         strSql = strSql + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,LC11,CU05||' '||CU88||' '||CU89||' '||CU90)),'',NVL(na04,LC15),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc07,LC06),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL22 & StrSQL223
    Else
         pub_QL05 = pub_QL05 & ";" & Label3 & "©Ò¦³µ{§Ç"   'Add By Sindy 2010/10/4
'                        strSQL = " SELECT PA75,pa77,NVL(cu06,PA26),pa11,NVL(na04,PA09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,PA07,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(cu06,SP08),sp11,NVL(na04,SP09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp07,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(cu06,LC11),'',NVL(na04,LC15),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc07,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,PA26),pa11,NVL(na04,PA09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,PA07,CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,SP08),sp11,NVL(na04,SP09),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',sp07,CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,LC11),'',NVL(na04,LC15),NVL(cpm13,CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',lc07,CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
'92.03.27 nick §ï¥»©Ò®×¸¹
'                        strSQL = " SELECT PA75,pa77,NVL(cu06,Decode(CU05,Null,PA26,CU05||' '||CU88||' '||CU89||' '||CU90)),pa11,NVL(na04,PA09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,Nvl(PA07,PA06),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select sp26,sp27,NVL(cu06,Decode(CU05,Null,SP08,CU05||' '||CU88||' '||CU89||' '||CU90)),sp11,NVL(na04,SP09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp07,SP06),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL33
'         strSQL = strSQL + " union all select lc22,lc23,NVL(cu06,Decode(CU05,Null,LC11,CU05||' '||CU88||' '||CU89||' '||CU90)),'',NVL(na04,LC15),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc07,LC06),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strsql44
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,PA26,CU05||' '||CU88||' '||CU89||' '||CU90)),pa11,NVL(na04,PA09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,Nvl(PA07,PA06),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,SP08,CU05||' '||CU88||' '||CU89||' '||CU90)),sp11,NVL(na04,SP09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14,sp21,'','',Nvl(sp07,SP06),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  CP09>'C') AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & StrSQL22 & StrSQL222
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,LC11,CU05||' '||CU88||' '||CU89||' '||CU90)),'',NVL(na04,LC15),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'','',Nvl(lc07,LC06),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & StrSQL22 & StrSQL223
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
                        strSql = " SELECT PA75,pa77,NVL(cu06,Decode(CU05,Null,PA26,CU05||' '||CU88||' '||CU89||' '||CU90)),pa11,NVL(na04,PA09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,Nvl(PA07,PA06),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select sp26,sp27,NVL(cu06,Decode(CU05,Null,SP08,CU05||' '||CU88||' '||CU89||' '||CU90)),sp11,NVL(na04,SP09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp07,SP06),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL33
         strSql = strSql + " union all select lc22,lc23,NVL(cu06,Decode(CU05,Null,LC11,CU05||' '||CU88||' '||CU89||' '||CU90)),'',NVL(na04,LC15),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc07,LC06),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL44
         strSql = strSql + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,PA26,CU05||' '||CU88||' '||CU89||' '||CU90)),pa11,NVL(na04,PA09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',pa48,PA01||'-'||PA02||decode(pa03||pa04,'000','','-'||PA03||'-'||PA04),PA22,pa25,decode(pa17,'Y','¦s‘Ò ','N','®ø·À',''),ptm06,Nvl(PA07,PA06),CP05,pA57,cp09,cp01,cp02,cp03,cp04 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
         strSql = strSql + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,SP08,CU05||' '||CU88||' '||CU89||' '||CU90)),sp11,NVL(na04,SP09),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',sp29,SP01||'-'||SP02||decode(sp03||sp04,'000','','-'||SP03||'-'||SP04),SP14,sp21,'','',Nvl(sp07,SP06),CP05,sp15,cp09,cp01,cp02,cp03,cp04 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+)  AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL3 & strSQL22 & StrSQL222
         strSql = strSql + " union all select CP44,CP45,NVL(cu06,Decode(CU05,Null,LC11,CU05||' '||CU88||' '||CU89||' '||CU90)),'',NVL(na04,LC15),NVL(Nvl(cpm13,CPM10),CP10)," & SQLDate("cp27") & ",' ',lc17,LC01||'-'||LC02||decode(lc03||lc04,'000','','-'||LC03||'-'||LC04),'',0,'','',Nvl(lc07,LC06),CP05,lc08,cp09,cp01,cp02,cp03,cp04 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cp01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL4 & strSQL22 & StrSQL223
    End If
         'strSQL = strSQL + " union all select PA75,pa77,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,PA26),pa11,NVL(na04,PA09),NVL(cpm10,CP10)," & SQLDate("cp27") & "," & SQLDate("np08") & ",pa48,CP01||'-'||CP02||'-'||CP03||'-'||CP04,PA22,pa25,decode(pa17,'Y','Yes','N','Cancelled',''),ptm05,PA06,' ',pA57,cp09 FROM PATENT,CASEPROGRESS,PATENTTRADEMARKMAP,CASEPROPERTYMAP,NEXTPROGRESS,CUSTOMER,NATION WHERE NP01=CP09 AND cP01=CPM01(+) AND CP10=CPM02(+) AND '1'=ptm01(+) AND PA08=PTM02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),NULL,'0',SUBSTR(PA26,9,1))=CU02(+) AND PA09=NA01(+)  " & StrSQL5
         'strSQL = strSQL + " union all select sp26,sp27,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,SP08),sp11,NVL(na04,SP09),NVL(cpm10,CP10)," & SQLDate("cp27") & "," & SQLDate("np08") & ",sp29,CP01||'-'||CP02||'-'||CP03||'-'||CP04,SP14,sp21,'','',sp06,'',sp15,cp09 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,NEXTPROGRESS,CUSTOMER,NATION WHERE NP01=CP09 AND cP01=CPM01(+) AND CP10=CPM02(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR9SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),NULL,'0',SUBSTR(SP09,9,1)) = CU02(+) AND SP09=NA01(+) " & StrSQL5
         'strSQL = strSQL + " union all select lc22,lc23,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,LC11),'',NVL(na04,TM10),NVL(cpm10,CP10)," & SQLDate("cp27") & "," & SQLDate("np08") & ",lc17,CP01||'-'||CP02||'-'||CP03||'-'||CP04,'','','','',lc06,'',lc08,cp09 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,NEXTPROGRESS,CUSTOMER,NATION WHERE NP01=CP09 AND Cp01=CPM01(+) AND CP10=CPM02(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),NULL,'0',SUBSTR(LC11,9,1)) = CU02(+) AND LC15=NA01(+) " & StrSQL5
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    k = 0
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        With adoRecordset
            .MoveFirst
            DoEvents
            Do While .EOF = False
                For i = 0 To 17
                    strTemp(i) = CheckStr(.Fields(i))
                Next i
                If strTemp(16) = "Y" Then
                    strTemp(9) = "*" + strTemp(9)
                    strTemp(16) = "ÇVÆê"
                Else
                    strTemp(16) = ""
                End If
                Select Case CheckSys(CheckStr(.Fields(18)))
                Case "1", "5"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>411 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (411,997,998,995,996,999,1204) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "3", "4", "7", "8"
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & " and np07<>6001 AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & strNpSqlOfNoSalesDuty & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "2", "6"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>305 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (305,997,998,995,996) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                
                
                End Select

                strSql = "INSERT INTO R050316_J values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & " ','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
                k = k + 1
                DoEvents
                .MoveNext
            Loop
        End With
        '¥Ø«eµ{§Ç
        If Val(txt1(5)) = 1 Then
            'Add By Cheng 2002/03/28
            '¥ý±NµLµo¤å¤éªº¸ê®Æ¶ñ¤J¤@­È(MaxDate), ¥H¨Ï¦¹Äæ¦ì­È³Ì¤j
            cnnConnection.Execute "Update R050316_J Set R013007='MaxDate' Where R013007 is null Or length(R013007)=0"
            '¥ý§ìµLµo¤å¤é¥B¦¬¤å¤é, ¦¬¤å¸¹³Ì¤jªÌ, ­Y³£¦³µo¤å¤é«h§ì³Ì¤jªÌ
            strSql = "SELECT R013010,MAX(R013007) FROM R050316_J WHERE ((R013018<'B' or R013018>'C')) AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_J WHERE ((R013018<'B' or R013018>'C')) AND (R013010 NOT IN " & strSQL1 & " OR R013007 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            '¦A±Nµo¤å¤éªº¸ê®Æ¬°MaxDateªÌ, §ó·s¬° NULL
            cnnConnection.Execute "Update R050316_J Set R013007=NULL Where R013007='MaxDate'"
            
            strSql = "SELECT R013010,MAX(R013016) FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            '
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013016 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MIN(R013008) FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013008 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "'  GROUP BY R013010 "
            'CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_J WHERE "
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSQL1 = " ("
                strSQL2 = " ("
                Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
                Loop
                strSQL1 = strSQL1 + ") "
                strSQL2 = strSQL2 + ") "
                cnnConnection.Execute "DELETE FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013018 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
                'cnnConnection.Execute "DELETE FROM R050316_E WHERE SUBSTR(R013012,1,1)='C' AND R013012<>'" & chgsql(CheckStr(adoRecordset1.Fields(0))) & "' AND ID='" & strUserNum & "' "
            End If
            CheckOC2
        End If
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/4
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    CheckOC
Case Else
End Select

'³øªí»y¤å
Select Case txt1(6)
Case 1     '¤¤¤å
    PrintDataCp
Case 2     '­^¤å
    PrintDataEp
Case 3     '¤é¤å
    PrintDataJp
Case Else
End Select

Screen.MousePointer = vbDefault
End Sub

Private Sub ProcessT()
'Modify By Cheng 2002/02/27
'****************************************************
'¥N²z¤H¦WºÙ, ¥Ó½Ð¤H¦WºÙ, ¦a§}, ®×¥ó©Ê½è¦WºÙ, ®×¥ó¦WºÙ
'­Y¤¤¤å³øªí  ¤¤-->­^-->¤é
'­Y­^¤å³øªí  ­^-->¤é-->NULL
'­Y¤é¤å³øªí  ¤é-->­^-->NULL
'****************************************************
ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/4 ²M°£¬d¸ß¦Lªí°O¿ýÀÉÄæ¦ì
Screen.MousePointer = vbHourglass
cnnConnection.Execute "delete from R050316_C WHERE ID='" & strUserNum & "' "     '»y¤å¬°¤¤¤å      ¹ï¤º
cnnConnection.Execute "DELETE FROM R050316_E WHERE ID='" & strUserNum & "' "     '»y¤å¬°­^¤å    ¹ï¥~
cnnConnection.Execute "DELETE FROM R050316_J WHERE ID='" & strUserNum & "' "     '»y¤å¬°¤é¤å   ¹ï¥~
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
strSQL11 = " AND tm01=CP01(+) AND tm02=CP02(+) AND tm03=CP03(+) AND tm04=CP04(+)  "
strSQL22 = ""
strSQL55 = ""
StrSQL221 = " AND CP01=tm01(+) AND CP02=tm02(+) AND CP03=tm03(+) AND CP04=tm04(+) "
If Len(txt1(0)) <> 0 Then
    strSQL11 = strSQL11 & " and tm01 in (" & SQLGrpStr(txt1(0), 2) & ") "
    StrSQL221 = StrSQL221 & " and cp01 in (" & SQLGrpStr(txt1(0), 2) & ") "
    pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/10/4
End If
If Len(txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10>='" & txt1(1) & "' "
End If
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND TM10<='" & txt1(2) & "' "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label2(3) & txt1(1) & "-" & txt1(2)  'Add By Sindy 2010/10/4
End If
If Len(txt1(4)) <> 0 Then
    strSQL11 = strSQL11 + " AND (tm44>='" & GetNewFagent(txt1(3)) & "' AND tm44<='" & GetNewFagent(txt1(4)) & "') "
    strSQL22 = strSQL22 + " AND (CP44>='" & GetNewFagent(txt1(3)) & "' AND CP44<='" & GetNewFagent(txt1(4)) & "') "
    strSQL55 = strSQL55 + " AND (CP44>='" & GetNewFagent(txt1(3)) & "' AND CP44<='" & GetNewFagent(txt1(4)) & "') "
    pub_QL05 = pub_QL05 & ";" & Label2(0) & txt1(3) & "-" & txt1(4)  'Add By Sindy 2010/10/4
End If
If Len(txt1(8)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(8)))
   StrSQL3 = StrSQL3 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(8)))
   StrSQL4 = StrSQL4 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(8)))
   strSQL5 = strSQL5 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(8)))
End If
If Len(txt1(7)) <> 0 Then
   strSQL1 = strSQL1 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(7)))
   StrSQL3 = StrSQL3 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(7)))
   StrSQL4 = StrSQL4 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(7)))
   strSQL5 = strSQL5 + " AND CP27<=" & Val(ChangeTStringToWString(txt1(7)))
'Add By Cheng 2002/03/19
Else
   If Len(txt1(8)) <> 0 Then
      strSQL1 = strSQL1 + " AND CP27<=" & Val(ChangeTStringToWString(ServerDate - 19110000))
      StrSQL3 = StrSQL3 + " AND CP27<=" & Val(ChangeTStringToWString(ServerDate - 19110000))
      StrSQL4 = StrSQL4 + " AND CP27<=" & Val(ChangeTStringToWString(ServerDate - 19110000))
      strSQL5 = strSQL5 + " AND CP27<=" & Val(ChangeTStringToWString(ServerDate - 19110000))
   End If
End If
If Len(txt1(8)) <> 0 Or Len(txt1(7)) <> 0 Then
   pub_QL05 = pub_QL05 & ";µo¤å¤é¡G" & txt1(8) & "-" & txt1(7)   'Add By Sindy 2010/10/4
End If
'Add By Cheng 2003/07/17
'¬O§_§t®Ö»é
If Me.txt1(9).Text <> "" Then
    strSQL2 = strSQL2 + " AND ( TM16<>'2' OR TM16 IS NULL ) "
    pub_QL05 = pub_QL05 & ";" & Label6(0) & "¤£§t"  'Add By Sindy 2010/10/4
End If

CheckOC
Select Case txt1(6)
Case 1 '¤¤¤å³øªí
    pub_QL05 = pub_QL05 & ";" & Label6(1) & "¤¤¤å"   'Add By Sindy 2010/10/4
    If txt1(5) = "1" Then '¥Ø«eµ{§Ç
         pub_QL05 = pub_QL05 & ";" & Label3 & "¥Ø«eµ{§Ç"   'Add By Sindy 2010/10/4
'                        strSQL = " SELECT tm44,tm45,NVL(cu04,tm23),tm12,NVL(na03,tm10),NVL(decode(tm10,'000',cpm03,cpm04),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm05,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+)  " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu04,tm23),tm12,NVL(na03,tm10),NVL(decode(tm10,'000',cpm03,cpm04),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm05,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+)  " & strSQL1 & StrSQL22
'92.03.27 nick §ï¥»©Ò®×¸¹
'                        strSQL = " SELECT tm44,tm45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm23),tm12,NVL(na03,tm10),NVL(Nvl(decode(tm10,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+)  " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm23),tm12,NVL(na03,tm10),NVL(Nvl(decode(tm10,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+)  " & strSQL1 & StrSQL22 & StrSQL221
         'Modify By Sindy 2011/3/15 +TM17
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
         'Modified by Lydia 2018/06/05 ­×§ïÅã¥Ü®×¥ó©Ê½è '020',CPM04,CPM03 => '000',CPM03,CPM04
                        strSql = " SELECT tm44,tm45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm23),tm12,NVL(na03,tm10),NVL(Nvl(decode(tm10,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+)  " & strSQL1 & strSQL11
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm23),tm12,NVL(na03,tm10),NVL(Nvl(decode(tm10,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+)  " & strSQL1 & strSQL22 & StrSQL221
         'end 2018/06/05
    Else '©Ò¦³µ{§Ç
         pub_QL05 = pub_QL05 & ";" & Label3 & "©Ò¦³µ{§Ç"   'Add By Sindy 2010/10/4
'                        strSQL = " SELECT tm44,tm45,NVL(cu04,tm23),tm12,NVL(na03,tm10),NVL(decode(tm10,'000',cpm03,cpm04),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm05,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu04,tm23),tm12,NVL(na03,tm10),NVL(decode(tm10,'000',cpm03,cpm04),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm05,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'92.03.27 nick §ï¥»©Ò®×¸¹
'                        strSQL = " SELECT tm44,tm45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm23),tm12,NVL(na03,tm10),NVL(Nvl(decode(tm10,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm23),tm12,NVL(na03,tm10),NVL(Nvl(decode(tm10,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
         'Modify By Sindy 2011/3/15 +TM17
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
         'Modified by Lydia 2018/06/05 ­×§ïÅã¥Ü®×¥ó©Ê½è '020',CPM04,CPM03 => '000',CPM03,CPM04
                        strSql = " SELECT tm44,tm45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm23),tm12,NVL(na03,tm10),NVL(Nvl(decode(tm10,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu04,Null,Decode(CU05,Null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90),CU04),tm23),tm12,NVL(na03,tm10),NVL(Nvl(decode(tm10,'000',cpm03,cpm04),Nvl(CPM10,CPM13)),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm05,Nvl(TM06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
         'end 2018/06/05
    End If
    adoRecordset.CursorLocation = adUseClient
    k = 0
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        With adoRecordset
            .MoveFirst
            DoEvents
            Do While .EOF = False
                For i = 0 To 17
                    strTemp(i) = CheckStr(.Fields(i))
                Next i
                If strTemp(16) = "Y" Then
                    strTemp(9) = "*" + strTemp(9)
                    strTemp(16) = "³¬¨÷"
                Else
                    strTemp(16) = ""
                End If
                CheckOC2
                strSql = "select min(pd05) from pridate where pd01='" & CheckStr(.Fields(18)) & "' and pd02='" & CheckStr(.Fields(19)) & "' and pd03='" & CheckStr(.Fields(20)) & "' and pd04='" & CheckStr(.Fields(21)) & "' "
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 Then
                     strTemp(6) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                Else
                     strTemp(6) = ""
                End If
                Select Case CheckSys(CheckStr(.Fields(18)))
                Case "1", "5"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>411 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "3", "4", "7", "8"
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & " and np07<>6001 AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & strNpSqlOfNoSalesDuty & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "2", "6"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>305 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify By Sindy 2011/3/15 FCT,T,TF©µ®i(102)©M²Ä¤G´Á(716)±M¥ÎÅv¶·¦s¦b(TM17=Y)
                     If (CheckStr(.Fields(18)) = "T" Or CheckStr(.Fields(18)) = "FCT" Or CheckStr(.Fields(18)) = "TF") And _
                        CheckStr("" & .Fields("TM17")) <> "Y" Then
                        'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                        'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (305,997,998,995,996,102,716) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                        strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (102,716) " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     '2011/3/15 End
                     Else
                        'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                        'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (305,997,998,995,996) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                        strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     End If
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                End Select
                strSql = "INSERT INTO R050316_C values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
                k = k + 1
                DoEvents
                .MoveNext
            Loop
        End With
        '¥Ø«eµ{§Ç
        If Val(txt1(5)) = 1 Then
            'Add By Cheng 2002/03/28
            '¥ý±NµLµo¤å¤éªº¸ê®Æ¶ñ¤J¤@­È(MaxDate), ¥H¨Ï¦¹Äæ¦ì­È³Ì¤j
            cnnConnection.Execute "Update R050316_C Set R013007='MaxDate' Where R013007 is null Or length(R013007)=0"
            '¥ý§ìµLµo¤å¤é¥B¦¬¤å¤é, ¦¬¤å¸¹³Ì¤jªÌ, ­Y³£¦³µo¤å¤é«h§ì³Ì¤jªÌ
            strSql = "SELECT R013010,MAX(R013007) FROM R050316_C WHERE ((R013018<'B' or R013018>'C')) AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_C WHERE ((R013018<'B' or R013018>'C')) AND (R013010 NOT IN " & strSQL1 & " OR R013007 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            '¦A±Nµo¤å¤éªº¸ê®Æ¬°MaxDateªÌ, §ó·s¬° NULL
            cnnConnection.Execute "Update R050316_C Set R013007=NULL Where R013007='MaxDate'"
            
            strSql = "SELECT R013010,MAX(R013016) FROM R050316_C WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_C WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013016 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MIN(R013008) FROM R050316_C WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_C WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013008 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_C WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "'  GROUP BY R013010 "
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_c WHERE "
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSQL1 = " ("
                strSQL2 = " ("
                Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
                Loop
                strSQL1 = strSQL1 + ") "
                strSQL2 = strSQL2 + ") "
                cnnConnection.Execute "DELETE FROM R050316_C WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013018 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
        End If
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/4
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    CheckOC
Case 2 '­^¤å³øªí
    'strSQL = " SELECT tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15," & SQLDate("tm11") & ",tm12,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",tm29,tm45,tm06,CP09,tm44 FROM trademark,CASEPROGRESS,patenttrademarkmap,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cP01=CPM01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND tm08=PTM02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL1
    'strSQL = strSQL + " union all select TM01||'-'||TM02||'-'||TM03||'-'||TM04,TM15," & SQLDate("TM11") & ",TM12,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",TM29,TM45,TM06,CP09,TM44 FROM TRADEMARK,CASEPROGRESS,patenttrademarkmap,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cp01=CPM01(+) AND CP10=CPM02(+) AND '2'=ptm02(+) AND TM08=PTM02(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL2
    'strSQL = strSQL + " union all select SP01||'-'||SP02||'-'||SP03||'-'||SP04,SP14," & SQLDate("SP10") & ",SP11,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",SP15,SP27,SP06,CP09,SP26 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cP01=CPM01(+) AND CP10=CPM02(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL3
    'strSQL = strSQL + " union all select LC01||'-'||LC02||'-'||LC03||'-'||LC04,'',0,'',CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",LC08,LC23,LC06,CP09,LC22 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cp01=CPM01(+) AND CP10=CPM02(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL4
    'strSQL = strSQL + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,tm15," & SQLDate("tm11") & ",tm12,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",tm29,CP45,tm06,CP09,CP44 FROM trademark,CASEPROGRESS,patenttrademarkmap,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cP01=CPM01(+) AND CP10=CPM02(+) AND '2'=ptm01(+) AND tm08=PTM02(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL5
    'strSQL = strSQL + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,TM15," & SQLDate("TM11") & ",TM12,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",TM29,CP45,TM06,CP09,CP44 FROM TRADEMARK,CASEPROGRESS,patenttrademarkmap,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cp01=CPM01(+) AND CP10=CPM02(+) AND '2'=ptm02(+) AND TM08=PTM02(+) AND CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL5
    'strSQL = strSQL + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,SP14," & SQLDate("SP10") & ",SP11,CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",SP15,CP45,SP06,CP09,CP44 FROM SERVICEPRACTICE,CASEPROGRESS,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND cP01=CPM01(+) AND CP10=CPM02(+) AND CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL5
    'strSQL = strSQL + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,'',0,'',CPM10," & SQLDate("CP27") & "," & SQLDate("NP08") & ",LC08,CP45,LC06,CP09,CP44 FROM LAWCASE,CASEPROGRESS,CASEPROPERTYMAP,NEXTPROGRESS WHERE NP01=CP09 AND Cp01=CPM01(+) AND CP10=CPM02(+) AND CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND (CP09<'B' or  CP09>'C') " & StrSQL5
    '89/11/23   §ï nick
    pub_QL05 = pub_QL05 & ";" & Label6(1) & "­^¤å"   'Add By Sindy 2010/10/4
    If txt1(5) = "1" Then
         pub_QL05 = pub_QL05 & ";" & Label3 & "¥Ø«eµ{§Ç"   'Add By Sindy 2010/10/4
'Modify By Cheng 2002/02/27
'                        strSQL = " SELECT tm44,tm45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,tm23),tm12,NVL(na04,tm10),NVL(cpm10,CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm06,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,tm23),tm12,NVL(na04,tm10),NVL(cpm10,CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm06,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'92.03.27 nick §ï¥»©Ò®×¸¹
'                        strSQL = " SELECT tm44,tm45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm10,CPM13),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm06,TM07),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm10,CPM13),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm06,TM07),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
         'Modify By Sindy 2011/3/15 +TM17
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
                        strSql = " SELECT tm44,tm45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm10,CPM13),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(TM05, Nvl(tm06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm10,CPM13),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(TM05, Nvl(tm06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
    Else
         pub_QL05 = pub_QL05 & ";" & Label3 & "©Ò¦³µ{§Ç"   'Add By Sindy 2010/10/4
'Modify By Cheng 2002/02/27
'                        strSQL = " SELECT tm44,tm45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,tm23),tm12,NVL(na04,tm10),NVL(cpm10,CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm06,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,tm23),tm12,NVL(na04,tm10),NVL(cpm10,CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm06,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'92.03.27 nick §ï¥»©Ò®×¸¹
'                        strSQL = " SELECT tm44,tm45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm10,CPM13),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm06,TM07),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm10,CPM13),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm06,TM07),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
         'Modify By Sindy 2011/3/15 +TM17
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
                        strSql = " SELECT tm44,tm45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm10,CPM13),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(TM05, Nvl(tm06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(CU05,Null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm10,CPM13),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(TM05, Nvl(tm06,TM07)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
    End If
    adoRecordset.CursorLocation = adUseClient
    k = 0
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        With adoRecordset
            .MoveFirst
            DoEvents
            Do While .EOF = False
                For i = 0 To 17
                    strTemp(i) = CheckStr(.Fields(i))
                Next i
                If strTemp(16) = "Y" Then
                    strTemp(9) = "*" + strTemp(9)
                    strTemp(16) = "Closed"
                Else
                    strTemp(16) = ""
                End If
                CheckOC2
                strSql = "select min(pd05) from pridate where pd01='" & CheckStr(.Fields(18)) & "' and pd02='" & CheckStr(.Fields(19)) & "' and pd03='" & CheckStr(.Fields(20)) & "' and pd04='" & CheckStr(.Fields(21)) & "' "
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 Then
                     strTemp(6) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                Else
                     strTemp(6) = ""
                End If
                Select Case CheckSys(CheckStr(.Fields(18)))
                Case "1", "5"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>411 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (411,997,998,995,996,999,1204) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "3", "4", "7", "8"
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & " and np07<>6001 AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & strNpSqlOfNoSalesDuty & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "2", "6"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>305 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify By Sindy 2011/3/15 FCT,T,TF©µ®i(102)©M²Ä¤G´Á(716)±M¥ÎÅv¶·¦s¦b(TM17=Y)
                     If (CheckStr(.Fields(18)) = "T" Or CheckStr(.Fields(18)) = "FCT" Or CheckStr(.Fields(18)) = "TF") And _
                        CheckStr("" & .Fields("TM17")) <> "Y" Then
                        'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                        'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (305,997,998,995,996,102,716) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                        strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (102,716) " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     '2011/3/15 End
                     Else
                        'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                        'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (305,997,998,995,996) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                        strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     End If
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                End Select
                strSql = "INSERT INTO R050316_E values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & "','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
                k = k + 1
                DoEvents
                .MoveNext
            Loop
        End With
        '¥Ø«eµ{§Ç
        If Val(txt1(5)) = 1 Then
            'Add By Cheng 2002/03/28
            '¥ý±NµLµo¤å¤éªº¸ê®Æ¶ñ¤J¤@­È(MaxDate), ¥H¨Ï¦¹Äæ¦ì­È³Ì¤j
            cnnConnection.Execute "Update R050316_E Set R013007='MaxDate' Where R013007 is null Or length(R013007)=0"
            '¥ý§ìµLµo¤å¤é¥B¦¬¤å¤é, ¦¬¤å¸¹³Ì¤jªÌ, ­Y³£¦³µo¤å¤é«h§ì³Ì¤jªÌ
            strSql = "SELECT R013010,MAX(R013007) FROM R050316_E WHERE ((R013018<'B' or R013018>'C')) AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_E WHERE ((R013018<'B' or R013018>'C')) AND (R013010 NOT IN " & strSQL1 & " OR R013007 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            '¦A±Nµo¤å¤éªº¸ê®Æ¬°MaxDateªÌ, §ó·s¬° NULL
            cnnConnection.Execute "Update R050316_E Set R013007=NULL Where R013007='MaxDate'"
            
            strSql = "SELECT R013010,MAX(R013016) FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013016 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MIN(R013008) FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013008 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "'  GROUP BY R013010 "
            'CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_E WHERE "
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSQL1 = " ("
                strSQL2 = " ("
                Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
                Loop
                strSQL1 = strSQL1 + ") "
                strSQL2 = strSQL2 + ") "
                cnnConnection.Execute "DELETE FROM R050316_E WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013018 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
        End If
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/4
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    CheckOC
Case 3 '¤é¤å³øªí
    pub_QL05 = pub_QL05 & ";" & Label6(1) & "¤é¤å"   'Add By Sindy 2010/10/4
    If txt1(5) = "1" Then
         pub_QL05 = pub_QL05 & ";" & Label3 & "¥Ø«eµ{§Ç"   'Add By Sindy 2010/10/4
'                        strSQL = " SELECT tm44,tm45,NVL(cu06,tm23),tm12,NVL(na04,tm10),NVL(cpm13,CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm07,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,tm23),tm12,NVL(na04,tm10),NVL(cpm13,CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm07,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'92.03.27 nick §ï¥»©Ò®×¸¹
'                        strSQL = " SELECT tm44,tm45,NVL(Decode(cu06,Null,CU05||' '||cu88||' '||CU89||' '||CU90,CU06),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm13,CPM10),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm07,TM06),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu06,Null,CU05||' '||cu88||' '||CU89||' '||CU90,CU06),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm13,CPM10),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm07,TM06),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
         'Modify By Sindy 2011/3/15 +TM17
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
                        strSql = " SELECT tm44,tm45,NVL(Decode(cu06,Null,CU05||' '||cu88||' '||CU89||' '||CU90,CU06),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm13,CPM10),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(TM05, Nvl(tm07,TM06)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu06,Null,CU05||' '||cu88||' '||CU89||' '||CU90,CU06),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm13,CPM10),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(TM05, Nvl(tm07,TM06)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
    Else
         pub_QL05 = pub_QL05 & ";" & Label3 & "©Ò¦³µ{§Ç"   'Add By Sindy 2010/10/4
'                        strSQL = " SELECT tm44,tm45,NVL(cu06,tm23),tm12,NVL(na04,tm10),NVL(cpm13,CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm07,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(cu06,tm23),tm12,NVL(na04,tm10),NVL(cpm13,CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,tm07,CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
'92.03.27 nick §ï¥»©Ò®×¸¹
'                        strSQL = " SELECT tm44,tm45,NVL(Decode(cu06,Null,CU05||' '||cu88||' '||CU89||' '||CU90,CU06),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm13,CPM10),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm07,TM06),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
'         strSQL = strSQL + " union all select CP44,CP45,NVL(Decode(cu06,Null,CU05||' '||cu88||' '||CU89||' '||CU90,CU06),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm13,CPM10),CP10),' ',' ',tm35,tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm15,tm21||'='||tm22,tm14,tm09,Nvl(tm07,TM06),CP05,tm29,cp09,cp01,cp02,cp03,cp04 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  CP09>'C') AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & StrSQL22 & StrSQL221
         'Modify By Sindy 2011/3/15 +TM17
         'Modify by Amy 2017/11/10 ­ì:CP09<'B' or  CP09>'C' §ï­ç°£DÃþ
                        strSql = " SELECT tm44,tm45,NVL(Decode(cu06,Null,CU05||' '||cu88||' '||CU89||' '||CU90,CU06),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm13,CPM10),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(TM05, Nvl(tm07,TM06)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL11
         strSql = strSql + " union all select CP44,CP45,NVL(Decode(cu06,Null,CU05||' '||cu88||' '||CU89||' '||CU90,CU06),tm23),tm12,NVL(na04,tm10),NVL(Nvl(cpm13,CPM10),CP10),' ',' ',tm35,tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04),tm15,tm21||'='||tm22,tm14,tm09,Nvl(TM05, Nvl(tm07,TM06)),CP05,tm29,cp09,cp01,cp02,cp03,cp04,TM17 FROM trademark,CASEPROGRESS,CASEPROPERTYMAP,CUSTOMER,NATION WHERE cP01=CPM01(+) AND CP10=CPM02(+) AND (CP09<'B' or  (CP09>'C' and  CP09<'D')) AND SUBSTR(tm23,1,8)=CU01(+) AND DECODE(SUBSTR(tm23,9,1),NULL,'0',SUBSTR(tm23,9,1))=CU02(+) AND tm10=NA01(+) " & strSQL1 & strSQL22 & StrSQL221
    End If
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    k = 0
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
        With adoRecordset
            .MoveFirst
            DoEvents
            Do While .EOF = False
                For i = 0 To 17
                    strTemp(i) = CheckStr(.Fields(i))
                Next i
                If strTemp(16) = "Y" Then
                    strTemp(9) = "*" + strTemp(9)
                    strTemp(16) = "ÇVÆê"
                Else
                    strTemp(16) = ""
                End If
                CheckOC2
                strSql = "select min(pd05) from pridate where pd01='" & CheckStr(.Fields(18)) & "' and pd02='" & CheckStr(.Fields(19)) & "' and pd03='" & CheckStr(.Fields(20)) & "' and pd04='" & CheckStr(.Fields(21)) & "' "
                adoRecordset1.CursorLocation = adUseClient
                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If adoRecordset1.RecordCount <> 0 Then
                     strTemp(6) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                Else
                     strTemp(6) = ""
                End If
                Select Case CheckSys(CheckStr(.Fields(18)))
                Case "1", "5"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>411 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (411,997,998,995,996,999,1204) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "3", "4", "7", "8"
                     'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & " and np07<>6001 AND NP06 IS NULL "
                     strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' AND NP08>=" & Val(GetTodayDate) & strNpSqlOfNoSalesDuty & " AND NP06 IS NULL "
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                Case "2", "6"
                     'edit by nick 2004/09/20
                     'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07<>305 AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     'Modify by Morgan 2009/7/13 +995,996
                     'Modify By Sindy 2011/3/15 FCT,T,TF©µ®i(102)©M²Ä¤G´Á(716)±M¥ÎÅv¶·¦s¦b(TM17=Y)
                     If (CheckStr(.Fields(18)) = "T" Or CheckStr(.Fields(18)) = "FCT" Or CheckStr(.Fields(18)) = "TF") And _
                        CheckStr("" & .Fields("TM17")) <> "Y" Then
                        'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                        'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (305,997,998,995,996,102,716) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                        strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (102,716) " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     '2011/3/15 End
                     Else
                        'Modify by Morgan 2011/6/10 ±Æ°£µ{§ÇºÞ¨îªº®×¥ó©Ê½è§ï¥Î strNpSqlOfNoSalesDuty ±`¼Æ
                        'strSQL = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' and NP07 not in (305,997,998,995,996) AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                        strSql = "select  MIN(np08)  from nextprogress where np02='" & CheckStr(.Fields(18)) & "' and np03='" & CheckStr(.Fields(19)) & "' and np04='" & CheckStr(.Fields(20)) & "' and np05='" & CheckStr(.Fields(21)) & "' " & strNpSqlOfNoSalesDuty & " AND NP08>=" & Val(GetTodayDate) & " AND NP06 IS NULL "
                     End If
                     CheckOC2
                     adoRecordset1.CursorLocation = adUseClient
                     adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If adoRecordset1.RecordCount <> 0 Then
                        strTemp(7) = ChangeTStringToTDateString(ChangeWStringToTString(CheckStr(adoRecordset1.Fields(0))))
                     Else
                        strTemp(7) = ""
                     End If
                     CheckOC2
                End Select
                strSql = "INSERT INTO R050316_J values ('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & ChgSQL(strTemp(9)) & "','" & ChgSQL(strTemp(10)) & "','" & ChgSQL(strTemp(11)) & "','" & ChgSQL(strTemp(12)) & " ','" & ChgSQL(strTemp(13)) & "','" & ChgSQL(strTemp(14)) & "','" & ChgSQL(strTemp(15)) & "','" & ChgSQL(strTemp(16)) & "','" & ChgSQL(strTemp(17)) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
                k = k + 1
                DoEvents
                .MoveNext
            Loop
        End With
        '¥Ø«eµ{§Ç
        If Val(txt1(5)) = 1 Then
            'Add By Cheng 2002/03/28
            '¥ý±NµLµo¤å¤éªº¸ê®Æ¶ñ¤J¤@­È(MaxDate), ¥H¨Ï¦¹Äæ¦ì­È³Ì¤j
            cnnConnection.Execute "Update R050316_J Set R013007='MaxDate' Where R013007 is null Or length(R013007)=0"
            '¥ý§ìµLµo¤å¤é¥B¦¬¤å¤é, ¦¬¤å¸¹³Ì¤jªÌ, ­Y³£¦³µo¤å¤é«h§ì³Ì¤jªÌ
            strSql = "SELECT R013010,MAX(R013007) FROM R050316_J WHERE ((R013018<'B' or R013018>'C')) AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_J WHERE ((R013018<'B' or R013018>'C')) AND (R013010 NOT IN " & strSQL1 & " OR R013007 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            '¦A±Nµo¤å¤éªº¸ê®Æ¬°MaxDateªÌ, §ó·s¬° NULL
            cnnConnection.Execute "Update R050316_J Set R013007=NULL Where R013007='MaxDate'"
            
            strSql = "SELECT R013010,MAX(R013016) FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 & " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 & " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013016 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MIN(R013008) FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "' GROUP BY R013010 "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 Then
               strSQL1 = " ("
               strSQL2 = " ("
               Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
               Loop
               strSQL1 = strSQL1 + ") "
               strSQL2 = strSQL2 + ") "
               cnnConnection.Execute "DELETE FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013008 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
            End If
            CheckOC2
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND ID='" & strUserNum & "'  GROUP BY R013010 "
            'CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            strSql = "SELECT R013010,MAX(R013018) FROM R050316_J WHERE "
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strSQL1 = " ("
                strSQL2 = " ("
                Do While adoRecordset1.EOF = False
                  strSQL1 = strSQL1 + " '" & ChgSQL(CheckStr(adoRecordset1.Fields(0))) & "' "
                  strSQL2 = strSQL2 + " '" & CheckStr(adoRecordset1.Fields(1)) & "' "
                  adoRecordset1.MoveNext
                  If adoRecordset1.EOF = False Then
                     strSQL1 = strSQL1 + ","
                     strSQL2 = strSQL2 + ","
                  End If
                Loop
                strSQL1 = strSQL1 + ") "
                strSQL2 = strSQL2 + ") "
                cnnConnection.Execute "DELETE FROM R050316_J WHERE (R013018<'B' or R013018>'C') AND (R013010 NOT IN " & strSQL1 & " OR R013018 NOT IN " & strSQL2 & ") AND ID='" & strUserNum & "' "
                'cnnConnection.Execute "DELETE FROM R050316_E WHERE SUBSTR(R013012,1,1)='C' AND R013012<>'" & chgsql(CheckStr(adoRecordset1.Fields(0))) & "' AND ID='" & strUserNum & "' "
            End If
            CheckOC2
        End If
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/4
        ShowNoData
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    CheckOC
Case Else
End Select
Select Case txt1(6)
Case 1     '¤¤¤å
    PrintDataCt
Case 2     '­^¤å
    PrintDataEt
Case 3     '¤é¤å
    PrintDataJt
Case Else
End Select
Screen.MousePointer = vbDefault

End Sub


Private Sub PrintDataCp()
strSql = "SELECT DISTINCT R013001 FROM R050316_C WHERE ID='" & strUserNum & "' GROUP BY R013001 "
CheckOC
Page = 1
SeekTempPrint = ""
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    'Add By Cheng 2002/02/27
    RefreshColData 1, "¥N²z¤H"
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst
        Do While .EOF = False
            strTemp(20) = CheckStr(.Fields(0))
            'Modify By Cheng 2002/02/27
'            strSQL = "SELECT FA04,'','','',FA17,'','','','',FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA07,FA52 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            'Modify by Morgan 2011/5/26 +FA70
            strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA07,FA52,Decode(FA17,Null,decode(FA32,null,FA70)) FA70 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                'StrTemp5(0) = GetNewFagent(strTemp(20)) + "  " + CheckStr(adoRecordset1.Fields(0))      ªô¤p©j»¡­n§ï¦^¨Ó  ¤£¨q¥N¸¹ 90/01/17
                StrTemp5(0) = CheckStr(adoRecordset1.Fields(0))
                StrTemp5(1) = CheckStr(adoRecordset1.Fields(1))
                StrTemp5(2) = CheckStr(adoRecordset1.Fields(2))
                StrTemp5(3) = CheckStr(adoRecordset1.Fields(3))
                StrTemp5(4) = CheckStr(adoRecordset1.Fields(4))
                StrTemp5(6) = CheckStr(adoRecordset1.Fields(5))
                StrTemp5(8) = CheckStr(adoRecordset1.Fields(6))
                StrTemp5(10) = CheckStr(adoRecordset1.Fields(7))
                StrTemp5(11) = CheckStr(adoRecordset1.Fields(8))
                StrTemp5(5) = CheckStr(adoRecordset1.Fields(9))
                StrTemp5(7) = CheckStr(adoRecordset1.Fields(10))
                StrTemp5(9) = CheckStr(adoRecordset1.Fields(11))
                StrTemp5(12) = CheckStr(adoRecordset1.Fields(12))
                StrTemp5(13) = CheckStr(adoRecordset1.Fields(13))
                StrTemp5(14) = CheckStr(adoRecordset1.Fields("FA70")) 'Add by Morgan 2011/5/26
            Else
                For i = 0 To 14
                    StrTemp5(i) = ""
                Next i
            End If
            CheckOC2
            PrintTitleCp
            'Modify By Cheng 2002/03/27
            'µo¤å¤é¥Ñ¤p¨ì¤j, NULL¦b«á¥[¥H±Æ§Ç
'            strSQL = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_C,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,cp05 desc,cp09 desc "
            strSql = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_C,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,r013007,cp09 desc "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                adoRecordset1.MoveFirst
                s = adoRecordset1.RecordCount
                Do While adoRecordset1.EOF = False
                    For i = 0 To 16
                        'If i <> 10 Then
                        '   strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1)) & "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
                        'Else
                           strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1))
                        'End If
                    Next i
                    If SeekTempPrint = "" Then
                        SeekTempPrint = strTemp(8)
                        strTemp(0) = StrToStr(strTemp(0), 15)
                        strTemp(1) = StrToStr(strTemp(1), 10)
                        strTemp(2) = StrToStr(strTemp(2), 10)
                        strTemp(3) = StrToStr(strTemp(3), 6)
                        strTemp(4) = StrToStr(strTemp(4), 38)
                        strTemp(7) = StrToStr(strTemp(7), 15)
                        strTemp(8) = StrToStr(strTemp(8), 7.5)
                        strTemp(9) = Format(StrToStr(strTemp(9) & String(20, " "), 10) & " ", "@@@@@@@@@@@@@@@@@@@@") & "/(" & StrToStr(Format(ChangeWStringToWDateString(strTemp(10)) & String(10, " "), "@@@@@@@@@@"), 5) & ")/" & strTemp(11)
                        strTemp(10) = StrToStr(strTemp(12), 10)
                        strTemp(11) = StrToStr(strTemp(13), 25)
                        strTemp(13) = strTemp(6)
                        strTemp(12) = ""
                        strTemp(6) = StrToStr(strTemp(15), 11)
                    Else
                        If SeekTempPrint <> strTemp(8) Then
                           Printer.Line (0, iPrint)-(19200, iPrint)
                           SeekTempPrint = strTemp(8)
                           iPrint = iPrint + 300
                           If iPrint >= 14200 Then
                              Page = Page + 1
                              Printer.NewPage
                              PrintTitleCp
                           End If
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 38)
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = Format(StrToStr(strTemp(9) & String(20, " "), 10) & " ", "@@@@@@@@@@@@@@@@@@@@") & "/(" & StrToStr(Format(ChangeWStringToWDateString(strTemp(10)) & String(10, " "), "@@@@@@@@@@"), 5) & ")/" & strTemp(11)
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 25)
                           strTemp(13) = strTemp(6)
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                        Else
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 38)
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = Format(StrToStr(strTemp(9) & String(20, " "), 10) & " ", "@@@@@@@@@@@@@@@@@@@@") & "/(" & StrToStr(Format(ChangeWStringToWDateString(strTemp(10)) & String(10, " "), "@@@@@@@@@@"), 5) & ")/" & strTemp(11)
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 25)
                           strTemp(13) = strTemp(6)
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                           For i = 0 To 3
                              strTemp(i) = ""
                           Next i
                           For i = 6 To 13
                              strTemp(i) = ""
                           Next i
                        End If
                    End If
'                    strTemp(0) = StrToStr(strTemp(0), 15)
'                    strTemp(1) = StrToStr(strTemp(1), 10)
'                    strTemp(2) = StrToStr(strTemp(2), 10)
'                    strTemp(3) = StrToStr(strTemp(3), 6)
'                    strTemp(4) = StrToStr(strTemp(4), 38)
 '                   strTemp(7) = StrToStr(strTemp(7), 15)
 '                   strTemp(8) = StrToStr(strTemp(8), 7.5)
 '                   strTemp(9) = Format(StrToStr(strTemp(9) & String(20, " "), 10) & " ", "@@@@@@@@@@@@@@@@@@@@") & "/(" & StrToStr(Format(ChangeWStringToWDateString(strTemp(10)) & String(10, " "), "@@@@@@@@@@"), 5) & ")/" & strTemp(11)
 '                   strTemp(10) = StrToStr(strTemp(12), 10)
 '                   strTemp(11) = StrToStr(strTemp(13), 25)
 '                   strTemp(13) = strTemp(6)
 '                   strTemp(12) = ""
 '                   strTemp(6) = StrToStr(strTemp(15), 11)
                    If iPrint >= 14200 Then
                        Page = Page + 1
                        Printer.Line (0, iPrint)-(19200, iPrint)

                        Printer.NewPage
                        PrintTitleCp
                    End If
                    PrintDatilCp
                    adoRecordset1.MoveNext
                Loop
                Printer.Line (0, iPrint)-(19200, iPrint)

                iPrint = iPrint + 100
                If iPrint >= 14200 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitleCp
                End If
                
                strSql = "SELECT distinct DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_c WHERE R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) "
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                Printer.CurrentX = PLeft(5) - 3000
                Printer.CurrentY = iPrint
                Printer.Print "Á`­p¡G" & Format(IntTot, "##0")
                strSql = "SELECT distinct DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_c WHERE R013001='" & strTemp(20) & "' AND R013017='³¬¨÷' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) "
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                End If
                Printer.CurrentX = PLeft(5)
                Printer.CurrentY = iPrint
                Printer.Print "³¬¨÷¥ó¼Æ¡G" & Format(IntTot, "##0")
                iPrint = iPrint + 300
                CheckOC3
                'Add by Amy 2014/12/11
                Page = Page + 1
                Printer.NewPage
                SeekTempPrint = ""
                'end 2014/12/11
            End If
            CheckOC2
            .MoveNext
        Loop
    End With
End If
CheckOC
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(250, "-")
Printer.EndDoc
ShowPrintOk
'add by nickc 2007/05/05
IsHavePrintP = True
End Sub

Private Sub PrintDataCt()
strSql = "SELECT DISTINCT R013001 FROM R050316_C WHERE ID='" & strUserNum & "' GROUP BY R013001 "
CheckOC
Page = 1
SeekTempPrint = ""
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    'Add By Cheng 2002/02/27
    RefreshColData 1, "¥N²z¤H"
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst
        Do While .EOF = False
            strTemp(20) = CheckStr(.Fields(0))
            'Modify By Cheng 2002/02/27
'            strSQL = "SELECT FA04,'','','',FA17,'','','','',FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA07,FA52 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            'Modify by Morgan 2011/5/26 +FA70
            strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA07,FA52,Decode(FA17,Null,decode(FA32,null,FA70)) FA70 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                'StrTemp5(0) = GetNewFagent(strTemp(20)) + "  " + CheckStr(adoRecordset1.Fields(0))      ªô¤p©j»¡­n§ï¦^¨Ó  ¤£¨q¥N¸¹ 90/01/17
                StrTemp5(0) = CheckStr(adoRecordset1.Fields(0))
                StrTemp5(1) = CheckStr(adoRecordset1.Fields(1))
                StrTemp5(2) = CheckStr(adoRecordset1.Fields(2))
                StrTemp5(3) = CheckStr(adoRecordset1.Fields(3))
                StrTemp5(4) = CheckStr(adoRecordset1.Fields(4))
                StrTemp5(6) = CheckStr(adoRecordset1.Fields(5))
                StrTemp5(8) = CheckStr(adoRecordset1.Fields(6))
                StrTemp5(10) = CheckStr(adoRecordset1.Fields(7))
                StrTemp5(11) = CheckStr(adoRecordset1.Fields(8))
                StrTemp5(5) = CheckStr(adoRecordset1.Fields(9))
                StrTemp5(7) = CheckStr(adoRecordset1.Fields(10))
                StrTemp5(9) = CheckStr(adoRecordset1.Fields(11))
                StrTemp5(12) = CheckStr(adoRecordset1.Fields(12))
                StrTemp5(13) = CheckStr(adoRecordset1.Fields(13))
                StrTemp5(14) = CheckStr(adoRecordset1.Fields("FA70")) 'Add by Morgan 2011/5/26
            Else
                For i = 0 To 14
                    StrTemp5(i) = ""
                Next i
            End If
            CheckOC2
            PrintTitleCt
            'Modify By Cheng 2002/03/27
            'µo¤å¤é¥Ñ¤p¨ì¤j, NULL¦b«á¥[¥H±Æ§Ç
'            strSQL = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_C,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,cp05 desc,cp09 desc "
            strSql = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_C,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,r013007,cp09 desc "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                adoRecordset1.MoveFirst
                s = adoRecordset1.RecordCount
                Do While adoRecordset1.EOF = False
                    For i = 0 To 16
                        'If i <> 10 Then
                        '  strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1)) & "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
                        'Else
                           strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1))
                        'End If
                    Next i
                    If SeekTempPrint = "" Then
                        SeekTempPrint = strTemp(8)
                        strTemp(0) = StrToStr(strTemp(0), 15)
                        strTemp(1) = StrToStr(strTemp(1), 10)
                        strTemp(2) = StrToStr(strTemp(2), 10)
                        strTemp(3) = StrToStr(strTemp(3), 6)
                        strTemp(4) = StrToStr(strTemp(4), 28)
                        strTemp(7) = StrToStr(strTemp(7), 15)
                        strTemp(8) = StrToStr(strTemp(8), 7.5)
                        strTemp(9) = Format(StrToStr(strTemp(9) & String(20, " "), 10) & " ", "@@@@@@@@@@@@@@@@@@@@") & "/(" & StrToStr(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 1)) & String(10, " "), "@@@@@@@@@@"), 5) & "-" & StrToStr(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 2)) & String(10, " "), "@@@@@@@@@@"), 5) & ")/" & Format(ChangeTStringToTDateString(ChangeWStringToTString(strTemp(11))) & " ", "@@@@@@@@@@")
                        strTemp(10) = StrToStr(strTemp(12), 10)
                        strTemp(11) = StrToStr(strTemp(13), 15)
                        strTemp(13) = strTemp(6)
                        strTemp(12) = ""
                        strTemp(6) = StrToStr(strTemp(15), 11)
                    Else
                        If SeekTempPrint <> strTemp(8) Then
                           Printer.Line (0, iPrint)-(19200, iPrint)
                           SeekTempPrint = strTemp(8)
                           iPrint = iPrint + 300
                           If iPrint >= 14200 Then
                              Page = Page + 1
                              Printer.NewPage
                              PrintTitleCt
                           End If
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 28)
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = Format(StrToStr(strTemp(9) & String(20, " "), 10) & " ", "@@@@@@@@@@@@@@@@@@@@") & "/(" & StrToStr(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 1)) & String(10, " "), "@@@@@@@@@@"), 5) & "-" & StrToStr(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 2)) & String(10, " "), "@@@@@@@@@@"), 5) & ")/" & Format(ChangeTStringToTDateString(ChangeWStringToTString(strTemp(11))) & " ", "@@@@@@@@@@")
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 15)
                           strTemp(13) = strTemp(6)
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                        Else
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 28)
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = Format(StrToStr(strTemp(9) & String(20, " "), 10) & " ", "@@@@@@@@@@@@@@@@@@@@") & "/(" & StrToStr(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 1)) & String(10, " "), "@@@@@@@@@@"), 5) & "-" & StrToStr(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 2)) & String(10, " "), "@@@@@@@@@@"), 5) & ")/" & Format(ChangeTStringToTDateString(ChangeWStringToTString(strTemp(11))) & " ", "@@@@@@@@@@")
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 15)
                           strTemp(13) = strTemp(6)
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                           For i = 0 To 3
                              strTemp(i) = ""
                           Next i
                           For i = 6 To 13
                              strTemp(i) = ""
                           Next i
                        End If
                    End If
'                    strTemp(0) = StrToStr(strTemp(0), 15)
'                    strTemp(1) = StrToStr(strTemp(1), 10)
'                    strTemp(2) = StrToStr(strTemp(2), 10)
'                    strTemp(3) = StrToStr(strTemp(3), 6)
'                    strTemp(4) = StrToStr(strTemp(4), 38)
 '                   strTemp(7) = StrToStr(strTemp(7), 15)
 '                   strTemp(8) = StrToStr(strTemp(8), 7.5)
 '                   strTemp(9) = Format(StrToStr(strTemp(9) & String(20, " "), 10) & " ", "@@@@@@@@@@@@@@@@@@@@") & "/(" & StrToStr(Format(ChangeWStringToWDateString(strTemp(10)) & String(10, " "), "@@@@@@@@@@"), 5) & ")/" & strTemp(11)
 '                   strTemp(10) = StrToStr(strTemp(12), 10)
 '                   strTemp(11) = StrToStr(strTemp(13), 25)
 '                   strTemp(13) = strTemp(6)
 '                   strTemp(12) = ""
 '                   strTemp(6) = StrToStr(strTemp(15), 11)
                    If iPrint >= 14200 Then
                        Page = Page + 1
                        Printer.Line (0, iPrint)-(19200, iPrint)

                        Printer.NewPage
                        PrintTitleCt
                    End If
                    PrintDatilCt
                    adoRecordset1.MoveNext
                Loop
                Printer.Line (0, iPrint)-(19200, iPrint)

                iPrint = iPrint + 100
                If iPrint >= 14200 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitleCt
                End If
                
                strSql = "SELECT distinct DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_c WHERE R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) "
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                Printer.CurrentX = PLeft(5) - 3000
                Printer.CurrentY = iPrint
                Printer.Print "Á`­p¡G" & Format(IntTot, "##0")
                strSql = "SELECT distinct DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_c WHERE R013001='" & strTemp(20) & "' AND R013017='³¬¨÷' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) "
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                End If
                Printer.CurrentX = PLeft(5)
                Printer.CurrentY = iPrint
                Printer.Print "³¬¨÷¥ó¼Æ¡G" & Format(IntTot, "##0")
                iPrint = iPrint + 300
                CheckOC3
                'Add by Amy 2014/12/11
                Page = Page + 1
                Printer.NewPage
                SeekTempPrint = ""
                'end 2014/12/11
            End If
            CheckOC2
            .MoveNext
        Loop
    End With
End If
CheckOC
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(250, "-")
Printer.EndDoc
ShowPrintOk
End Sub


Private Sub PrintTitleCp()
GetPleftCp
iPrint = 500
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.FontName = "²Ó©úÅé"
Printer.CurrentX = 9000 - (Printer.TextWidth(GetTitleNick & "¥N²z¤H®×¥óÁ`Ã¯") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "¥N²z¤H®×¥óÁ`Ã¯"
iPrint = iPrint + 500
'Printer.CurrentX = 9000 - (Printer.TextWidth("Prepared by Tai E International Patent & Law Office") / 2)
'Printer.CurrentY = iPrint
'Printer.Print "Prepared by Tai E International Patent & Law Office"
'iPrint = iPrint + 500
Printer.Font.Size = 10
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 300
'USER ½T»{®æ¦¡¨Ã¤£¥Î USER ID
'89/11/14
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Print "USER ID¡G" & strUserNum
'iPrint = iPrint + 300
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(GetTaiwanTodayDate)
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "¦¬¥ó¤H¡G"
If Len(StrTemp5(12)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(12)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(13)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(13)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(0)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(0)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(1)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(1)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(2)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(2)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(3)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(3)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(12)) = 0 And Len(StrTemp5(13)) = 0 And Len(StrTemp5(0)) = 0 And Len(StrTemp5(1)) = 0 And Len(StrTemp5(2)) = 0 And Len(StrTemp5(3)) = 0 Then
   iPrint = iPrint + 300
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "¦a§}¡G" & StrTemp5(4)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "¹q¸Ü¡G" & StrTemp5(5)
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "Page¡G" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("¦a§}¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(6)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "E-mail¡G" & StrTemp5(7)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("¦a§}¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(8)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "¶Ç¯u¡G" & StrTemp5(9)
iPrint = iPrint + 300
If Len(StrTemp5(10)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦a§}¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(10)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(11)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦a§}¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(11)
   iPrint = iPrint + 300
End If
'Add by Morgan 2011/5/26
If Len(StrTemp5(14)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦a§}¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(14)
   iPrint = iPrint + 300
End If
'end 2011/5/26
If Len(StrTemp5(10)) = 0 And Len(strTemp(11)) = 0 Then
   iPrint = iPrint + 300
End If
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "©¼©Ò®×¸¹"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
'If intPCaseKind = 2 Then
Printer.Print "¥Ó½Ð¤H"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "¥Ó½Ð®×¸¹"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "¥Ó½Ð°ê®a"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "®×¥ó©Ê½è"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "µo¤å¤é"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "¬O§_³¬¨÷"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "«È¤á®×¥ó®×¸¹"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "¥»©Ò®×¸¹"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "±M§Q¸¹¼Æ¡@¡@/(±M§QÅv¨ì´Á¤é)/±M§QÅv¬O§_¦s¦b"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "±M§QºØÃþ"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "®×¥ó¦WºÙ"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "³Æµù"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "³Ìªñ´Á­­"
iPrint = iPrint + 300
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
End Sub

Private Sub PrintDatilCt()
For i = 0 To 6
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
For i = 7 To 13
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
Next i
'Printer.Line (0, iPrint + 150)-(19200, iPrint + 150)
iPrint = iPrint + 300
End Sub

Private Sub PrintTitleCt()
GetPleftCt
iPrint = 500
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.FontName = "²Ó©úÅé"
Printer.CurrentX = 9000 - (Printer.TextWidth(GetTitleNick & "¥N²z¤H®×¥óÁ`Ã¯") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "¥N²z¤H®×¥óÁ`Ã¯"
iPrint = iPrint + 500
'Printer.CurrentX = 9000 - (Printer.TextWidth("Prepared by Tai E International Patent & Law Office") / 2)
'Printer.CurrentY = iPrint
'Printer.Print "Prepared by Tai E International Patent & Law Office"
'iPrint = iPrint + 500
Printer.Font.Size = 10
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 300
'USER ½T»{®æ¦¡¨Ã¤£¥Î USER ID
'89/11/14
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Print "USER ID¡G" & strUserNum
'iPrint = iPrint + 300
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(GetTaiwanTodayDate)
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "¦¬¥ó¤H¡G"
If Len(StrTemp5(12)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(12)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(13)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(13)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(0)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(0)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(1)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(1)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(2)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(2)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(3)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦¬¥ó¤H¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(3)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(12)) = 0 And Len(StrTemp5(13)) = 0 And Len(StrTemp5(0)) = 0 And Len(StrTemp5(1)) = 0 And Len(StrTemp5(2)) = 0 And Len(StrTemp5(3)) = 0 Then
   iPrint = iPrint + 300
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "¦a§}¡G" & StrTemp5(4)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "¹q¸Ü¡G" & StrTemp5(5)
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "Page¡G" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("¦a§}¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(6)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "E-mail¡G" & StrTemp5(7)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("¦a§}¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(8)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "¶Ç¯u¡G" & StrTemp5(9)
iPrint = iPrint + 300
If Len(StrTemp5(10)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦a§}¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(10)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(11)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦a§}¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(11)
   iPrint = iPrint + 300
End If
'Add by Morgan 2011/5/26
If Len(StrTemp5(14)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦a§}¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(14)
   iPrint = iPrint + 300
End If
'end 2011/5/26
If Len(StrTemp5(10)) = 0 And Len(strTemp(11)) = 0 Then
   iPrint = iPrint + 300
End If
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "©¼©Ò®×¸¹"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
'If intPCaseKind = 2 Then
Printer.Print "¥Ó½Ð¤H"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "¥Ó½Ð®×¸¹"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "¥Ó½Ð°ê®a"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "®×¥ó©Ê½è"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "Àu¥ýÅv¤é"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "¬O§_³¬¨÷"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "«È¤á®×¥ó®×¸¹"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "¥»©Ò®×¸¹"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "¼f©w¸¹¼Æ¡@¡@¡@¡@ ¡@¡@/(±M¥Î´Á¶¡)¡@¡@ ¡@¡@¡@¡@/¤½§i¤é¡@¡@¡@    "
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "°Ó«~Ãþ§O"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "°Ó¼Ð¦WºÙ"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "³Æµù"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "³Ìªñ´Á­­"
iPrint = iPrint + 300
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
End Sub

Private Sub PrintDatilCp()
For i = 0 To 6
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
For i = 7 To 13
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
Next i
'Printer.Line (0, iPrint + 150)-(19200, iPrint + 150)
iPrint = iPrint + 300
End Sub


Private Sub GetPleftCp()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 3100
PLeft(2) = 5200
PLeft(3) = 7300
PLeft(4) = 8600
PLeft(5) = 16300
PLeft(6) = 17900
PLeft(7) = 0
PLeft(8) = 3100
PLeft(9) = 4800
'Modify By Cheng 2002/11/15
'PLeft(10) = 9000 + 200
PLeft(10) = 9000 + 700
PLeft(11) = 11100 + 200
PLeft(12) = 16400 + 200
PLeft(13) = 17900
End Sub

Private Sub GetPleftCt()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 3100
PLeft(2) = 5200
PLeft(3) = 7300
PLeft(4) = 8600
PLeft(5) = 16300 - 2000
PLeft(6) = 17900
PLeft(7) = 0
PLeft(8) = 3100
PLeft(9) = 4800
PLeft(10) = 9000 + 200 + 1500
PLeft(11) = 11100 + 200 + 1500
PLeft(12) = 16400 + 200
PLeft(13) = 17900
End Sub


Private Sub PrintDataEp()
strSql = "SELECT DISTINCT R013001 FROM R050316_E WHERE ID='" & strUserNum & "' GROUP BY R013001 "
CheckOC
Page = 1
SeekTempPrint = ""
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    'Add By Cheng 2002/02/27
    RefreshColData 2, "¥N²z¤H"
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst
        Do While .EOF = False
            strTemp(20) = CheckStr(.Fields(0))
            'Modify By Cheng 2002/02/27
'            strSQL = "SELECT FA05,FA63,FA64,FA65,FA18,FA19,FA20,FA21,FA22,FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA08,FA53 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            'Modify by Morgan 2011/5/26 +FA70
            strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA08,FA53,decode(FA32,null,FA70) FA70 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                'StrTemp5(0) = GetNewFagent(strTemp(20)) + "  " + CheckStr(adoRecordset1.Fields(0))      ªô¤p©j»¡­n§ï¦^¨Ó  ¤£¨q¥N¸¹ 90/01/17
                StrTemp5(0) = CheckStr(adoRecordset1.Fields(0))
                StrTemp5(1) = CheckStr(adoRecordset1.Fields(1))
                StrTemp5(2) = CheckStr(adoRecordset1.Fields(2))
                StrTemp5(3) = CheckStr(adoRecordset1.Fields(3))
                StrTemp5(4) = CheckStr(adoRecordset1.Fields(4))
                StrTemp5(6) = CheckStr(adoRecordset1.Fields(5))
                StrTemp5(8) = CheckStr(adoRecordset1.Fields(6))
                StrTemp5(10) = CheckStr(adoRecordset1.Fields(7))
                StrTemp5(11) = CheckStr(adoRecordset1.Fields(8))
                StrTemp5(5) = CheckStr(adoRecordset1.Fields(9))
                StrTemp5(7) = CheckStr(adoRecordset1.Fields(10))
                StrTemp5(9) = CheckStr(adoRecordset1.Fields(11))
                StrTemp5(12) = CheckStr(adoRecordset1.Fields(12))
                StrTemp5(13) = CheckStr(adoRecordset1.Fields(13))
                StrTemp5(14) = CheckStr(adoRecordset1.Fields("FA70")) 'Add by Morgan 2011/5/26
            Else
                For i = 0 To 14
                    StrTemp5(i) = ""
                Next i
            End If
            CheckOC2
            PrintTitleEp
            'strSQL = "SELECT * FROM R050316_E WHERE R013001='" & chgsql(strTemp(14)) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) DESC"
            'Modify By Cheng 2002/03/27
            'µo¤å¤é¥Ñ¤p¨ì¤j, NULL¦b«á¥[¥H±Æ§Ç
'            strSQL = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_e,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,cp05 desc,cp09 desc "
            strSql = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_e,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,r013007,cp09 desc "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                adoRecordset1.MoveFirst
                s = adoRecordset1.RecordCount
                Do While adoRecordset1.EOF = False
                    For i = 0 To 16
                        'If i <> 10 Then
                        '   strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1)) & "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
                        'Else
                           strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1))
                        'End If
                    Next i
                    If SeekTempPrint = "" Then
                        SeekTempPrint = strTemp(8)
                        strTemp(0) = StrToStr(strTemp(0), 15)
                        strTemp(1) = StrToStr(strTemp(1), 10)
                        strTemp(2) = StrToStr(strTemp(2), 10)
                        strTemp(3) = StrToStr(strTemp(3), 6)
                        strTemp(4) = StrToStr(strTemp(4), 38)
                        strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                        strTemp(7) = StrToStr(strTemp(7), 15)
                        strTemp(8) = StrToStr(strTemp(8), 7.5)
                        strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(strTemp(10)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & strTemp(11)
                        strTemp(10) = StrToStr(strTemp(12), 10)
                        strTemp(11) = StrToStr(strTemp(13), 25)
                        strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                        strTemp(12) = ""
                        strTemp(6) = StrToStr(strTemp(15), 11)
                    Else
                        If SeekTempPrint <> strTemp(8) Then
                           Printer.Line (0, iPrint)-(19200, iPrint)
                           SeekTempPrint = strTemp(8)
                           iPrint = iPrint + 300
                           If iPrint >= 14200 Then
                              Page = Page + 1
                              Printer.NewPage
                              PrintTitleEp
                           End If
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 38)
                           strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(strTemp(10)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & strTemp(11)
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 25)
                           strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                        Else
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 38)
                           strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(strTemp(10)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & strTemp(11)
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 25)
                           strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                           For i = 0 To 3
                              strTemp(i) = ""
                           Next i
                           For i = 6 To 13
                              strTemp(i) = ""
                           Next i
                        End If
                    End If
                    If iPrint >= 14200 Then
                        Page = Page + 1
                        Printer.Line (0, iPrint)-(19200, iPrint)

                        Printer.NewPage
                        PrintTitleEp
                    End If
                    PrintDatilEp
                    adoRecordset1.MoveNext
                Loop
                Printer.Line (0, iPrint)-(19200, iPrint)

                iPrint = iPrint + 100
                If iPrint >= 14200 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitleEp
                End If
                
                strSql = "SELECT distinct DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_e WHERE R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) "
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                End If
                Printer.CurrentX = PLeft(5) - 3000
                Printer.CurrentY = iPrint
                Printer.Print "Total¡G" & Format(IntTot, "##0")
                strSql = "SELECT distinCt DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_e WHERE R013001='" & strTemp(20) & "' AND R013017='Closed' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010)"
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                End If
                Printer.CurrentX = PLeft(5)
                Printer.CurrentY = iPrint
                Printer.Print "File Closed¡G" & Format(IntTot, "##0")
                iPrint = iPrint + 300
                CheckOC3
                'Add by Amy 2014/12/11
                Page = Page + 1
                Printer.NewPage
                SeekTempPrint = ""
                'end 2014/12/11
            End If
            CheckOC2
            .MoveNext
        Loop
    End With
End If
CheckOC
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(250, "-")
Printer.EndDoc
ShowPrintOk
'add by nickc 2007/05/25
IsHavePrintP = True
End Sub

Private Sub PrintDataEt()
strSql = "SELECT DISTINCT R013001 FROM R050316_E WHERE ID='" & strUserNum & "' GROUP BY R013001 "
CheckOC
Page = 1
SeekTempPrint = ""
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    'Add By Cheng 2002/02/27
    RefreshColData 2, "¥N²z¤H"
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst
        Do While .EOF = False
            strTemp(20) = CheckStr(.Fields(0))
            'Modify By Cheng 2002/02/27
'            strSQL = "SELECT FA05,FA63,FA64,FA65,FA18,FA19,FA20,FA21,FA22,FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA08,FA53 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            'Modify by Morgan 2011/5/26 +FA70
            strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA08,FA53,decode(FA32,null,FA70) FA70 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                'StrTemp5(0) = GetNewFagent(strTemp(20)) + "  " + CheckStr(adoRecordset1.Fields(0))      ªô¤p©j»¡­n§ï¦^¨Ó  ¤£¨q¥N¸¹ 90/01/17
                StrTemp5(0) = CheckStr(adoRecordset1.Fields(0))
                StrTemp5(1) = CheckStr(adoRecordset1.Fields(1))
                StrTemp5(2) = CheckStr(adoRecordset1.Fields(2))
                StrTemp5(3) = CheckStr(adoRecordset1.Fields(3))
                StrTemp5(4) = CheckStr(adoRecordset1.Fields(4))
                StrTemp5(6) = CheckStr(adoRecordset1.Fields(5))
                StrTemp5(8) = CheckStr(adoRecordset1.Fields(6))
                StrTemp5(10) = CheckStr(adoRecordset1.Fields(7))
                StrTemp5(11) = CheckStr(adoRecordset1.Fields(8))
                StrTemp5(5) = CheckStr(adoRecordset1.Fields(9))
                StrTemp5(7) = CheckStr(adoRecordset1.Fields(10))
                StrTemp5(9) = CheckStr(adoRecordset1.Fields(11))
                StrTemp5(12) = CheckStr(adoRecordset1.Fields(12))
                StrTemp5(13) = CheckStr(adoRecordset1.Fields(13))
                StrTemp5(14) = CheckStr(adoRecordset1.Fields("FA70")) 'Add by Morgan 2011/5/26
            Else
                For i = 0 To 14
                    StrTemp5(i) = ""
                Next i
            End If
            CheckOC2
            PrintTitleEt
            'strSQL = "SELECT * FROM R050316_E WHERE R013001='" & chgsql(strTemp(14)) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) DESC"
            'Modify By Cheng 2002/03/27
            'µo¤å¤é¥Ñ¤p¨ì¤j, NULL¦b«á¥[¥H±Æ§Ç
'            strSQL = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_e,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,cp05 desc,cp09 desc "
            strSql = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_e,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,r013007,cp09 desc "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                adoRecordset1.MoveFirst
                s = adoRecordset1.RecordCount
                Do While adoRecordset1.EOF = False
                    For i = 0 To 16
                        'If i <> 10 Then
                        '   strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1)) & "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
                        'Else
                           strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1))
                        'End If
                    Next i
                    If SeekTempPrint = "" Then
                        SeekTempPrint = strTemp(8)
                        strTemp(0) = StrToStr(strTemp(0), 15)
                        strTemp(1) = StrToStr(strTemp(1), 10)
                        strTemp(2) = StrToStr(strTemp(2), 10)
                        strTemp(3) = StrToStr(strTemp(3), 6)
                        strTemp(4) = StrToStr(strTemp(4), 28)
                        strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                        strTemp(7) = StrToStr(strTemp(7), 15)
                        strTemp(8) = StrToStr(strTemp(8), 7.5)
                        strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 1)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "-" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 2)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & Format(ChangeWStringToWDateString(strTemp(11)), "mmm dd,yyyy")
                        strTemp(10) = StrToStr(strTemp(12), 10)
                        strTemp(11) = StrToStr(strTemp(13), 15)
                        strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                        strTemp(12) = ""
                        strTemp(6) = StrToStr(strTemp(15), 11)
                    Else
                        If SeekTempPrint <> strTemp(8) Then
                           Printer.Line (0, iPrint)-(19200, iPrint)
                           SeekTempPrint = strTemp(8)
                           iPrint = iPrint + 300
                           If iPrint >= 14200 Then
                              Page = Page + 1
                              Printer.NewPage
                              PrintTitleEt
                           End If
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 28)
                           strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 1)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "-" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 2)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & Format(ChangeWStringToWDateString(strTemp(11)), "mmm dd,yyyy")
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 15)
                           strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                        Else
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 28)
                           strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 1)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "-" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 2)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & Format(ChangeWStringToWDateString(strTemp(11)), "mmm dd,yyyy")
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 15)
                           strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                           For i = 0 To 3
                              strTemp(i) = ""
                           Next i
                           For i = 6 To 13
                              strTemp(i) = ""
                           Next i
                        End If
                    End If
                    If iPrint >= 14200 Then
                        Page = Page + 1
                        Printer.Line (0, iPrint)-(19200, iPrint)

                        Printer.NewPage
                        PrintTitleEt
                    End If
                    PrintDatilEt
                    adoRecordset1.MoveNext
                Loop
                Printer.Line (0, iPrint)-(19200, iPrint)

                iPrint = iPrint + 100
                If iPrint >= 14200 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitleEt
                End If
                
                strSql = "SELECT distinct DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_e WHERE R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) "
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                End If
                Printer.CurrentX = PLeft(5) - 3000
                Printer.CurrentY = iPrint
                Printer.Print "Total¡G" & Format(IntTot, "##0")
                strSql = "SELECT distinCt DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_e WHERE R013001='" & strTemp(20) & "' AND R013017='Closed' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010)"
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                End If
                Printer.CurrentX = PLeft(5)
                Printer.CurrentY = iPrint
                Printer.Print "File Closed¡G" & Format(IntTot, "##0")
                iPrint = iPrint + 300
                CheckOC3
                'Add by Amy 2014/12/11
                Page = Page + 1
                Printer.NewPage
                SeekTempPrint = ""
                'end 2014/12/11
            End If
            CheckOC2
            .MoveNext
        Loop
    End With
End If
CheckOC
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(250, "-")
Printer.EndDoc
ShowPrintOk
End Sub


Private Sub PrintDataJp()
strSql = "SELECT DISTINCT R013001 FROM R050316_J WHERE ID='" & strUserNum & "' GROUP BY R013001 "
CheckOC
Page = 1
SeekTempPrint = ""
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    'Add By Cheng 2002/02/27
    RefreshColData 3, "¥N²z¤H"
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst
        Do While .EOF = False
            strTemp(20) = CheckStr(.Fields(0))
            'Modify By Cheng 2002/02/27
'            strSQL = "SELECT FA06,'','','',FA23,'','','','',FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA09,FA54 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            'Modify by Morgan 2011/5/26 +FA70
            strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA09,FA54,Decode(FA23,Null,decode(FA32,null,FA70)) FA70 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                'StrTemp5(0) = GetNewFagent(strTemp(20)) + "  " + CheckStr(adoRecordset1.Fields(0))      ªô¤p©j»¡­n§ï¦^¨Ó  ¤£¨q¥N¸¹ 90/01/17
                StrTemp5(0) = CheckStr(adoRecordset1.Fields(0))
                StrTemp5(1) = CheckStr(adoRecordset1.Fields(1))
                StrTemp5(2) = CheckStr(adoRecordset1.Fields(2))
                StrTemp5(3) = CheckStr(adoRecordset1.Fields(3))
                StrTemp5(4) = CheckStr(adoRecordset1.Fields(4))
                StrTemp5(6) = CheckStr(adoRecordset1.Fields(5))
                StrTemp5(8) = CheckStr(adoRecordset1.Fields(6))
                StrTemp5(10) = CheckStr(adoRecordset1.Fields(7))
                StrTemp5(11) = CheckStr(adoRecordset1.Fields(8))
                StrTemp5(5) = CheckStr(adoRecordset1.Fields(9))
                StrTemp5(7) = CheckStr(adoRecordset1.Fields(10))
                StrTemp5(9) = CheckStr(adoRecordset1.Fields(11))
                StrTemp5(12) = CheckStr(adoRecordset1.Fields(12))
                StrTemp5(13) = CheckStr(adoRecordset1.Fields(13))
                StrTemp5(14) = CheckStr(adoRecordset1.Fields("FA70")) 'Add by Morgan 2011/5/26
            Else
                For i = 0 To 14
                    StrTemp5(i) = ""
                Next i
            End If
            CheckOC2
            PrintTitleJp
            'strSQL = "SELECT * FROM R050316_J WHERE R013001='" & chgsql(strTemp(14)) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) DESC"
            'Modify By Cheng 2002/03/27
            'µo¤å¤é¥Ñ¤p¨ì¤j, NULL¦b«á¥[¥H±Æ§Ç
'            strSQL = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_j,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,cp05 desc,cp09 desc "
            strSql = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_j,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,r013007,cp09 desc "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                adoRecordset1.MoveFirst
                s = adoRecordset1.RecordCount
                Do While adoRecordset1.EOF = False
                    For i = 0 To 16
                        'If i <> 10 Then
                        '   strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1)) & "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
                        'Else
                           strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1))
                        'End If
                    Next i
                    If SeekTempPrint = "" Then
                        SeekTempPrint = strTemp(8)
                        strTemp(0) = StrToStr(strTemp(0), 15)
                        strTemp(1) = StrToStr(strTemp(1), 10)
                        strTemp(2) = StrToStr(strTemp(2), 10)
                        strTemp(3) = StrToStr(strTemp(3), 6)
                        strTemp(4) = StrToStr(strTemp(4), 38)
                        strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                        strTemp(7) = StrToStr(strTemp(7), 15)
                        strTemp(8) = StrToStr(strTemp(8), 7.5)
                        strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(strTemp(10)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & strTemp(11)
                        strTemp(10) = StrToStr(strTemp(12), 10)
                        strTemp(11) = StrToStr(strTemp(13), 25)
                        strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                        strTemp(12) = ""
                        strTemp(6) = StrToStr(strTemp(15), 11)
                    Else
                        If SeekTempPrint <> strTemp(8) Then
                           Printer.Line (0, iPrint)-(19200, iPrint)
                           SeekTempPrint = strTemp(8)
                           iPrint = iPrint + 300
                           If iPrint >= 14200 Then
                              Page = Page + 1
                              Printer.NewPage
                              PrintTitleJp
                           End If
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 38)
                           strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(strTemp(10)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & strTemp(11)
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 25)
                           strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                        Else
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 38)
                           strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(strTemp(10)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & strTemp(11)
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 25)
                           strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                           For i = 0 To 3
                              strTemp(i) = ""
                           Next i
                           For i = 6 To 13
                              strTemp(i) = ""
                           Next i
                        End If
                    End If
                    If iPrint >= 14200 Then
                        Page = Page + 1
                        Printer.Line (0, iPrint)-(19200, iPrint)

                        Printer.NewPage
                        PrintTitleJp
                    End If
                    PrintDatilJp
                    adoRecordset1.MoveNext
                Loop
                Printer.Line (0, iPrint)-(19200, iPrint)
                iPrint = iPrint + 100
                If iPrint >= 14200 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitleJp
                End If
                strSql = "SELECT distinct DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_J WHERE R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) "
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                End If
                Printer.CurrentX = PLeft(5) - 3000
                Printer.CurrentY = iPrint
                Printer.Print "¦X­p¡G" & Format(IntTot, "##0")
                strSql = "SELECT distinCt DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_J WHERE R013001='" & strTemp(20) & "' AND R013017='ÇVÆê' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010)"
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                End If
                Printer.CurrentX = PLeft(5)
                Printer.CurrentY = iPrint
                Printer.Print "¥XÄ@’ÇÇU³¬Âê¡G" & Format(IntTot, "##0")
                iPrint = iPrint + 300
                CheckOC3
                'Add by Amy 2014/12/11
                Page = Page + 1
                Printer.NewPage
                SeekTempPrint = ""
                'end 2014/12/11
            End If
            CheckOC2
            .MoveNext
        Loop
    End With
End If
CheckOC
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(250, "-")
Printer.EndDoc
ShowPrintOk
'add by nickc 2007/05/25
IsHavePrintP = True
End Sub

Private Sub PrintDataJt()
strSql = "SELECT DISTINCT R013001 FROM R050316_J WHERE ID='" & strUserNum & "' GROUP BY R013001 "
CheckOC
Page = 1
SeekTempPrint = ""
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    'Add By Cheng 2002/02/27
    RefreshColData 3, "¥N²z¤H"
    With adoRecordset
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/10/4
        .MoveFirst
        Do While .EOF = False
            strTemp(20) = CheckStr(.Fields(0))
            'Modify By Cheng 2002/02/27
'            strSQL = "SELECT FA06,'','','',FA23,'','','','',FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA09,FA54 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            'Modify by Morgan 2011/5/26 +FA70
            strSql = "SELECT " & m_ColAgName & "," & m_ColAgAdd & ",FA12||DECODE(FA13,NULL,'',','||FA13),FA16,FA14||DECODE(FA15,NULL,'',','||FA15),FA09,FA54,Decode(FA23,Null,decode(FA32,null,FA70)) FA70 FROM FAGENT WHERE FA01='" & Mid(GetNewFagent(strTemp(20)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(strTemp(20)), 9, 1) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                'StrTemp5(0) = GetNewFagent(strTemp(20)) + "  " + CheckStr(adoRecordset1.Fields(0))      ªô¤p©j»¡­n§ï¦^¨Ó  ¤£¨q¥N¸¹ 90/01/17
                StrTemp5(0) = CheckStr(adoRecordset1.Fields(0))
                StrTemp5(1) = CheckStr(adoRecordset1.Fields(1))
                StrTemp5(2) = CheckStr(adoRecordset1.Fields(2))
                StrTemp5(3) = CheckStr(adoRecordset1.Fields(3))
                StrTemp5(4) = CheckStr(adoRecordset1.Fields(4))
                StrTemp5(6) = CheckStr(adoRecordset1.Fields(5))
                StrTemp5(8) = CheckStr(adoRecordset1.Fields(6))
                StrTemp5(10) = CheckStr(adoRecordset1.Fields(7))
                StrTemp5(11) = CheckStr(adoRecordset1.Fields(8))
                StrTemp5(5) = CheckStr(adoRecordset1.Fields(9))
                StrTemp5(7) = CheckStr(adoRecordset1.Fields(10))
                StrTemp5(9) = CheckStr(adoRecordset1.Fields(11))
                StrTemp5(12) = CheckStr(adoRecordset1.Fields(12))
                StrTemp5(13) = CheckStr(adoRecordset1.Fields(13))
                StrTemp5(14) = CheckStr(adoRecordset1.Fields("FA70")) 'Add by Morgan 2011/5/26
            Else
                For i = 0 To 14
                    StrTemp5(i) = ""
                Next i
            End If
            CheckOC2
            PrintTitleJt
            'strSQL = "SELECT * FROM R050316_J WHERE R013001='" & chgsql(strTemp(14)) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) DESC"
            'Modify By Cheng 2002/03/27
            'µo¤å¤é¥Ñ¤p¨ì¤j, NULL¦b«á¥[¥H±Æ§Ç
'            strSQL = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_j,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,cp05 desc,cp09 desc "
            strSql = "SELECT r013001,r013002,r013003,r013004,r013005,r013006,r013007,r013008,r013009,r013010,r013011,r013012,r013013,r013014,r013015,r013016,r013017,r013018,cp05,cp09 FROM R050316_j,caseprogress WHERE r013018=cp09(+) and R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' ORDER BY DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) Desc,r013007,cp09 desc "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                adoRecordset1.MoveFirst
                s = adoRecordset1.RecordCount
                Do While adoRecordset1.EOF = False
                    For i = 0 To 16
                        'If i <> 10 Then
                        '   strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1)) & "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
                        'Else
                           strTemp(i) = CheckStr(adoRecordset1.Fields(i + 1))
                        'End If
                    Next i
                    If SeekTempPrint = "" Then
                        SeekTempPrint = strTemp(8)
                        strTemp(0) = StrToStr(strTemp(0), 15)
                        strTemp(1) = StrToStr(strTemp(1), 10)
                        strTemp(2) = StrToStr(strTemp(2), 10)
                        strTemp(3) = StrToStr(strTemp(3), 6)
                        strTemp(4) = StrToStr(strTemp(4), 28)
                        strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                        strTemp(7) = StrToStr(strTemp(7), 15)
                        strTemp(8) = StrToStr(strTemp(8), 7.5)
                        strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 1)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "-" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 2)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & Format(ChangeWStringToWDateString(strTemp(11)), "mmm dd,yyyy")
                        strTemp(10) = StrToStr(strTemp(12), 10)
                        strTemp(11) = StrToStr(strTemp(13), 15)
                        strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                        strTemp(12) = ""
                        strTemp(6) = StrToStr(strTemp(15), 11)
                    Else
                        If SeekTempPrint <> strTemp(8) Then
                           Printer.Line (0, iPrint)-(19200, iPrint)
                           SeekTempPrint = strTemp(8)
                           iPrint = iPrint + 300
                           If iPrint >= 14200 Then
                              Page = Page + 1
                              Printer.NewPage
                              PrintTitleJt
                           End If
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 28)
                           strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 1)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "-" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 2)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & Format(ChangeWStringToWDateString(strTemp(11)), "mmm dd,yyyy")
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 15)
                           strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                        Else
                           strTemp(0) = StrToStr(strTemp(0), 15)
                           strTemp(1) = StrToStr(strTemp(1), 10)
                           strTemp(2) = StrToStr(strTemp(2), 10)
                           strTemp(3) = StrToStr(strTemp(3), 6)
                           strTemp(4) = StrToStr(strTemp(4), 28)
                           strTemp(5) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(5)))), "mmm dd,yyyy")
                           strTemp(7) = StrToStr(strTemp(7), 15)
                           strTemp(8) = StrToStr(strTemp(8), 7.5)
                           strTemp(9) = StrToStr(strTemp(9) & String(20, " "), 10) & "/" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 1)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "-" & StrToStr(Format(Format(ChangeWStringToWDateString(StringTwoString(strTemp(10), 2)), "mmm dd,yyyy") & String(11, " "), "@@@@@@@@@@@"), 5.5) & "/" & Format(ChangeWStringToWDateString(strTemp(11)), "mmm dd,yyyy")
                           strTemp(10) = StrToStr(strTemp(12), 10)
                           strTemp(11) = StrToStr(strTemp(13), 15)
                           strTemp(13) = Format(ChangeWStringToWDateString(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(6)))), "mmm dd,yyyy")
                           strTemp(12) = ""
                           strTemp(6) = StrToStr(strTemp(15), 11)
                           For i = 0 To 3
                              strTemp(i) = ""
                           Next i
                           For i = 6 To 13
                              strTemp(i) = ""
                           Next i
                        End If
                    End If
                    If iPrint >= 14200 Then
                        Page = Page + 1
                        Printer.Line (0, iPrint)-(19200, iPrint)

                        Printer.NewPage
                        PrintTitleJt
                    End If
                    PrintDatilJt
                    adoRecordset1.MoveNext
                Loop
                Printer.Line (0, iPrint)-(19200, iPrint)
                iPrint = iPrint + 100
                If iPrint >= 14200 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitleJt
                End If
                strSql = "SELECT distinct DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_J WHERE R013001='" & strTemp(20) & "' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) "
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                End If
                Printer.CurrentX = PLeft(5) - 3000
                Printer.CurrentY = iPrint
                Printer.Print "¦X­p¡G" & Format(IntTot, "##0")
                strSql = "SELECT distinCt DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010) FROM R050316_J WHERE R013001='" & strTemp(20) & "' AND R013017='ÇVÆê' AND ID='" & strUserNum & "' group by DECODE(SUBSTR(R013010,1,1),'*',SUBSTR(R013010,2,LENGTH(R013010)-1),R013010)"
                CheckOC3
                AdoRecordSet3.CursorLocation = adUseClient
                AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                IntTot = 0
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
                        IntTot = AdoRecordSet3.RecordCount
                End If
                End If
                Printer.CurrentX = PLeft(5)
                Printer.CurrentY = iPrint
                Printer.Print "¥XÄ@’ÇÇU³¬Âê¡G" & Format(IntTot, "##0")
                iPrint = iPrint + 300
                CheckOC3
                'Add by Amy 2014/12/11
                Page = Page + 1
                Printer.NewPage
                SeekTempPrint = ""
                'end 2014/12/11
            End If
            CheckOC2
            .MoveNext
        Loop
    End With
End If
CheckOC
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(250, "-")
Printer.EndDoc
ShowPrintOk
End Sub


Private Sub PrintTitleEt()
GetPleftEt
iPrint = 500
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.FontName = "²Ó©úÅé"
Printer.CurrentX = 9000 - (Printer.TextWidth("Status Report On TradeMark/Applications") / 2)
Printer.CurrentY = iPrint
Printer.Print "Status Report On TradeMark/Applications"
iPrint = iPrint + 500
Printer.CurrentX = 9000 - (Printer.TextWidth("Prepared by Tai E International Patent & Law Office") / 2)
Printer.CurrentY = iPrint
Printer.Print "Prepared by Tai E International Patent & Law Office"
iPrint = iPrint + 500
Printer.Font.Size = 10
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 300
'USER ½T»{­^¤å®æ¦¡¨Ã¤£¥Î USER ID
'89/11/14
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Print "USER ID¡G" & strUserNum
'iPrint = iPrint + 300
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "Date¡G" & Format(ChangeWStringToWDateString(GetTodayDate), "mmm dd,yyyy")
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "To¡G"
If Len(StrTemp5(12)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(12)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(13)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(13)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(0)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(0)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(1)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(1)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(2)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(2)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(3)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(3)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(12)) = 0 And Len(StrTemp5(13)) = 0 And Len(StrTemp5(0)) = 0 And Len(StrTemp5(1)) = 0 And Len(StrTemp5(2)) = 0 And Len(StrTemp5(3)) = 0 Then
   iPrint = iPrint + 300
End If

Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "Address¡G" & StrTemp5(4)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "Tel¡G" & StrTemp5(5)
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "Page¡G" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("Address¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(6)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "E-mail¡G" & StrTemp5(7)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("Address¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(8)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "Fax¡G" & StrTemp5(9)
iPrint = iPrint + 300
If Len(StrTemp5(10)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("Address¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(10)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(11)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("Address¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(11)
   iPrint = iPrint + 300
End If
'Add by Morgan 2011/5/26
If Len(StrTemp5(14)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("Address¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(14)
   iPrint = iPrint + 300
End If
'end 2011/5/26

If Len(StrTemp5(10)) = 0 And Len(strTemp(11)) = 0 Then
   iPrint = iPrint + 300
End If
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "Your Ref"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
'If intPCaseKind = 2 Then
Printer.Print "Applicant"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "Appln. No."
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "Country"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "Status"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "Relevant Filing Date"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "File Closed"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "Case No."
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "Our Ref"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "Approved No.¡@¡@¡@¡@ /(Duration)¡@¡@¡@ ¡@¡@¡@/Publication Date"
'Printer.Print "¼f©w¸¹¼Æ¡@¡@¡@¡@ ¡@¡@/(±M¥Î´Á¶¡)¡@¡@ ¡@¡@¡@¡@/¤½§i¤é¡@¡@¡@    "
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "Class(es)"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "Mark"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "Remarks"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "Next Due Day"
iPrint = iPrint + 300
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
End Sub

Private Sub PrintTitleEp()
GetPleftEp
iPrint = 500
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.FontName = "²Ó©úÅé"
Printer.CurrentX = 9000 - (Printer.TextWidth("Status Report On Patents/Applications") / 2)
Printer.CurrentY = iPrint
Printer.Print "Status Report On Patents/Applications"
iPrint = iPrint + 500
Printer.CurrentX = 9000 - (Printer.TextWidth("Prepared by Tai E International Patent & Law Office") / 2)
Printer.CurrentY = iPrint
Printer.Print "Prepared by Tai E International Patent & Law Office"
iPrint = iPrint + 500
Printer.Font.Size = 10
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 300
'USER ½T»{­^¤å®æ¦¡¨Ã¤£¥Î USER ID
'89/11/14
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Print "USER ID¡G" & strUserNum
'iPrint = iPrint + 300
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "Date¡G" & Format(ChangeWStringToWDateString(GetTodayDate), "mmm dd,yyyy")
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "To¡G"
If Len(StrTemp5(12)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(12)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(13)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(13)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(0)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(0)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(1)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(1)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(2)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(2)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(3)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(3)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(12)) = 0 And Len(StrTemp5(13)) = 0 And Len(StrTemp5(0)) = 0 And Len(StrTemp5(1)) = 0 And Len(StrTemp5(2)) = 0 And Len(StrTemp5(3)) = 0 Then
   iPrint = iPrint + 300
End If

Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "Address¡G" & StrTemp5(4)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "Tel¡G" & StrTemp5(5)
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "Page¡G" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("Address¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(6)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "E-mail¡G" & StrTemp5(7)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("Address¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(8)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "Fax¡G" & StrTemp5(9)
iPrint = iPrint + 300
If Len(StrTemp5(10)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("Address¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(10)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(11)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("Address¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(11)
   iPrint = iPrint + 300
End If
'Add by Morgan 2011/5/26
If Len(StrTemp5(14)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("Address¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(14)
   iPrint = iPrint + 300
End If
'end 2011/5/26
If Len(StrTemp5(10)) = 0 And Len(strTemp(11)) = 0 Then
   iPrint = iPrint + 300
End If
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "Your Ref"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
'If intPCaseKind = 2 Then
Printer.Print "Applicant"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "Appln. No."
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "Country"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "Status"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "Date Of Filing"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "File Closed"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "Case No."
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "Our Ref"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "Pat. No.¡@¡@/Expiry Date/Patent Right"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "Type Of Patent"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "Title"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "Remarks"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "Next Due Day"
iPrint = iPrint + 300
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
End Sub


Private Sub PrintTitleJt()
Dim strCmpTitle As String 'Add by Amy 2020/03/30

GetPleftEt
iPrint = 500
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.FontName = "²Ó©úÅé"
Printer.CurrentX = 9000 - (Printer.TextWidth("®×¥óÇUª¬ªp³ø§i") / 2)
Printer.CurrentY = iPrint
Printer.Print "®×¥óÇUª¬ªp³ø§i"
iPrint = iPrint + 500
'Modifed by Lydia 2018/03/16
'Printer.CurrentX = 9000 - (Printer.TextWidth("¥x¤@ï»Ú‘®§Qªk«ß¨Æ°È©Ò  §@¦¨") / 2)
'Modify by Amy 2020/04/01
'strCmpTitle = Taie_Jpn_Title & "  §@¦¨"
strCmpTitle = CompNameQuery(2, 3) & "  §@¦¨"
Printer.CurrentX = 9000 - (Printer.TextWidth(strCmpTitle) / 2)
Printer.CurrentY = iPrint
'Modifed by Lydia 2018/03/16
'Printer.Print "¥x¤@ï»Ú‘®§Qªk«ß¨Æ°È©Ò  §@¦¨"
Printer.Print strCmpTitle 'Taie_Jpn_Title & "  §@¦¨"
'end 2020/04/01
iPrint = iPrint + 500
Printer.Font.Size = 10
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 300
'USER ½T»{­^¤å®æ¦¡¨Ã¤£¥Î USER ID
'89/11/14
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Print "USER ID¡G" & strUserNum
'iPrint = iPrint + 300
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "Date¡G" & Format(ChangeWStringToWDateString(GetTodayDate), "mmm dd,yyyy")
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "To¡G"
If Len(StrTemp5(12)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(12)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(13)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(13)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(0)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(0)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(1)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(1)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(2)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(2)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(3)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(3)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(12)) = 0 And Len(StrTemp5(13)) = 0 And Len(StrTemp5(0)) = 0 And Len(StrTemp5(1)) = 0 And Len(StrTemp5(2)) = 0 And Len(StrTemp5(3)) = 0 Then
   iPrint = iPrint + 300
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "¦í©Ò¡G" & StrTemp5(4)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "Tel¡G" & StrTemp5(5)
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "Page¡G" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("¦í©Ò¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(6)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "E-mail¡G" & StrTemp5(7)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("¦í©Ò¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(8)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "Fax¡G" & StrTemp5(9)
iPrint = iPrint + 300
If Len(StrTemp5(10)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦í©Ò¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(10)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(11)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦í©Ò¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(11)
   iPrint = iPrint + 300
End If
'Add by Morgan 2011/5/26
If Len(StrTemp5(14)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦í©Ò¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(14)
   iPrint = iPrint + 300
End If
'end 2011/5/26
If Len(StrTemp5(10)) = 0 And Len(strTemp(11)) = 0 Then
   iPrint = iPrint + 300
End If
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "¶Q¾ã²zµfî"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
'If intPCaseKind = 2 Then
Printer.Print "¥XÄ@¤H"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "¥XÄ@µfî"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "ï¦W"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "ª¬ªp"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "Àu¥ýÛ¤é"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "¥XÄ@’ÇÇU³¬Âê"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "Case No."
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "’J¾ã²zµfî"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "µn¿ýµfî¡@¡@¡@¡@ ¡@¡@/¦s‘Ò´Á¶¡¡@¡@¡@¡@¡@¡@¡@ /¤½§i¦~¤ë¤é¡@¡@¡@"
'Printer.Print "¼f©w¸¹¼Æ¡@¡@¡@¡@ ¡@¡@/(±M¥Î´Á¶¡)¡@¡@ ¡@¡@¡@¡@/¤½§i¤é¡@¡@¡@    "
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "°Ó¼ÐÂ¤À"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "°Ó¼Ð¦W‘b"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "µù"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "¦¸ÇU´Á­­"
iPrint = iPrint + 300
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
End Sub

Private Sub PrintTitleJp()
Dim strCmpTitle As String 'Add by Amy 2020/03/30

GetPleftEp
iPrint = 500
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.FontName = "²Ó©úÅé"
Printer.CurrentX = 9000 - (Printer.TextWidth("®×¥óÇUª¬ªp³ø§i") / 2)
Printer.CurrentY = iPrint
Printer.Print "®×¥óÇUª¬ªp³ø§i"
iPrint = iPrint + 500
'Modifed by Lydia 2018/03/16
'Printer.CurrentX = 9000 - (Printer.TextWidth("¥x¤@ï»Ú‘®§Qªk«ß¨Æ°È©Ò  §@¦¨") / 2)
'Moidfy by Amy 2020/04/01
'strCmpTitle = Taie_Jpn_Title & "  §@¦¨"
strCmpTitle = CompNameQuery(2, 3) & "  §@¦¨"
Printer.CurrentX = 9000 - (Printer.TextWidth(strCmpTitle) / 2)
Printer.CurrentY = iPrint
'Modifed by Lydia 2018/03/16
'Printer.Print "¥x¤@ï»Ú‘®§Qªk«ß¨Æ°È©Ò  §@¦¨"
Printer.Print strCmpTitle 'Taie_Jpn_Title & "  §@¦¨"
'end 2020/04/01
iPrint = iPrint + 500
Printer.Font.Size = 10
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 300
'USER ½T»{­^¤å®æ¦¡¨Ã¤£¥Î USER ID
'89/11/14
'Printer.CurrentX = 0
'Printer.CurrentY = iPrint
'Printer.Print "USER ID¡G" & strUserNum
'iPrint = iPrint + 300
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "Date¡G" & Format(ChangeWStringToWDateString(GetTodayDate), "mmm dd,yyyy")
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "To¡G"
If Len(StrTemp5(12)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(12)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(13)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(13)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(0)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(0)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(1)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(1)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(2)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(2)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(3)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("To¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(3)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(12)) = 0 And Len(StrTemp5(13)) = 0 And Len(StrTemp5(0)) = 0 And Len(StrTemp5(1)) = 0 And Len(StrTemp5(2)) = 0 And Len(StrTemp5(3)) = 0 Then
   iPrint = iPrint + 300
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "¦í©Ò¡G" & StrTemp5(4)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "Tel¡G" & StrTemp5(5)
Printer.CurrentX = 16000
Printer.CurrentY = iPrint
Printer.Print "Page¡G" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("¦í©Ò¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(6)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "E-mail¡G" & StrTemp5(7)
iPrint = iPrint + 300
Printer.CurrentX = 0 + Printer.TextWidth("¦í©Ò¡G")
Printer.CurrentY = iPrint
Printer.Print StrTemp5(8)
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "Fax¡G" & StrTemp5(9)
iPrint = iPrint + 300
If Len(StrTemp5(10)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦í©Ò¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(10)
   iPrint = iPrint + 300
End If
If Len(StrTemp5(11)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦í©Ò¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(11)
   iPrint = iPrint + 300
End If
'Add by Morgan 2011/5/26
If Len(StrTemp5(14)) <> 0 Then
   Printer.CurrentX = 0 + Printer.TextWidth("¦í©Ò¡G")
   Printer.CurrentY = iPrint
   Printer.Print StrTemp5(14)
   iPrint = iPrint + 300
End If
'end 2011/5/26
If Len(StrTemp5(10)) = 0 And Len(strTemp(11)) = 0 Then
   iPrint = iPrint + 300
End If
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "¶Q¾ã²zµfî"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
'If intPCaseKind = 2 Then
Printer.Print "¥XÄ@¤H"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "¥XÄ@µfî"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "ï¦W"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "ª¬ªp"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "’}°e¤é"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "¥XÄ@’ÇÇU³¬Âê"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "Case No."
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "’J¾ã²zµfî"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "¯S³\Ûµfî¡@¡@/¦s‘Ò´Á¶¡ÇU’Ó¤FÇU¤é/¦s‘Ò"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "¤u·~©Ò¦³ÛÇUºØ“E"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "¦W‘b"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "µù"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "¦¸ÇU´Á­­"
iPrint = iPrint + 300
Printer.Line (0, iPrint)-(19200, iPrint)
iPrint = iPrint + 100
End Sub


Private Sub PrintDatilEp()
For i = 0 To 6
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
For i = 7 To 13
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
Next i
'Printer.Line (0, iPrint + 150)-(19200, iPrint + 150)
iPrint = iPrint + 300
End Sub

Private Sub PrintDatilEt()
For i = 0 To 6
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
For i = 7 To 13
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
Next i
'Printer.Line (0, iPrint + 150)-(19200, iPrint + 150)
iPrint = iPrint + 300
End Sub

Private Sub PrintDatilJp()
For i = 0 To 6
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
For i = 7 To 13
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
Next i
'Printer.Line (0, iPrint + 150)-(19200, iPrint + 150)
iPrint = iPrint + 300
End Sub

Private Sub PrintDatilJt()
For i = 0 To 6
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
For i = 7 To 13
   Printer.CurrentX = PLeft(i)
   Printer.CurrentY = iPrint
   Printer.Print strTemp(i)
Next i
'Printer.Line (0, iPrint + 150)-(19200, iPrint + 150)
iPrint = iPrint + 300
End Sub


Private Sub GetPleftEp()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 3100
PLeft(2) = 5200
PLeft(3) = 7300
PLeft(4) = 8600
PLeft(5) = 16300
PLeft(6) = 17900
PLeft(7) = 0
PLeft(8) = 3100
PLeft(9) = 4800
PLeft(10) = 9000 + 200
PLeft(11) = 11100 + 200
PLeft(12) = 16400 + 200
PLeft(13) = 17900

End Sub

Private Sub GetPleftEt()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 3100
PLeft(2) = 5200
PLeft(3) = 7300
PLeft(4) = 8600
PLeft(5) = 16300 - 2000
PLeft(6) = 17900
PLeft(7) = 0
PLeft(8) = 3100
PLeft(9) = 4800
PLeft(10) = 9000 + 200 + 1500
PLeft(11) = 11100 + 200 + 1500
PLeft(12) = 16400 + 200
PLeft(13) = 17900

End Sub

Private Sub GetPleftJp()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 3100
PLeft(2) = 5200
PLeft(3) = 7300
PLeft(4) = 8600
PLeft(5) = 16300
PLeft(6) = 17900
PLeft(7) = 0
PLeft(8) = 3100
PLeft(9) = 4800
PLeft(10) = 9000 + 200
PLeft(11) = 11100 + 200
PLeft(12) = 16400 + 200
PLeft(13) = 17900

End Sub

Private Sub GetPleftJt()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 3100
PLeft(2) = 5200
PLeft(3) = 7300
PLeft(4) = 8600
PLeft(5) = 16300 - 2000
PLeft(6) = 17900
PLeft(7) = 0
PLeft(8) = 3100
PLeft(9) = 4800
PLeft(10) = 9000 + 200 + 1500
PLeft(11) = 11100 + 200 + 1500
PLeft(12) = 16400 + 200
PLeft(13) = 17900

End Sub

Private Sub Form_Load()
   
   MoveFormToCenter Me
   txt1(0) = GetSystemKindByNick
   
   SeekPrintL = Printer.Orientation
   PUB_SetPrinter Me.Name, Combo1, , , SeekPrint     'Modified by Morgan 2017/11/21 ³]©w¦Lªí¾÷§ï©I¥s¤½¥Î¨ç¼Æ,­ìµ{¦¡²¾°£
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
DoEvents
Set frm050316 = Nothing
End Sub

Private Sub opt_Click(Index As Integer)
    'Add By Cheng 2002/11/15
    Select Case Index
    Case 0
        Me.opt(0).Value = True
        Me.opt(1).Value = False
    Case 1
        Me.opt(0).Value = False
        Me.opt(1).Value = True
    End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " ¨S¦³ " & strTemp2(i) & " ªºÅv­­ ", , "Åv­­°ÝÃD")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
Case 2
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
Case 7, 8
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
    If Index = 7 Then
      If RunNick(txt1(Index + 1), txt1(Index)) Then
         txt1(Index + 1).SetFocus
         txt1_GotFocus (Index + 1)
         Exit Sub
      End If
    End If
Case 4
     If RunNick(txt1(3), txt1(4)) Then
         txt1(3).SetFocus
         txt1_GotFocus (3)
         Exit Sub
     End If
     If Mid(Trim(txt1(3)), 1, 6) <> Mid(Trim(txt1(4)), 1, 6) Then
         s = MsgBox("¥N²z¤H½s¸¹«e¤»½X¥²¶·¬Û¦P!!", , "USER ¿é¤J¿ù»~")
         If Len(txt1(3)) = 0 Then txt1(3).SetFocus: txt1_GotFocus (3)
         Exit Sub
     End If
Case 5
     Select Case Trim(txt1(5))
     Case "1", "2", ""
     Case Else
          s = MsgBox("¤º®e¿é¤J¿ù»~,¥u¯à 1 ©Î 2 !!", , "USER ¿é¤J¿ù»~")
          txt1(5).SetFocus
          txt1(5).SelStart = 0
          txt1(5).SelLength = Len(txt1(5))
          Exit Sub
     End Select
Case 6
     Select Case Trim(txt1(6))
     Case "1", "2", "3", ""
     Case Else
          s = MsgBox("³øªí»y¤å¿é¤J¿ù»~,¥u¯à 1 ¨ì 3 !!", , "USER ¿é¤J¿ù»~")
          txt1(6).SetFocus
          txt1(6).SelStart = 0
          txt1(6).SelLength = Len(txt1(6))
          Exit Sub
     End Select
     'If RunNick(txt1(5), txt1(6)) Then
     ' txt1(5).SetFocus
     ' txt1_GotFocus (5)
     ' Exit Sub
    'End If
Case Else
End Select
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
'Add By Cheng 2003/07/17
Select Case Index
Case 9
    If KeyAscii <> 8 And KeyAscii <> 78 Then
        KeyAscii = 0
    End If
End Select
End Sub

Private Sub RefreshColData(Index As Integer, strKind As String)
Select Case Index
Case 1 '¤¤¤å³øªí
   If strKind = "¥Ó½Ð¤H" Then
      m_ColCustName = "Nvl(CU04,Nvl(CU05,CU06)),Decode(CU04,Null,Decode(CU05,Null,CU88,''),''),Decode(CU04,Null,Decode(CU05,Null,CU89,''),''),Decode(CU04,Null,Decode(CU05,Null,CU90,''),'')"
      m_ColCustAdd = "Nvl(CU23,Nvl(CU65,Nvl(CU24,CU29))),Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66),''),Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67),''),Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68),''),Decode(CU23,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU28),CU69),'')"
   Else '¥N²z¤H
      m_ColAgName = "Nvl(FA04,Nvl(FA05,FA06)),Decode(FA04,Null,Decode(FA05,Null,'',FA63),''),Decode(FA04,Null,Decode(FA05,Null,'',FA64),''),Decode(FA04,Null,Decode(FA05,Null,'',FA65),'')"
      m_ColAgAdd = "Nvl(FA17,Nvl(FA32,Nvl(FA18,FA23))),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA19),FA33),''),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA20),FA34),''),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA21),FA35),''),Decode(FA17,Null,Decode(FA32,Null,Decode(FA18,Null,'',FA22),FA36),'')"
   End If
Case 2 '­^¤å³øªí
   If strKind = "¥Ó½Ð¤H" Then
      m_ColCustName = "Nvl(CU05,CU06),Decode(CU05,Null,'',CU88),Decode(CU05,Null,'',CU89),Decode(CU05,Null,'',CU90)"
      m_ColCustAdd = "Nvl(CU65,Nvl(CU24,CU29)),Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66),Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67),Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68),Decode(CU65,Null,Decode(CU24,Null,'',CU28),CU69)"
   Else '¥N²z¤H
      m_ColAgName = "Nvl(FA05,FA06),Decode(FA05,Null,'',FA63),Decode(FA05,Null,'',FA64),Decode(FA05,Null,'',FA65)"
      m_ColAgAdd = "Nvl(FA32,Nvl(FA18,FA23)),Decode(FA32,Null,Decode(FA18,Null,'',FA19),FA33),Decode(FA32,Null,Decode(FA18,Null,'',FA20),FA34),Decode(FA32,Null,Decode(FA18,Null,'',FA21),FA35),Decode(FA32,Null,Decode(FA18,Null,'',FA22),FA36)"
   End If
Case 3 '¤é¤å³øªí
   If strKind = "¥Ó½Ð¤H" Then
      m_ColCustName = "Nvl(CU06,CU05),Decode(CU06,Null,CU88,''),Decode(CU06,Null,CU89,''),Decode(CU06,Null,CU90,'')"
      m_ColCustAdd = "Nvl(CU29,Nvl(CU65,CU24)),Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU25),CU66),''),Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU26),CU67),''),Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU27),CU68),''),Decode(CU29,Null,Decode(CU65,Null,Decode(CU24,Null,'',CU28),CU69),'')"
   Else '¥N²z¤H
      m_ColAgName = "Nvl(FA06,FA05),Decode(FA06,Null,FA63,''),Decode(FA06,Null,FA64,''),Decode(FA06,Null,FA65,'')"
      m_ColAgAdd = "Nvl(FA23,Nvl(FA32,FA18)),Decode(FA23,Null,Decode(FA32,Null,FA19,FA33),''),Decode(FA23,Null,Decode(FA32,Null,FA20,FA34),''),Decode(FA23,Null,Decode(FA32,Null,FA21,FA35),''),Decode(FA23,Null,Decode(FA32,Null,FA22,FA36),'')"
   End If
End Select
End Sub
