VERSION 5.00
Begin VB.Form frm030619 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "°Ó¼Ð¤½³ø¶}©Ý¨ç¦C¦L : ¥»©Ò®×¥ó"
   ClientHeight    =   3510
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5655
   Begin VB.FileListBox File2 
      Height          =   450
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTBD01 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   264
      Left            =   1260
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1830
      Width           =   1092
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Text            =   "C:\temp\XmlTrans"
      Top             =   1290
      Width           =   3675
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4560
      TabIndex        =   3
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3765
      TabIndex        =   2
      Top             =   90
      Width           =   756
   End
   Begin VB.Label Label5 
      Caption         =   "¨Ã¥B¹q¸£¤£¥i¥H³]©w¿Ã¹õ«OÅ@¸Ë¸m"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   780
      TabIndex        =   9
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "¤½³ø¨÷´Á¡G"
      Height          =   210
      Left            =   330
      TabIndex        =   7
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "(               ¤H)"
      Height          =   210
      Left            =   2400
      TabIndex        =   6
      Top             =   1860
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ª`·N¡G·íµ{¦¡¥¿¦b°õ¦æ®É¡A½Ð¼È®É¤£­n¨Ï¥ÎWord"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   2730
      Width           =   5100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "°Ó¼Ð¹ÏÀÉ®×¸ô®|¡G"
      Height          =   180
      Left            =   330
      TabIndex        =   4
      Top             =   1350
      Width           =   1440
   End
End
Attribute VB_Name = "frm030619"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/11 Form2.0¤w­×§ï (µL»Ý­×§ï)
'Memo By Sindy 2012/12/5 ´¼Åv¤H­ûÄæ¤w­×§ï
Option Explicit

Dim m_AppAddr As String '°Ó¼Ðµù¥U¤H¦a§}
Dim m_AppName As String '°Ó¼Ðµù¥U¤H
Dim m_AppAddrZip As String '°Ó¼Ðµù¥U¤H¦a§}¶l»¼°Ï¸¹
Dim bolRetry As Boolean '¬O§_¤wµo¥Í¿ù»~¥B­«¸Õ

'¥[¤J¥Nªí¹Ï¥Î
Const msoBringInFrontOfText = 4
Const msoFalse = 0
Const msoLineSolid = 1
Const msoLineSingle = 1
Const msoTrue = -1
Const msoPictureAutomatic = 1

Dim m_intFileCnt As Integer
Dim m_WordFilePath As String
Dim m_TM01 As String, m_TM02 As String, m_TM03 As String, m_TM04 As String
Dim custarea As String   '·~°È°Ï
Dim custsales As String  '´¼Åv¤H­û
Dim strP22 As String 'Add By Sindy 2015/5/13
Dim m_TaieCustAddr As String 'Add By Sindy 2019/10/23


Private Function Process() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strTime As String
Dim fs As Object
Dim i As Integer
Dim strSubject As String  'Add By Sindy 2015/5/13
Dim bolConnect As Boolean
Dim intRow As Integer, intWordRow As Integer
Dim strSales As String
   
   On Error GoTo ErrHnd
   
   Process = False
   
   strTime = time()
   
   If Right(txtPath(0), 1) = "\" Then txtPath(0) = Left(txtPath(0), Len(txtPath(0)) - 1)
   File2.path = txtPath(0).Text & "\imagesdata"
   File2.Refresh
   If File2.ListCount = 0 Then
      MsgBox "§ä¤£¨ì°Ó¼Ð¹ÏÀÉ¡I"
      Exit Function
   End If
   
   m_WordFilePath = "c:\temp\WordFile"
   
   Set fs = CreateObject("Scripting.FileSystemObject")
   fs.DeleteFolder m_WordFilePath, True
NotFolder76:
   fs.CreateFolder m_WordFilePath
   
   '²£¥ÍWordÀÉ
   bolRetry = True
   Screen.MousePointer = vbHourglass
   cnnConnection.BeginTrans: bolConnect = True
   For i = 1 To 2
      If i = 1 Then
         '­n±H¶}©Ý¨çªº«È¤á...µL?
         'Modify By Sindy 2013/3/1 + where tbnp08='T'
         'Modify By Sindy 2018/12/13 + and tbd16='1' : ¤½³ø¶}©Ý
         strSql = "select tbor03,count(tbor01) from tmbulletinowner,tmbulletindata,Trademark " & _
                  "Where tbor02=1 and tbd16='1' and tbd04=TM12 and tm44 is null " & _
                  "and tbor01=tbd02 and tbor06=tbd03 and tbd15='A' and (tbd14<>'N' or tbd14 is null) " & _
                  "and ltrim(rtrim(tbor03)) not in(select ltrim(rtrim(tbnp01)) from tmbulletinnp where tbnp08='T') " & _
                  "and instr(tbor03,'?')=0 " & _
                  "and instr(tbor05,'?')=0 " & _
                  "group by tbor03 order by tbor03 "
      ElseIf i = 2 Then
         '­n±H¶}©Ý¨çªº«È¤á...¦³?
         rsTmp.Close
         'Modify By Sindy 2013/3/1 + where tbnp08='T'
         'Modify By Sindy 2018/12/13 + and tbd16='1' : ¤½³ø¶}©Ý
         strSql = "select tbor03,count(tbor01) from tmbulletinowner,tmbulletindata,Trademark " & _
                  "Where tbor02=1 and tbd16='1' and tbd04=TM12 and tm44 is null " & _
                  "and tbor01=tbd02 and tbor06=tbd03 and tbd15='A' and (tbd14<>'N' or tbd14 is null) " & _
                  "and ltrim(rtrim(tbor03)) not in(select ltrim(rtrim(tbnp01)) from tmbulletinnp where tbnp08='T') " & _
                  "and (instr(tbor03,'?')>0 or instr(tbor05,'?')>0) " & _
                  "and tbor10<>'Y' " & _
                  "group by tbor03 order by tbor03 "
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         m_intFileCnt = 0
         intWordRow = 0
         For intRow = 1 To rsTmp.RecordCount '¤»­ûÀô¦³­­¤½¥q   " & rsTmp.Fields("tbor03") & "
            'Modify By Sindy 2018/12/13 + and tbd16='1' : ¤½³ø¶}©Ý
            strSql = "select * from tmbulletindata,tmbulletinowner,Trademark " & _
                     "Where tbd02 = tbor01 and tbd03 = tbor06 and tbd16='1' and tbor02=1 and tbd15='A' " & _
                     "and tbd04=TM12 and tm44 is null " & _
                     "and tbor03='" & rsTmp.Fields("tbor03") & "' " & _
                     "order by tbd03 asc,tbd02 asc "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               m_AppName = RsTemp.Fields("tbor03") '°Ó¼Ðµù¥U¤H
               m_AppAddrZip = "" '¶l»¼°Ï¸¹ 'Add By Sindy 2012/1/16
               m_AppAddr = RsTemp.Fields("tbor05") '°Ó¼Ðµù¥U¤H¦a§}
               m_TaieCustAddr = "" 'Add By Sindy 2019/10/23
               'Add By Sindy 2012/1/16 ¥»©Ò®×¤á§ìÁpµ¸¦a§}
               'Modify By Sindy 2018/12/13 + and tbd16='1' : ¤½³ø¶}©Ý
               strSql = "select TM01,TM02,TM03,TM04,TM11,TM23,tbor03,CU30,CU31 " & _
                          "From tmbulletindata,tmbulletinowner,Trademark,customer " & _
                         "Where tbd02=tbor01 and tbd03=tbor06 and tbd16='1' and tbd15='A' " & _
                           "and tbor02=1 " & _
                           "and tbor03='" & rsTmp.Fields("tbor03") & "' " & _
                           "and tbd04=TM12 and TM10='000' and TM28='1' " & _
                           "and cu01=substr(TM23,1,8) and cu02=substr(TM23,9,1) " & _
                        "order by TM11 desc "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  'Add By Sindy 2020/11/16 ¯S®í«È¤á¦h®×®É,­n¥t¥~¥X¹q¤lÀÉ
                  If RsTemp.RecordCount > 1 And _
                     InStr(Pub_GetSpecMan("°Ó¼Ð¤½³ø¶}©Ý¨ç¯S®í«È¤á"), RsTemp.Fields("TM23")) > 0 Then
                     strSql = "update tmbulletinowner set TBor10='Y' where tbor03='" & RsTemp.Fields("tbor03") & "'"
                     cnnConnection.Execute strSql
                     GoTo WordEditEnd
                  End If
                  '2020/11/16 END
                  
'                 m_AppAddrZip = PUB_ChangeZIPToSir("" & RsTemp.Fields("CU30")) '¶l»¼°Ï¸¹
'                 m_AppAddr = "" & RsTemp.Fields("CU31") '°Ó¼Ðµù¥U¤H¦a§}
                  'Modify By Sindy 2019/10/23 ¦]¥»©Ò«È¤áÁÙ­n§ì«È¤á±µ¬¢¤Hªº¸ê°T
'                     »O¥_¥««H¸q°ÏªQ´¼¸ô1¸¹24¼Ó
'                     µØ·s¬ì§ÞªÑ¥÷¦³­­¤½¥q
'                     ±i¹Å¬À¤p©j ¶v±Ò(T-219231)
                  m_MySt(1) = RsTemp.Fields("TM01")
                  m_MySt(2) = RsTemp.Fields("TM02")
                  m_MySt(3) = RsTemp.Fields("TM03")
                  m_MySt(4) = RsTemp.Fields("TM04")
                  m_TaieCustAddr = ExceptFieldData2("¤¤¥Ó¶}µ¡¶l¸¹") & vbCrLf
                  m_TaieCustAddr = m_TaieCustAddr & ExceptFieldData2("¤¤¥Ó¶}µ¡¦a§}")
                  '2019/10/23 END
                  
               'Modify By Sindy 2014/3/19 ¦]µo¥Í m_AppAddr = "557«n§ë¿¤¦Ë¤sÂí©µ¤s¨½·ç¤s«Ñ138¸¹"
               Else
                  'ÀË¬d¦a§}«eÀY¬O§_¬°¶l»¼°Ï¸¹
                  If IsNumeric(Left(m_AppAddr, 3)) = True Then
                     If IsNumeric(Left(m_AppAddr, 5)) = True Then
                        m_AppAddrZip = Left(m_AppAddr, 5)
                        m_AppAddr = Mid(m_AppAddr, 6)
                     Else
                        m_AppAddrZip = Left(m_AppAddr, 3)
                        m_AppAddr = Mid(m_AppAddr, 4)
                     End If
                  End If
               End If
               '2014/3/19 END
               '2012/1/16 End
               
               'Add By Sindy 2013/6/3
               If m_AppAddrZip = "" Then
                  m_AppAddrZip = PUB_ChangeZIPToSir(Left(PUB_AddrChangeZIPCode(m_AppAddr), 3))
               End If
               '2013/6/3 End
               
               ' ¦C¦L©w½Z
               If WordEdit() = False Then
                  GoTo ErrHnd
               'Add By Sindy 2020/11/16
               Else
                  intWordRow = intWordRow + 1
               '2020/11/16 END
               End If
            End If
            
            'If (intRow Mod 100) = 0 Or intRow = rsTmp.RecordCount Then
            If (intWordRow Mod 100) = 0 Or intRow = rsTmp.RecordCount Then
               g_WordAp.Documents.Save
               g_WordAp.Documents.Close
               bolRetry = True
            End If
'            If (intRow Mod 40) = 0 Then
'               Exit For
'            End If
            
WordEditEnd:
            rsTmp.MoveNext
         Next intRow
      End If
   Next i
   rsTmp.Close
   If bolRetry = False Then
      g_WordAp.Documents.Save
      g_WordAp.Documents.Close
'      g_WordAp.Visible = True
'      g_WordAp.WindowState = wdWindowStateMaximize
   End If
   cnnConnection.CommitTrans: bolConnect = False
   
   'Add By Sindy 2020/11/16
   '²£¥Í¯S®í«È¤á©w½Z-¨Ì´¼Åv¤H­û¤À¶}¹q¤lÀÉ
   bolRetry = True
   Screen.MousePointer = vbHourglass
   cnnConnection.BeginTrans: bolConnect = True
   strSql = "select cu13,st02,tbor03,count(tbor01) from tmbulletinowner,tmbulletindata,Trademark,customer,staff " & _
            "Where tbor02=1 and tbd16='1' and tbd04=TM12 and tm44 is null " & _
            "and tbor01=tbd02 and tbor06=tbd03 and tbd15='A' and (tbd14<>'N' or tbd14 is null) " & _
            "and tbor10='Y' " & _
            "and cu01=substr(TM23,1,8) and cu02=substr(TM23,9,1) " & _
            "and cu13=st01(+) " & _
            "group by cu13,st02,tbor03 order by cu13 asc,tbor03 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      intWordRow = 0
      For intRow = 1 To rsTmp.RecordCount
         'Modify By Sindy 2018/12/13 + and tbd16='1' : ¤½³ø¶}©Ý
         strSql = "select * from tmbulletindata,tmbulletinowner,Trademark " & _
                  "Where tbd02 = tbor01 and tbd03 = tbor06 and tbd16='1' and tbor02=1 and tbd15='A' " & _
                  "and tbd04=TM12 and tm44 is null " & _
                  "and tbor03='" & rsTmp.Fields("tbor03") & "' " & _
                  "order by tbd03 asc,tbd02 asc "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            m_AppName = RsTemp.Fields("tbor03") '°Ó¼Ðµù¥U¤H
            m_AppAddrZip = "" '¶l»¼°Ï¸¹ 'Add By Sindy 2012/1/16
            m_AppAddr = RsTemp.Fields("tbor05") '°Ó¼Ðµù¥U¤H¦a§}
            m_TaieCustAddr = "" 'Add By Sindy 2019/10/23
            'Add By Sindy 2012/1/16 ¥»©Ò®×¤á§ìÁpµ¸¦a§}
            'Modify By Sindy 2018/12/13 + and tbd16='1' : ¤½³ø¶}©Ý
            strSql = "select TM01,TM02,TM03,TM04,TM11,TM23,tbor03,CU30,CU31 " & _
                       "From tmbulletindata,tmbulletinowner,Trademark,customer " & _
                      "Where tbd02=tbor01 and tbd03=tbor06 and tbd16='1' and tbd15='A' " & _
                        "and tbor02=1 " & _
                        "and tbor03='" & rsTmp.Fields("tbor03") & "' " & _
                        "and tbd04=TM12 and TM10='000' and TM28='1' " & _
                        "and cu01=substr(TM23,1,8) and cu02=substr(TM23,9,1) " & _
                     "order by TM11 desc "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
'              m_AppAddrZip = PUB_ChangeZIPToSir("" & RsTemp.Fields("CU30")) '¶l»¼°Ï¸¹
'              m_AppAddr = "" & RsTemp.Fields("CU31") '°Ó¼Ðµù¥U¤H¦a§}
               'Modify By Sindy 2019/10/23 ¦]¥»©Ò«È¤áÁÙ­n§ì«È¤á±µ¬¢¤Hªº¸ê°T
'                     »O¥_¥««H¸q°ÏªQ´¼¸ô1¸¹24¼Ó
'                     µØ·s¬ì§ÞªÑ¥÷¦³­­¤½¥q
'                     ±i¹Å¬À¤p©j ¶v±Ò(T-219231)
               m_MySt(1) = RsTemp.Fields("TM01")
               m_MySt(2) = RsTemp.Fields("TM02")
               m_MySt(3) = RsTemp.Fields("TM03")
               m_MySt(4) = RsTemp.Fields("TM04")
               m_TaieCustAddr = ExceptFieldData2("¤¤¥Ó¶}µ¡¶l¸¹") & vbCrLf
               m_TaieCustAddr = m_TaieCustAddr & ExceptFieldData2("¤¤¥Ó¶}µ¡¦a§}")
               '2019/10/23 END

            'Modify By Sindy 2014/3/19 ¦]µo¥Í m_AppAddr = "557«n§ë¿¤¦Ë¤sÂí©µ¤s¨½·ç¤s«Ñ138¸¹"
            Else
               'ÀË¬d¦a§}«eÀY¬O§_¬°¶l»¼°Ï¸¹
               If IsNumeric(Left(m_AppAddr, 3)) = True Then
                  If IsNumeric(Left(m_AppAddr, 5)) = True Then
                     m_AppAddrZip = Left(m_AppAddr, 5)
                     m_AppAddr = Mid(m_AppAddr, 6)
                  Else
                     m_AppAddrZip = Left(m_AppAddr, 3)
                     m_AppAddr = Mid(m_AppAddr, 4)
                  End If
               End If
            End If
            '2014/3/19 END
            '2012/1/16 End
            
            'Add By Sindy 2013/6/3
            If m_AppAddrZip = "" Then
               m_AppAddrZip = PUB_ChangeZIPToSir(Left(PUB_AddrChangeZIPCode(m_AppAddr), 3))
            End If
            '2013/6/3 End
            
            If strSales <> "" And strSales <> "" & rsTmp.Fields("st02") Then
               g_WordAp.Documents.Save
               g_WordAp.Documents.Close
               bolRetry = True
            End If
            
            ' ¦C¦L©w½Z
            If WordEdit("" & rsTmp.Fields("st02")) = False Then
               GoTo ErrHnd
            'Add By Sindy 2020/11/16
            Else
               intWordRow = intWordRow + 1
            '2020/11/16 END
            End If
         End If
         
         'If (intRow Mod 100) = 0 Or intRow = rsTmp.RecordCount Then
         If (intWordRow Mod 100) = 0 Or intRow = rsTmp.RecordCount Then
            g_WordAp.Documents.Save
            g_WordAp.Documents.Close
            bolRetry = True
         End If
'            If (intRow Mod 40) = 0 Then
'               Exit For
'            End If
         
         strSales = "" & rsTmp.Fields("st02")
         rsTmp.MoveNext
      Next intRow
   End If
   rsTmp.Close
   If bolRetry = False Then
      g_WordAp.Documents.Save
      g_WordAp.Documents.Close
'      g_WordAp.Visible = True
'      g_WordAp.WindowState = wdWindowStateMaximize
   End If
   cnnConnection.CommitTrans: bolConnect = False
   
   Screen.MousePointer = vbDefault
   
   'Add By Sindy 2015/5/13 ³qª¾µ{§Ç¶}©Ý¨ç¹q¤lÀÉ¤w²£¥Í§¹²¦
   If strP22 <> "" Then
      strSubject = "°Ó¼Ð¤½³ø¶}©Ý¨ç¹q¤lÀÉ¤w²£¥Í§¹²¦¡I"
      PUB_SendMail strUserNum, strP22, "", strSubject, strSubject, , , , , , , , , , , False
   End If
   '2015/5/13 END
   
   Process = True
   MsgBox "§@·~§¹¦¨¡I½Ð¦Ü" & m_WordFilePath & "\¸ê®Æ§¨¤¤¦C¦L¶}©Ý¨ç¡C¡]ªá¶O®É¶¡¡G" & strTime & "  " & time() & "¡^"
   
   Set rsTmp = Nothing
   Set g_WordAp = Nothing
   Exit Function
   
ErrHnd:
   If Err.Number = 76 Then
      On Error GoTo 0 'Add By Sindy 2020/4/16
      GoTo NotFolder76
   ElseIf Err.Number = 70 Then
      MsgBox Err.Description, vbCritical
   ElseIf Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
      '³qª¾µ{§Ç¶}©Ý¨ç¹q¤lÀÉ²£¥Í¦³»~
      If strP22 <> "" Then
         strSubject = Me.Caption & "¡A¹q¤lÀÉ²£¥Í¦³»~¡I"
         PUB_SendMail strUserNum, strUserNum, "", strSubject, strSubject, , , , , , , , , , , False
      End If
   End If
   If bolConnect = True Then
      cnnConnection.RollbackTrans
   End If
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   Set g_WordAp = Nothing
   'Resume
End Function

Private Function WordEdit(Optional strSales As String = "") As Boolean
   'Add by Morgan 2011/10/26 +«HÀY
   Dim stFileName As String
   Dim iPicNo As Integer
   Dim iPicNo2 As Integer
   Dim oShape
   
   'Added by Morgan 2020/3/30
   If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é Then
      PUB_GetLetterPicID "1", "T", iPicNo, iPicNo2
   Else
   'end 2020/3/30
      iPicNo = 12
      iPicNo2 = 11
   End If 'Added by Morgan 2020/3/30
   
   'end 2011/10/26
   Dim rsTmp As New ADODB.Recordset
   Dim i As Integer, j As Integer
   Dim strTBD01 As String, strTBD01_2 As String
   Dim strTBD02 As String
   Dim strTBD03 As String, strTBD03_2 As String
   Dim strTBD04 As String
   Dim strTBD05 As String
   Dim strTBD06 As String
   Dim strTBD07 As String
   Dim strTBD08 As String
   Dim strTBD09 As String
   Dim strTBD10 As String
   Dim strTBD11 As String
   Dim strTBD12 As String
   Dim strTBD13 As String
   Dim bolIsTit As Boolean
   Dim strTemp As String
   Dim intSpecRow As Integer
   Dim strTitle As String
   
On Error GoTo ERRORSECTION1
   
   WordEdit = True
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize 'wdWindowStateMinimize  wdWindowStateMaximize
   With g_WordAp
   
      If bolRetry = True Then
         'Add By Sindy 2020/11/16
         If strSales <> "" Then
            g_WordAp.Documents.add.SaveAs m_WordFilePath & "\°Ó¼Ð¤½³ø" & Left(txtTBD01, 2) & "¨÷" & Right(txtTBD01, 2) & "´Á" & "¶}©Ý¨ç-" & strSales & ".doc"
         Else
         '2020/11/16 END
            m_intFileCnt = m_intFileCnt + 1
            g_WordAp.Documents.add.SaveAs m_WordFilePath & "\°Ó¼Ð¤½³ø" & Left(txtTBD01, 2) & "¨÷" & Right(txtTBD01, 2) & "´Á" & "¶}©Ý¨ç" & Format(m_intFileCnt, "00") & ".doc"
         End If
'         'Add by Morgan 2011/10/26 +«HÀY
'         If PUB_ReadDB2File(stFileName, iPicNo) = True Then
'            '¤Á´«¬°¾ã­¶¼Ò¦¡
'            If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
'               .ActiveWindow.ActivePane.View.Type = wdPageView
'            Else
'               .ActiveWindow.View.Type = wdPageView
'            End If
'            .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader '­¶­º
'            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
'            oShape.ZOrder 4
'            oShape.LockAnchor = True
'            oShape.LockAspectRatio = -1
'            oShape.Width = .CentimetersToPoints(21)
'            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'            oShape.Left = .CentimetersToPoints(0)
'            oShape.Top = .CentimetersToPoints(0.5)
'            If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
'               .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter '­¶§À
'               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
'               oShape.ZOrder 4
'               oShape.LockAnchor = True
'               oShape.LockAspectRatio = -1
'               oShape.Width = .CentimetersToPoints(21)
'               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'               oShape.Left = .CentimetersToPoints(0)
'               oShape.Top = .CentimetersToPoints(27)
'            End If
'            .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
'            .Selection.EndKey Unit:=wdStory
'         End If
'         'end 2011/10/26
      End If
   
      If bolRetry = False Then
         '¸õ­¶
         '.Selection.EndKey Unit:=wdStory
         .Selection.InsertBreak Type:=wdPageBreak
         .Selection.MoveUp Unit:=wdLine, Count:=1
         .Selection.TypeBackspace
         .Selection.EndKey Unit:=wdStory
      End If
      
      .Selection.Font.Name = "¼Ð·¢Åé"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 14
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      'Modify by Morgan 2008/7/3
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2.5)
      'end 2008/7/3
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      
            
      'Add By Sindy Modify 2011/11/29
      'Add by Morgan 2011/7/12 ¦]¬°²Ä 2 ­¶¥H«á¤£­n¦³«HÀY¬G§ï¦^©ñ¦b¥»¤å
      If PUB_ReadDB2File(stFileName, iPicNo) = True Then
         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
         oShape.ZOrder 4
         oShape.LockAnchor = True
         oShape.LockAspectRatio = -1
         oShape.Width = .CentimetersToPoints(21)
         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
         oShape.Left = .CentimetersToPoints(0)
         oShape.Top = .CentimetersToPoints(0.5)
         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
            oShape.ZOrder 4
            oShape.LockAnchor = True
            oShape.LockAspectRatio = -1
            oShape.Width = .CentimetersToPoints(21)
            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
            oShape.Left = .CentimetersToPoints(0)
            'oShape.Top = .CentimetersToPoints(27.3)
            oShape.Top = .CentimetersToPoints(27)
         End If
         .Selection.EndKey Unit:=wdStory
      End If
      
      
      'Add by Morgan 2008/7/17 °t¦X·sªº¶}µ¡©w½Z§ï©T©w¦æ°ª
      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
      .Selection.ParagraphFormat.LineSpacing = 15
      'end 2008/7/17
      
      
      .Selection.TypeParagraph 'Add by Morgan 2008/6/11 CFT «HÀY¤ñ¸û°ª
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      
'      '³]©w¦r«¬ª©­±(°Ñ·Ó©w½Z)
'      '.Selection.Font.Name = "Times New Roman"
'      .Selection.Font.Name = "¼Ð·¢Åé"
'      .Selection.PageSetup.Orientation = wdOrientPortrait
'      .Selection.Orientation = wdTextOrientationHorizontal
'      .Selection.Font.Size = 14
'      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(3.175)
'      .Selection.PageSetup.RightMargin = .CentimetersToPoints(3.175)
'      .Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
'      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
'      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      '¾a¥ª
      '.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      '¸m¥k
      '.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
      '¸m¤¤
      '.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
      '¤£­n¤À´²¹ï»ô
      '.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
      
      '¾a¥ª
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
      'Modify By Sindy 2019/10/23
      If m_TaieCustAddr <> "" Then
         .Selection.TypeText m_TaieCustAddr
      Else
      '2019/10/23 END
         If m_AppAddrZip = "" Then
            .Selection.TypeParagraph
         End If
         .Selection.TypeText getAddrData
      End If
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeText "­P¡G" & m_AppName
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      .Selection.TypeText "·q±ÒªÌ¡G"
      .Selection.TypeParagraph
      .Selection.TypeParagraph
      
      .Selection.TypeText "¡@¡@®¥ÁH¡I¥xºÝ¡þ¶Q¤½¥q¤§°Ó¼Ð¤wÀò­ãµù¥U¡I±Nµù¥U¤½§i¤T­Ó¤ë¡C¨Ìªk¥xºÝ¡þ¶Q¤½¥q¦Ûµù¥U¤½§i¤§¤é°_¨ú±o°Ó¼ÐÅv¡A±M¥Î´Á¶¡10¦~¡C"
      .Selection.TypeParagraph
      
      .Selection.Font.Size = 10
      
      m_TM01 = "": m_TM02 = "": m_TM03 = "": m_TM04 = "": custarea = "": custsales = ""
      'Modify By Sindy 2018/12/13 + and tbd16='1' : ¤½³ø¶}©Ý
      strSql = "select tmbulletindata.*,tmbulletinowner.*,tm01,tm02,tm03,tm04 from tmbulletindata,tmbulletinowner,trademark " & _
               "Where tbd02 = tbor01 and tbd03 = tbor06 and tbd16='1' and tbor02=1 and tbd15='A' " & _
               "and tbor03='" & m_AppName & "' " & _
               "and tbd04=tm12(+) and tm44 is null " & _
               "order by tbd02 "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         For i = 1 To rsTmp.RecordCount
            'µ§¼Æ¬°°¸¼Æ®É,±µ¤U¤@­¶
            If (i Mod 2) = 0 Then
               '¸õ­¶
               .Selection.EndKey Unit:=wdStory
               .Selection.InsertBreak Type:=wdPageBreak
            End If
            
            If m_TM01 = "" Then
               m_TM01 = Trim("" & rsTmp.Fields("TM01"))
               m_TM02 = Trim("" & rsTmp.Fields("TM02"))
               m_TM03 = Trim("" & rsTmp.Fields("TM03"))
               m_TM04 = Trim("" & rsTmp.Fields("TM04"))
            End If
            intSpecRow = 0: strTitle = ""
            strTBD01 = "" & rsTmp.Fields("TBD01")
            strTBD01_2 = Left(txtTBD01, 2) & "¨÷" & Format(Right(txtTBD01, 2), "00") & "´Á¡@" & ChangeWStringToTDateString(ChgTMBM07ToDate(strTBD01))
            strTBD02 = "" & rsTmp.Fields("TBD02")
            strTBD03 = "" & rsTmp.Fields("TBD03")
            If strTBD03 = "7" Or strTBD03 = "8" Then
               strTitle = "¼Ð³¹"
            Else
               strTitle = "°Ó¼Ð"
            End If
            strTBD03_2 = GetTradeMarkName(strTBD03, 0)
            strTBD04 = "" & rsTmp.Fields("TBD04")
            strTBD05 = "" & rsTmp.Fields("TBD05")
            strTBD06 = "" & rsTmp.Fields("TBD06")
            strTBD07 = "" & rsTmp.Fields("TBD07")
            strTBD08 = "" & rsTmp.Fields("TBD08")
            strTBD09 = "" & rsTmp.Fields("TBD09")
            strTBD10 = "" & rsTmp.Fields("TBD10")
            strTBD11 = "" & rsTmp.Fields("TBD11")
            strTBD12 = "" & rsTmp.Fields("TBD12")
            strTBD13 = "" & rsTmp.Fields("TBD13")
            .Selection.TypeText "----------------------------------------------------------------------------------------------"
            .Selection.TypeParagraph
            .Selection.TypeText "µù¥U" & strTBD03_2 & "²Ä" & strTBD02 & "¸¹¡@¥Ó½Ð®×¸¹¡G" & strTBD04 & "¡@" & strTBD01_2 & "¡@°Ó¼Ð¹Ï¼Ë¡G" & strTBD05
            .Selection.TypeParagraph
            .Selection.TypeText "¥Ó½Ð¤é´Á¡G" & strTBD06 '& "|#¥k¥Nªí¹Ï#|"
            AddInPicToWordR g_WordAp, strTBD12 '´¡¤J¹ÏÀÉ
            .Selection.TypeParagraph
            If strTBD13 <> "" Then
               .Selection.TypeText "Àu¥ýÅv¤é¡G" & strTBD13
               .Selection.TypeParagraph
               intSpecRow = intSpecRow + 1
            End If
            .Selection.TypeText strTitle & "¦WºÙ¡G" & strTBD07
            .Selection.TypeParagraph
            '°Ó¼ÐÅv¤H¸ê®Æ
            strSql = "select * from tmbulletinowner Where tbor01='" & strTBD02 & "' and tbor06='" & strTBD03 & "' order by tbor02 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               For j = 1 To RsTemp.RecordCount
                  bolIsTit = False '©|µL¼ÐÃD
                  If j > 1 Then intSpecRow = intSpecRow + 1
                  If "" & RsTemp.Fields("tbor03") <> "" Then
                     If bolIsTit = False Then
                        .Selection.TypeText strTitle & "Åv¤H¡G" & "" & RsTemp.Fields("tbor03")
                        bolIsTit = True '¦³¼ÐÃD¤F
                     Else
                        .Selection.TypeText "¡@¡@¡@¡@¡@" & "" & RsTemp.Fields("tbor03")
                     End If
                     .Selection.TypeParagraph
                  End If
                  If "" & RsTemp.Fields("tbor04") <> "" Then
                     If bolIsTit = False Then
                        .Selection.TypeText strTitle & "Åv¤H¡G" & "" & RsTemp.Fields("tbor04")
                        bolIsTit = True '¦³¼ÐÃD¤F
                        intSpecRow = intSpecRow + 1
                     Else
                        .Selection.TypeText "¡@¡@¡@¡@¡@" & "" & RsTemp.Fields("tbor04")
                     End If
                     .Selection.TypeParagraph
                  End If
                  If "" & RsTemp.Fields("tbor05") <> "" Then
                     If bolIsTit = False Then
                        .Selection.TypeText strTitle & "Åv¤H¡G" & "" & RsTemp.Fields("tbor05")
                        bolIsTit = True '¦³¼ÐÃD¤F
                     Else
                        .Selection.TypeText "¡@¡@¡@¡@¡@" & "" & RsTemp.Fields("tbor05")
                     End If
                     .Selection.TypeParagraph
                  End If
                  RsTemp.MoveNext
               Next j
            End If
            '°Ó¼ÐÅv¤H¸ê®Æ End
            If strTBD08 <> "" Then
               .Selection.TypeText "¥N²z¤H¡G" & strTBD08
               .Selection.TypeParagraph
            End If
            .Selection.TypeText "Åv§Q´Á¶¡¡G" & strTBD09
            .Selection.TypeParagraph
            .Selection.TypeText "¼f¬d¤H­û¡G" & strTBD10
            .Selection.TypeParagraph
            If intSpecRow = 0 Then
               .Selection.TypeParagraph
               .Selection.TypeParagraph
            ElseIf intSpecRow = 1 Then
               .Selection.TypeParagraph
            End If
            If strTBD11 <> "" Then
               .Selection.TypeText strTBD11
               .Selection.TypeParagraph
               .Selection.TypeParagraph
            End If
            '°Ó«~¸ê®Æ
            strSql = "select * from tmbulletingoods Where tbg01='" & strTBD02 & "' and tbg07='" & strTBD03 & "' order by tbg02 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               For j = 1 To RsTemp.RecordCount
                  If "" & RsTemp.Fields("tbg03") <> "" Then
                     .Selection.TypeText "" & RsTemp.Fields("tbg03")
                     .Selection.TypeParagraph
                  End If
                  strTemp = Trim("" & RsTemp.Fields("tbg04")) & _
                            Trim("" & RsTemp.Fields("tbg05")) & _
                            Trim("" & RsTemp.Fields("tbg06")) & _
                            Trim("" & RsTemp.Fields("tbg08")) & _
                            Trim("" & RsTemp.Fields("tbg09")) & _
                            Trim("" & RsTemp.Fields("tbg10"))
                  If strTemp <> "" Then
                     If strTBD03 = "7" Then 'ÃÒ©ú¼Ð³¹
                        .Selection.TypeText "ÃÒ©ú¤º®e¡G" & strTemp
                     ElseIf strTBD03 = "8" Then '¹ÎÅé¼Ð³¹
                        .Selection.TypeText "ªí¹ü¤º®e¡G" & strTemp
                     Else
                        .Selection.TypeText "°Ó«~©ÎªA°È¦WºÙ¡G" & strTemp
                     End If
                     .Selection.TypeParagraph
                  End If
                  RsTemp.MoveNext
               Next j
            End If
            '°Ó«~¸ê®Æ End
            
            '¥[µù:¤w²£¥Í©w½Z
            strSql = "update TMBulletinData set TBD14='Y' where TBD02='" & strTBD02 & "' and TBD03='" & strTBD03 & "'"
            cnnConnection.Execute strSql
            
            rsTmp.MoveNext
         Next i
         .Selection.TypeText "----------------------------------------------------------------------------------------------"
         .Selection.TypeParagraph
      End If
      
      .Selection.Font.Size = 14
      
      .Selection.TypeText "°Ó¼Ð©óµù¥U«áÀ³¨Ï¥Î¡A§_«h³sÄò¤T¦~µL¥¿·í¨Æ¥Ñ¥¼¨Ï¥Î¡A°Ó¼ÐÅv±N³Q¼o¤î¡F¬°«K©ó©Ý®i¥~¾P¥«³õ¡A©y©ó°ê¤º°Ó¼Ðµù¥U«á¡A¥Ó½Ð¤j³°¤Î¨ä¥L¦U°ê°Ó¼Ð¤§µù¥U¡C­Õ¨Ï¡@¥xºÝ¡þ¶Q¤½¥q¹ï°Ó¼Ð¤§¨Ï¥Î¡A©|¦³½èºÃ¡A·q¬è¤£§[¨Ó¹q©Î»YÁ{¬¢¸ß¡A¥»©Ò¤G¦Ê¦h¦ì±M·~¤H¤hºÜ¸Û¬°±z´£¨Ñ³Ì§¹µ½ªºªA°È¡I"
      .Selection.TypeParagraph
      
      .Selection.TypeParagraph
'      .Selection.TypeParagraph
      .Selection.TypeText "¡@¡@¡@¡@­B¦¹¡@¡@¶¶¹|"
      .Selection.TypeParagraph
      .Selection.TypeText "°Ó¡@¸R"
      .Selection.TypeParagraph
      'Modified by Morgan 2020/3/30 ¨Æ°È©Ò¦WºÙ§ï¥Î¨ç¼Æ§ì
      '.Selection.TypeText "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¥x¤@°ê»Ú±M§Q°Ó¼Ð¨Æ°È©Ò¡@·q¤W"
      .Selection.TypeText "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@" & PUB_GetCompName2("1") & "¡@·q¤W"
      'end 2020/3/30
      If m_TM01 <> "" Then
         .Selection.TypeParagraph
         Call GetSales
         If custarea = "" Then
            .Selection.TypeText "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@" & custsales
         Else
            .Selection.TypeText "¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@¡@" & custarea & "¡@" & custsales
         End If
         '.Selection.TypeParagraph
      End If
'      .Selection.TypeParagraph
'      .Selection.WholeStory
'      ChgWordFormat g_WordAp, .Selection.Text
   End With
   
'   PhaseIndent    '½Õ¾ã­º¦æ¥Y±Æ
'   g_WordAp.Visible = True
'   g_WordAp.WindowState = wdWindowStateMaximize
   bolRetry = False
   Exit Function
   
ERRORSECTION1:
   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91, 462:
            Set g_WordAp = New Word.Application
            g_WordAp.Documents.add
            If bolRetry = False Then
               bolRetry = True
               Resume
            End If
         Case Else:
            MsgBox "¿ù»~ : " & Err.Description, vbCritical
            WordEdit = False
      End Select
   End If
End Function

Private Sub AddInPicToWordR(ByRef oWord As Word.Application, strFileName As String)

   Dim bytes() As Byte
   Dim file_num As Integer
   Dim rsPic As New ADODB.Recordset
   Dim IsWmf As Boolean
   'Add by Morgan 2007/11/6
   Dim stSQL As String
   Dim intR As Integer
   Dim stFileName As String
   Dim oShape 'Added by Lydia 2016/09/29

On Error GoTo ErrHnd

   With oWord
'      .Selection.MoveDown
      
'      '¶}±Ò¬d¸ß
'      'Modify by Morgan 2007/12/14 ¬°¥[³t¥u©¹«e²¾3­¶
'      '.Selection.HomeKey Unit:=wdStory
'      .Selection.GoTo what:=wdGoToPage, which:=wdGoToPrevious, Count:=3
'      'end 2007/12/14
'      .Selection.Find.ClearFormatting
'      .Selection.Find.Text = "|#¥k¥Nªí¹Ï#|"
'      .Selection.Find.Replacement.Text = ""
'      .Selection.Find.Forward = True
'      .Selection.Find.Wrap = wdFindContinue
'      .Selection.Find.Format = False
'      .Selection.Find.MatchCase = False
'      .Selection.Find.MatchWholeWord = False
'      .Selection.Find.MatchWildcards = False
'      .Selection.Find.MatchSoundsLike = False
'      .Selection.Find.MatchAllWordForms = False
'      .Selection.Find.MatchByte = True
'      .Selection.Find.Execute
'      '.Selection.Select
'      .Selection.Delete 'Add by Morgan 2007/11/8

      If InStr(strFileName, "imagesdata") = 0 Then
         strFileName = "imagesdata/" & strFileName
      End If

      '´¡¤J¹Ï¤ùÀÉ®×
      .ChangeFileOpenDirectory txtPath(0) & "\imagesdata\"
      'Add By Sindy 2012/10/17 ÀË¬dÀÉ®×¬O§_¦s¦b
      If FileExists(Replace(txtPath(0) & "\" & strFileName, "/", "\")) = False Then Exit Sub
      '2012/10/17 End
      '«ü©wÀÉ¦W
      'Modified by Lydia 2016/09/29 ¥ÎÂÂ¼gªk·|³y¦¨Word2010¥X¿ù
      '.ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:= _
      'txtPath(0) & "\" & strFileName, LinkToFile:= _
      'False, SaveWithDocument:=True
      '.ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
      Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=txtPath(0) & "\" & strFileName, LinkToFile:=False, SaveWithDocument:=True)
      oShape.Select
      
'
'         .Selection.ShapeRange.ZOrder msoBringInFrontOfText
'         .Selection.ShapeRange.Fill.Visible = msoTrue
'         .Selection.ShapeRange.Fill.Solid
'         'add by nickc 2007/12/07 ­×¥¿©³¦â
'         .Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
'         .Selection.ShapeRange.Fill.Transparency = 0.5
'         .Selection.ShapeRange.Line.Weight = 0.75
'         .Selection.ShapeRange.Line.DashStyle = msoLineSolid
'         .Selection.ShapeRange.Line.Style = msoLineSingle
'         .Selection.ShapeRange.Line.Transparency = 0#
'         .Selection.ShapeRange.Line.Visible = msoFalse   'msoTrue.µe®Ø½u
'         .Selection.ShapeRange.Line.ForeColor.RGB = RGB(0, 0, 0)
'         .Selection.ShapeRange.Line.BackColor.RGB = RGB(255, 255, 255)
'         .Selection.ShapeRange.LockAspectRatio = msoTrue
'
'         '©w¸q¤j¤p
'         'Âê©w³Ì°ª ¹Ï°Ï
'         '¹Ï¤j¤p
'         .Selection.ShapeRange.Height = 230
'         If .Selection.ShapeRange.Width > 150 Then
'            .Selection.ShapeRange.Width = 150
'         End If
'
'         .Selection.ShapeRange.PictureFormat.Brightness = 0.5
'         .Selection.ShapeRange.PictureFormat.Contrast = 0.5
'         .Selection.ShapeRange.PictureFormat.ColorType = msoPictureAutomatic
'         .Selection.ShapeRange.PictureFormat.CropLeft = 0#
'         .Selection.ShapeRange.PictureFormat.CropRight = 0#
'         .Selection.ShapeRange.PictureFormat.CropTop = 0#
'         .Selection.ShapeRange.PictureFormat.CropBottom = 0#
'         '³]¦ì¸m¬Û¹ï©óÃä¬É,¤£³z©ú
'         .Selection.ShapeRange.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'         .Selection.ShapeRange.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'         .Selection.ShapeRange.Fill.Transparency = 0#
'
'         .Selection.ShapeRange.LockAnchor = False
'         '¹Ï»\¤å
'         .Selection.ShapeRange.WrapFormat.Type = wdWrapSquare 'wdWrapNone.¹Ï»\¤å wdWrapSquare.¤å¦rÂ¶¹Ï
'         .Selection.ShapeRange.WrapFormat.Side = wdWrapBoth
'         .Selection.ShapeRange.WrapFormat.DistanceTop = .CentimetersToPoints(0)
'         .Selection.ShapeRange.WrapFormat.DistanceBottom = .CentimetersToPoints(0)
'         .Selection.ShapeRange.WrapFormat.DistanceLeft = .CentimetersToPoints(0.32)
'         .Selection.ShapeRange.WrapFormat.DistanceRight = .CentimetersToPoints(0.32)
'         '²¾¨ì«ü©w¦ì¸m
'         '¥Ó½Ð¤H¤@¤U­±¨º¤@¦æ
'         .Selection.ShapeRange.Left = .CentimetersToPoints(12) '11.2
'         '.Selection.ShapeRange.Top = .CentimetersToPoints(6.6)
      
      '©w¸q¤j¤p
      'Âê©w³Ì°ª ¹Ï°Ï
      '¹Ï¤j¤p
      'Modified by Lydia 2016/09/29
'      .Selection.ShapeRange.LockAspectRatio = msoTrue
'      .Selection.ShapeRange.Height = 230
'      If .Selection.ShapeRange.Width > 150 Then
'         .Selection.ShapeRange.Width = 150
'      End If
'      '²¾¨ì«ü©w¦ì¸m
'      .Selection.ShapeRange.Left = .CentimetersToPoints(12) '11.2
'      '.Selection.ShapeRange.Top = .CentimetersToPoints(1)
'      .Selection.ShapeRange.LockAnchor = False
'      '¹Ï»\¤å
'      .Selection.ShapeRange.WrapFormat.Type = wdWrapSquare 'wdWrapNone.¹Ï»\¤å wdWrapSquare.¤å¦rÂ¶¹Ï
      oShape.LockAspectRatio = msoTrue
      oShape.Height = 230
      If oShape.Width > 150 Then
         oShape.Width = 150
      End If
      '²¾¨ì«ü©w¦ì¸m
      oShape.Left = .CentimetersToPoints(12)
      oShape.LockAnchor = False
      '¹Ï»\¤å
      oShape.WrapFormat.Type = wdWrapSquare
      
      .Selection.EndKey Unit:=wdStory
   End With
   Exit Sub
   
'Add by Morgan 2008/7/16 ¥[§PÂ_­Y¿ù»~¬°µLªk§R°£ÀÉ®×®ÉÄ~Äò(¤U¦¸¶]¾ã§å©w½Z®É·|§R)
ErrHnd:
   If (pub_OS = 1 And Err.Number = 75) Or (pub_OS <> 1 And Err.Number = 70) Then 'Err.Number = 5152
      Resume Next
   Else
      Err.Raise Err.Number
   End If
End Sub

'½Õ¾ã­º¦æ¥Y±Æ
Sub PhaseIndent()
    g_WordAp.Selection.WholeStory
    With g_WordAp.Selection.ParagraphFormat
        .LeftIndent = g_WordAp.CentimetersToPoints(1)
        .RightIndent = g_WordAp.CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 15
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = g_WordAp.CentimetersToPoints(-1)
        .OutlineLevel = wdOutlineLevelBodyText
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = True
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub

Private Function getAddrData() As String
Dim strAddrData As String
Dim m_line As Variant
Dim ii As Integer
   
   '¦a§}
   If m_AppAddr = "" Then
      m_AppAddr = String(20, "¡@")
   Else
      m_AppAddr = ToWide(Trim(CheckStr(m_AppAddr)))
   End If
   '¦¬¥ó¤H
   'Modify By Sindy 2016/11/16 ¦]¡¨§õ ¿ªÚ¡¨ ¿¬°Ãø¦r,µ{¦¡¸ÌÅªSql¤£¥iTrim±¼
   'm_AppName = Trim(CheckStr(m_AppName))
   m_AppName = CheckStr(m_AppName)
   '2016/11/16 END
   If m_AppAddrZip <> "" Then
      strAddrData = m_AppAddrZip & vbCrLf & m_AppAddr & vbCrLf & m_AppName & "¡@¶v±Ò"
   Else
      strAddrData = m_AppAddr & vbCrLf & m_AppName & "¡@¶v±Ò"
   End If
   If strAddrData <> "" Then
      m_line = Split(strAddrData, vbCrLf)
      For ii = 0 To UBound(m_line)
         strAddrData = m_line(ii)
         Do While strAddrData <> StrToStr(strAddrData, 17)
               If InStr(1, m_line(ii), StrToStr(strAddrData, 17)) = 1 Then
                   m_line(ii) = Mid(m_line(ii), 1, InStr(1, m_line(ii), StrToStr(strAddrData, 17)) - 1) & StrToStr(strAddrData, 17) & vbCrLf & Replace(m_line(ii), StrToStr(strAddrData, 17), "")
               Else
                   m_line(ii) = Mid(m_line(ii), 1, InStr(1, m_line(ii), StrToStr(strAddrData, 17)) - 1) & StrToStr(strAddrData, 17) & vbCrLf & Replace(Mid(m_line(ii), InStr(1, m_line(ii), StrToStr(strAddrData, 17))), StrToStr(strAddrData, 17), "")
               End If
               strAddrData = Replace(strAddrData, StrToStr(strAddrData, 17), "")
         Loop
      Next ii
      strAddrData = Join(m_line, vbCrLf)
      m_line = Split(strAddrData, vbCrLf)
      For ii = 0 To UBound(m_line)
           m_line(ii) = m_line(ii)
      Next ii
      strAddrData = Join(m_line, vbCrLf)
      m_line = Split(strAddrData, vbCrLf)
      If UBound(m_line) < 3 Then
           strAddrData = strAddrData & vbCrLf
      End If
   End If
   
   getAddrData = strAddrData
End Function

Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         'ÀË¬d¸ê®Æ
         If txtPath(0).Text = "" Then
            MsgBox "ÀÉ®×¸ô®|¤£¥iªÅ¥Õ¡I", vbExclamation
            txtPath(0).SetFocus
            Exit Sub
         End If
         Process
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   If QueryData = False Then
      cmdOK(0).Enabled = False
   Else
      cmdOK(0).Enabled = True
   End If
   
   If Pub_StrUserSt03 = "M51" Then
      txtPath(0).Enabled = True
   Else
      txtPath(0).Enabled = False
   End If
   
   'Add By Sindy 2015/5/13 °Ó¼Ð³Bµ{§Ç¤H­û
   strExc(0) = "select st01 from staff where st03='P22' and st04='1' and substr(st01,1,1) in(" & ST01CodeNum1 & ")"
   intI = 1
   strP22 = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strP22 = strP22 & ";" & RsTemp.Fields("st01")
         RsTemp.MoveNext
      Loop
   End If
   RsTemp.Close
   If strP22 <> "" Then
      strP22 = Mid(strP22, 2)
   End If
   '2015/5/13 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030619 = Nothing
End Sub

Private Sub txtPath_GotFocus(Index As Integer)
   TextInverse txtPath(Index)
End Sub

Private Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   QueryData = False
   Label3 = "( 0 µ§)"
   txtTBD01 = ""
   
   'Modify By Sindy 2013/3/1 + where tbnp08='T'
   'Modify By Sindy 2018/12/13 + and tbd16='1' : ¤½³ø¶}©Ý
   strSql = "select count(distinct tbor03) from tmbulletinowner,tmbulletindata,Trademark " & _
            "Where tbor02=1 " & _
            "and tbor01=tbd02 and tbor06=tbd03 and tbd16='1' and tbd15='A' and (tbd14<>'N' or tbd14 is null) " & _
            "and tbd04=TM12 and tm44 is null " & _
            "and ltrim(rtrim(tbor03)) not in(select ltrim(rtrim(tbnp01)) from tmbulletinnp where tbnp08='T') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields(0)) Then
         If Val(rsTmp.Fields(0)) > 0 Then
            QueryData = True
            Label3 = "( " & rsTmp.Fields(0) & " µ§)"
            
            rsTmp.Close
            'Modify By Sindy 2018/12/13 + and tbd16='1' : ¤½³ø¶}©Ý
            strSql = "SELECT distinct tbd01 FROM TMBulletinData WHERE tbd15='A' and tbd16='1' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenDynamic
            If rsTmp.RecordCount > 0 Then
               txtTBD01 = "" & rsTmp.Fields(0)
            End If
         End If
      End If
   End If
   Set rsTmp = Nothing
   
   If QueryData = False Then
      MsgBox "µL¸ê®Æ¡I", vbOKOnly, "°Ó¼Ð¤½³ø¶}©Ý¨ç¦C¦L"
   End If
End Function

'±N¤½³ø¨÷´ÁÂà´«¬°¤é´Á
Private Function ChgTMBM07ToDate(strData As String)
Dim strYY As String
Dim strMM As String
Dim strDD As String
'920101 : 3001, 920116 : 3002 ...(¨C¦~·|¦³24´Á)

strYY = (Val(Mid(strData, 1, Len(strData) - 2)) + 62)
strMM = Format(Right(strData, 2) / 2, "00")
If Right(strData, 2) Mod 2 <> 0 Then
    strDD = "01"
Else
    strDD = "16"
End If
ChgTMBM07ToDate = DBDATE(strYY & strMM & strDD)
End Function

Private Sub GetSales()
Dim stCP13 As String, stCP12 As String
   
   '·~°È°Ï¤Î´¼Åv¤H­û
   stCP13 = PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
   stCP12 = GetSalesArea(stCP13)
   custarea = GetDepartmentName(stCP12)
   custsales = GetStaffName(stCP13)
   
   '68096¤¤¤T§ù°ÆÁ`ªº«È¤á©w½Z¯S§O±±¨î(©w½Z¥æªô¯À½¬)
   '¸Ó«È¤á©Ò¦³®×¥ó³Ì«á¦¬¤å´¼Åv¤H­û¦bÂ¾«h¤£¦L·~°È°Ï¦Ó´¼Åv¤H­û§ï¬°¤¤¤T§ù°ÆÁ`¡]¢æ¢æ¢æ¡^
   '                          Â÷Â¾«h¥¿±`¦C¦L
   If stCP13 = "68096" Then
      strExc(0) = "select st02 from staff,(select max(cp05||cp13) cp13 from ( " & _
                  "      Select cp05,cp13 From patent, caseprogress Where pa26='" & GetPrjPeopleNum1(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04) & "' and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 and cp09<'B' " & _
                  "union Select cp05,cp13 From trademark, caseprogress Where tm23='" & GetPrjPeopleNum1(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04) & "' and tm01=cp01 and tm02=cp02 and tm03=cp03 and tm04=cp04 and cp09<'B' " & _
                  "union Select cp05,cp13 From lawcase, caseprogress Where lc11='" & GetPrjPeopleNum1(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04) & "' and lc01=cp01 and lc02=cp02 and lc03=cp03 and lc04=cp04 and cp09<'B' " & _
                  "union Select cp05,cp13 From servicepractice, caseprogress Where sp08='" & GetPrjPeopleNum1(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04) & "' and sp01=cp01 and sp02=cp02 and sp03=cp03 and sp04=cp04 and cp09<'B' " & _
                  "union Select cp05,cp13 From hirecase, caseprogress Where hc05='" & GetPrjPeopleNum1(m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04) & "' and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and cp09<'B' " & _
                  ")) aa where substr(aa.cp13,9)=st01(+) and st04='1'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         custarea = ""
         custsales = custsales & "¡]" & RsTemp.Fields(0).Value & "¡^"
      End If
   End If
End Sub
