VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030616 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "°Ó¼Ð¤½³øÂàÀÉ§@·~"
   ClientHeight    =   5730
   ClientLeft      =   40
   ClientTop       =   280
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8950
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Caption         =   "Frame2"
      Height          =   1395
      Left            =   360
      TabIndex        =   23
      Top             =   2310
      Width           =   4125
      Begin VB.TextBox txtPath3 
         Height          =   264
         Left            =   1410
         TabIndex        =   27
         Text            =   "C:\temp\XmlTrans"
         Top             =   240
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.OptionButton Option1 
         Caption         =   "¤½³ø ÂàÀÉ"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   300
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.TextBox txtTBD17 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   264
         Left            =   1410
         MaxLength       =   5
         TabIndex        =   7
         Top             =   960
         Width           =   1092
      End
      Begin VB.TextBox txtTMBM07 
         Height          =   264
         Left            =   1410
         MaxLength       =   5
         TabIndex        =   0
         Top             =   570
         Width           =   1092
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ºM¤T ÂàÀÉ"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2010
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "FTP¥Nªí¹Ï¸ô®|¡G"
         Height          =   180
         Left            =   75
         TabIndex        =   28
         Top             =   300
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label10 
         Caption         =   "¶}©Ý¤½³ø¦~¤ë¡G"
         Height          =   210
         Left            =   180
         TabIndex        =   26
         Top             =   990
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "¤½³ø¨÷´Á¡G"
         Height          =   210
         Left            =   540
         TabIndex        =   25
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "(               µ§)"
         Height          =   210
         Left            =   2550
         TabIndex        =   24
         Top             =   600
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdTemp 
      Caption         =   "¸ÉÂà¥Ó½Ð¤H¦WºÙ"
      Height          =   405
      Left            =   6840
      TabIndex        =   10
      Top             =   3330
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   345
      Left            =   7980
      TabIndex        =   3
      Top             =   750
      Width           =   345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÀË¬d¼f©w¸¹¼Æ¸õ¸¹"
      Height          =   405
      Left            =   4650
      TabIndex        =   9
      Top             =   3330
      Width           =   1905
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   30
      TabIndex        =   17
      Top             =   5250
      Width           =   8895
      Begin VB.TextBox Text2 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H00FF0000&
         Height          =   300
         Left            =   30
         TabIndex        =   19
         Top             =   120
         Width           =   8820
      End
   End
   Begin VB.FileListBox File2 
      Height          =   180
      Left            =   1560
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   735
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   405
      Left            =   1020
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   864
      _ExtentY        =   723
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frm030616.frx":0000
   End
   Begin VB.TextBox txtPath2 
      Height          =   315
      Left            =   1410
      TabIndex        =   2
      Text            =   "C:\temp\XmlTrans"
      Top             =   1830
      Width           =   6555
   End
   Begin VB.TextBox txtPath1 
      Height          =   315
      Left            =   1410
      TabIndex        =   1
      Text            =   "D:"
      Top             =   750
      Width           =   6555
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "«þ¨©¤½³ø¸ê®Æ(&C)"
      Height          =   400
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdTransFile 
      Caption         =   "ÂàÀÉ(&T)"
      Height          =   400
      Left            =   4650
      TabIndex        =   8
      Top             =   2850
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7260
      TabIndex        =   11
      Top             =   120
      Width           =   912
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "µù¡G·í¸ô®|§ä¤£¨ì®É¡A½Ðª½±µ¶i¸ê®Æ§¨ÂI¿ïXMLÀÉ®×¡A¦A«ö«þ¨©§Y¥i¡C"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   1410
      TabIndex        =   22
      Top             =   1140
      Width           =   5595
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "D:\°Ó¼Ð¤½³ø39¨÷1´Á\xml\RegContent\039001_RegContent0_01494137.xml"
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   1410
      TabIndex        =   21
      Top             =   1560
      Width           =   5490
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "[ÂI¿ïXMLÀÉ®×ªº½d¨Ò]"
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   1410
      TabIndex        =   20
      Top             =   1350
      Width           =   1755
   End
   Begin VB.Label Label2 
      Caption         =   "ÂàÀÉ¤¤, ½Ðµy­Ô . . ."
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   15.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   30
      TabIndex        =   18
      Top             =   4890
      Width           =   8895
   End
   Begin VB.Label Label6 
      Caption         =   "³Æµù¡G¦P®É§ó·s°Ó¼Ð°ò¥»ÀÉªº¼f©w¸¹"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   120
      TabIndex        =   16
      Top             =   4110
      Width           =   3495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "¥úºÐ¥Øªº¸ô®|¡G"
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   1890
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "¥úºÐ¨Ó·½¸ô®|¡G"
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   780
      Width           =   1260
   End
   Begin MSForms.TextBox txtChkWord 
      Height          =   300
      Left            =   5310
      TabIndex        =   29
      Top             =   4680
      Visible         =   0   'False
      Width           =   3380
      VariousPropertyBits=   679495707
      MaxLength       =   100
      Size            =   "5962;529"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm030616"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/3/3 Form2.0¤w­×§ï
'Memo By Sindy 2012/12/5 ´¼Åv¤H­ûÄæ¤w­×§ï
Option Explicit

Dim m_bolCharQ  As Boolean, m_strCharQNote As String
Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iLine As Integer, iLine2 As Integer
Dim m_PrintRpt1 As Boolean, m_PrintRpt2 As Boolean, m_PrintRpt3 As Boolean
Dim ff1 As Integer, FF2 As Integer, ff3 As Integer
Dim m_strFileName1 As String, m_strFileName2 As String, m_strFileName3 As String
Dim strErrTxt As String
Dim bolIsTaiwanCase As Boolean, bolIsChinaCase As Boolean
Dim strP22 As String 'Add By Sindy 2015/5/13
'Add By Sindy 2017/4/21
Dim bolTaieCase As Boolean
Dim bolTaieCase01 As Boolean
Dim strAChinese As String, strAChinese1 As String
Dim strAddress1 As String
Dim strTMBM05 As String, strTMBM06 As String, strTMBM08 As String
Dim strTMBM06_temp1 As String
Dim strTBD02 As String, strTBD03 As String, strTBD04 As String, strTBD05 As String, strTBD06 As String
Dim strTBD07 As String, strTBD08 As String, strTBD09 As String, strTBD10 As String, strTBD11 As String
Dim strTBD12 As String, strTBD13 As String, strTBD14 As String, strTBD15 As String
Dim strTM01 As String, strTM02 As String, strTM03 As String, strTM04 As String, strTM09 As String
Dim dblChar As Double, dblLastEnd As Double
Dim strText As String, strTitNM As String
Dim dblStar As Double, dblEnd As Double, intCol As Integer, strData As String
Dim strOurAgentName As String, strMsg As String
Dim i As Integer, j As Integer
Dim intApp As Integer, strTMBMApp(1 To 10) As String
'2017/4/21 END
Dim adoStream As ADODB.Stream 'Add By Sindy 2022/3/3
Dim m_strTextBox As String 'Add by Sindy 2022/3/3
Dim m_strText As String 'Add By Sindy 2024/5/17

Private Sub cmdCopy_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim fs As Object, strTime As String
Dim DeleteFilePathErr As Boolean
Dim bolCopyToSer As Boolean 'Add By Sindy 2023/8/2
   
On Error GoTo ErrHnd
   
   strTime = time()
   DeleteFilePathErr = False
   
   If IsEmptyText(txtTMBM07) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤J¤½³ø¨÷´Á¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      txtTMBM07.SetFocus
      Exit Sub
   End If
   If IsEmptyText(txtPath1) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤J¥úºÐ¨Ó·½¸ô®|¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      txtPath1.SetFocus
      Exit Sub
   End If
   If IsEmptyText(txtPath2) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤J¥úºÐ¥Øªº¸ô®|¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      txtPath2.SetFocus
      Exit Sub
   End If
   
   If Right(Trim(txtPath1), 1) = "\" Then txtPath1 = Left(txtPath1, Len(txtPath1) - 1)
   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   
   '¹LÂo¥úºÐ¾÷¸ô®|¤º®e
   If InStr(UCase(txtPath1), "XML") > 0 Then
      txtPath1 = Left(txtPath1, InStr(UCase(txtPath1), "XML") - 2)
   End If
   
   File2.path = txtPath1.Text & "\xml\RegContent"
   File2.Refresh
   If File2.ListCount = 0 Then
      MsgBox "¥úºÐ¨Ó·½¸ô®|¤¤µL" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¤½³ø¸ê®Æ¡I"
      txtPath1.SetFocus
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   'Add By Sindy 2018/12/19
   'ºM¤T,ÀË¬d¬O§_¦³¥L´Áªº¸ê®Æ,­Y¦³¥ý§R°£¸ê®Æ§¨
   Dim pErrMsg As String
   Dim pFileName As String, stFileName As String
   Dim hConnection As Long
   Dim pFtpPath As String
   Dim pData As WIN32_FIND_DATA
   Dim hFind As Long
   Dim LRet   As Long
   Dim strTBD17 As String
   If Option1(1).Value = True Then
      'ÀË¬d¬O§_¦³»Ý­n²M°£ªº¸ê®Æ
      strSql = "SELECT tbd16,tbd17 FROM tmbulletindata" & _
               " WHERE tbd16='2' and tbd17<>" & Left(DBDATE(txtTBD17 & "01"), 6) & _
               " group by tbd16,tbd17"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strTBD17 = RsTemp.Fields(1)
      End If
      If strTBD17 <> "" Then
         hConnection = PUB_GetFtpConnect(pErrMsg, , , Pub_GetSpecMan("FTP_TM31"))
         If hConnection <> 0 Then
            'FTPªºAPI¦^¶Ç­È¦ü¥G³£¬O¼Æ­È(¤£ºÞ«Å§i¬°¦ó),µLªk¥Î¥¬ªL§PÂ_
            'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
            'pFtpPath = "\\" & Replace(UCase(txtPath3), "\\SALE1\", "") & "\" & strTBD17 - 191100 & "\"
            pFtpPath = "\\" & Replace(UCase(txtPath3), "\\" & strSale1Path & " \", "") & "\" & strTBD17 - 191100 & "\"
            pFtpPath = Replace(pFtpPath, "\", "/")
            If FtpSetCurrentDirectory(hConnection, pFtpPath) = 1 Then
               '§R°£¥Ø¿ý¤º¥þ³¡ÀÉ®×
               If pFileName = "" Then
                  pData.cFileName = String(MAX_PATH, 0)
                  hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
                  If hFind <> 0 Then
                     Do
                        stFileName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
                        If stFileName <> "." And stFileName <> ".." Then
                           If FtpDeleteFile(hConnection, stFileName) <> 1 Then
                              pErrMsg = pFileName & "ÀÉ®×§R°£¥¢±Ñ¡I"
                              GoTo ErrHnd
                           End If
                        End If
                        LRet = InternetFindNextFile(hFind, pData)
                     Loop While LRet <> 0
                     
                     If FtpRemoveDirectory(hConnection, pFtpPath) <> 1 Then
                        pErrMsg = pFileName & "¥Ø¿ý§R°£¥¢±Ñ¡I"
                        GoTo ErrHnd
                     End If
                     
                     InternetCloseHandle hFind
                     hFind = 0
                  End If
               End If
            End If
         End If
      End If
   End If
   '2018/12/19 END
   
   Set fs = CreateObject("Scripting.FileSystemObject")
   DeleteFilePathErr = True
   fs.DeleteFolder txtPath2, True
NotFolder76:
   fs.CreateFolder txtPath2
   fs.CreateFolder txtPath2 & "\imagesdata"
   fs.CreateFolder txtPath2 & "\RegContent"
   fs.CreateFolder txtPath2 & "\Reject" 'Add By Sindy 2012/9/21
   fs.CreateFolder txtPath2 & "\Revocation" 'Add By Sindy 2012/9/21
   fs.CreateFolder txtPath2 & "\DelProduct" 'Add By Sindy 2024/6/27
   fs.CopyFile txtPath1 & "\imagesdata\*.*", txtPath2 & "\imagesdata\"
   fs.CopyFile txtPath1 & "\xml\RegContent\*.*", txtPath2 & "\RegContent\"
   fs.CopyFile txtPath1 & "\xml\Reject\*.*", txtPath2 & "\Reject\" 'Add By Sindy 2012/9/21
   fs.CopyFile txtPath1 & "\xml\Revocation\*.*", txtPath2 & "\Revocation\" 'Add By Sindy 2012/9/21
   fs.CopyFile txtPath1 & "\xml\DelProduct\*.*", txtPath2 & "\DelProduct\" 'Add By Sindy 2024/6/27
   
   'Add By Sindy 2023/8/1 «þ¨©°Ó¼Ð¹ÏÀÉ¨ìServer¤W,¥H³Æµù¥UÃÒ©w½Z¨Ï¥Î
   If Option1(0).Value = True Then
      If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or UCase(pub_DbTerminalName) <> UCase(¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ) Then
         nResponse = MsgBox("¦³­n«þ¨©°Ó¼Ð¹ÏÀÉ¨ìServer¤W¡A¥H³Æµù¥UÃÒ©w½Z¨Ï¥Î¶Ü¡H", vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbNo Then
            bolCopyToSer = False
         Else
            bolCopyToSer = True
         End If
      Else
         bolCopyToSer = True
      End If
      If bolCopyToSer = True Then
         hConnection = PUB_GetFtpConnect(pErrMsg, , , Pub_GetSpecMan("FTP_TM31"))
         If hConnection <> 0 Then
            'FTPªºAPI¦^¶Ç­È¦ü¥G³£¬O¼Æ­È(¤£ºÞ«Å§i¬°¦ó),µLªk¥Î¥¬ªL§PÂ_
            'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
            'pFtpPath = "\\" & Replace(UCase(Pub_GetSpecMan("¤º°Ó¶}©Ý¸ê®Æ¦s©ñ¸ô®|")), "\\SALE1\", "") & "\XmlTrans\imagesdata\"
            pFtpPath = "\\" & Replace(UCase(Pub_GetSpecMan("¤º°Ó¶}©Ý¸ê®Æ¦s©ñ¸ô®|")), "\\" & UCase(strSale1Path) & "\", "") & "\XmlTrans\imagesdata\"
            pFtpPath = Replace(pFtpPath, "\", "/")
            If FtpSetCurrentDirectory(hConnection, pFtpPath) = 1 Then
               '§R°£¥Ø¿ý¤º¥þ³¡ÀÉ®×
               If pFileName = "" Then
                  pData.cFileName = String(MAX_PATH, 0)
                  hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
                  If hFind <> 0 Then
                     Do
                        stFileName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
                        If stFileName <> "." And stFileName <> ".." Then
                           If FtpDeleteFile(hConnection, stFileName) <> 1 Then
                              pErrMsg = pFileName & "ÀÉ®×§R°£¥¢±Ñ¡I"
                              GoTo ErrHnd
                           End If
                        End If
                        LRet = InternetFindNextFile(hFind, pData)
                     Loop While LRet <> 0
                     
                     If FtpRemoveDirectory(hConnection, pFtpPath) <> 1 Then
                        pErrMsg = pFileName & "¥Ø¿ý§R°£¥¢±Ñ¡I"
                        GoTo ErrHnd
                     End If
                     
                     InternetCloseHandle hFind
                     hFind = 0
                  End If
               End If
            End If
            '¥Nªí¹Ï»Ý­n¤W¶ÇFTP
            'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
            'Call FtpPutFileTM31(txtPath1 & "\imagesdata", "\\" & Replace(UCase(Pub_GetSpecMan("¤º°Ó¶}©Ý¸ê®Æ¦s©ñ¸ô®|")), "\\SALE1\", "") & "\XmlTrans\imagesdata\")
            Call FtpPutFileTM31(txtPath1 & "\imagesdata", "\\" & Replace(UCase(Pub_GetSpecMan("¤º°Ó¶}©Ý¸ê®Æ¦s©ñ¸ô®|")), "\\" & UCase(strSale1Path) & "\", "") & "\XmlTrans\imagesdata\")
         End If
      End If
   End If
   '2023/8/1 END
   
   'Add By Sindy 2018/12/19
   'ºM¤T,¥Nªí¹Ï»Ý­n¤W¶ÇFTP
   If Option1(1).Value = True Then
      'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
      'Call FtpPutFileTM31(txtPath1 & "\imagesdata", "\\" & Replace(UCase(txtPath3), "\\SALE1\", "") & "\" & txtTBD17 & "\")
      Call FtpPutFileTM31(txtPath1 & "\imagesdata", "\\" & Replace(UCase(txtPath3), "\\" & strSale1Path & "\", "") & "\" & txtTBD17 & "\")
   End If
   '2018/12/19 END
   
   Screen.MousePointer = vbDefault
   MsgBox "«þ¨©§¹²¦¡I(«þ¨©ªá¶O®É¶¡¡G" & strTime & "  " & time() & ")"
   Exit Sub
   
ErrHnd:
   If Err.Number = 76 And DeleteFilePathErr = True Then
      GoTo NotFolder76
   ElseIf Err.Number = 68 Or Err.Number = 76 Then
      MsgBox "¥úºÐ¨Ó·½¸ô®|¤¤µL" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¤½³ø¸ê®Æ¡I"
      txtPath1.SetFocus
   Else
      MsgBox Err.Description & IIf(pErrMsg <> "", "(" & pErrMsg & ")", "")
   End If
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2018/12/18
'FTPÀÉ®×¤W¶Ç(TM31)
Private Function FtpPutFileTM31(pLocalPath As String, pFtpPath As String, _
   Optional pErrMsg As String, Optional pRaiseErr As Boolean = True) As Boolean
   
   Dim stDir As String, stFileName As String
   Dim hConnection As Long
   Dim hFind As Long
   Dim pData As WIN32_FIND_DATA
   Dim dwInternetFlags As Integer
   Dim pFtpSrv As String
   Dim dblFCnt As Double
   
   pFtpSrv = Pub_GetSpecMan("FTP_TM31")
   pFtpPath = Replace(pFtpPath, "\", "/")
   'stDir = "//" & Mid(pFtpPath, InStr(3, pFtpPath, "/") + 1)
   stDir = Left(pFtpPath, InStrRev(pFtpPath, "/") - 1)
   'stFileName = Mid(pFtpPath, InStrRev(pFtpPath, "/") + 1)
   
   hConnection = PUB_GetFtpConnect(pErrMsg, , , pFtpSrv)
   
   If hConnection <> 0 Then
      '¤Á´«©Ò¦b¦a¥Ø¿ý
      If PUB_SetFtpDirectory(hConnection, stDir, pErrMsg, pRaiseErr) = False Then GoTo OutPort
      
      'dwInternetFlags = FTP_TRANSFER_TYPE_BINARY
      dwInternetFlags = 2 'INTERNET_FLAG_TRANSFER_BINARY
      
      File2.path = pLocalPath 'txtPath1.Text & "\imagesdata"
      File2.Refresh
      '­Y¦³¦PÀÉ¦W,«h·|ª½±µÂÐ»\
      For dblFCnt = 0 To File2.ListCount - 1
         If FtpPutFile(hConnection, pLocalPath & "\" & File2.List(dblFCnt), File2.List(dblFCnt), dwInternetFlags, 0) <> 1 Then
            pErrMsg = pLocalPath & "ÀÉ®×¤W¶Ç¥¢±Ñ¡I"
            GoTo OutPort
         End If
      Next dblFCnt
      FtpPutFileTM31 = True
   End If
   
OutPort:
   If Err.Number <> 0 Then pErrMsg = Err.Description
   If hConnection <> 0 Then InternetCloseHandle (hConnection)
   
   If FtpPutFileTM31 = False And pRaiseErr = True Then
      Err.Raise 999, , pErrMsg & IIf(Err.Number <> 0, "(" & Err.Number & ")", "")
   End If
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

'Add By Sindy 2017/4/21 ¸ÉÂà¥Ó½Ð¤H¦WºÙ
Private Sub cmdTemp_Click()
Dim strTit As String
Dim nResponse
Dim dblFCnt As Double
Dim strTime As String, strTotRow As String
Dim dblMaxWidth As Double
Dim strSubject As String
   
On Error GoTo ErrHand
   
   strTime = time()
   
   '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
   If TxtValidate = False Then Exit Sub
   
   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   File2.path = txtPath2.Text & "\RegContent"
   File2.Refresh
   If Val(Val(Left(File2.List(0), 3)) & Mid(File2.List(0), 5, 2)) <> Val(txtTMBM07) Then
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2.Text & "\RegContent" & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡I", vbExclamation, "¸ÉÂà¥Ó½Ð¤H¦WºÙ"
      txtPath2.SetFocus
      Exit Sub
   End If
   
   If IsRecordExist = True Then
      strTit = "¸ÉÂà¥Ó½Ð¤H¦WºÙ"
      strMsg = "¤½³ø¨÷´Á" & txtTMBM07 & "¤w¦³¸ê®Æ¦s¦b¡A½T©w¬O§_­n­«·sÂàÀÉ¡H"
      nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
      If nResponse = vbNo Then Exit Sub
   Else
      MsgBox "¸Ó´Á¤½³ø¸ê®Æ©|¥¼¶×¤J¡AµLªk§ó·s¥Ó½Ð¤H¡I", vbExclamation, "¸ÉÂà¥Ó½Ð¤H¦WºÙ"
      txtTMBM07.SetFocus
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   strOurAgentName = GetTOurAgentName()
   m_PrintRpt1 = False: m_PrintRpt2 = False: m_PrintRpt3 = False: iLine = 0: iLine2 = 0
   strTotRow = File2.ListCount
   Me.Height = 6120
   dblMaxWidth = 8820
   Text2.Width = 0
   For dblFCnt = 0 To File2.ListCount - 1
      'Add by Sindy 2022/3/3
      If strSrvDate(1) >= Form20¤W½u¤é Then
         adoStream.LoadFromFile (txtPath2.Text & "\RegContent\" & File2.List(dblFCnt))
         m_strTextBox = adoStream.ReadText
      Else
      '2022/3/3 END
         RichTextBox1.LoadFile (txtPath2.Text & "\RegContent\" & File2.List(dblFCnt))
         m_strTextBox = RichTextBox1.Text
      End If
      dblLastEnd = InStr(m_strTextBox, "</RegContent>")
      
      Text2.Width = dblMaxWidth / Val(strTotRow) * (dblFCnt + 1): DoEvents
      
      cnnConnection.BeginTrans
      
      If ReadXmlData = False Then GoTo ErrHand
      
      'Add By Sindy 2017/4/25 ¥Ó½Ð¤H1~10
      If intApp > 0 Then
         strSql = "update TMBulletin set" & _
                  " TMBM09=" & CNULL(strTMBMApp(1)) & _
                  ",TMBM10=" & CNULL(strTMBMApp(2)) & _
                  ",TMBM11=" & CNULL(strTMBMApp(3)) & _
                  ",TMBM12=" & CNULL(strTMBMApp(4)) & _
                  ",TMBM13=" & CNULL(strTMBMApp(5)) & _
                  ",TMBM14=" & CNULL(strTMBMApp(6)) & _
                  ",TMBM15=" & CNULL(strTMBMApp(7)) & _
                  ",TMBM16=" & CNULL(strTMBMApp(8)) & _
                  ",TMBM17=" & CNULL(strTMBMApp(9)) & _
                  ",TMBM18=" & CNULL(strTMBMApp(10)) & _
                  " where TMBM01='" & strTBD02 & "' and TMBM02='" & strTBD03 & "'"
         cnnConnection.Execute strSql
      End If
      '2017/4/25 END
      
      cnnConnection.CommitTrans
      
      'Add By Sindy 2017/4/26 «D¥xÆW®×ªº°Ó¼ÐÅv¤H¤¤¤å¦³?®É,»Ý¦C¦L²M³æ
      'If bolIsChinaCase = False And intApp > 0 Then
      If intApp > 0 Then
         txtChkWord = strTMBMApp(1)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(1), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(1), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(1), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(1), strAddress1, strTBD03)
         txtChkWord = strTMBMApp(2)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(2), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(2), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(2), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(2), "", strTBD03)
         txtChkWord = strTMBMApp(3)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(3), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(3), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(3), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(3), "", strTBD03)
         txtChkWord = strTMBMApp(4)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(4), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(4), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(4), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(4), "", strTBD03)
         txtChkWord = strTMBMApp(5)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(5), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(5), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(5), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(5), "", strTBD03)
         txtChkWord = strTMBMApp(6)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(6), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(6), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(6), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(6), "", strTBD03)
         txtChkWord = strTMBMApp(7)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(7), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(7), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(7), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(7), "", strTBD03)
         txtChkWord = strTMBMApp(8)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(8), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(8), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(8), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(8), "", strTBD03)
         txtChkWord = strTMBMApp(9)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(9), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(9), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(9), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(9), "", strTBD03)
         txtChkWord = strTMBMApp(10)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(10), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(10), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(10), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(10), "", strTBD03)
      End If
      '2017/4/26 END
   Next dblFCnt
   
   strMsg = ""
   If m_PrintRpt1 = True Then
'      Close ff1
      strMsg = m_strFileName1
      'Add By Sindy 2024/5/17
      If Dir(PUB_Getdesktop & "\" & m_strFileName1) <> "" Then
         Kill PUB_Getdesktop & "\" & m_strFileName1
         Sleep 100
      End If
      Call PUB_SaveTextAsUTF8(PUB_Getdesktop & "\" & m_strFileName1, m_strText)
      '2024/5/17 END
      strMsg = "½Ð¦Ü¤U¦C¦ì¸m¦C¦LÀË®Öªí¡G" & PUB_Getdesktop & "\" & strMsg
   End If
   
   Screen.MousePointer = vbDefault
   Call IsRecordExist '²£¥Íµ§¼Æ
   
   MsgBox "ÂàÀÉ§¹²¦¡I(ÂàÀÉªá¶O®É¶¡¡G" & strTime & "  " & time() & ")" & vbCrLf & strMsg, vbInformation, "¸ÉÂà¥Ó½Ð¤H¦WºÙ"
   Me.Height = 5000
   
   Exit Sub
   
ErrHand:
   If Err.Number = -2147217900 Then 'ORA-00917: ¿òº|³rÂI
      '¼gLog
      Call ReadTxt3(strSql)
      '±µµÛµo¥Í¿ù»~³¯­z¦¡ªº¤U­Ó³¯­z¦¡¶}©l°õ¦æ
      Resume Next
   End If
   Screen.MousePointer = vbDefault
   
   If Err.Number = 76 Then
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2.Text & "\RegContent" & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡I", vbExclamation, "¸ÉÂà¥Ó½Ð¤H¦WºÙ"
      txtPath2.SetFocus
   Else
      cnnConnection.RollbackTrans
      
      If Err.Number = -2147217873 Then
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¼f©w¸¹¼Æ¡]" & strTBD02 & "¡^°Ó¼ÐºØÃþ¡]" & strTBD03 & "¡^" & vbCrLf & strErrTxt & ": ¹H¤Ï¥²¶·¬°°ß¤@ªº­­¨î±ø¥ó", vbExclamation, "¸ÉÂà¥Ó½Ð¤H¦WºÙ"
      Else
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¼f©w¸¹¼Æ¡]" & strTBD02 & "¡^°Ó¼ÐºØÃþ¡]" & strTBD03 & "¡^" & vbCrLf & strErrTxt & Err.Description, vbExclamation, "¸ÉÂà¥Ó½Ð¤H¦WºÙ"
      End If
   End If
End Sub

'Add By Sindy 2017/4/21
Private Function ReadXmlData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strChar As String
Dim strSDate As String
Dim dblSubChar As Double
Dim strChineseNM As String, strEnglishNM As String
   
   ReadXmlData = True
   
   bolTaieCase = False '«D¥»©Ò®×¸¹
   bolTaieCase01 = False 'Add By Sindy 2012/2/1
   bolIsTaiwanCase = False '«D¥xÆW®×
   bolIsChinaCase = False  '«D¤j³°®×
   m_bolCharQ = False '­Y¦r¤¸¸Ì¦³?«h¬°True,±ý²£¥Í²M³æ
   m_strCharQNote = ""
   strErrTxt = "": strAChinese = "": strAChinese1 = "": strAddress1 = ""
   strTMBM05 = "": strTMBM06 = "": strTMBM06_temp1 = "": strTMBM08 = ""
   strTBD02 = "": strTBD03 = "": strTBD04 = "": strTBD05 = "": strTBD06 = ""
   strTBD07 = "": strTBD08 = "": strTBD09 = "": strTBD10 = "": strTBD11 = ""
   strTBD12 = "": strTBD13 = "": strTBD14 = "": strTBD15 = ""
   strTM01 = "": strTM02 = "": strTM03 = "": strTM04 = "" 'Add By Sindy 2015/1/16
   strTM09 = "" 'Add By Sindy 2015/10/27
   
   'Åª¨ú¤½³ø¸ê®Æ
   dblLastEnd = InStr(m_strTextBox, "</RegContent>")
   For dblChar = 1 To dblLastEnd
      strText = "CaseNo": strTitNM = "¥Ó½Ð®×¸¹"
      dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
      If dblStar < dblChar Then Exit For
      For intCol = 1 To 14
         strData = ""
         If intCol = 1 Then
            strText = "CaseNo": strTitNM = "¥Ó½Ð®×¸¹"
         ElseIf intCol = 2 Then strText = "RegisterNo": strTitNM = "¼f©w¸¹¼Æ"
         ElseIf intCol = 3 Then strText = "Trademark_Name": strTitNM = "°Ó¼Ð¦WºÙ"
         ElseIf intCol = 4 Then strText = "Trademark_Design": strTitNM = "°Ó¼ÐºA¼Ë"
         ElseIf intCol = 5 Then strText = "Filing_Date": strTitNM = "¥Ó½Ð¤é´Á"
         ElseIf intCol = 6 Then strText = "Censor": strTitNM = "¼f¬d¤H­û"
         ElseIf intCol = 7 Then strText = "Priority_Date": strTitNM = "Àu¥ýÅv"
         ElseIf intCol = 8 Then strText = "SDate": strTitNM = "Åv§Q´Á¶¡°_¤é"
         ElseIf intCol = 9 Then strText = "EDate": strTitNM = "Åv§Q´Á¶¡¨´¤é"
         ElseIf intCol = 10 Then strText = "Word_Description": strTitNM = "¤å¦r´y­z"
         ElseIf intCol = 11 Then strText = "Mark_Type": strTitNM = "°Ó¼ÐºØÃþ"
         ElseIf intCol = 12 Then
            strText = "BChinese": strTitNM = "¥N²z¤H"
            If dblEnd <= 0 Then dblEnd = 1
            For dblSubChar = dblEnd To dblLastEnd
               dblStar = InStr(dblSubChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
               If dblStar < dblSubChar Then Exit For
               '***** ¸ÑªRXML *****
               If GetXmlData(dblSubChar, strText, strTitNM, True, strData, dblEnd) = False Then
               '***** End
                  Exit For
               Else
                  '©T©wªº¥N²z¤H¹ï·Óªí
'                  If strData = "?ªö§D" Then
'                     strData = "üÚªö§D"
'                  ElseIf strData = "ªL?¥Í" Then strData = "ªL­}¥Í"
'                  ElseIf strData = "¤ý«T?" Then strData = "¤ý«Tû["
'                  ElseIf strData = "©ö¤å?" Then strData = "©ö¤å®p"
'                  ElseIf strData = "±i¤å?" Then strData = "±i¤å®p"
'                  ElseIf strData = "?ªÚ½÷" Then strData = "üÚªÚ½÷"
'                  ElseIf strData = "ÀF•K®õ" Or strData = "ÀF?®õ" Then strData = "ÀF±Ò®õ" 'Add By Sindy 2014/5/26
'                  End If
                  'Add By Sindy 2018/2/1
                  If InStr(strData, "¡]") > 0 Then
                     strData = Trim(Mid(strData, 1, InStr(strData, "¡]") - 1))
                  End If
                  '2018/2/1 END
                  
                  'Add By Sindy 2017/10/30 ¼W¥[¤ñ¹ï¥N²z¤H
                  'Modify By Sindy 2023/8/1
'                  strData = ReplaceMadeWord(strData, "?") 'Modify By Sindy 2018/5/21 ÀË¬d³y¦r
'                  strData = PUB_FilterBulletinSpecWord("2", strData, "")
                  '2023/8/1 END
                  '2017/10/30 END
                  
                  If strTMBM06_temp1 = "" Then strTMBM06_temp1 = strData '°O¿ý²Ä¤@¦ì¥X¦W¥N²z¤H
                  strTBD08 = strTBD08 & strData & "¡@"
                  '©|¥¼Åª¨ú¨ì¥N²z¤H¦WºÙ®É
                  'Modify By Sindy 2020/1/9
                  'If Trim(strTMBM06) = "" And strData <> "" Then
                  If strData <> "" Then
                  '2020/1/9 END
                     If Trim(strTMBM06) = "" Then
                        '¥H¥Ó½Ð®×¸¹§ìtm12 ¥B tm10='000'¡Atm28='1' ±N¥N²z¤H¦Û°Ê¤W'01'
   '                        strSql = "select cp09 from caseprogress,(SELECT TM01,TM02,TM03,TM04 FROM trademark WHERE tm12='" & strTBD04 & "' AND tm10='000' and tm28='1') " & _
   '                                 "Where CP01=TM01 And cp02=TM02 And cp03=TM03 And cp04=TM04 " & _
   '                                 "and instr('101,308',cp10)>0 and cp27 is not null "
                        'Modify By Sindy 2015/10/27 +,TM09
                        strSql = "SELECT TM01,TM02,TM03,TM04,TM09 FROM trademark " & _
                                  "WHERE tm12='" & strTBD04 & "' AND tm10='000' and tm28='1' "
                        intI = 1
                        Set rsTmp = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           If rsTmp.RecordCount > 0 Then
                              'Add By Sindy 2015/1/16
                              strTM01 = rsTmp.Fields("TM01")
                              strTM02 = rsTmp.Fields("TM02")
                              strTM03 = rsTmp.Fields("TM03")
                              strTM04 = rsTmp.Fields("TM04")
                              '2015/1/16 END
                              strTM09 = "" & rsTmp.Fields("TM09") 'Add By Sindy 2015/10/27
                              bolTaieCase = True '¥»©Ò®×¸¹
                              If InStr(1, strOurAgentName, strData) > 0 Or strData = "ÀF±Ò®õ" Then 'Add By Sindy 2014/5/26 +Or strData = "ÀF±Ò®õ"
                                 strTMBM06 = GetTAgentName("01")
                                 bolTaieCase01 = True 'Add By Sindy 2012/2/1
   '                              Else
   '                                 strMsg = rsTmp.Fields("TM01") & "-" & rsTmp.Fields("TM02") & "-" & rsTmp.Fields("TM03") & "-" & rsTmp.Fields("TM04") & "¬°¥»©Ò®×¥ó¦ý¥N²z¤H¨Ã«D¥»©Ò"
   '                                 Call ReadTxt1(strTBD02, strTBD04, strMsg, "", "", "", strTBD03)
                              End If
                           End If
                        End If
                        rsTmp.Close
                     End If
                     'If Trim(strTMBM06) = "" Then
                     '¨ú±o¤w¦³½s¦Cªº¥N²z¤H¦WºÙ
                     strSql = "SELECT TA02,TA03 FROM TAGENT " & _
                                   "WHERE TA01 = 'T' AND " & _
                                       "replace(replace(TA03,'¡@',''),' ','')='" & Trim(strData) & "' "
                     rsTmp.CursorLocation = adUseClient
                     rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsTmp.RecordCount > 0 Then
                        rsTmp.MoveFirst
                        If Trim(strTMBM06) = "" Then
                           If IsNull(rsTmp.Fields("TA02")) = False Then
                              strTMBM06 = rsTmp.Fields("TA03")
                           End If
                        End If
                     Else
                        'Modify By Sindy 2020/1/9
                        '·s¼W°ê¤º¤½³ø¥N²z¤HÀÉ
                        strSql = "INSERT INTO TAgent (TA01,TA02,TA03,TA04,TA05) " & _
                                 "VALUES ('T','" & GetFreeAgentCode & "','" & Trim(strData) & "','" & Trim(strData) & "'," & DBDATE(GetTA05) & ")"
                        cnnConnection.Execute strSql
                        '2020/1/9 END
                     End If
                     rsTmp.Close
                     'End If
                  End If
               End If
               dblSubChar = dblEnd
            Next dblSubChar
            If strTBD08 <> "" Then strTBD08 = Trim(strTBD08)
            '©|¥¼Åª¨ú¨ì¥N²z¤H¦WºÙ®É
            If Trim(strTMBM06) = "" And strTMBM06_temp1 <> "" Then
               strTMBM06 = strTMBM06_temp1
               'Modify By Sindy 2020/1/9 Mark,§ï«e­±³vµ§µL¸ê®Æ,«hinsert
'               If InStr(strTMBM06_temp1, "?") = 0 Then
'                  '·s¼W°ê¤º¤½³ø¥N²z¤HÀÉ
'                  strSql = "INSERT INTO TAgent (TA01,TA02,TA03,TA04,TA05) " & _
'                           "VALUES ('T','" & GetFreeAgentCode & "','" & Trim(strTMBM06) & "','" & Trim(strTMBM06) & "'," & DBDATE(GetTA05) & ")"
'                  cnnConnection.Execute strSql
'               End If
            End If
            'Modify By Sindy 2012/2/1 ¬°¥»©Ò®×¥ó¦ý¥N²z¤H¨Ã«D¥»©Ò
            If bolTaieCase = True And bolTaieCase01 = False Then
               'strMsg = rsTmp3.Fields("TM01") & "-" & rsTmp3.Fields("TM02") & "-" & rsTmp3.Fields("TM03") & "-" & rsTmp3.Fields("TM04") & "¬°¥»©Ò®×¥ó¦ý¥N²z¤H¨Ã«D¥»©Ò"
               strMsg = strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04 & "¬°¥»©Ò®×¥ó¦ý¥N²z¤H¨Ã«D¥»©Ò"
               Call ReadTxt1(strTBD02, strTBD04, strMsg, "", "", "", strTBD03)
            End If
            '2012/2/1 End
         ElseIf intCol = 13 Then
            If dblEnd <= 0 Then dblEnd = 1
            For dblSubChar = dblEnd To dblLastEnd
               strText = "AChinese": strTitNM = "°Ó¼ÐÅv¤H¤¤¤å"
               dblStar = InStr(dblSubChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
               If dblStar < dblSubChar Then Exit For
               For j = 1 To 2
                  strData = ""
                  If j = 1 Then
                     strText = "AChinese": strTitNM = "°Ó¼ÐÅv¤H¤¤¤å"
                  ElseIf j = 2 Then
                     strText = "Address": strTitNM = "°Ó¼ÐÅv¤H¦a§}"
                  End If
                  dblStar = InStr(dblSubChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
                  If dblStar < dblSubChar Then Exit For
                  '***** ¸ÑªRXML *****
                  If GetXmlData(dblSubChar, strText, strTitNM, False, strData, dblEnd) = False Then
                  '***** End
                     Exit For
                  End If
                  If j = 1 Then '°Ó¼ÐÅv¤H¤¤¤å
                     'Modify By Sindy 2017/7/3
                     '©m¦W¦³³y¦r¦³¹Ï¤ù
                     'strData=¸âµú<img align="absmiddle" height="18px" width="27px" file="106203003/106203003-009.TIF" alt="¨ä¥L«D¹Ï¦¡ ed10999.png" img-content="tif" orientation="portrait" inline="yes" giffile="106203003/106203003-009.png"></img>
                     If InStr(strData, "<") > 0 Then
                        strData = Left(strData, InStr(strData, "<") - 1)
                     End If
                     '2017/7/3 END
                     strAChinese = strData
                     If strAChinese1 = "" Then strAChinese1 = strData
                  ElseIf j = 2 Then '°Ó¼ÐÅv¤H¦a§}
                     'Add By Sindy 2014/2/17 ¹LÂo±¼«e­±ªº¶l»¼°Ï¸¹
                     If IsNumeric(Left(strData, 5)) = True Then
                        strData = Mid(strData, 6)
                     ElseIf IsNumeric(Left(strData, 3)) = True Then
                        strData = Mid(strData, 4)
                     End If
                     '2014/2/17 END
                     If strAddress1 = "" Then strAddress1 = strData
                     If strData <> "" Then
                        If strTMBM05 = "" Then
                           '¥ý¥Î¥þ¦W¤ñ¹ï¦a°Ï
                           If GetNationNo(strData) <> "" Then
                              strTMBM05 = strData
                              Exit For
                           End If
                           '³v¦r¤ñ¹ï
                           For i = 1 To Len(strData)
                              strChar = Left(strData, i)
                              strChar = Replace(strChar, "»O", "¥x")
                              If GetNationNo(strChar) <> "" Then
                                 strTMBM05 = strChar
                                 Exit For
                              End If
                              '[¯S¨Ò]³B²z¥xÆW¦a°Ï¦WºÙ
                              If Len(strChar) = 3 Then
                                 strChar = Left(strChar, 2) & "¿¤"
                                 If GetNationNo(strChar) <> "" Then
                                    strTMBM05 = strChar
                                    Exit For
                                 End If
                              End If
                           Next i
                           '¼Ò½k¤ñ¹ï¦a°Ï¦WºÙ
                           If strTMBM05 = "" Or strTMBM05 = "¤¤°ê¤j³°" Then
                              If strAChinese <> "" Then
                                 strChar = GetNationLike(strAChinese)
                                 If strChar <> "" Then
                                    strTMBM05 = strChar
                                    Exit For
                                 End If
                              End If
                           ElseIf strTMBM05 <> "" Then
                              Exit For
                           End If
                        End If
                     End If
                  End If
               Next j
               dblSubChar = dblEnd
            Next dblSubChar
            '§ï¨îªº¥|¿¤¥«½Ð§ï¬°§ï¨î«áªº·s¥_¥«,¥x¤¤¥«,¥x«n¥«,°ª¶¯¥«
            If strTMBM05 = "¥x¥_¿¤" Then strTMBM05 = "·s¥_¥«"
            If strTMBM05 = "¥x¤¤¿¤" Then strTMBM05 = "¥x¤¤¥«"
            If strTMBM05 = "¥x«n¿¤" Then strTMBM05 = "¥x«n¥«"
            If strTMBM05 = "°ª¶¯¿¤" Then strTMBM05 = "°ª¶¯¥«"
         ElseIf intCol = 14 Then strText = "FileName": strTitNM = "°Ó¼Ð¹Ï¸ô®|ÀÉ¦W"
         End If
         '***** ¸ÑªRXML *****
         If GetXmlData(dblChar, strText, strTitNM, True, strData, dblEnd) = False Then GoTo ReadNextCol1
         '***** End
         If strData <> "" Then
            If intCol = 1 Then '¥Ó½Ð®×¸¹
               strTBD04 = strData
               'ª§Ä³¹ï³y®×¥ó
               strSql = "SELECT TM01,TM02,TM03,TM04 FROM trademark " & _
                        "WHERE tm12='" & strTBD04 & "' AND tm10='000' and tm28<>'1' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strTBD14 = "N"
               End If
            ElseIf intCol = 2 Then strTBD02 = strData '¼f©w¸¹¼Æ
            ElseIf intCol = 3 Then strTBD07 = strData
            ElseIf intCol = 4 Then strTBD05 = strData
            ElseIf intCol = 5 Then strTBD06 = strData
            ElseIf intCol = 6 Then strTBD10 = strData
            ElseIf intCol = 7 Then strTBD13 = strData
            ElseIf intCol = 8 Then strSDate = strData
            ElseIf intCol = 9 Then 'Åv§Q´Á¶¡
               strTBD09 = strSDate & "~" & strData
            ElseIf intCol = 10 Then strTBD11 = strData
            ElseIf intCol = 11 Then '°Ó¼ÐºØÃþ
               strTBD03 = strData
               If strTBD03 = "0" Then
                  strTBD03 = "1" '°Ó¼Ð
               ElseIf strTBD03 = "1" Then
                  strTBD03 = "9" '¹ÎÅé°Ó¼Ð
               ElseIf strTBD03 = "2" Then
                  strTBD03 = "7" 'ÃÒ©ú¼Ð³¹
               ElseIf strTBD03 = "3" Then
                  strTBD03 = "8" '¹ÎÅé¼Ð³¹
               End If
            ElseIf intCol = 14 Then strTBD12 = strData
            End If
         End If
ReadNextCol1:
      Next intCol
      dblChar = dblEnd
   Next dblChar
   
   'Add By Sindy 2017/4/21 ¸³¨Æªø­n¥Ó½Ð¤H¸ê®Æ°µ¶}©Ý¥Î
   strText = "RegContentOwner": strTitNM = "¥Ó½Ð¤H"
   dblStar = InStr(m_strTextBox, "<" & strText & ">")
   'dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
   dblLastEnd = InStr(m_strTextBox, "</RegContent>")
   intApp = 0
   For j = 1 To 10
      strTMBMApp(j) = ""
   Next j
   If dblStar > 0 Then
      For dblChar = dblStar To dblLastEnd
         strChineseNM = "": strEnglishNM = ""
         For j = 1 To 2
            strData = ""
            If j = 1 Then
               dblChar = InStr(dblChar, m_strTextBox, "<AChinese")
               strText = "AChinese": strTitNM = "°Ó¼ÐÅv¤H¤¤¤å¦WºÙ"
            ElseIf j = 2 Then
               dblChar = InStr(dblChar, m_strTextBox, "<AEnglish")
               strText = "AEnglish": strTitNM = "°Ó¼ÐÅv¤H­^¤å¦WºÙ"
            End If
            If dblChar <= 0 Then Exit For
            dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
            If dblStar < dblChar Then Exit For
            If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
            '***** ¸ÑªRXML *****
            If GetXmlData(dblChar, strText, strTitNM, False, strData, dblEnd) = False Then
            '***** End
               Exit For
            End If
            If j = 1 Then '°Ó¼ÐÅv¤H¤¤¤å¦WºÙ
               'Modify By Sindy 2017/7/3
               '©m¦W¦³³y¦r¦³¹Ï¤ù
               'strData=¸âµú<img align="absmiddle" height="18px" width="27px" file="106203003/106203003-009.TIF" alt="¨ä¥L«D¹Ï¦¡ ed10999.png" img-content="tif" orientation="portrait" inline="yes" giffile="106203003/106203003-009.png"></img>
               If InStr(strData, "<") > 0 Then
                  strData = Left(strData, InStr(strData, "<") - 1)
               End If
               '2017/7/3 END
               'Modify By Sindy 2023/8/1
'               strData = ReplaceMadeWord(strData, "?") 'Modify By Sindy 2018/5/21 ÀË¬d³y¦r
'               strChineseNM = PUB_FilterBulletinSpecWord("1", strData, strTMBM05)
               strChineseNM = strData
               '2023/8/1 END
            ElseIf j = 2 Then '°Ó¼ÐÅv¤H­^¤å¦WºÙ
               strEnglishNM = strData
            End If
            dblChar = dblEnd
         Next j
         If strChineseNM <> "" Or strEnglishNM <> "" Then
            intApp = intApp + 1
            '¸ê®Æ®w¥u¦s10¦ì¥Ó½Ð¤H
            If intApp >= 11 Then
               Exit For
            End If
            If strChineseNM <> "" Then
               strTMBMApp(intApp) = strChineseNM
            Else
               If strEnglishNM <> "" Then
                  strTMBMApp(intApp) = strEnglishNM
               End If
            End If
         Else
            Exit For
         End If
      Next dblChar
   End If
   '2017/4/21 End
   
   Set rsTmp = Nothing
End Function

Private Sub cmdTransFile_Click()
   'Modify By Sindy 2018/12/13
   If Option1(1).Value = True And txtTBD17 <> "" Then 'ºM¤T ÂàÀÉ
      Call ExeTransFile_Three
   Else
   '2018/12/13 END
      Call ExeTransFile
   End If
End Sub

'ºM¤TÂàÀÉ
Private Sub ExeTransFile_Three()
Dim strTit As String
Dim nResponse
Dim dblFCnt As Double
Dim intSeqno As Integer
Dim strTBG02 As String, strTBG03 As String, strTBG04 As String
Dim strTBG04_1 As String, strTBG04_2 As String, strTBG04_3 As String
Dim strTBG04_4 As String, strTBG04_5 As String, strTBG04_6 As String
Dim strTBOR03 As String, strTBOR04 As String, strTBOR05 As String
Dim rsTmp As New ADODB.Recordset
Dim rsTmp3 As New ADODB.Recordset 'Add By Sindy 2015/1/16
Dim strTime As String, strTotRow As String
Dim dblMaxWidth As Double
Dim strTo As String, strSubject As String, strContext As String 'Add By Sindy 2012/8/24
Dim strUpdTM29 As String  'add by sonia 2018/3/12
Dim bolChkedReadTxt1 As Boolean 'Add By Sindy 2018/12/18
Dim strTBOR09 As String 'Add By Sindy 2018/12/18

On Error GoTo ErrHand
   
   strTime = time()
   
   '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
   If TxtValidate = False Then Exit Sub
   
   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   File2.path = txtPath2.Text & "\RegContent"
   File2.Refresh
   If Val(Val(Left(File2.List(0), 3)) & Mid(File2.List(0), 5, 2)) <> Val(txtTMBM07) Then
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2.Text & "\RegContent" & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡I"
      txtPath2.SetFocus
      Exit Sub
   End If
   
'   If IsRecordExist_Three = True Then
'      strTit = "¸ß°Ý"
'      strMsg = "¤½³ø¨÷´Á" & txtTMBM07 & "¤w¦³¸ê®Æ¦s¦b¡A½T©w¬O§_­n­«·sÂàÀÉ¡H"
'      nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
'      If nResponse = vbNo Then Exit Sub
'
'      strSql = "delete from TMBulletinData where tbd16='2' and tbd01=" & CNULL(txtTMBM07)
'      cnnConnection.Execute strSql
'      strSql = "delete from TMBulletinOwner where tbor07='2' and tbor08=" & CNULL(txtTMBM07)
'      cnnConnection.Execute strSql
'      strSql = "delete from TMBulletinGoods where tbg11='2' and tbg12=" & CNULL(txtTMBM07)
'      cnnConnection.Execute strSql
'   Else
'      '¦³ºM¤T,¶}©Ý¤½³ø¦~¤ë«D·í´Á,«h¥þ³¡§R°£
'      strSql = "SELECT count(*) FROM TMBulletinData " & _
'               "WHERE tbd16='2' and tbd17<>" & Left(DBDATE(GetTA05), 6)
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         If RsTemp.Fields(0) > 0 Then
            strSql = "delete from TMBulletinData where tbd16='2'"
            cnnConnection.Execute strSql
            strSql = "delete from TMBulletinOwner where tbor07='2'"
            cnnConnection.Execute strSql
            strSql = "delete from TMBulletinGoods where tbg11='2'"
            cnnConnection.Execute strSql
'         End If
'      End If
'   End If
   
   Screen.MousePointer = vbHourglass
   
   strOurAgentName = GetTOurAgentName()
   m_PrintRpt1 = False: m_PrintRpt2 = False: m_PrintRpt3 = False: iLine = 0: iLine2 = 0
   strTotRow = File2.ListCount
   Me.Height = 6120
   dblMaxWidth = 8820
   Text2.Width = 0
   For dblFCnt = 0 To File2.ListCount - 1
      'Add by Sindy 2022/3/3
      If strSrvDate(1) >= Form20¤W½u¤é Then
         adoStream.LoadFromFile (txtPath2.Text & "\RegContent\" & File2.List(dblFCnt))
         m_strTextBox = adoStream.ReadText
      Else
      '2022/3/3 END
         RichTextBox1.LoadFile (txtPath2.Text & "\RegContent\" & File2.List(dblFCnt))
         m_strTextBox = RichTextBox1.Text
      End If
      
      dblLastEnd = InStr(m_strTextBox, "</RegContent>")
      
      Text2.Width = dblMaxWidth / Val(strTotRow) * (dblFCnt + 1): DoEvents
      
      cnnConnection.BeginTrans
      
      If ReadXmlData = False Then GoTo ErrHand
      
      'Åª¨ú°Ó¼ÐÅv¤H¸ê®Æ
      intSeqno = 0
      strTBOR03 = "": strTBOR04 = "": strTBOR05 = ""
      For dblChar = 1 To dblLastEnd
         strText = "AChinese": strTitNM = "°Ó¼ÐÅv¤H¤¤¤å"
         dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
         If dblStar < dblChar Then Exit For
         For intCol = 1 To 3
            strData = ""
            If intCol = 1 Then
               strText = "AChinese": strTitNM = "°Ó¼ÐÅv¤H¤¤¤å"
            ElseIf intCol = 2 Then strText = "AEnglish": strTitNM = "°Ó¼ÐÅv¤H­^¤å"
            ElseIf intCol = 3 Then strText = "Address": strTitNM = "°Ó¼ÐÅv¤H¦a§}"
            End If
            '***** ¸ÑªRXML *****
            If GetXmlData(dblChar, strText, strTitNM, True, strData, dblEnd) = False Then GoTo ReadNextCol2
            '***** End
            If strData <> "" Then
               If intCol = 1 Then
                  'Modify By Sindy 2023/8/1
'                  strData = ReplaceMadeWord(strData, "?") 'Modify By Sindy 2018/5/21 ÀË¬d³y¦r
'                  strTBOR03 = PUB_FilterBulletinSpecWord("1", strData, strTMBM05)
                  strTBOR03 = strData
                  '2023/8/1 EMD
               ElseIf intCol = 2 Then strTBOR04 = strData
               ElseIf intCol = 3 Then
                  strTBOR05 = strData
               End If
            End If
ReadNextCol2:
         Next intCol
         dblChar = dblEnd
         '§Ç¸¹
         intSeqno = intSeqno + 1
         '·s¼WTable
         strErrTxt = "°Ó¼ÐÅv¤H.TMBulletinOwner"
         '«D¥»©Ò®×¥ó¥B°Ó¼ÐÅv¤H¬°¥xÆWªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
'         If bolTaieCase = False And bolIsTaiwanCase = True Then
         '°Ó¼ÐÅv¤H¬°¥xÆW®×ªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
         'Modify By Sindy 2012/1/5 +¤j³°®×
         'If bolIsTaiwanCase = True Or bolIsChinaCase = True Then
         'Modify By Sindy 2018/12/25 ºM¤T¥u§ì¥xÆW®×«D¥»©Ò®×¥ó
         If bolIsTaiwanCase = True And bolTaieCase = False Then
            If strTBOR05 <> "" Then
               strTBOR09 = Left(PUB_AddrChangeZIPCode(strTBOR05), 3)
            Else
               strTBOR09 = ""
            End If
            'Modify By Sindy 2018/12/10 + ,TBOR07,TBOR08,TBOR09
            strSql = "insert into TMBulletinOwner (TBOR01,TBOR02,TBOR03,TBOR04,TBOR05,TBOR06" & _
                     ",TBOR07,TBOR08,TBOR09) " & _
                     "values(" & CNULL(strTBD02) & "," & intSeqno & "," & CNULL(strTBOR03) & _
                     "," & CNULL(strTBOR04) & "," & CNULL(strTBOR05) & "," & CNULL(strTBD03) & _
                     ",'2'," & CNULL(txtTMBM07) & "," & CNULL(strTBOR09) & ")"
            cnnConnection.Execute strSql
         End If
         
'         'If intSeqno > 1 Then 'Modify By Sindy 2017/9/18 Mark
'            '«D¥»©Ò®×¥ó¥B¬°¥xÆW®×ªº°Ó¼ÐÅv¤H¤¤¤å,¦a§}¦³?®É,»Ý¦C¦L²M³æ
''            If bolTaieCase = False And bolIsTaiwanCase = True And _
''               (InStr(strTBOR03, "?") > 0 Or InStr(strTBOR05, "?") > 0) Then
'            '¬°¥xÆW®×ªº°Ó¼ÐÅv¤H¤¤¤å,¦a§}¦³?®É,»Ý¦C¦L²M³æ
'            'Modify By Sindy 2012/1/5 +¤j³°®×
'            'If (bolIsTaiwanCase = True Or bolIsChinaCase = True) And
'            'Modify By Sindy 2017/6/16 ¨ú®ø¤j³°®×¸ê®Æ±¾¦b¶}©ÝÀË®Öªí¸Ì,§ï¥u©ñ¸ê®ÆÀË®Öªí
'            If bolIsTaiwanCase = True And _
'               (InStr(strTBOR03, "?") > 0 Or InStr(strTBOR05, "?") > 0) Then
'               Call ReadTxt2(strTBD02, strTBD04, strTMBM05, strTMBM06, strTBOR03, strTBOR05, strTBD03)
'            End If
'         'End If
      Next dblChar
      
'      '¦a°Ï¦WºÙ¬°ªÅ¥Õ©Î¤¤°ê¤j³°,¥N²z¤H¦WºÙ¦³?,°Ó¼ÐºØÃþ«D1,7,8,9®É,»Ý¦C¦L²M³æ
'      'Modify By Sindy 2015/10/16 +Or strTMBM05 = "¤¤µØ¥Á°ê" Or strTMBM05 = "¥xÆW"
'      bolChkedReadTxt1 = False
'      If strTMBM05 = "" Or _
'         strTMBM05 = "¤¤°ê¤j³°" Or strTMBM05 = "¤¤µØ¥Á°ê" Or strTMBM05 = "¥xÆW" Or _
'         InStr(strTMBM06, "?") > 0 Or _
'         (strTBD03 <> "1" And strTBD03 <> "7" And strTBD03 <> "8" And strTBD03 <> "9") Then
'         Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strAChinese1, strAddress1, strTBD03)
'         bolChkedReadTxt1 = True
'      End If
'      'Add By Sindy 2017/4/26 «D¥xÆW®×ªº°Ó¼ÐÅv¤H¤¤¤å¦³?®É,»Ý¦C¦L²M³æ
'      'If bolIsChinaCase = False And intApp > 0 Then
'      '¤£¤À°êÄy
'      'Modify By Sindy 2017/6/16 ¥u¼g¤J«D¥xÆWªº¥Ó½Ð¤H¸ê®Æ
'      If intApp > 0 And bolIsTaiwanCase = False Then
'         If InStr(strTMBMApp(1), "?") > 0 Or InStr(Left(strTMBMApp(1), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(1), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(1), 5), "Äy") > 0 Then
'            If bolChkedReadTxt1 = False Then '¥H§K­«ÂÐÀË¬d
'               Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(1), strAddress1, strTBD03)
'            End If
'         End If
'         If InStr(strTMBMApp(2), "?") > 0 Or InStr(Left(strTMBMApp(2), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(2), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(2), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(2), "", strTBD03)
'         If InStr(strTMBMApp(3), "?") > 0 Or InStr(Left(strTMBMApp(3), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(3), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(3), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(3), "", strTBD03)
'         If InStr(strTMBMApp(4), "?") > 0 Or InStr(Left(strTMBMApp(4), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(4), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(4), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(4), "", strTBD03)
'         If InStr(strTMBMApp(5), "?") > 0 Or InStr(Left(strTMBMApp(5), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(5), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(5), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(5), "", strTBD03)
'         If InStr(strTMBMApp(6), "?") > 0 Or InStr(Left(strTMBMApp(6), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(6), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(6), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(6), "", strTBD03)
'         If InStr(strTMBMApp(7), "?") > 0 Or InStr(Left(strTMBMApp(7), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(7), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(7), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(7), "", strTBD03)
'         If InStr(strTMBMApp(8), "?") > 0 Or InStr(Left(strTMBMApp(8), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(8), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(8), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(8), "", strTBD03)
'         If InStr(strTMBMApp(9), "?") > 0 Or InStr(Left(strTMBMApp(9), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(9), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(9), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(9), "", strTBD03)
'         If InStr(strTMBMApp(10), "?") > 0 Or InStr(Left(strTMBMApp(10), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(10), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(10), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(10), "", strTBD03)
'      End If
'      '2017/4/26 END
      
      'Åª¨ú°Ó«~¸ê®Æ
      strTBG02 = "": strTBG03 = "": strTBG04 = ""
      For dblChar = 1 To dblLastEnd
         strText = "Class": strTitNM = "°Ó«~Ãþ§O"
         dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
         If dblStar < dblChar Then Exit For
         For intCol = 1 To 3
            strData = ""
            If intCol = 1 Then
               strText = "Class": strTitNM = "°Ó«~Ãþ§O"
            ElseIf intCol = 2 Then strText = "Enforcement_Rules": strTitNM = "°Ó«~©ÎªA°ÈÃþ§O"
            ElseIf intCol = 3 Then strText = "Goods_Denomination": strTitNM = "°Ó«~¦WºÙ"
            End If
            '***** ¸ÑªRXML *****
            If GetXmlData(dblChar, strText, strTitNM, True, strData, dblEnd) = False Then GoTo ReadNextCol3
            '***** End
            If strData <> "" Then
               If intCol = 1 Then
                  If Val(strData) <= 0 Then
                     strData = ""
                  Else
                     strData = Val(strData)
                  End If
                  If Len(strData) = 1 Then strData = "0" & strData
                  strTBG02 = strData
                  strTMBM08 = strTMBM08 & strData & ","
               ElseIf intCol = 2 Then strTBG03 = strData
               ElseIf intCol = 3 Then strTBG04 = strData
               End If
            End If
ReadNextCol3:
         Next intCol
         
         dblChar = dblEnd
         'ºI¨ú°Ó«~¦WºÙ
         strTBG04_1 = "": strTBG04_2 = "": strTBG04_3 = ""
         strTBG04_4 = "": strTBG04_5 = "": strTBG04_6 = ""
         If Len(strTBG04) < 2000 Then
            strTBG04_1 = strTBG04
         ElseIf Len(strTBG04) > 2000 Then
            strTBG04_1 = Mid(strTBG04, 1, 2000)
            If Len(strTBG04) < 4000 Then
               strTBG04_2 = Mid(strTBG04, 2001, Len(strTBG04))
            ElseIf Len(strTBG04) > 4000 Then
               strTBG04_2 = Mid(strTBG04, 2001, 2000)
               If Len(strTBG04) < 6000 Then
                  strTBG04_3 = Mid(strTBG04, 4001, Len(strTBG04))
               ElseIf Len(strTBG04) > 6000 Then
                  strTBG04_3 = Mid(strTBG04, 4001, 2000)
                  If Len(strTBG04) < 8000 Then
                     strTBG04_4 = Mid(strTBG04, 6001, Len(strTBG04))
                  ElseIf Len(strTBG04) > 8000 Then
                     strTBG04_4 = Mid(strTBG04, 6001, 2000)
                     If Len(strTBG04) < 10000 Then
                        strTBG04_5 = Mid(strTBG04, 8001, Len(strTBG04))
                     ElseIf Len(strTBG04) > 10000 Then
                        strTBG04_5 = Mid(strTBG04, 8001, 2000)
                        If Len(strTBG04) < 12000 Then
                           strTBG04_6 = Mid(strTBG04, 10001, Len(strTBG04))
                        ElseIf Len(strTBG04) > 12000 Then
                           strTBG04_6 = Mid(strTBG04, 10001, 2000)
                        End If
                     End If
                  End If
               End If
            End If
         End If
         strErrTxt = "°Ó«~ÀÉ.TMBulletinGoods"
         '·s¼WTable
         '«D¥»©Ò®×¥ó¥B°Ó«~¸ê®Æ¬°¥xÆW®×ªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
'         If bolTaieCase = False And bolIsTaiwanCase = True Then
         '°Ó«~¸ê®Æ¬°¥xÆW®×ªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
         'Modify By Sindy 2012/1/5 +¤j³°®×
         'If bolIsTaiwanCase = True Or bolIsChinaCase = True Then
         'Modify By Sindy 2018/12/25 ºM¤T¥u§ì¥xÆW®×«D¥»©Ò®×¥ó
         If bolIsTaiwanCase = True And bolTaieCase = False Then
            'If strTBG02 <> "" Then
               If strTBG02 = "" Then strTBG02 = " " '¼Ð³¹¨S¦³°Ó«~Ãþ§O,¦ý¦³¼Ð³¹¤º®e
               'Modify By Sindy 2018/12/10 + ,TBG11,TBG12
               strSql = "insert into TMBulletinGoods (TBG01,TBG02,TBG03,TBG04,TBG05,TBG06,TBG07" & _
                        ",TBG08,TBG09,TBG10,TBG11,TBG12) " & _
                        "values(" & CNULL(strTBD02) & ",'" & strTBG02 & "'," & CNULL(strTBG03) & _
                        "," & CNULL(strTBG04_1) & "," & CNULL(strTBG04_2) & "," & CNULL(strTBG04_3) & _
                        "," & CNULL(strTBD03) & "," & CNULL(strTBG04_4) & "," & CNULL(strTBG04_5) & _
                        "," & CNULL(strTBG04_6) & ",'2'," & CNULL(txtTMBM07) & ")"
               cnnConnection.Execute strSql
               'Add By Sindy 2014/12/1
               If Trim(strTBG02) = "" Then '¼Ð³¹¨S¦³°Ó«~Ãþ§O,¦ý¦³¼Ð³¹¤º®e
                  Exit For
               End If
               '2014/12/1 END
            'End If
         End If
      Next dblChar
      If strTMBM08 <> "" Then strTMBM08 = Left(strTMBM08, Len(strTMBM08) - 1)
      
'      'Add By Sindy 2015/10/27 ¤ñ¹ï°Ó«~Ãþ§O¸ê®Æ
'      If bolTaieCase = True Then
'         If Trim(strTMBM08) <> Trim(strTM09) Then
'            strMsg = strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04 & "¤½³ø°Ó«~Ãþ§O¬°:" & strTMBM08 & " ¥»©Ò°Ó«~Ãþ§O¬°:" & strTM09
'            Call ReadTxt1(strTBD02, strTBD04, strMsg, "", "", "", strTBD03)
'         End If
'      End If
'      '2015/10/27 End
      
'      'Add By Sindy 2017/4/25 ¥Ó½Ð¤H1~10
'      If intApp > 0 Then
'         strSql = "update TMBulletin set" & _
'                  " TMBM09=" & CNULL(strTMBMApp(1)) & _
'                  ",TMBM10=" & CNULL(strTMBMApp(2)) & _
'                  ",TMBM11=" & CNULL(strTMBMApp(3)) & _
'                  ",TMBM12=" & CNULL(strTMBMApp(4)) & _
'                  ",TMBM13=" & CNULL(strTMBMApp(5)) & _
'                  ",TMBM14=" & CNULL(strTMBMApp(6)) & _
'                  ",TMBM15=" & CNULL(strTMBMApp(7)) & _
'                  ",TMBM16=" & CNULL(strTMBMApp(8)) & _
'                  ",TMBM17=" & CNULL(strTMBMApp(9)) & _
'                  ",TMBM18=" & CNULL(strTMBMApp(10)) & _
'                  " where TMBM01='" & strTBD02 & "' and TMBM02='" & strTBD03 & "'"
'         cnnConnection.Execute strSql
'      End If
'      '2017/4/25 END
      
      '«D¥»©Ò®×¥ó¥B¤½³ø¸ê®Æ¬°¥xÆW®×ªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
'      If bolTaieCase = False And bolIsTaiwanCase = True Then
      '¤½³ø¸ê®Æ¬°¥xÆW®×ªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
      'Modify By Sindy 2012/1/5 +¤j³°®×
      'If bolIsTaiwanCase = True Or bolIsChinaCase = True Then
      'Modify By Sindy 2018/12/25 ºM¤T¥u§ì¥xÆW®×«D¥»©Ò®×¥ó
      If bolIsTaiwanCase = True And bolTaieCase = False Then
         strErrTxt = "°Ó¼Ð¤½³ø«D¥»©Ò¸ê®ÆÀÉ.TMBulletinData"
         If bolIsTaiwanCase = True Then
            strTBD15 = "A"
         ElseIf bolIsChinaCase = True Then
            strTBD15 = "B"
         End If
         'Modify By Sindy 2018/12/10 + ,TBD16,TBD17
         strSql = "insert into TMBulletinData (TBD01,TBD02,TBD03,TBD04,TBD05,TBD06,TBD07,TBD08" & _
                  ",TBD09,TBD10,TBD11,TBD12,TBD13,TBD14,TBD15,TBD16,TBD17) " & _
                  "values(" & CNULL(txtTMBM07) & "," & CNULL(strTBD02) & "," & CNULL(strTBD03) & _
                  "," & CNULL(strTBD04) & "," & CNULL(strTBD05) & "," & CNULL(strTBD06) & _
                  "," & CNULL(strTBD07) & "," & CNULL(strTBD08) & "," & CNULL(strTBD09) & _
                  "," & CNULL(strTBD10) & "," & CNULL(strTBD11) & "," & CNULL(strTBD12) & _
                  "," & CNULL(strTBD13) & "," & CNULL(strTBD14) & "," & CNULL(strTBD15) & _
                  ",'2'," & Val(txtTBD17) + 191100 & ")"
         cnnConnection.Execute strSql
'         '¸ê®Æ¸Ì¦³?®É,»Ý¦C¦L²M³æ¤G
'         If m_bolCharQ = True Then
'            For i = 1 To 6
'               strTemp(i) = ""
'            Next i
'            strTemp(1) = strTBD02
'            strTemp(2) = strTBD04
'            strTemp(3) = GetTradeMarkName(strTBD03, 0)
'            strTemp(4) = m_strCharQNote
'            If iLine2 > 54 Or iLine2 = 0 Then
'               If iLine2 > 0 Then Printer.NewPage
'               PrintTitle2 '¦C¦LªíÀY
'            End If
'            PrintDetail2 '¦C¦L©ú²Ó
'         End If
      End If
      
      cnnConnection.CommitTrans
   Next dblFCnt
   
   '±N¬°ª§Ä³¹ï³y®×¥óªº°Ó¼ÐÅv¤H¸ê®Æ§¡¤@¨Ö³]¬°N¤£¦C¦L¶}©Ý¨ç
   cnnConnection.BeginTrans
   strSql = "SELECT distinct tbor03 FROM TMBulletinData,TMBulletinOwner " & _
            "WHERE tbor01=tbd02 and tbor06=tbd03 and tbd14='N' and tbd01='" & txtTMBM07 & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strSql = "update TMBulletinData " & _
                  "set TBD14='N' " & _
                  "WHERE tbd02||tbd03 in (SELECT tbor01||tbor06 FROM TMBulletinOwner WHERE tbor03='" & RsTemp.Fields(0) & "') "
         cnnConnection.Execute strSql
         RsTemp.MoveNext
      Loop
   End If
   cnnConnection.CommitTrans
   
   'If m_PrintRpt = True Then Printer.EndDoc
   strMsg = ""
   If m_PrintRpt1 = True Then
      'Close ff1
      'Add By Sindy 2024/5/17
      If Dir(PUB_Getdesktop & "\" & m_strFileName1) <> "" Then
         Kill PUB_Getdesktop & "\" & m_strFileName1
         Sleep 100
      End If
      Call PUB_SaveTextAsUTF8(PUB_Getdesktop & "\" & m_strFileName1, m_strText)
      '2024/5/17 END
'      strMsg = m_strFileName1
   End If
'   If m_PrintRpt2 = True Then
'      Close FF2
'      If strMsg <> "" Then strMsg = strMsg & " ¤Î "
'      strMsg = strMsg & m_strFileName2
'   End If
   If m_PrintRpt3 = True Then
      'Add By Sindy 2016/2/17
      Print #ff3, "¥H¤W¸ê®Æ½Ð³qª¾¹q¸£¤¤¤ß¨ó§UÂà¤J¡A¨ä¥L¸ê®Æ¤w¶×¤J§¹²¦¡I"
      '2016/2/17 End
      Close ff3
      If strMsg <> "" Then strMsg = strMsg & " ¤Î "
      strMsg = strMsg & m_strFileName3
   End If
   If m_PrintRpt1 = True Or m_PrintRpt2 = True Or m_PrintRpt3 = True Then
      'MsgBox "½Ð¦Ü¤U¦C¦ì¸m¦C¦LÀË®Öªí¡G" & PUB_Getdesktop & "\" & strMsg
      strMsg = "½Ð¦Ü¤U¦C¦ì¸m¦C¦LÀË®Öªí¡G" & PUB_Getdesktop & "\" & strMsg
   End If
   
   Screen.MousePointer = vbDefault
   Call IsRecordExist_Three '²£¥Íµ§¼Æ
   
   'Add By Sindy 2015/5/13 ³qª¾µ{§Ç¤wÂàÀÉ§¹²¦
   If strP22 <> "" Then
      strSubject = "°Ó¼Ð¤½³ø¡]ºM¤T¡^¤wÂàÀÉ§¹²¦¡I"
      PUB_SendMail strUserNum, strP22, "", strSubject, strSubject, , , , , , , , , , , False
   End If
   '2015/5/13 END
   
   MsgBox "ÂàÀÉ§¹²¦¡I(ÂàÀÉªá¶O®É¶¡¡G" & strTime & "  " & time() & ")" & vbCrLf & strMsg
   Me.Height = 5000
   
   Set rsTmp = Nothing
   Set rsTmp3 = Nothing
   Exit Sub
   
ErrHand:
   If Err.Number = -2147217900 Then 'ORA-00917: ¿òº|³rÂI
      '¼gLog
      Call ReadTxt3(strSql)
      '±µµÛµo¥Í¿ù»~³¯­z¦¡ªº¤U­Ó³¯­z¦¡¶}©l°õ¦æ
      Resume Next
   End If
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   
   If Err.Number = 76 Then
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2.Text & "\RegContent" & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡I"
      txtPath2.SetFocus
   Else
      cnnConnection.RollbackTrans
      If Err.Number = -2147217873 Then
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¼f©w¸¹¼Æ¡]" & strTBD02 & "¡^°Ó¼ÐºØÃþ¡]" & strTBD03 & "¡^" & vbCrLf & strErrTxt & ": ¹H¤Ï¥²¶·¬°°ß¤@ªº­­¨î±ø¥ó"
      Else
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¼f©w¸¹¼Æ¡]" & strTBD02 & "¡^°Ó¼ÐºØÃþ¡]" & strTBD03 & "¡^" & vbCrLf & strErrTxt & Err.Description & vbCrLf & strSql
      End If
      'Add By Sindy 2015/5/13 ³qª¾µ{§ÇÂàÀÉ¦³»~
      If strP22 <> "" Then
         strSubject = "°Ó¼Ð¤½³ø¡]ºM¤T¡^ÂàÀÉ¦³»~¡I"
         PUB_SendMail strUserNum, strP22, "", strSubject, strSubject, , , , , , , , , , , False
      End If
      '2015/5/13 END
   End If
End Sub

'¤½³øÂàÀÉ
Private Sub ExeTransFile()
Dim strTit As String
Dim nResponse
Dim dblFCnt As Double
Dim intSeqno As Integer
Dim strTBG02 As String, strTBG03 As String, strTBG04 As String
Dim strTBG04_1 As String, strTBG04_2 As String, strTBG04_3 As String
Dim strTBG04_4 As String, strTBG04_5 As String, strTBG04_6 As String
Dim strTBOR03 As String, strTBOR04 As String, strTBOR05 As String
Dim rsTmp As New ADODB.Recordset
Dim rsTmp3 As New ADODB.Recordset 'Add By Sindy 2015/1/16
Dim strTime As String, strTotRow As String
Dim dblMaxWidth As Double
Dim strTo As String, strSubject As String, strContext As String 'Add By Sindy 2012/8/24
Dim strUpdTM29 As String  'add by sonia 2018/3/12
Dim bolChkedReadTxt1 As Boolean
Dim strTA05 As String
   
On Error GoTo ErrHand
   
   strTime = time()
   strTA05 = DBDATE(GetTA05)
   
   '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
   If TxtValidate = False Then Exit Sub
   
   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   File2.path = txtPath2.Text & "\RegContent"
   File2.Refresh
   If Val(Val(Left(File2.List(0), 3)) & Mid(File2.List(0), 5, 2)) <> Val(txtTMBM07) Then
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2.Text & "\RegContent" & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡I"
      txtPath2.SetFocus
      Exit Sub
   End If
   
   If IsRecordExist = True Then
      strTit = "¸ß°Ý"
      strMsg = "¤½³ø¨÷´Á" & txtTMBM07 & "¤w¦³¸ê®Æ¦s¦b¡A½T©w¬O§_­n­«·sÂàÀÉ¡H"
      nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
      If nResponse = vbNo Then Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   strSql = "delete from TMBulletin where TMBM07='" & txtTMBM07 & "'"
   cnnConnection.Execute strSql

   strSql = "delete from TMBulletinData where tbd16='1'"
   cnnConnection.Execute strSql
   strSql = "delete from TMBulletinOwner where tbor07='1'"
   cnnConnection.Execute strSql
   strSql = "delete from TMBulletinGoods where tbg11='1'"
   cnnConnection.Execute strSql
   
   strOurAgentName = GetTOurAgentName()
   m_PrintRpt1 = False: m_PrintRpt2 = False: m_PrintRpt3 = False: iLine = 0: iLine2 = 0
   strTotRow = File2.ListCount
   Me.Height = 6120
   dblMaxWidth = 8820
   Text2.Width = 0
   For dblFCnt = 0 To File2.ListCount - 1
      strExc(10) = File2.List(dblFCnt)
      
'      If dblFCnt = 0 Then
'         strExc(10) = "052003_RegContent0_02434451_1.xml"
'      ElseIf dblFCnt = 1 Then
'         strExc(10) = "052003_RegContent0_02434885_1.xml"
'      ElseIf dblFCnt = 2 Then
'         strExc(10) = "052003_RegContent0_02434452_1.xml"
'      End If
      
      'Add by Sindy 2022/3/3
      If strSrvDate(1) >= Form20¤W½u¤é Then
         adoStream.LoadFromFile (txtPath2.Text & "\RegContent\" & strExc(10))
         m_strTextBox = adoStream.ReadText
      Else
      '2022/3/3 END
         RichTextBox1.LoadFile (txtPath2.Text & "\RegContent\" & strExc(10))
         m_strTextBox = RichTextBox1.Text
      End If
           
      dblLastEnd = InStr(m_strTextBox, "</RegContent>")
      
      Text2.Width = dblMaxWidth / Val(strTotRow) * (dblFCnt + 1): DoEvents
      
      cnnConnection.BeginTrans
      
      'Modify By Sindy 2017/4/21 ¿W¥ß¥X¨Ó¤@­Ó¨ç¼Æ
      If ReadXmlData = False Then GoTo ErrHand
      '2017/4/21 END
      If ChkDataErr(strTBD02, strTBD03, strTBD04) = True Then GoTo ErrHand
      
      '«D¥»©Ò®×¥ó¥B¬°¥xÆW®×ªº°Ó¼ÐÅv¤H¤¤¤å,¦a§}¦³?®É,»Ý¦C¦L²M³æ
'      If bolTaieCase = False And bolIsTaiwanCase = True And _
'         (InStr(strAChinese1, "?") > 0 Or InStr(strAddress1, "?") > 0) Then
      '¬°¥xÆW®×ªº°Ó¼ÐÅv¤H¤¤¤å,¦a§}¦³?®É,»Ý¦C¦L²M³æ
      'Modify By Sindy 2012/1/5 +¤j³°®×
      'If (bolIsTaiwanCase = True Or bolIsChinaCase = True) And
      'Modify By Sindy 2017/6/16 ¨ú®ø¤j³°®×¸ê®Æ±¾¦b¶}©ÝÀË®Öªí¸Ì,§ï¥u©ñ¸ê®ÆÀË®Öªí
      
      'Modify By Sindy 2017/9/18 ²Î¤@¦b¤U­±(Åª¨ú°Ó¼ÐÅv¤H¸ê®Æ)ÀË®Ö
'      If bolIsTaiwanCase = True And _
'         (InStr(strAChinese1, "?") > 0 Or InStr(strAddress1, "?") > 0) Then
'         '(InStr(strAChinese1, "?") > 0 Or InStr(strAddress1, "?") > 0) Then
'         Call ReadTxt2(strTBD02, strTBD04, strTMBM05, strTMBM06, strAChinese1, strAddress1, strTBD03)
'      End If

'      If strTMBM05 = "" Or _
'         strTMBM05 = "¤¤°ê¤j³°" Or _
'         InStr(strTMBM06, "?") > 0 Or _
'         (bolTaieCase = False And bolIsTaiwanCase = True And _
'         (InStr(strAChinese1, "?") > 0 Or InStr(strAddress1, "?") > 0)) Then
'         Call ReadTxt(strTBD02, strTBD04, strTMBM05, strTMBM06, strAChinese1, strAddress1, bolTaieCase, bolIsTaiwanCase)
'         'ª½±µ¦C¦L³øªí
'         For i = 1 To 6
'            strTemp(i) = ""
'         Next i
'         strTemp(1) = strTBD02
'         strTemp(2) = strTBD04
'         strTemp(3) = strTMBM05
'         strTemp(4) = strTMBM06
'         strTemp(5) = strAChinese1
'         strTemp(6) = strAddress1
'         If iLine > 37 Or iLine = 0 Then
'            If iLine > 0 Then Printer.NewPage
'            PrintTitle '¦C¦LªíÀY
'         End If
'         PrintDetail '¦C¦L©ú²Ó
'      End If
      'Åª¨ú°Ó¼ÐÅv¤H¸ê®Æ
      intSeqno = 0
      strTBOR03 = "": strTBOR04 = "": strTBOR05 = ""
      For dblChar = 1 To dblLastEnd
         strText = "AChinese": strTitNM = "°Ó¼ÐÅv¤H¤¤¤å"
         dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
         If dblStar < dblChar Then Exit For
         For intCol = 1 To 3
            strData = ""
            If intCol = 1 Then
               strText = "AChinese": strTitNM = "°Ó¼ÐÅv¤H¤¤¤å"
            ElseIf intCol = 2 Then strText = "AEnglish": strTitNM = "°Ó¼ÐÅv¤H­^¤å"
            ElseIf intCol = 3 Then strText = "Address": strTitNM = "°Ó¼ÐÅv¤H¦a§}"
            End If
            '***** ¸ÑªRXML *****
            If GetXmlData(dblChar, strText, strTitNM, True, strData, dblEnd) = False Then GoTo ReadNextCol2
            '***** End
            If strData <> "" Then
               If intCol = 1 Then
                  'Modify By Sindy 2023/8/1
'                  strData = ReplaceMadeWord(strData, "?") 'Modify By Sindy 2018/5/21 ÀË¬d³y¦r
'                  strTBOR03 = PUB_FilterBulletinSpecWord("1", strData, strTMBM05)
                  strTBOR03 = strData
                  '2023/8/1 END
               ElseIf intCol = 2 Then strTBOR04 = strData
               ElseIf intCol = 3 Then
                  strTBOR05 = strData
               End If
            End If
ReadNextCol2:
         Next intCol
         dblChar = dblEnd
         '§Ç¸¹
         intSeqno = intSeqno + 1
         '·s¼WTable
         strErrTxt = "°Ó¼ÐÅv¤H.TMBulletinOwner"
         '«D¥»©Ò®×¥ó¥B°Ó¼ÐÅv¤H¬°¥xÆWªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
'         If bolTaieCase = False And bolIsTaiwanCase = True Then
         '°Ó¼ÐÅv¤H¬°¥xÆW®×ªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
         'Modify By Sindy 2012/1/5 +¤j³°®×
         'Modify By Sindy 2018/12/25 ¥xÆW®×­­¨î¬°¥»©Ò®×¥ó
         If (bolIsTaiwanCase = True And bolTaieCase = True) Or bolIsChinaCase = True Then
            'Modify By Sindy 2018/12/10 + ,TBOR07,TBOR08
            strSql = "insert into TMBulletinOwner (TBOR01,TBOR02,TBOR03,TBOR04,TBOR05,TBOR06" & _
                     ",TBOR07,TBOR08) " & _
                     "values(" & CNULL(strTBD02) & "," & intSeqno & "," & CNULL(strTBOR03) & _
                     "," & CNULL(strTBOR04) & "," & CNULL(strTBOR05) & "," & CNULL(strTBD03) & _
                     ",'1'," & CNULL(txtTMBM07) & ")"
            cnnConnection.Execute strSql
         End If
         'If intSeqno > 1 Then 'Modify By Sindy 2017/9/18 Mark
            '«D¥»©Ò®×¥ó¥B¬°¥xÆW®×ªº°Ó¼ÐÅv¤H¤¤¤å,¦a§}¦³?®É,»Ý¦C¦L²M³æ
'            If bolTaieCase = False And bolIsTaiwanCase = True And _
'               (InStr(strTBOR03, "?") > 0 Or InStr(strTBOR05, "?") > 0) Then
            '¬°¥xÆW®×ªº°Ó¼ÐÅv¤H¤¤¤å,¦a§}¦³?®É,»Ý¦C¦L²M³æ
            'Modify By Sindy 2012/1/5 +¤j³°®×
            'If (bolIsTaiwanCase = True Or bolIsChinaCase = True) And
            'Modify By Sindy 2017/6/16 ¨ú®ø¤j³°®×¸ê®Æ±¾¦b¶}©ÝÀË®Öªí¸Ì,§ï¥u©ñ¸ê®ÆÀË®Öªí
            'Modify By Sindy 2019/2/18 ¥xÆW®×­­¨î¥u¬°¥»©Ò®×¥ó + And bolTaieCase = True)
            'If bolIsTaiwanCase = True And
            If (bolIsTaiwanCase = True And bolTaieCase = True) And _
               (InStr(strTBOR03, "?") > 0 Or InStr(strTBOR05, "?") > 0) Then
               'Add By Sindy 2024/5/17 §ïÀË¬d¸ê®Æ®wÄæ¦ì­È;¦]¤w¥i¥H¦s¸U°ê½X
               strSql = "SELECT * FROM TMBulletinOwner" & _
                        " WHERE tbor01=" & CNULL(strTBD02) & _
                        " and tbor02=" & intSeqno & _
                        " and tbor06='1'" & _
                        " and (instr(tbor03,'?')>0 or instr(tbor05,'?')>0)"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
               '2024/5/17 END
                  Call ReadTxt2(strTBD02, strTBD04, strTMBM05, strTMBM06, strTBOR03, strTBOR05, strTBD03)
               End If
            End If
         'End If
      Next dblChar
      '¦a°Ï¦WºÙ¬°ªÅ¥Õ©Î¤¤°ê¤j³°,¥N²z¤H¦WºÙ¦³?,°Ó¼ÐºØÃþ«D1,7,8,9®É,»Ý¦C¦L²M³æ
      'Modify By Sindy 2015/10/16 +Or strTMBM05 = "¤¤µØ¥Á°ê" Or strTMBM05 = "¥xÆW"
      bolChkedReadTxt1 = False
      txtChkWord = strTMBM06
      If strTMBM05 = "" Or _
         strTMBM05 = "¤¤°ê¤j³°" Or strTMBM05 = "¤¤µØ¥Á°ê" Or strTMBM05 = "¥xÆW" Or _
         InStr(txtChkWord, "?") > 0 Or _
         (strTBD03 <> "1" And strTBD03 <> "7" And strTBD03 <> "8" And strTBD03 <> "9") Then
         Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strAChinese1, strAddress1, strTBD03)
         bolChkedReadTxt1 = True
      End If
      'Add By Sindy 2017/4/26 «D¥xÆW®×ªº°Ó¼ÐÅv¤H¤¤¤å¦³?®É,»Ý¦C¦L²M³æ
      'If bolIsChinaCase = False And intApp > 0 Then
      '¤£¤À°êÄy
      'Modify By Sindy 2017/6/16 ¥u¼g¤J«D¥xÆWªº¥Ó½Ð¤H¸ê®Æ
      'Modify By Sindy 2019/3/4 ½Õ¾ãÀË¬d¥Ó½Ð¤H¸ê®Æªº½d³ò±ø¥ó
      'If intApp > 0 And bolIsTaiwanCase = False Then
      If intApp > 0 And Not (bolIsTaiwanCase = True And bolTaieCase = True) Then
         txtChkWord = strTMBMApp(1)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(1), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(1), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(1), 5), "Äy") > 0 Then
            If bolChkedReadTxt1 = False Then '¥H§K­«ÂÐÀË¬d
               Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(1), strAddress1, strTBD03)
            End If
         End If
         txtChkWord = strTMBMApp(2)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(2), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(2), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(2), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(2), "", strTBD03)
         txtChkWord = strTMBMApp(3)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(3), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(3), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(3), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(3), "", strTBD03)
         txtChkWord = strTMBMApp(4)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(4), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(4), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(4), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(4), "", strTBD03)
         txtChkWord = strTMBMApp(5)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(5), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(5), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(5), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(5), "", strTBD03)
         txtChkWord = strTMBMApp(6)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(6), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(6), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(6), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(6), "", strTBD03)
         txtChkWord = strTMBMApp(7)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(7), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(7), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(7), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(7), "", strTBD03)
         txtChkWord = strTMBMApp(8)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(8), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(8), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(8), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(8), "", strTBD03)
         txtChkWord = strTMBMApp(9)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(9), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(9), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(9), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(9), "", strTBD03)
         txtChkWord = strTMBMApp(10)
         If InStr(txtChkWord, "?") > 0 Or InStr(Left(strTMBMApp(10), 5), "°Ó") > 0 Or InStr(Left(strTMBMApp(10), 5), "°Ï") > 0 Or InStr(Left(strTMBMApp(10), 5), "Äy") > 0 Then Call ReadTxt1(strTBD02, strTBD04, strTMBM05, strTMBM06, strTMBMApp(10), "", strTBD03)
      End If
      '2017/4/26 END
      
      'Åª¨ú°Ó«~¸ê®Æ
      strTBG02 = "": strTBG03 = "": strTBG04 = ""
      For dblChar = 1 To dblLastEnd
         strText = "Class": strTitNM = "°Ó«~Ãþ§O"
         dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
         If dblStar < dblChar Then Exit For
         For intCol = 1 To 3
            strData = ""
            If intCol = 1 Then
               strText = "Class": strTitNM = "°Ó«~Ãþ§O"
            ElseIf intCol = 2 Then strText = "Enforcement_Rules": strTitNM = "°Ó«~©ÎªA°ÈÃþ§O"
            ElseIf intCol = 3 Then strText = "Goods_Denomination": strTitNM = "°Ó«~¦WºÙ"
            End If
            '***** ¸ÑªRXML *****
            If GetXmlData(dblChar, strText, strTitNM, True, strData, dblEnd) = False Then GoTo ReadNextCol3
            '***** End
            If strData <> "" Then
               If intCol = 1 Then
                  If Val(strData) <= 0 Then
                     strData = ""
                  Else
                     strData = Val(strData)
                  End If
                  If Len(strData) = 1 Then strData = "0" & strData
                  strTBG02 = strData
                  strTMBM08 = strTMBM08 & strData & ","
               ElseIf intCol = 2 Then strTBG03 = strData
               ElseIf intCol = 3 Then strTBG04 = strData
               End If
            End If
ReadNextCol3:
         Next intCol
         
         dblChar = dblEnd
         'ºI¨ú°Ó«~¦WºÙ
         strTBG04_1 = "": strTBG04_2 = "": strTBG04_3 = ""
         strTBG04_4 = "": strTBG04_5 = "": strTBG04_6 = ""
         If Len(strTBG04) < 2000 Then
            strTBG04_1 = strTBG04
         ElseIf Len(strTBG04) > 2000 Then
            strTBG04_1 = Mid(strTBG04, 1, 2000)
            If Len(strTBG04) < 4000 Then
               strTBG04_2 = Mid(strTBG04, 2001, Len(strTBG04))
            ElseIf Len(strTBG04) > 4000 Then
               strTBG04_2 = Mid(strTBG04, 2001, 2000)
               If Len(strTBG04) < 6000 Then
                  strTBG04_3 = Mid(strTBG04, 4001, Len(strTBG04))
               ElseIf Len(strTBG04) > 6000 Then
                  strTBG04_3 = Mid(strTBG04, 4001, 2000)
                  If Len(strTBG04) < 8000 Then
                     strTBG04_4 = Mid(strTBG04, 6001, Len(strTBG04))
                  ElseIf Len(strTBG04) > 8000 Then
                     strTBG04_4 = Mid(strTBG04, 6001, 2000)
                     If Len(strTBG04) < 10000 Then
                        strTBG04_5 = Mid(strTBG04, 8001, Len(strTBG04))
                     ElseIf Len(strTBG04) > 10000 Then
                        strTBG04_5 = Mid(strTBG04, 8001, 2000)
                        If Len(strTBG04) < 12000 Then
                           strTBG04_6 = Mid(strTBG04, 10001, Len(strTBG04))
                        ElseIf Len(strTBG04) > 12000 Then
                           strTBG04_6 = Mid(strTBG04, 10001, 2000)
                        End If
                     End If
                  End If
               End If
            End If
         End If
         strErrTxt = "°Ó«~ÀÉ.TMBulletinGoods"
         '·s¼WTable
         '«D¥»©Ò®×¥ó¥B°Ó«~¸ê®Æ¬°¥xÆW®×ªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
'         If bolTaieCase = False And bolIsTaiwanCase = True Then
         '°Ó«~¸ê®Æ¬°¥xÆW®×ªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
         'Modify By Sindy 2012/1/5 +¤j³°®×
         'Modify By Sindy 2018/12/25 ¥xÆW®×­­¨î¬°¥»©Ò®×¥ó
         'Modify By Sindy 2024/12/19 ¹Å¶²»Ý¨D±ý¦b¿é¤Jµù¥UÃÒ®É,¼W¥[ÀË¬d¤½³ø°Ó«~¸ê®Æ,©Ò¥H¥»©Ò®×¥ó§¡­n¼g¤J¨t²Î
         'If (bolIsTaiwanCase = True And bolTaieCase = True) Or bolIsChinaCase = True Then
         If bolTaieCase = True Or bolIsChinaCase = True Then
         '2024/12/19 END
            'If strTBG02 <> "" Then
               If strTBG02 = "" Then strTBG02 = " " '¼Ð³¹¨S¦³°Ó«~Ãþ§O,¦ý¦³¼Ð³¹¤º®e
               'Modify By Sindy 2018/12/10 + ,TBG11,TBG12
               strSql = "insert into TMBulletinGoods (TBG01,TBG02,TBG03,TBG04,TBG05,TBG06,TBG07" & _
                        ",TBG08,TBG09,TBG10,TBG11,TBG12) " & _
                        "values(" & CNULL(strTBD02) & ",'" & strTBG02 & "'," & CNULL(strTBG03) & _
                        "," & CNULL(strTBG04_1) & "," & CNULL(strTBG04_2) & "," & CNULL(strTBG04_3) & _
                        "," & CNULL(strTBD03) & "," & CNULL(strTBG04_4) & "," & CNULL(strTBG04_5) & _
                        "," & CNULL(strTBG04_6) & ",'1'," & CNULL(txtTMBM07) & ")"
               cnnConnection.Execute strSql
               'Add By Sindy 2014/12/1
               If Trim(strTBG02) = "" Then '¼Ð³¹¨S¦³°Ó«~Ãþ§O,¦ý¦³¼Ð³¹¤º®e
                  Exit For
               End If
               '2014/12/1 END
            'End If
         End If
      Next dblChar
      If strTMBM08 <> "" Then strTMBM08 = Left(strTMBM08, Len(strTMBM08) - 1)
      
      'Add By Sindy 2015/10/27 ¤ñ¹ï°Ó«~Ãþ§O¸ê®Æ
      If bolTaieCase = True Then
         If Trim(strTMBM08) <> Trim(strTM09) Then
            strMsg = strTM01 & "-" & strTM02 & "-" & strTM03 & "-" & strTM04 & "¤½³ø°Ó«~Ãþ§O¬°:" & strTMBM08 & " ¥»©Ò°Ó«~Ãþ§O¬°:" & strTM09
            Call ReadTxt1(strTBD02, strTBD04, strMsg, "", "", "", strTBD03)
         End If
      End If
      '2015/10/27 End
      
      '·s¼WTable
      strErrTxt = "°Ó¼Ð¤½³øÀÉ.TMBulletin"
      strSql = "insert into TMBulletin (TMBM01,TMBM02,TMBM03,TMBM04,TMBM05,TMBM06,TMBM07,TMBM08) " & _
               "values(" & CNULL(strTBD02) & "," & CNULL(strTBD03) & ",''," & CNULL(strTBD04) & "," & CNULL(strTMBM05) & "," & CNULL(strTMBM06) & "," & CNULL(txtTMBM07) & "," & CNULL(strTMBM08) & ")"
      cnnConnection.Execute strSql
      'Add By Sindy 2017/4/25 ¥Ó½Ð¤H1~10
      If intApp > 0 Then
         strSql = "update TMBulletin set" & _
                  " TMBM09=" & CNULL(strTMBMApp(1)) & _
                  ",TMBM10=" & CNULL(strTMBMApp(2)) & _
                  ",TMBM11=" & CNULL(strTMBMApp(3)) & _
                  ",TMBM12=" & CNULL(strTMBMApp(4)) & _
                  ",TMBM13=" & CNULL(strTMBMApp(5)) & _
                  ",TMBM14=" & CNULL(strTMBMApp(6)) & _
                  ",TMBM15=" & CNULL(strTMBMApp(7)) & _
                  ",TMBM16=" & CNULL(strTMBMApp(8)) & _
                  ",TMBM17=" & CNULL(strTMBMApp(9)) & _
                  ",TMBM18=" & CNULL(strTMBMApp(10)) & _
                  " where TMBM01='" & strTBD02 & "' and TMBM02='" & strTBD03 & "'"
         cnnConnection.Execute strSql
      End If
      '2017/4/25 END
      
      '«D¥»©Ò®×¥ó¥B¤½³ø¸ê®Æ¬°¥xÆW®×ªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
'      If bolTaieCase = False And bolIsTaiwanCase = True Then
      '¤½³ø¸ê®Æ¬°¥xÆW®×ªÌ¡A¤~¶··s¼W¸ê®Æ¦ÜDB¡A±ý²£¥Í©w½Z©Ò¥Î
      'Modify By Sindy 2012/1/5 +¤j³°®×
      'Modify By Sindy 2018/12/25 ¥xÆW®×­­¨î¬°¥»©Ò®×¥ó
      If (bolIsTaiwanCase = True And bolTaieCase = True) Or bolIsChinaCase = True Then
         strErrTxt = "°Ó¼Ð¤½³ø«D¥»©Ò¸ê®ÆÀÉ.TMBulletinData"
         If bolIsTaiwanCase = True Then
            strTBD15 = "A"
         ElseIf bolIsChinaCase = True Then
            strTBD15 = "B"
         End If
         'Modify By Sindy 2018/12/10 + ,TBD16,TBD17
         strSql = "insert into TMBulletinData (TBD01,TBD02,TBD03,TBD04,TBD05,TBD06,TBD07,TBD08" & _
                  ",TBD09,TBD10,TBD11,TBD12,TBD13,TBD14,TBD15,TBD16,TBD17) " & _
                  "values(" & CNULL(txtTMBM07) & "," & CNULL(strTBD02) & "," & CNULL(strTBD03) & _
                  "," & CNULL(strTBD04) & "," & CNULL(strTBD05) & "," & CNULL(strTBD06) & _
                  "," & CNULL(strTBD07) & "," & CNULL(strTBD08) & "," & CNULL(strTBD09) & _
                  "," & CNULL(strTBD10) & "," & CNULL(strTBD11) & "," & CNULL(strTBD12) & _
                  "," & CNULL(strTBD13) & "," & CNULL(strTBD14) & "," & CNULL(strTBD15) & _
                  ",'1'," & Left(strTA05, 6) & ")"
         cnnConnection.Execute strSql
'         '¸ê®Æ¸Ì¦³?®É,»Ý¦C¦L²M³æ¤G
'         If m_bolCharQ = True Then
'            For i = 1 To 6
'               strTemp(i) = ""
'            Next i
'            strTemp(1) = strTBD02
'            strTemp(2) = strTBD04
'            strTemp(3) = GetTradeMarkName(strTBD03, 0)
'            strTemp(4) = m_strCharQNote
'            If iLine2 > 54 Or iLine2 = 0 Then
'               If iLine2 > 0 Then Printer.NewPage
'               PrintTitle2 '¦C¦LªíÀY
'            End If
'            PrintDetail2 '¦C¦L©ú²Ó
'         End If
      End If
      
      cnnConnection.CommitTrans
   Next dblFCnt
   
   '§ó·s¼f©w¸¹§@·~
   If File2.ListCount > 0 Then
      'MsgBox "¶}©l°õ¦æ§ó·s¼f©w¸¹§@·~¡A½Ðµy­Ô..."
      Label2.Caption = "§ó·s¼f©w¸¹¤¤¡A½Ðµy­Ô . . ."
'      frm030603.Hide
'      frm030603.textTMBM07 = txtTMBM07
'      frm030603.bolNotShowMsg = True
'      frm030603.cmdOK_Click
'      Unload frm030603
      Call frm030603_Process(txtTMBM07) 'Modify By Sindy 2018/12/17
   End If
   
   '±N¬°ª§Ä³¹ï³y®×¥óªº°Ó¼ÐÅv¤H¸ê®Æ§¡¤@¨Ö³]¬°N¤£¦C¦L¶}©Ý¨ç
   cnnConnection.BeginTrans
   strSql = "SELECT distinct tbor03 FROM TMBulletinData,TMBulletinOwner " & _
            "WHERE tbor01=tbd02 and tbor06=tbd03 and tbd14='N' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strSql = "update TMBulletinData " & _
                  "set TBD14='N' " & _
                  "WHERE tbd02||tbd03 in (SELECT tbor01||tbor06 FROM TMBulletinOwner WHERE tbor03='" & RsTemp.Fields(0) & "') "
         cnnConnection.Execute strSql
         RsTemp.MoveNext
      Loop
   End If
   cnnConnection.CommitTrans
   
   Dim strCP09 As String, strCP10 As String
   Dim m_TM01 As String, m_TM02 As String, m_TM03 As String, m_TM04 As String
   Dim m_CP09 As String, m_CP14 As String, m_CP36 As String, m_CP37 As String, m_CP38 As String, m_CP39 As String
   Dim m_cp40 As String, m_CP41 As String, m_CP42 As String
   Dim strCaseNo As String, strCause As String, strCauseText As String
   Dim m_TM05 As String, m_TM28 As String
   Dim m_TM23 As String, m_TM78 As String, m_TM79 As String, m_TM80 As String, m_TM81 As String
   Dim m_CP06 As String, m_CP07 As String 'Add By Sindy 2012/10/1
   'Add By Sindy 2013/9/16
   Dim m_CP13 As String, m_caseID As String, m_custnm As String, m_st02 As String, m_TM15 As String
   Dim m_TM14_dt As String
   '2013/9/16 END
   Dim m_TM29 As String     'add by sonia 2018/3/12
   
   'strTA05 = DBDATE(GetTA05)
   'Add By Sindy 2012/8/24 ­Y¬°¥Ó´_®×¥ó,­n³vµ§µoe-mail³qª¾¬ÛÃö¤H­û
   'TMBM0 CASEID          TM05                                                                             TM12                           TM14_DT                                                                          TM15                 TM23      CUSTNM                                                                           CP36                 CP37                                                                             CP38                                                                             CP39                                                                             CP40                                                                             CP41                                                                             CP42                                                                             SQLDATET(CP27)                                                                   CP14   CP13   ST02
   '----- --------------- -------------------------------------------------------------------------------- ------------------------------ -------------------------------------------------------------------------------- -------------------- --------- -------------------------------------------------------------------------------- -------------------- -------------------------------------------------------------------------------- -------------------------------------------------------------------------------- -------------------------------------------------------------------------------- -------------------------------------------------------------------------------- -------------------------------------------------------------------------------- -------------------------------------------------------------------------------- -------------------------------------------------------------------------------- ------ ------ ------------
   '3907  T-178269-0-00   »Ê°O¤Î¹Ï                                                                         100029446                                                                                                       01512234             X67698000 §d°®»Ê                                                                           100029446            »Ê°O¤Î¹Ï                                                                                                                                                                                                                                           »Ê°Oµ©¥J¦Ì¿|©± ½²¥É¯u                                                                                                                                                                                                                              101/01/11                                                                        98019  83001  ªL¤h³ó
   '3907  T-178710-0-00   ¸Ö¤H°sµ¢LE CAVEAU DU POETE / POET CELLAR                                         100042841                      101/04/01                                                                        01512147             X67628010 ±i·ç®e                                                                           100042841            ¸Ö¤H°sµ¢LE CAVEAU DU POETE POET CELLAR                                                                                                                                                                                                             ¸­§Ó°í                                                                                                                                                                                                                                             101/02/15                                                                        98019  99032  ¿c§Ó»Ê
   strSql = "select TM01||'-'||TM02||'-'||TM03||'-'||TM04 as caseID,TM05,TM12,sqldatet(TM14) as TM14_dt,TM15,TM23,decode(cu04,null,decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) as custnm,CP36,CP37,CP38,CP39,CP40,CP41,CP42,sqldatet(CP27),cp14,cp13,st02,cp09,TM01,TM02,TM03,TM04 " & _
            "From caseprogress, Trademark, staff, customer, TMBulletin " & _
            "Where CP01 in('T','FCT') and CP10 in('202','210') and CP24 is null and not CP27 is null " & _
            "and CP01=TM01(+) and CP02=TM02(+) and CP03=TM03(+) and CP04=TM04(+) " & _
            "and TM28<>'1' and TM29 is null and TM10='000' " & _
            "and cp13=st01(+) " & _
            "and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
            "and TMBM07='" & txtTMBM07 & "' and TMBM04=TM12 " & _
            "order by TM01,TM02,TM03,TM04 "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      cnnConnection.BeginTrans
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         'Add By Sindy 2012/9/21 §ó·s¸ê®Æ
         strSql = "update caseprogress set cp24='2',cp25=" & strTA05 & " where cp09='" & rsTmp.Fields("cp09") & "'"
         cnnConnection.Execute strSql
         strCP09 = AutoNo("C", 6)
         strCP10 = "1004" '®×¥ó©Ê½è¬°±Ñ¶D
         m_TM01 = "" & rsTmp.Fields("tm01")
         m_TM02 = "" & rsTmp.Fields("tm02")
         m_TM03 = "" & rsTmp.Fields("tm03")
         m_TM04 = "" & rsTmp.Fields("tm04")
         m_CP09 = "" & rsTmp.Fields("cp09")
         m_CP14 = "" & rsTmp.Fields("cp14")
         'Add By Sindy 2015/1/16
         m_CP13 = "" & rsTmp.Fields("cp13")
         m_caseID = "" & rsTmp.Fields("caseID")
         m_TM05 = "" & rsTmp.Fields("tm05")
         m_custnm = "" & rsTmp.Fields("custnm")
         m_st02 = "" & rsTmp.Fields("st02")
         m_TM15 = "" & rsTmp.Fields("TM15")
         m_TM14_dt = "" & rsTmp.Fields("TM14_dt")
         '2015/1/16 END
         m_CP36 = "" & rsTmp.Fields("cp36")
         m_CP37 = "" & rsTmp.Fields("cp37")
         m_CP38 = "" & rsTmp.Fields("cp38")
         m_CP39 = "" & rsTmp.Fields("cp39")
         m_cp40 = "" & rsTmp.Fields("cp40")
         m_CP41 = "" & rsTmp.Fields("cp41")
         m_CP42 = "" & rsTmp.Fields("cp42")
         'Add By Sindy 2012/10/1
         m_CP07 = Format(DateAdd("M", 3, DateSerial(Left(strTA05, 4), Mid(strTA05, 5, 2), Right(strTA05, 2))), "YYYYMMDD") 'ªk©w´Á­­
         'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
         If Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
            m_CP06 = PUB_GetOurDeadline(DBDATE(m_CP07))
         Else
         '2014/10/6 END
            m_CP06 = Format(DateAdd("D", -4, DateSerial(Left(m_CP07, 4), Mid(m_CP07, 5, 2), Right(m_CP07, 2))), "YYYYMMDD") '¥»©Ò´Á­­
         End If
         '2012/10/1 End
         'Modify By Sindy 2012/10/1 +CP06,CP07
         strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP27) " & _
                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strTA05 & "," & m_CP06 & "," & m_CP07 & "," & _
                          "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & m_CP14 & "'," & _
                          "'N','N','N'," & _
                          "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
                          "'" & m_cp40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "'," & DBDATE("111111") & ")"
         cnnConnection.Execute strSql
         '2012/9/21 End
         'Add By Sindy 2012/10/1 ·s¼W¤U¤@µ{§Ç601²§Ä³
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','601'," & _
                          m_CP06 & "," & m_CP07 & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & GetNextProgressNo() & ")"
         cnnConnection.Execute strSql
         '2012/10/1 End
         'Modify By Sindy 2012/9/21 ¤£¥²µoµ¹ÅÇ,§ïµo´¼Åv¤H­û
         'strTo = "67002" & ";" & "" & rsTmp.Fields("cp14") & ";" & Pub_GetSpecMan("P1")
         'modify by sonia 2016/10/19 ¥[µo69008
         'Modify By Sindy 2019/5/16
         'strTo = "67002;69008" & ";" & m_CP14 & ";" & m_CP13
         'modify by sonia 2020/5/5 ¨ú®ø67002
         'Modify By Sindy 2021/11/10 ³¯­z·N¨£®Ñ¤§ª§Ä³®×, ¹ï³y®×¥ó¤w®Ö­ã©Î®Ö»é, ½Ð°µ«áÄò³B²z¡A
         'MAIL³¡¤ÀªL¸g²z§ï¬°98003ªL«ß®v¤Î86048ªL©Ó¼z¡C
         strTo = "98003;86048;" & m_CP14 & ";" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))
         '2019/5/16 END
         strSubject = m_caseID & " ³¯­z·N¨£®Ñ¤§ª§Ä³®×, ¹ï³y®×¥ó¤w®Ö­ã, ½Ð°µ«áÄò³B²z !"
         strContext = "¥»©Ò®×¸¹¡G" & m_caseID & vbCrLf & _
                      "®×¥ó¦WºÙ¡G" & m_TM05 & vbCrLf & _
                      "¥Ó ½Ð ¤H¡G" & m_custnm & vbCrLf & _
                      "´¼Åv¤H­û¡G" & m_st02 & vbCrLf & _
                      "¼f ©w ¸¹¡G" & m_TM15 & vbCrLf & _
                      "¤½ §i ¤é¡G" & m_TM14_dt & vbCrLf & vbCrLf & _
                      "¹ï³y¸¹¼Æ¡G" & m_CP36 & vbCrLf & _
                      "¹ï³y®×¥ó¦WºÙ¡G" & m_CP37 & m_CP38 & m_CP39 & vbCrLf & _
                      "¹ï³y¤¤¤å¡G" & m_cp40 & m_CP41 & m_CP42 & vbCrLf
         PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , , , , , , False
         rsTmp.MoveNext
      Loop
      cnnConnection.CommitTrans
   End If
   rsTmp.Close
   '2012/8/24 End
   
   'Add By Sindy 2012/9/21
   'ÀË¯Á®Ö»é¤½§i¸ê®Æ
   'strTA05 = DBDATE(GetTA05)
   Label2.Caption = "ÀË¯Á®Ö»é¤½§i¸ê®Æ¡A½Ðµy­Ô . . ."
   File2.path = txtPath2.Text & "\Reject"
   File2.Refresh
   strTotRow = File2.ListCount
   Me.Height = 6120
   dblMaxWidth = 8820
   Text2.Width = 0
   If File2.ListCount > 0 Then
      cnnConnection.BeginTrans
      For dblFCnt = 0 To File2.ListCount - 1
         'Add by Sindy 2022/3/3
         If strSrvDate(1) >= Form20¤W½u¤é Then
            adoStream.LoadFromFile (txtPath2.Text & "\Reject\" & File2.List(dblFCnt))
            m_strTextBox = adoStream.ReadText
         Else
         '2022/3/3 END
            RichTextBox1.LoadFile (txtPath2.Text & "\Reject\" & File2.List(dblFCnt))
            m_strTextBox = RichTextBox1.Text
         End If
         dblLastEnd = InStr(m_strTextBox, "</Reject>")
         Text2.Width = dblMaxWidth / Val(strTotRow) * (dblFCnt + 1): DoEvents
         
         strText = "CaseNo": strTitNM = "¥Ó½Ð®×¸¹"
         dblStar = InStr(1, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
         '***** ¸ÑªRXML *****
         If GetXmlData(1, strText, strTitNM, True, strData, dblEnd) = False Then
            GoTo ReadNextCol4 'Exit For
         End If
         '***** End
         If strData <> "" Then
            strSql = "select TM01||'-'||TM02||'-'||TM03||'-'||TM04 as caseID,TM05,TM12,sqldatet(TM14) as TM14_dt,TM15,TM23,decode(cu04,null,decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90),cu04) as custnm,CP36,CP37,CP38,CP39,CP40,CP41,CP42,sqldatet(CP27),cp14,cp13,st02,cp09,TM01,TM02,TM03,TM04 " & _
                     "From caseprogress, Trademark, staff, customer " & _
                     "Where CP01 in('T','FCT') and CP10 in('202','210') and CP24 is null and not CP27 is null " & _
                     "and CP01=TM01 and CP02=TM02 and CP03=TM03 and CP04=TM04 " & _
                     "and TM28<>'1' and TM29 is null and TM10='000' " & _
                     "and cp13=st01(+) " & _
                     "and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
                     "and TM12='" & strData & "' "
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               rsTmp.MoveFirst
               Do While Not rsTmp.EOF
                  '§ó·s¸ê®Æ
                  strSql = "update caseprogress set cp24='1',cp25=" & strTA05 & " where cp09='" & rsTmp.Fields("cp09") & "'"
                  cnnConnection.Execute strSql
                  strCP09 = AutoNo("C", 6)
                  strCP10 = "1003" '®×¥ó©Ê½è¬°³Ó¶D
                  m_TM01 = "" & rsTmp.Fields("tm01")
                  m_TM02 = "" & rsTmp.Fields("tm02")
                  m_TM03 = "" & rsTmp.Fields("tm03")
                  m_TM04 = "" & rsTmp.Fields("tm04")
                  m_CP09 = "" & rsTmp.Fields("cp09")
                  m_CP14 = "" & rsTmp.Fields("cp14")
                  m_CP36 = "" & rsTmp.Fields("cp36")
                  m_CP37 = "" & rsTmp.Fields("cp37")
                  m_CP38 = "" & rsTmp.Fields("cp38")
                  m_CP39 = "" & rsTmp.Fields("cp39")
                  m_cp40 = "" & rsTmp.Fields("cp40")
                  m_CP41 = "" & rsTmp.Fields("cp41")
                  m_CP42 = "" & rsTmp.Fields("cp42")
                  'Add By Sindy 2013/9/16
                  m_CP13 = "" & rsTmp.Fields("cp13")
                  m_caseID = "" & rsTmp.Fields("caseID")
                  m_TM05 = "" & rsTmp.Fields("tm05")
                  m_custnm = "" & rsTmp.Fields("custnm")
                  m_st02 = "" & rsTmp.Fields("st02")
                  m_TM15 = "" & rsTmp.Fields("TM15")
                  m_TM14_dt = "" & rsTmp.Fields("TM14_dt")
                  '2013/9/16 END
                  strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP27) " & _
                           "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strTA05 & "," & _
                                   "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & m_CP14 & "'," & _
                                   "'N','N','N'," & _
                                   "'" & m_CP36 & "','" & ChgSQL(m_CP37) & "','" & ChgSQL(m_CP38) & "','" & m_CP39 & "'," & _
                                   "'" & m_cp40 & "','" & ChgSQL(m_CP41) & "','" & m_CP42 & "','" & m_CP09 & "'," & DBDATE("111111") & ")"
                  cnnConnection.Execute strSql
                  'µoMail
                  'modify by sonia 2016/10/19 ¥[µo69008
                  'Modify By Sindy 2019/5/16
                  'strTo = "67002;69008" & ";" & m_CP14 & ";" & m_CP13
                  'modify by sonia 2020/5/5 ¨ú®ø67002
                  'Modify By Sindy 2021/11/10 ³¯­z·N¨£®Ñ¤§ª§Ä³®×, ¹ï³y®×¥ó¤w®Ö­ã©Î®Ö»é, ½Ð°µ«áÄò³B²z¡A
                  'MAIL³¡¤ÀªL¸g²z§ï¬°98003ªL«ß®v¤Î86048ªL©Ó¼z¡C
                  strTo = "98003;86048;" & m_CP14 & ";" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))
                  '2019/5/16 END
                  strSubject = m_caseID & " ³¯­z·N¨£®Ñ¤§ª§Ä³®×, ¹ï³y®×¥ó¤w®Ö»é, ½Ð°µ«áÄò³B²z !"
                  strContext = "¥»©Ò®×¸¹¡G" & m_caseID & vbCrLf & _
                               "®×¥ó¦WºÙ¡G" & m_TM05 & vbCrLf & _
                               "¥Ó ½Ð ¤H¡G" & m_custnm & vbCrLf & _
                               "´¼Åv¤H­û¡G" & m_st02 & vbCrLf & _
                               "¼f ©w ¸¹¡G" & m_TM15 & vbCrLf & _
                               "¤½ §i ¤é¡G" & m_TM14_dt & vbCrLf & vbCrLf & _
                               "¹ï³y¸¹¼Æ¡G" & m_CP36 & vbCrLf & _
                               "¹ï³y®×¥ó¦WºÙ¡G" & m_CP37 & m_CP38 & m_CP39 & vbCrLf & _
                               "¹ï³y¤¤¤å¡G" & m_cp40 & m_CP41 & m_CP42 & vbCrLf
                  PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , , , , , , False
                  rsTmp.MoveNext
               Loop
            End If
            rsTmp.Close
         End If
ReadNextCol4:
      Next dblFCnt
      cnnConnection.CommitTrans
   End If
   'ÀË¯ÁºM¾P¤½§i¸ê®Æ
   'strTA05 = DBDATE(GetTA05)
   Label2.Caption = "ÀË¯ÁºM¾P¤½§i¸ê®Æ¡A½Ðµy­Ô . . ."
   File2.path = txtPath2.Text & "\Revocation"
   File2.Refresh
   strTotRow = File2.ListCount
   Me.Height = 6120
   dblMaxWidth = 8820
   Text2.Width = 0
   If File2.ListCount > 0 Then
      cnnConnection.BeginTrans
      For dblFCnt = 0 To File2.ListCount - 1
         'Add by Sindy 2022/3/3
         If strSrvDate(1) >= Form20¤W½u¤é Then
            adoStream.LoadFromFile (txtPath2.Text & "\Revocation\" & File2.List(dblFCnt))
            m_strTextBox = adoStream.ReadText
         Else
         '2022/3/3 END
            RichTextBox1.LoadFile (txtPath2.Text & "\Revocation\" & File2.List(dblFCnt))
            m_strTextBox = RichTextBox1.Text
         End If
         dblLastEnd = InStr(m_strTextBox, "</Revocation>")
         Text2.Width = dblMaxWidth / Val(strTotRow) * (dblFCnt + 1): DoEvents
         
         strCaseNo = ""
         strCause = ""
         'Åª¨ú¤½³ø¸ê®Æ
         For dblChar = 1 To dblLastEnd
            strText = "RegisterNo": strTitNM = "µù¥U¸¹"
            dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
            If dblStar < dblChar Then Exit For
            For intCol = 1 To 2
               strData = ""
               If intCol = 1 Then
                  strText = "RegisterNo": strTitNM = "µù¥U¸¹"
               ElseIf intCol = 2 Then
                  strText = "Cause": strTitNM = "ºM¾P­ì¦]"
               End If
               '***** ¸ÑªRXML *****
               If GetXmlData(dblChar, strText, strTitNM, True, strData, dblEnd) = False Then GoTo ReadNextCol5
               '***** End
               If strData <> "" Then
                  If intCol = 1 Then '¥Ó½Ð®×¸¹
                     strCaseNo = strData
                  ElseIf intCol = 2 Then 'ºM¾P­ì¦]
                     strCause = strData
                  End If
               End If
ReadNextCol5:
            Next intCol
            dblChar = dblEnd
         Next dblChar
         If strCaseNo <> "" Then
            'modify by sonia 2018/3/12 +TM29
            strSql = "select TM01,TM02,TM03,TM04,TM05,TM23,TM28,TM78,TM79,TM80,TM81,TM29 " & _
                     "From Trademark " & _
                     "Where TM15='" & strCaseNo & "' "
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               '§ó·s¸ê®Æ
               strCP09 = AutoNo("C", 6)
               strCP10 = "1007" '®×¥ó©Ê½è¬°ºM¾P¤½§i
               m_TM01 = "" & rsTmp.Fields("tm01")
               m_TM02 = "" & rsTmp.Fields("tm02")
               m_TM03 = "" & rsTmp.Fields("tm03")
               m_TM04 = "" & rsTmp.Fields("tm04")
               m_TM05 = "" & rsTmp.Fields("tm05")
               m_TM28 = "" & rsTmp.Fields("tm28")
               m_TM23 = "" & rsTmp.Fields("tm23")
               m_TM78 = "" & rsTmp.Fields("tm78")
               m_TM79 = "" & rsTmp.Fields("tm79")
               m_TM80 = "" & rsTmp.Fields("tm80")
               m_TM81 = "" & rsTmp.Fields("tm81")
               m_TM29 = "" & rsTmp.Fields("tm29")   'add by sonia 2018/3/12
               'ABÃþ³Ì¤jµo¤å¤é¤§µo¤å¤H­û
               strSql = "select cp83 From caseprogress " & _
                        "Where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' " & _
                        "and cp09<'C' and cp27 in(select max(cp27) " & _
                        "From caseprogress " & _
                        "Where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' " & _
                        "and cp09<'C' and cp27 is not null)"
               intI = 1
               Set rsTmp3 = ClsLawReadRstMsg(intI, strSql)
               m_CP14 = ""
               If intI = 1 Then
                  m_CP14 = "" & rsTmp3.Fields(0)
                  'Add By Sindy 2020/12/23 ex:FCT-042862\T-203840
                  If ChkStaffST04(m_CP14, False) = True Then '¤wÂ÷Â¾«h±a¾Þ§@¤H­û
                     m_CP14 = strUserNum
                  End If
                  '2020/12/23 END
               End If
               rsTmp3.Close
               strCauseText = "¦¹®×©ó" & ChangeWStringToTDateString(strTA05) & "ºM¾P¤½§i¡A" & Trim(strCause)
               strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP64) " & _
                        "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strTA05 & "," & _
                                "'" & strCP09 & "','" & strCP10 & "'," & _
                                "'" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "'," & _
                                "'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & _
                                "'" & m_CP14 & "'," & _
                                "'N','N','N'," & _
                                strTA05 & ",'" & strCauseText & "')"
               cnnConnection.Execute strSql
               
               'add by sonia 2018/3/12  '¨÷©v©Ê½è¬°¥Ó½Ð®×®É­n¦P®É³¬¨÷T160591~3
               'strSql = "update trademark " & _
                        "set TM58=TM58||';" & strCauseText & "' " & _
                        "Where TM01='" & m_TM01 & "' and TM02='" & m_TM02 & "' and TM03='" & m_TM03 & "' and TM04='" & m_TM04 & "' "
               strUpdTM29 = ""
               If m_TM28 = "1" And m_TM29 = "" Then
                  strUpdTM29 = ",TM29='Y',TM31='86',TM30=" & strSrvDate(1)
               End If
               strSql = "update trademark " & _
                        "set TM58=TM58||';" & strCauseText & "'" & strUpdTM29 & _
                        " Where TM01='" & m_TM01 & "' and TM02='" & m_TM02 & "' and TM03='" & m_TM03 & "' and TM04='" & m_TM04 & "' "
               'end 2018/3/12
               cnnConnection.Execute strSql
               '¨÷©v©Ê½è«D¥Ó½Ð®×®É
               If m_TM28 <> "1" Then
                  strSql = "select * From caseprogress " & _
                           "Where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' " & _
                           "and (cp36 is not null or cp40 is not null or cp41 is not null or cp42 is not null) " & _
                           "order by cp05 desc "
                  intI = 1
                  Set rsTmp3 = ClsLawReadRstMsg(intI, strSql)
                  strTo = ""
                  If intI = 1 Then 'ÀË¬dCP¬O§_¦³¹ï³y
                     strText = ""
                     If m_TM23 <> "" Then
                        If strText <> "" Then strText = strText & vbCrLf & "¡@¡@¡@¡@¡@"
                        strText = strText & GetPrjPeople1(m_TM23, "1")
                     End If
                     If m_TM78 <> "" Then
                        If strText <> "" Then strText = strText & vbCrLf & "¡@¡@¡@¡@¡@"
                        strText = strText & GetPrjPeople1(m_TM78, "1")
                     End If
                     If m_TM79 <> "" Then
                        If strText <> "" Then strText = strText & vbCrLf & "¡@¡@¡@¡@¡@"
                        strText = strText & GetPrjPeople1(m_TM79, "1")
                     End If
                     If m_TM80 <> "" Then
                        If strText <> "" Then strText = strText & vbCrLf & "¡@¡@¡@¡@¡@"
                        strText = strText & GetPrjPeople1(m_TM80, "1")
                     End If
                     If m_TM81 <> "" Then
                        If strText <> "" Then strText = strText & vbCrLf & "¡@¡@¡@¡@¡@"
                        strText = strText & GetPrjPeople1(m_TM81, "1")
                     End If
                     'µoMail
                     'Modify By Sindy 2019/5/16
                     'strTo = "" & rsTmp3.Fields("cp13")
                     strTo = IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))
                     '2019/5/16 END
                     strSubject = m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & " ª§Ä³®×¥ó¡A¹ï³y°Ó¼Ð¤w¾DºM¾P¤½§i³qª¾ !"
                     strContext = "¥»©Ò®×¸¹¡G" & "" & m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & vbCrLf & _
                                  "¥»©Ò«È¤á¡G" & strText & vbCrLf & vbCrLf & _
                                  "®×¥ó¦WºÙ¡G" & m_TM05 & vbCrLf & _
                                  "¹ï³y¦WºÙ¡G" & "" & rsTmp3.Fields("CP40") & " " & rsTmp3.Fields("CP41") & " " & rsTmp3.Fields("CP42") & vbCrLf & _
                                  "¹ï³y¸¹¼Æ¡G" & "" & rsTmp3.Fields("CP36") & vbCrLf & vbCrLf & _
                                  "¦¹®×¥ó¹ï³y°Ó¼Ð¤w¾DºM¾P¤½§i¡A«áÄò¦V«È¤á³ø§i©Î½Ð´Ú¨Æ©y¥i³w¦æ³B²z©Î»P±M·~³¡ÁpÃ´¡I"
                     PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , , , , , , False
                  End If
                  rsTmp3.Close
               End If
            End If
            rsTmp.Close
         End If
      Next dblFCnt
      cnnConnection.CommitTrans
   End If
   '2012/9/21 End
   
   'Add By Sindy 2024/6/27 ÀË¯Á¡u³¡¤À¡vºM¾P¤½§i¸ê®Æ
   'strTA05 = DBDATE(GetTA05)
   Label2.Caption = "ÀË¯Á¡u³¡¤À¡vºM¾P¤½§i¸ê®Æ¡A½Ðµy­Ô . . ."
   File2.path = txtPath2.Text & "\DelProduct"
   File2.Refresh
   strTotRow = File2.ListCount
   Me.Height = 6120
   dblMaxWidth = 8820
   Text2.Width = 0
   If File2.ListCount > 0 Then
      cnnConnection.BeginTrans
      For dblFCnt = 0 To File2.ListCount - 1
         'Form20¤W½u
         adoStream.LoadFromFile (txtPath2.Text & "\DelProduct\" & File2.List(dblFCnt))
         m_strTextBox = adoStream.ReadText
         dblLastEnd = InStr(m_strTextBox, "</DelProduct>")
         Text2.Width = dblMaxWidth / Val(strTotRow) * (dblFCnt + 1): DoEvents
         
         strCaseNo = ""
         strCause = ""
         'Åª¨ú¤½³ø¸ê®Æ
         For dblChar = 1 To dblLastEnd
            strText = "RegisterNo": strTitNM = "µù¥U¸¹"
            dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
            If dblStar < dblChar Then Exit For
            For intCol = 1 To 2
               strData = ""
               If intCol = 1 Then
                  strText = "RegisterNo": strTitNM = "µù¥U¸¹"
               ElseIf intCol = 2 Then
                  strText = "Tgoods_Denomination": strTitNM = "³¡¤ÀºM¾P­ì¦]"
               End If
               '***** ¸ÑªRXML *****
               If GetXmlData(dblChar, strText, strTitNM, True, strData, dblEnd) = False Then GoTo ReadNextCol6
               '***** End
               If strData <> "" Then
                  If intCol = 1 Then '¥Ó½Ð®×¸¹
                     strCaseNo = strData
                  ElseIf intCol = 2 Then 'ºM¾P­ì¦]
                     strCause = strData
                  End If
               End If
ReadNextCol6:
            Next intCol
            dblChar = dblEnd
         Next dblChar
         If strCaseNo <> "" Then
            strSql = "select TM01,TM02,TM03,TM04,TM05,TM23,TM28,TM78,TM79,TM80,TM81,TM29 " & _
                     "From Trademark " & _
                     "Where TM15='" & strCaseNo & "' "
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               '§ó·s¸ê®Æ
               strCP09 = AutoNo("C", 6)
               strCP10 = "1008" '®×¥ó©Ê½è¬°³¡¤ÀºM¾P¤½§i
               m_TM01 = "" & rsTmp.Fields("tm01")
               m_TM02 = "" & rsTmp.Fields("tm02")
               m_TM03 = "" & rsTmp.Fields("tm03")
               m_TM04 = "" & rsTmp.Fields("tm04")
               m_TM05 = "" & rsTmp.Fields("tm05")
               m_TM28 = "" & rsTmp.Fields("tm28")
               m_TM23 = "" & rsTmp.Fields("tm23")
               m_TM78 = "" & rsTmp.Fields("tm78")
               m_TM79 = "" & rsTmp.Fields("tm79")
               m_TM80 = "" & rsTmp.Fields("tm80")
               m_TM81 = "" & rsTmp.Fields("tm81")
               m_TM29 = "" & rsTmp.Fields("tm29")
               'ABÃþ³Ì¤jµo¤å¤é¤§µo¤å¤H­û
               strSql = "select cp83 From caseprogress " & _
                        "Where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' " & _
                        "and cp09<'C' and cp27 in(select max(cp27) " & _
                        "From caseprogress " & _
                        "Where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' " & _
                        "and cp09<'C' and cp27 is not null)"
               intI = 1
               Set rsTmp3 = ClsLawReadRstMsg(intI, strSql)
               m_CP14 = ""
               If intI = 1 Then
                  m_CP14 = "" & rsTmp3.Fields(0)
                  If ChkStaffST04(m_CP14, False) = True Then '¤wÂ÷Â¾«h±a¾Þ§@¤H­û
                     m_CP14 = strUserNum
                  End If
               End If
               rsTmp3.Close
               strCauseText = "¦¹®×©ó" & ChangeWStringToTDateString(strTA05) & "³¡¤ÀºM¾P¤½§i¡A" & Trim(strCause)
               strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP64) " & _
                        "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strTA05 & "," & _
                                "'" & strCP09 & "','" & strCP10 & "'," & _
                                "'" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "'," & _
                                "'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & _
                                "'" & m_CP14 & "'," & _
                                "'N','N','N'," & _
                                strTA05 & ",'" & strCauseText & "')"
               cnnConnection.Execute strSql
               
               strSql = "update trademark " & _
                        "set TM58=TM58||';" & strCauseText & "'" & _
                        " Where TM01='" & m_TM01 & "' and TM02='" & m_TM02 & "' and TM03='" & m_TM03 & "' and TM04='" & m_TM04 & "' "
               cnnConnection.Execute strSql
               
               strSql = "select * From caseprogress " & _
                        "Where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' " & _
                        "and (cp36 is not null or cp40 is not null or cp41 is not null or cp42 is not null) " & _
                        "order by cp05 desc "
               intI = 1
               Set rsTmp3 = ClsLawReadRstMsg(intI, strSql)
               strTo = ""
               If intI = 1 Then 'ÀË¬dCP¬O§_¦³¹ï³y
                  strText = ""
                  If m_TM23 <> "" Then
                     If strText <> "" Then strText = strText & vbCrLf & "¡@¡@¡@¡@¡@"
                     strText = strText & GetPrjPeople1(m_TM23, "1")
                  End If
                  If m_TM78 <> "" Then
                     If strText <> "" Then strText = strText & vbCrLf & "¡@¡@¡@¡@¡@"
                     strText = strText & GetPrjPeople1(m_TM78, "1")
                  End If
                  If m_TM79 <> "" Then
                     If strText <> "" Then strText = strText & vbCrLf & "¡@¡@¡@¡@¡@"
                     strText = strText & GetPrjPeople1(m_TM79, "1")
                  End If
                  If m_TM80 <> "" Then
                     If strText <> "" Then strText = strText & vbCrLf & "¡@¡@¡@¡@¡@"
                     strText = strText & GetPrjPeople1(m_TM80, "1")
                  End If
                  If m_TM81 <> "" Then
                     If strText <> "" Then strText = strText & vbCrLf & "¡@¡@¡@¡@¡@"
                     strText = strText & GetPrjPeople1(m_TM81, "1")
                  End If
                  '¨÷©v©Ê½è«D¥Ó½Ð®×®É
                  If m_TM28 <> "1" Then '§Ú­Ì¥h¥´§O¤Hªº°Ó¼Ð®×~
                     'µoMail
                     strTo = IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))
                     strSubject = m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & " ª§Ä³®×¥ó¡A¹ï³y°Ó¼Ð¤w¾DºM¾P³¡¤À°Ó«~¤½§i³qª¾¡I"
                     strContext = "¥»©Ò®×¸¹¡G" & "" & m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & vbCrLf & _
                                  "¥»©Ò«È¤á¡G" & strText & vbCrLf & vbCrLf & _
                                  "®×¥ó¦WºÙ¡G" & m_TM05 & vbCrLf & _
                                  "¹ï³y¦WºÙ¡G" & "" & rsTmp3.Fields("CP40") & " " & rsTmp3.Fields("CP41") & " " & rsTmp3.Fields("CP42") & vbCrLf & _
                                  "¹ï³y¸¹¼Æ¡G" & "" & rsTmp3.Fields("CP36") & vbCrLf & vbCrLf & _
                                  "¦¹®×¥ó¹ï³y°Ó¼Ð¤w¾DºM¾P³¡¤À°Ó«~¤½§i¡A«áÄò¦V«È¤á³ø§i©Î½Ð´Ú¨Æ©y¥i³w¦æ³B²z©Î»P±M·~³¡ÁpÃ´¡I"
                     PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , , , , , , False
                  
                  '¨÷©v©Ê½è¥Ó½Ð®×®É: §Ú­Ìªº°Ó¼Ð®×³Q¤H¥´~
                  ElseIf m_TM28 = "1" And m_TM01 = "FCT" Then
                     'µoMail
                     strTo = IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))
                     strSubject = m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & " ª§Ä³®×¥ó¡A«È¤á°Ó¼Ð¤w¾DºM¾P³¡¤À°Ó«~¤½§i³qª¾¡I"
                     strContext = "¥»©Ò®×¸¹¡G" & "" & m_TM01 & "-" & m_TM02 & IIf(m_TM03 & m_TM04 = "000", "", "-" & m_TM03 & "-" & m_TM04) & vbCrLf & _
                                  "¥»©Ò«È¤á¡G" & strText & vbCrLf & vbCrLf & _
                                  "®×¥ó¦WºÙ¡G" & m_TM05 & vbCrLf & _
                                  "¹ï³y¦WºÙ¡G" & "" & rsTmp3.Fields("CP40") & " " & rsTmp3.Fields("CP41") & " " & rsTmp3.Fields("CP42") & vbCrLf & vbCrLf & _
                                  "¦¹®×¥ó«È¤á°Ó¼Ð¤w¾DºM¾P³¡¤À°Ó«~¤½§i¡A½Ð½T»{°Ó«~¨Ã­×¥¿°ò¥»ÀÉ°Ó«~¸ê®Æ¡A«áÄò¦V«È¤á³ø§i©Î½Ð´Ú¨Æ©y¥i³w¦æ³B²z©Î»P±M·~³¡ÁpÃ´¡I"
                     PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , , , , , , , , , , False
                  End If
               End If
               rsTmp3.Close
            End If
            rsTmp.Close
         End If
      Next dblFCnt
      cnnConnection.CommitTrans
   End If
   '2024/6/27 End
   
'   '¦a°Ï¦WºÙ¬°ªÅ¥Õ©Î¤¤°ê¤j³°,¤Î­Y¥N²z¤H¦WºÙ¦³?®É,»Ý¦C¦L²M³æ¤@
'   strSql = "SELECT * FROM Tmbulletin " & _
'                 "WHERE TMBM07 = '" & txtTMBM07 & "' order by TMBM01 asc "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   m_PrintRpt = False
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         If "" & rsTmp.Fields("TMBM05") = "" Or _
'            "" & rsTmp.Fields("TMBM05") = "¤¤°ê¤j³°" Or _
'            GetNationNo("" & rsTmp.Fields("TMBM05")) = "" Or _
'            InStr("" & rsTmp.Fields("TMBM06"), "?") > 0 Then
'            iLine2 = iLine2 + 1
'            For i = 1 To 6
'               strTemp(i) = ""
'            Next i
'            strTemp(1) = "" & rsTmp.Fields("TMBM01")
'            strTemp(2) = "" & rsTmp.Fields("TMBM04")
'            strTemp(3) = "" & rsTmp.Fields("TMBM05")
'            strTemp(4) = "" & rsTmp.Fields("TMBM06")
'   '         strTemp(5) = strAChinese1
'   '         strTemp(6) = strAddress1
'            If iLine > 37 Or iLine = 0 Then
'               If iLine > 0 Then Printer.NewPage
'               PrintTitle '¦C¦LªíÀY
'            End If
'            PrintDetail '¦C¦L©ú²Ó
'         End If
'         rsTmp.MoveNext
'      Loop
'   End If
'   If m_PrintRpt = True Then Printer.EndDoc: MsgBox iLine2 & "µ§"
'   rsTmp.Close
   
   'If m_PrintRpt = True Then Printer.EndDoc
   strMsg = ""
   If m_PrintRpt1 = True Then
'      Close ff1
      'Add By Sindy 2024/5/17
      If Dir(PUB_Getdesktop & "\" & m_strFileName1) <> "" Then
         Kill PUB_Getdesktop & "\" & m_strFileName1
         Sleep 100
      End If
      Call PUB_SaveTextAsUTF8(PUB_Getdesktop & "\" & m_strFileName1, m_strText)
      '2024/5/17 END
      strMsg = m_strFileName1
   End If
   If m_PrintRpt2 = True Then
      Close FF2
      If strMsg <> "" Then strMsg = strMsg & " ¤Î "
      strMsg = strMsg & m_strFileName2
   End If
   If m_PrintRpt3 = True Then
      'Add By Sindy 2016/2/17
      Print #ff3, "¥H¤W¸ê®Æ½Ð³qª¾¹q¸£¤¤¤ß¨ó§UÂà¤J¡A¨ä¥L¸ê®Æ¤w¶×¤J§¹²¦¡I"
      '2016/2/17 End
      Close ff3
      If strMsg <> "" Then strMsg = strMsg & " ¤Î "
      strMsg = strMsg & m_strFileName3
   End If
   If m_PrintRpt1 = True Or m_PrintRpt2 = True Or m_PrintRpt3 = True Then
      'MsgBox "½Ð¦Ü¤U¦C¦ì¸m¦C¦LÀË®Öªí¡G" & PUB_Getdesktop & "\" & strMsg
      strMsg = "½Ð¦Ü¤U¦C¦ì¸m¦C¦LÀË®Öªí¡G" & PUB_Getdesktop & "\" & strMsg
   End If
   
   Screen.MousePointer = vbDefault
   Call IsRecordExist '²£¥Íµ§¼Æ
   
   'Add By Sindy 2015/5/13 ³qª¾µ{§Ç¤wÂàÀÉ§¹²¦
   If strP22 <> "" Then
      strSubject = "°Ó¼Ð¤½³ø¤wÂàÀÉ§¹²¦¡I"
      PUB_SendMail strUserNum, strP22, "", strSubject, strSubject, , , , , , , , , , , False
      'Add By Sindy 2024/1/17
      strExc(10) = ""
      If m_strFileName1 <> "" Then
         If strExc(10) <> "" Then strExc(10) = strExc(10) & "*"
         strExc(10) = PUB_Getdesktop & "\" & m_strFileName1
      End If
      If m_strFileName2 <> "" Then
         If strExc(10) <> "" Then strExc(10) = strExc(10) & "*"
         strExc(10) = PUB_Getdesktop & "\" & m_strFileName2
      End If
      If m_strFileName3 <> "" Then
         If strExc(10) <> "" Then strExc(10) = strExc(10) & "*"
         strExc(10) = PUB_Getdesktop & "\" & m_strFileName3
      End If
      'Modify By Sindy 2025/4/25
      PUB_SendMail strUserNum, "97038", "", "[ÀË¬d¬O§_¦³?¸¹¦r¤¸]" & strSubject, " ªþ¤WÀË®Öªí!!", , strExc(10), , , , , , , , True, False
      '2024/1/17 END
   End If
   '2015/5/13 END
   
   MsgBox "ÂàÀÉ§¹²¦¡I(ÂàÀÉªá¶O®É¶¡¡G" & strTime & "  " & time() & ")" & vbCrLf & strMsg
   Me.Height = 5000
   
   Set rsTmp = Nothing
   Set rsTmp3 = Nothing
   Exit Sub
   
ErrHand:
   If Err.Number = -2147217900 Then 'ORA-00917: ¿òº|³rÂI
      '¼gLog
      Call ReadTxt3(strSql)
      '±µµÛµo¥Í¿ù»~³¯­z¦¡ªº¤U­Ó³¯­z¦¡¶}©l°õ¦æ
      Resume Next
   End If
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   
   If Err.Number = 76 Then
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2.Text & "\RegContent" & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡I"
      txtPath2.SetFocus
   Else
      cnnConnection.RollbackTrans
      If Err.Number = -2147217873 Then
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & _
                "·í´Á¤½³ø¼f©w¸¹¼Æ¡]" & strTBD02 & "¡^°Ó¼ÐºØÃþ¡]" & strTBD03 & "¡^" & vbCrLf & vbCrLf & _
                strErrTxt & ": ¹H¤Ï¥²¶·¬°°ß¤@ªº­­¨î±ø¥ó"
      Else
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & _
                "·í´Á¤½³ø¼f©w¸¹¼Æ¡]" & strTBD02 & "¡^°Ó¼ÐºØÃþ¡]" & strTBD03 & "¡^" & vbCrLf & vbCrLf & _
                strErrTxt & Err.Description & vbCrLf & strSql
      End If
      'Add By Sindy 2015/5/13 ³qª¾µ{§ÇÂàÀÉ¦³»~
      If strP22 <> "" Then
         strSubject = "°Ó¼Ð¤½³øÂàÀÉ¦³»~¡I"
         PUB_SendMail strUserNum, strP22, "", strSubject, strSubject, , , , , , , , , , , False
      End If
      '2015/5/13 END
   End If
End Sub

Private Function ChkDataErr(strTMBM01 As String, strTMBM02 As String, strTMBM04 As String) As Boolean
   ChkDataErr = False
   
   CheckOC3
   strSql = "select * from tmbulletin where tmbm04='" & strTMBM04 & "' "
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount <> 0 Then
      strErrTxt = "¥Ó½Ð®×¸¹" & strTMBM04 & "­«½Æ¡I" & _
                  "¡]¬°²Ä" & CheckStr(AdoRecordSet3.Fields("tmbm07")) & "¨÷´Á¡A¼f©w¸¹" & CheckStr(AdoRecordSet3.Fields("tmbm01")) & "¡A°Ó¼ÐºØÃþ" & CheckStr(AdoRecordSet3.Fields("tmbm02")) & "¡^" & vbCrLf & _
                  "½Ð¤WºôÀË¬d¸ê®Æ¬O§_¦³»~¡A­Y¬O¡A½Ð§ó§ï¤½³ø¸ê®Æ¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If

'   CheckOC3
'   strSql = "select * from tmbulletin where tmbm01||tmbm02='" & strTMBM01 & "'||'" & strTMBM02 & "' "
'   AdoRecordSet3.CursorLocation = adUseClient
'   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If AdoRecordSet3.RecordCount <> 0 Then
'      strErrTxt = "¼f©w¸¹¡Ï°Ó¼ÐºØÃþ­«½Æ¡I" & vbCrLf
'      ChkDataErr = True
'      Exit Function
'   End If
   
   'ÀË¬d°ò¥»ÀÉ­Y¤w¦³¼f©w¸¹¥B»PstrTMBM01¤£¦P«h´£¿ô¾Þ§@ªÌ
   CheckOC3
   strSql = "select * from trademark where tm10='000' and tm28='1' and tm12='" & strTMBM04 & "' "
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount <> 0 Then
      If "" & AdoRecordSet3.Fields("TM15") <> "" Then
         If "" & AdoRecordSet3.Fields("TM15") <> strTMBM01 Then
            strErrTxt = "¥Ó½Ð®×¸¹¡]" & strTMBM04 & "¡^¦b°ò¥»ÀÉ¤¤¼f©w¸¹¡]" & CheckStr(AdoRecordSet3.Fields("TM15")) & "¡^¡A»P¤½³ø¤£²Å¡A½Ð¦A½T»{ !" & vbCrLf
            ChkDataErr = True
            Exit Function
         End If
      End If
   End If
End Function

'¦a°Ï¦WºÙ¸ê®ÆÀË®Öªí
Private Sub ReadTxt1(strTBD02 As String, strTBD04 As String, _
   ByRef strTMBM05 As String, strTMBM06 As String, strAChinese1 As String, strAddress1 As String, strTBD03 As String)
Dim i As Integer
Dim rsTmp2 As New ADODB.Recordset
Dim bolWrite As Boolean 'Add By Sindy 2025/2/13
   
   'Modify By Sindy 2018/11/1 ¦]°Ó¼ÐÅv¤H¤¤¤å¤£±Æ°£¥~°ê°Ó¦r¼Ë, ¦¹²M³æ¥~°ê°Ó§¡¤£¥X²{
   If strTMBM05 <> "" Then
      '­Y¬°¥~°ê°Ó¶·­ç°£
      strSql = "SELECT na03 FROM NATION " & _
            "WHERE na03='" & strTMBM05 & "'" & _
            " and substr(na02,1,1)='C'"
      rsTmp2.CursorLocation = adUseClient
      rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp2.RecordCount > 0 Then
         rsTmp2.Close
         Exit Sub
      End If
      rsTmp2.Close
   End If
   '2018/11/1 END
   
   'Add By Sindy 2017/9/13
   '1.ÀË¬d¥Ó½Ð¤H¦WºÙ¤¤¦³¤j³°¦a°Ï¦r¼ËªÌ¥B¦WºÙ¤¤¦³(¬Ù)¦WºÙ(¦p:¼sªF), ´NÂk(¼sªF)¬Ù½s¸¹
   '2.ÀË¬d¥Ó½Ð¤H¦WºÙ¦rÀY¦³¤j³°ªº(¬Ù)¦WºÙ(¦p:¼sªF), ´NÂk(¼sªF)¬Ù½s¸¹
   '3.­Y¤W¦C§¡µL, ¦ý¦³¤j³°¦a°Ï¦r¼ËªÌ, ÂkB00
   If strTMBM05 = "¤¤°ê¤j³°" Or _
      (strTMBM05 = "" And InStr(strAChinese1, "¤j³°¦a°Ï") > 0) Then
      strSql = "SELECT na03 FROM NATION " & _
            "WHERE substr(na02,1,1)='B'" & _
            " and instr('" & strAChinese1 & "',decode(na03,'¤Ñ¬z¥«','¤Ñ¬z','¥_¨Ê¥«','¥_¨Ê','¤W®ü¥«','¤W®ü',na03))>0"
      rsTmp2.CursorLocation = adUseClient
      rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp2.RecordCount > 0 Then
         strTMBM05 = rsTmp2.Fields("NA03")
      ElseIf InStr(strAChinese1, "¤j³°¦a°Ï") > 0 Then
         strTMBM05 = "¤j³°¨ä¥L"
      End If
      rsTmp2.Close
   End If
   '2017/9/13 END
   
   For i = 1 To 7
      strTemp(i) = ""
   Next i
   strTemp(1) = Trim(strTBD02)
   strTemp(2) = Trim(strTBD04)
   strTemp(3) = Trim(strTMBM05)
   strTemp(4) = Trim(strTMBM06)
   strTemp(5) = Trim(strAChinese1)
   strTemp(6) = Trim(strAddress1)
   If Trim(strTBD03) <> "1" And Trim(strTBD03) <> "7" And Trim(strTBD03) <> "8" And Trim(strTBD03) <> "9" Then
      strTemp(7) = "*" & Trim(strTBD03)
      bolWrite = True 'Add By Sindy 2025/2/13
   Else
      strTemp(7) = Trim(strTBD03)
   End If
   
   'Modify By Sindy 2015/10/16 +Or strTemp(3) = "¤¤µØ¥Á°ê" Or strTemp(3) = "¥xÆW"
   If strTemp(3) = "" Or strTemp(3) = "¤¤°ê¤j³°" Or strTemp(3) = "¤¤µØ¥Á°ê" Or strTemp(3) = "¥xÆW" Then
      strTemp(3) = "*" & strTemp(3)
      bolWrite = True 'Add By Sindy 2025/2/13
   End If
   txtChkWord = strTemp(4)
   If InStr(txtChkWord, "?") > 0 Then
      strTemp(4) = "*" & strTemp(4)
      bolWrite = True 'Add By Sindy 2025/2/13
   End If
   
   'Add By Sindy 2025/2/13 ¸ê®Æ¦³°ÝÃD*¤~¼g¤JÀË®Öªí¤¤
   If bolWrite = True Then
   '2025/2/13 END
      strTemp(1) = convForm(CheckStr(strTemp(1)), 10)
      strTemp(2) = convForm(CheckStr(strTemp(2)), 10)
      If strTemp(4) = "" And strTemp(5) = "" And strTemp(6) = "" Then
         strTemp(3) = CheckStr(strTemp(3))
      Else
         strTemp(3) = convForm(CheckStr(strTemp(3)), 15)
      End If
      strTemp(4) = convForm(CheckStr(strTemp(4)), 12)
      strTemp(5) = convForm(CheckStr(strTemp(5)), 45)
      strTemp(6) = convForm(CheckStr(strTemp(6)), 40)
      strTemp(7) = convForm(CheckStr(strTemp(7)), 4)
      
      If m_PrintRpt1 = False Then
         m_PrintRpt1 = True
   '      If ff1 > 0 Then Close #ff1
   '      ff1 = FreeFile
         m_strFileName1 = "°Ó¼Ð¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¸ê®ÆÀË®Öªí.txt"
   '      Open PUB_Getdesktop & "\" & m_strFileName1 For Output As ff1
   '      Print #ff1, "³Æµù¡G§ï¦r«¬Fixedsys¼Ð·Ç11¸¹¦r¥H¾î¦¡¤W¤U¥ª¥k¦U10MM¦C¦L"
   '      Print #ff1, "¼f©w¸¹¼Æ   ¥Ó½Ð®×¸¹   ºØÃþ ¦a°Ï¦WºÙ        ¥N²z¤H¦WºÙ   °Ó¼ÐÅv¤H¤¤¤å                                  °Ó¼ÐÅv¤H¦a§}"
   '      Print #ff1, "                           ©Î ´£¿ô³Æµù"
   '      Print #ff1, "========== ========== ==== =============== ============ ============================================= ========================================"
         
         m_strText = "³Æµù¡G§ï¦r«¬Fixedsys¼Ð·Ç11¸¹¦r¥H¾î¦¡¤W¤U¥ª¥k¦U10MM¦C¦L" & vbCrLf
         m_strText = m_strText & "¼f©w¸¹¼Æ   ¥Ó½Ð®×¸¹   ºØÃþ ¦a°Ï¦WºÙ        ¥N²z¤H¦WºÙ   °Ó¼ÐÅv¤H¤¤¤å                                  °Ó¼ÐÅv¤H¦a§}" & vbCrLf
         m_strText = m_strText & "                           ©Î ´£¿ô³Æµù" & vbCrLf
         m_strText = m_strText & "========== ========== ==== =============== ============ ============================================= ========================================" & vbCrLf
      End If
      
   '   Print #ff1, strTemp(1) & " " & strTemp(2) & " " & strTemp(7) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6)
      m_strText = m_strText & strTemp(1) & " " & strTemp(2) & " " & strTemp(7) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6) & vbCrLf
   End If
End Sub

'¶}©Ý¸ê®ÆÀË®Öªí
Private Sub ReadTxt2(strTBD02 As String, strTBD04 As String, strTMBM05 As String, strTMBM06 As String, strAChinese1 As String, strAddress1 As String, strTBD03 As String)
Dim i As Integer
   
   If m_PrintRpt2 = False Then
      m_PrintRpt2 = True
      If FF2 > 0 Then Close #FF2
      FF2 = FreeFile
      m_strFileName2 = "°Ó¼Ð¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¶}©Ý¸ê®ÆÀË®Öªí.txt"
      Open PUB_Getdesktop & "\" & m_strFileName2 For Output As FF2
      Print #FF2, "³Æµù¡G§ï¦r«¬Fixedsys¼Ð·Ç11¸¹¦r¥H¾î¦¡¤W¤U¥ª¥k¦U10MM¦C¦L"
      Print #FF2, "¼f©w¸¹¼Æ   ¥Ó½Ð®×¸¹   ºØÃþ ¦a°Ï¦WºÙ        ¥N²z¤H¦WºÙ   °Ó¼ÐÅv¤H¤¤¤å                                  °Ó¼ÐÅv¤H¦a§}"
      Print #FF2, "========== ========== ==== =============== ============ ============================================= ========================================"
   End If
   For i = 1 To 7
      strTemp(i) = ""
   Next i
   strTemp(1) = Trim(strTBD02)
   strTemp(2) = Trim(strTBD04)
   strTemp(3) = Trim(strTMBM05)
   strTemp(4) = Trim(strTMBM06)
   strTemp(5) = Trim(strAChinese1)
   strTemp(6) = Trim(strAddress1)
   strTemp(7) = Trim(strTBD03)
   
   txtChkWord = strTemp(5)
   If InStr(txtChkWord, "?") > 0 Then
      strTemp(5) = "*" & strTemp(5)
   End If
   txtChkWord = strTemp(6)
   If InStr(txtChkWord, "?") > 0 Then
      strTemp(6) = "*" & strTemp(6)
   End If
   
   strTemp(1) = convForm(CheckStr(strTemp(1)), 10)
   strTemp(2) = convForm(CheckStr(strTemp(2)), 10)
   strTemp(3) = convForm(CheckStr(strTemp(3)), 15)
   strTemp(4) = convForm(CheckStr(strTemp(4)), 12)
   strTemp(5) = convForm(CheckStr(strTemp(5)), 45)
   strTemp(6) = convForm(CheckStr(strTemp(6)), 40)
   strTemp(7) = convForm(CheckStr(strTemp(7)), 4)
   Print #FF2, strTemp(1) & " " & strTemp(2) & " " & strTemp(7) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6)
End Sub

'·s¼W¥¢±Ñ°O¿ýÀÉ
Private Sub ReadTxt3(strSql As String)
   If m_PrintRpt3 = False Then
      m_PrintRpt3 = True
      If ff3 > 0 Then Close #ff3
      ff3 = FreeFile
      m_strFileName3 = "°Ó¼Ð¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "·s¼W¥¢±Ñ°O¿ýÀÉ.txt"
      Open PUB_Getdesktop & "\" & m_strFileName3 For Output As ff3
   End If
   Print #ff3, strSql
End Sub

'­­©w¦r¦êªø«×
'Remove by Lydia 2018/08/24 »PbasQuery­«½Æ
'Private Function convForm(ByVal p_InStr As String, ByVal p_Num As Integer, Optional ByVal p_Char As String = " ") As String
'   convForm = StrConv(LeftB(StrConv(p_InStr & String(p_Num, p_Char), vbFromUnicode), p_Num), vbUnicode)
'End Function

'ºI¨úXML¸ê®Æ
Private Function GetXmlData(dblChar As Double, strText As String, strTitNM As String, bolPrintChk As Boolean, ByRef strData As String, ByRef dblEnd As Double) As Boolean
Dim dblStar As Double
   
   GetXmlData = False
   strData = "": dblEnd = 0
   dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
   If dblStar <= dblChar Then
      Exit Function
   End If
   dblEnd = InStr(dblStar, m_strTextBox, "</" & strText & ">") - 1
   If dblStar >= dblEnd Or dblEnd <= 0 Then
      Exit Function
   End If
   strData = Trim(Mid(m_strTextBox, dblStar + 1, (dblEnd - dblStar)))
   strData = Replace(ChgSQL(strData), "amp;", "")
   GetXmlData = True
'   If bolPrintChk = True And InStr(strData, "?") > 0 Then
'      m_bolCharQ = True
'      If m_strCharQNote = "" Then
'         m_strCharQNote = strTitNM
'      Else
'         m_strCharQNote = m_strCharQNote & "," & strTitNM
'      End If
'   End If
End Function

Private Sub Command1_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim dblCnt As Double
   
   ' ¤½³ø¨÷´Á¤£¥iªÅ¥Õ
   If IsEmptyText(txtTMBM07) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤J¤½³ø¨÷´Á¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      txtTMBM07.SetFocus
      Exit Sub
   End If
   
   strSql = "select * from TMBulletin " & _
            "Where TMBM07='" & txtTMBM07 & "' order by TMBM01 asc "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      dblCnt = Val(RsTemp.Fields("TMBM01"))
      Do While Not RsTemp.EOF
         If Val(RsTemp.Fields("TMBM01")) <> dblCnt Then
            MsgBox "¸õ¸¹¡G" & dblCnt
            dblCnt = dblCnt + 1
            'Exit Sub
         End If
         dblCnt = dblCnt + 1
         RsTemp.MoveNext
      Loop
   End If
   MsgBox "ÀË¬d§¹²¦!"
End Sub

Private Sub Command2_Click()
Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.xml"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "XMLÀÉ®× (*.xml)|*.xml"
      .InitDir = PUB_Getdesktop
      '.MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         txtPath1.Text = .FileName
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Me.Height = 5000
   
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
   
   'Add By Sindy 2017/4/25
   If Pub_StrUserSt03 = "M51" Then
      cmdTemp.Visible = True
   End If
   '2017/4/25 END
   
   'Add By Sindy 2022/3/3
   Set adoStream = New ADODB.Stream
   adoStream.Charset = "UTF-8" '"UTF-8" Unicode
   adoStream.Open
   '2022/3/3 END
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2022/3/3
   adoStream.Close
   Set adoStream = Nothing
   '2022/3/3 END
   Set frm030616 = Nothing
End Sub

'Add By Sindy 2018/12/12
Private Sub Option1_Click(Index As Integer)
   If Pub_StrUserSt03 = "M51" Then
      Option1(0).Visible = True '¤½³ø
      Option1(1).Visible = True 'ºM¤T
      Label11.Visible = True
      txtPath3.Visible = True
   End If
   If Option1(0).Value = True Then '¤½³ø
      Label10.Visible = False
      txtTBD17.Visible = False
      Label6.Visible = True '³Æµù¡G¦P®É§ó·s°Ó¼Ð°ò¥»ÀÉªº¼f©w¸¹
      Command1.Visible = True 'ÀË¬d¼f©w¸¹¼Æ¸õ¸¹
   Else 'ºM¤T
      Label10.Visible = True
      txtTBD17.Visible = True
      Label6.Visible = False
      Command1.Visible = False
   End If
End Sub

Private Sub txtPath1_GotFocus()
   InverseTextBox txtPath1
End Sub

Private Sub txtPath2_GotFocus()
   InverseTextBox txtPath2
End Sub

Private Sub txtTMBM07_GotFocus()
   InverseTextBox txtTMBM07
End Sub

Private Sub txtTMBM07_LostFocus()
'   If txtTMBM07 <> "" Then
'      txtPath1 = "D:\°Ó¼Ð¤½³ø" & Val(Left(txtTMBM07, Len(txtTMBM07) - 2)) & "¨÷" & Val(Right(txtTMBM07, 2)) & "´Á"
'   End If
End Sub

' ¤½³ø¨÷´Á
Private Sub txtTMBM07_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Label3.Caption = "(               µ§)"
   Cancel = False
   If IsEmptyText(txtTMBM07) = False Then
      If IsNumeric(txtTMBM07) = False Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¤½³ø¨÷´Á¥u¥i¿é¤J¼Æ­È¸ê®Æ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTMBM07_GotFocus
         Exit Sub
      ElseIf Val(Right(Me.txtTMBM07.Text, 2)) < 1 Or Val(Right(Me.txtTMBM07.Text, 2)) > 24 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¤½³ø´Á¼Æ¿é¤J¿ù»~!!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTMBM07_GotFocus
         Exit Sub
      End If
      'Modify By Sindy 2018/12/13
      If Option1(1).Value = True And txtTBD17 <> "" Then 'ºM¤T ÂàÀÉ
         If Left(GetTA05, 5) <> txtTBD17 Then
            Cancel = True
            strTit = "ÀË®Ö¸ê®Æ"
            strMsg = "¤½³ø¨÷´Á¿é¤J¿ù»~!!!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            txtTMBM07_GotFocus
            Exit Sub
         End If
         Call IsRecordExist_Three
      Else
      '2018/12/13 END
         Call IsRecordExist
      End If
   End If
End Sub

Private Function TxtValidate() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim Cancel As Boolean

TxtValidate = False

' ¤½³ø¨÷´Á¤£¥iªÅ¥Õ
If IsEmptyText(txtTMBM07) = True Then
   strTit = "ÀË®Ö¸ê®Æ"
   strMsg = "½Ð¿é¤J¤½³ø¨÷´Á¡I"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   txtTMBM07.SetFocus
   Exit Function
End If

If Me.txtTMBM07.Enabled = True Then
   Cancel = False
   txtTMBM07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If IsEmptyText(txtPath2) = True Then
   strTit = "ÀË®Ö¸ê®Æ"
   strMsg = "½Ð¿é¤J¥úºÐ¥Øªº¸ô®|¡I"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   txtPath2.SetFocus
   Exit Function
End If

TxtValidate = True
End Function

'Add By Sindy 2018/12/13
' ÀË¬d°O¿ý¬O§_¤w¸g¦s¦b
Private Function IsRecordExist_Three() As Boolean
   Dim rsTmp2 As New ADODB.Recordset
   Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   IsRecordExist_Three = False
   
   strSql = "SELECT count(*) FROM TMBulletinData" & _
            " WHERE TBD16='2' AND TBD01=" & CNULL(txtTMBM07) & " AND TBD17=" & txtTBD17 + 191100
   
   ' Åª¨ú¸ê®Æ®w
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   IsRecordExist_Three = False
   Label3.Caption = "(               µ§)"
   ' ÀË¬dÅª¨úªº¸ê®Æµ§¼Æ
   If rsTmp2.RecordCount > 0 Then
      If rsTmp2.Fields(0) > 0 Then
         IsRecordExist_Three = True
         Label3.Caption = "(  " & rsTmp2.Fields(0) & "  µ§)"
      End If
   End If
   rsTmp2.Close
   
   Set rsTmp2 = Nothing
   Screen.MousePointer = vbDefault
End Function

' ÀË¬d°O¿ý¬O§_¤w¸g¦s¦b
Private Function IsRecordExist() As Boolean
   Dim rsTmp2 As New ADODB.Recordset
   Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   IsRecordExist = False
   
   strSql = "SELECT count(*) FROM TMBulletin WHERE TMBM07=" & CNULL(txtTMBM07)
   
   ' Åª¨ú¸ê®Æ®w
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   IsRecordExist = False
   Label3.Caption = "(               µ§)"
   ' ÀË¬dÅª¨úªº¸ê®Æµ§¼Æ
   If rsTmp2.RecordCount > 0 Then
      If rsTmp2.Fields(0) > 0 Then
         IsRecordExist = True
         Label3.Caption = "(  " & rsTmp2.Fields(0) & "  µ§)"
      End If
   End If
   rsTmp2.Close
   
   Set rsTmp2 = Nothing
   Screen.MousePointer = vbDefault
End Function

' ¨ú±o¤½³ø¥N²z¤Hªº¦WºÙ
Private Function GetTAgentName(ByVal strData As String) As String
Dim strSql As String
Dim rsTmp2 As New ADODB.Recordset
   
   GetTAgentName = Empty
   strSql = "SELECT * FROM TAGENT " & _
            "WHERE TA01 = 'T' AND " & _
                  "TA02 = '" & strData & "' "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      If IsNull(rsTmp2.Fields("TA03")) = False Then
         GetTAgentName = rsTmp2.Fields("TA03")
      End If
   End If
   rsTmp2.Close
   Set rsTmp2 = Nothing
End Function

' ¨ú±o¥X¦W¥N²z¤H¦WºÙ
Private Function GetTOurAgentName() As String
Dim strSql As String
Dim rsTmp2 As New ADODB.Recordset
   
   GetTOurAgentName = Empty
   strSql = "SELECT distinct ST02 FROM ouragent,staff " & _
            "where OA01 in('T','FCT') " & _
            "and OA02=ST01 "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      rsTmp2.MoveFirst
      Do While Not rsTmp2.EOF
         If Not IsNull(rsTmp2.Fields(0)) Then
            GetTOurAgentName = GetTOurAgentName & Trim(rsTmp2.Fields(0)) & ","
         End If
         rsTmp2.MoveNext
      Loop
   End If
   rsTmp2.Close
   Set rsTmp2 = Nothing
End Function

'¥N²z¤H¥N¸¹¦Û°Êµ¹¸¹
Private Function GetFreeAgentCode() As String
Dim rsTmp2 As New ADODB.Recordset
Dim strSql As String
   
   strSql = "SELECT max(to_number(ta02)) FROM TAgent WHERE TA01 = 'T'"
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      GetFreeAgentCode = Val(rsTmp2.Fields(0)) + 1
   Else
      GetFreeAgentCode = "01" '°ò¥»¤WÀ³¸Ó¤£·|Åª¨ì¦¹¬qµ{¦¡
   End If
   Set rsTmp2 = Nothing
End Function

'±N¤½³ø¨÷´ÁÂà´«¬°¤é´Á
Private Function ChgTMBM07ToDate()
Dim strYY As String
Dim strMM As String
Dim strDD As String
'920101 : 3001, 920116 : 3002 ...(¨C¦~·|¦³24´Á)

strYY = (Val(Mid(txtTMBM07, 1, Len(txtTMBM07) - 2)) + 62)
strMM = Format(Right(txtTMBM07, 2) / 2, "00")
If Right(txtTMBM07, 2) Mod 2 <> 0 Then
    strDD = "01"
Else
    strDD = "16"
End If
ChgTMBM07ToDate = DBDATE(strYY & strMM & strDD)
End Function

'¨ú±o¤½§i¤é
Private Function GetTA05() As String
Dim strTemp As String
   
   GetTA05 = CStr(Val(Left(Trim(txtTMBM07), 2)) + 62)
   Select Case CStr(Right(Trim(txtTMBM07), 2))
   Case "01"
      strTemp = "0101"
   Case "02"
      strTemp = "0116"
   Case "03"
      strTemp = "0201"
   Case "04"
      strTemp = "0216"
   Case "05"
      strTemp = "0301"
   Case "06"
      strTemp = "0316"
   Case "07"
      strTemp = "0401"
   Case "08"
      strTemp = "0416"
   Case "09"
      strTemp = "0501"
   Case "10"
      strTemp = "0516"
   Case "11"
      strTemp = "0601"
   Case "12"
      strTemp = "0616"
   Case "13"
      strTemp = "0701"
   Case "14"
      strTemp = "0716"
   Case "15"
      strTemp = "0801"
   Case "16"
      strTemp = "0816"
   Case "17"
      strTemp = "0901"
   Case "18"
      strTemp = "0916"
   Case "19"
      strTemp = "1001"
   Case "20"
      strTemp = "1016"
   Case "21"
      strTemp = "1101"
   Case "22"
      strTemp = "1116"
   Case "23"
      strTemp = "1201"
   Case "24"
      strTemp = "1216"
   End Select
   GetTA05 = GetTA05 & strTemp
End Function

' ¨ú±o°ê®aªº¥N½X
Private Function GetNationNo(ByRef strData As String) As String
Dim strSql As String
Dim rsTmp2 As New ADODB.Recordset
Dim arrData, i As Integer 'Add By Sindy 2013/3/19
   
   GetNationNo = Empty
   
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA03 = '" & strData & "' AND length(na01)=3 "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      If IsNull(rsTmp2.Fields("NA01")) = False Then
         GetNationNo = rsTmp2.Fields("NA01")
         strData = rsTmp2.Fields("NA03")
         If Left("" & rsTmp2.Fields("NA02"), 1) = "A" Then
            bolIsTaiwanCase = True
         ElseIf Left("" & rsTmp2.Fields("NA02"), 1) = "B" Then
            bolIsChinaCase = True
         End If
      End If
   End If
   rsTmp2.Close
   
   If GetNationNo = "" Then
'      strSql = "SELECT * FROM NATION " & _
'               "WHERE NA70 = '" & strData & "' "
      'Modify By Sindy 2013/3/5 NA70·|¦s©ñ¦h­Ó¤½³ø¦a°Ï¦WºÙ
      strSql = "SELECT * FROM NATION " & _
               "WHERE instr(NA70,'" & strData & "')>0 AND length(na01)=3 " & _
               "union SELECT * FROM NATION " & _
               "WHERE instr('" & strData & "',NA70)>0 AND length(na01)=3 "
      rsTmp2.CursorLocation = adUseClient
      rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp2.RecordCount > 0 Then
         'Modify By Sindy 2013/3/19
'         If IsNull(rsTmp2.Fields("NA01")) = False Then
'            GetNationNo = rsTmp2.Fields("NA01")
'            strData = rsTmp2.Fields("NA03")
'            If Left("" & rsTmp2.Fields("NA02"), 1) = "A" Then
'               bolIsTaiwanCase = True
'            ElseIf Left("" & rsTmp2.Fields("NA02"), 1) = "B" Then
'               bolIsChinaCase = True
'            End If
'         End If
         rsTmp2.MoveFirst
         Do While Not rsTmp2.EOF
            arrData = Split(rsTmp2.Fields("NA70"), ",")
            For i = 0 To UBound(arrData)
               If arrData(i) = strData Then
                  GetNationNo = rsTmp2.Fields("NA01")
                  strData = rsTmp2.Fields("NA03")
                  If Left("" & rsTmp2.Fields("NA02"), 1) = "A" Then
                     bolIsTaiwanCase = True
                  ElseIf Left("" & rsTmp2.Fields("NA02"), 1) = "B" Then
                     bolIsChinaCase = True
                  End If
               End If
            Next i
            rsTmp2.MoveNext
         Loop
         '2013/3/19 End
      End If
      rsTmp2.Close
   End If
   
   'Modify By Sindy 2025/2/13 NA70·|¦s©ñ¦h­Ó¤½³ø¦a°Ï¦WºÙ
   '¥x¥_¥«,»O¥_¥«
   If GetNationNo = "" Then
      strSql = "SELECT * FROM NATION WHERE na70 is not null and instr(na70,',')>0"
      rsTmp2.CursorLocation = adUseClient
      rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp2.RecordCount > 0 Then
         rsTmp2.MoveFirst
         Do While Not rsTmp2.EOF
            arrData = Split(rsTmp2.Fields("NA70"), ",")
            For i = 0 To UBound(arrData)
               If InStr(strData, arrData(i)) > 0 And arrData(i) <> "" Then
                  'Modify By Sindy 2025/2/17
                  'GetNationNo = rsTmp2.Fields("NA03")
                  GetNationNo = rsTmp2.Fields("NA01")
                  strData = rsTmp2.Fields("NA03")
                  '2025/2/17 END
                  If Left("" & rsTmp2.Fields("NA02"), 1) = "A" Then
                     bolIsTaiwanCase = True
                  ElseIf Left("" & rsTmp2.Fields("NA02"), 1) = "B" Then
                     bolIsChinaCase = True
                  End If
                  rsTmp2.Close
                  Set rsTmp2 = Nothing
                  Exit Function
               End If
            Next i
            rsTmp2.MoveNext
         Loop
      End If
      rsTmp2.Close
   End If
   '2025/2/13 END
   
'   If GetNationNo = "" Then
'      If strData = "«nÁú" Then
'         GetNationNo = "Áú°ê"
'      End If
'   End If
   Set rsTmp2 = Nothing
End Function

' ¼Ò½k¤ñ¹ï¯S®í¦a°Ï¦WºÙ
Private Function GetNationLike(ByVal strData As String) As String
Dim strSql As String
Dim rsTmp2 As New ADODB.Recordset
Dim arrData, i As Integer 'Add By Sindy 2013/3/19
   
   GetNationLike = Empty
   
   strSql = "SELECT NA02,NA03 FROM NATION WHERE instr('" & strData & "',na03)>0 AND length(na01)=3 order by length(na03) desc "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      rsTmp2.MoveFirst
      GetNationLike = rsTmp2.Fields("NA03")
      If Left("" & rsTmp2.Fields("NA02"), 1) = "A" Then
         bolIsTaiwanCase = True
      ElseIf Left("" & rsTmp2.Fields("NA02"), 1) = "B" Then
         bolIsChinaCase = True
      End If
      rsTmp2.Close
      Set rsTmp2 = Nothing
      Exit Function
   End If
   rsTmp2.Close
   
   'Modify By Sindy 2013/3/5 NA70·|¦s©ñ¦h­Ó¤½³ø¦a°Ï¦WºÙ
   'strSql = "SELECT NA02,NA03,NA70 FROM NATION WHERE instr('" & strData & "',na70)>0 order by length(na70) desc "
   strSql = "SELECT NA02,NA03,NA70 FROM NATION WHERE instr('" & strData & "',na70)>0 and instr(na70,',')=0 AND length(na01)=3 order by length(na70) desc" 'Modify By Sindy 2013/3/19
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      rsTmp2.MoveFirst
      GetNationLike = rsTmp2.Fields("NA03")
      If Left("" & rsTmp2.Fields("NA02"), 1) = "A" Then
         bolIsTaiwanCase = True
      ElseIf Left("" & rsTmp2.Fields("NA02"), 1) = "B" Then
         bolIsChinaCase = True
      End If
      rsTmp2.Close
      Set rsTmp2 = Nothing
      Exit Function
   End If
   rsTmp2.Close
   'Add By Sindy 2013/3/19
   strSql = "SELECT NA02,NA03,NA70 FROM NATION WHERE instr(na70,',')>0 AND length(na01)=3 "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      rsTmp2.MoveFirst
      Do While Not rsTmp2.EOF
         arrData = Split(rsTmp2.Fields("NA70"), ",")
         For i = 0 To UBound(arrData)
            If InStr(strData, arrData(i)) > 0 Then
               GetNationLike = rsTmp2.Fields("NA03")
               If Left("" & rsTmp2.Fields("NA02"), 1) = "A" Then
                  bolIsTaiwanCase = True
               ElseIf Left("" & rsTmp2.Fields("NA02"), 1) = "B" Then
                  bolIsChinaCase = True
               End If
               rsTmp2.Close
               Set rsTmp2 = Nothing
               Exit Function
            End If
         Next i
         rsTmp2.MoveNext
      Loop
   End If
   rsTmp2.Close
   '2013/3/19 End
   
   '°w¹ï¤j³°¦a°Ï
   strSql = "SELECT na03 FROM NATION WHERE na02='B00' and na03 like '%¥«' and instr('" & strData & "',replace(na03,'¥«',''))>0 AND length(na01)=3 "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      rsTmp2.MoveFirst
      GetNationLike = rsTmp2.Fields("NA03")
      rsTmp2.Close
      Set rsTmp2 = Nothing
      Exit Function
   End If
   rsTmp2.Close
   
   Set rsTmp2 = Nothing
End Function

'Sub GetPleft()
'PLeft(1) = 500
'PLeft(2) = 1800
'PLeft(3) = 3200
'PLeft(4) = 5000
'PLeft(5) = 6500
'PLeft(6) = 12000
'End Sub
'
'Sub PrintTitle()
'If m_PrintRpt = False Then
'   Printer.EndDoc
'   Printer.Orientation = 2 '1.ª½¦L 2.¾î¦L
'   m_PrintRpt = True
'End If
'
'GetPleft
'iLine = 1
'
'Printer.Font.Size = 16
'Printer.Font.Underline = False
'Printer.FontBold = False
'
'Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("°Ó¼Ð¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¸ê®ÆÀË®Öªí¤@") / 2)
'Printer.CurrentY = iLine * 300
'Printer.Print "°Ó¼Ð¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¸ê®ÆÀË®Öªí"
'
'Printer.Font.Size = 12
'Printer.Font.Underline = False
'Printer.FontBold = False
'
'iLine = iLine + 1
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = 900
'Printer.Print "¦C¦L¤H­û¡G" & strUserName
'Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(strSrvDate(2))) - 500
'Printer.CurrentY = 900
'Printer.Print "¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(strSrvDate(2))
'iLine = iLine + 1
'Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(strSrvDate(2))) - 500
'Printer.CurrentY = 1200
'Printer.Print "­¶¡@¡@¦¸¡G" & Printer.Page
'
'iLine = 5
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = iLine * 300
'Printer.Print "¼f©w¸¹¼Æ"
'Printer.CurrentX = PLeft(2)
'Printer.CurrentY = iLine * 300
'Printer.Print "¥Ó½Ð®×¸¹"
'Printer.CurrentX = PLeft(3)
'Printer.CurrentY = iLine * 300
'Printer.Print "¦a°Ï¦WºÙ"
'Printer.CurrentX = PLeft(4)
'Printer.CurrentY = iLine * 300
'Printer.Print "¥N²z¤H¦WºÙ"
'Printer.CurrentX = PLeft(5)
'Printer.CurrentY = iLine * 300
'Printer.Print "°Ó¼ÐÅv¤H¤¤¤å"
'Printer.CurrentX = PLeft(6)
'Printer.CurrentY = iLine * 300
'Printer.Print "°Ó¼ÐÅv¤H¦a§}"
'
'iLine = iLine + 1
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = iLine * 300
'Printer.Print String(205, "-")
'iLine = iLine + 1
'End Sub
'
'Sub PrintDetail()
'Dim m_j As Integer
'   For m_j = 1 To 6
'      Printer.CurrentX = PLeft(m_j)
'      Printer.CurrentY = iLine * 300
'      Printer.Print strTemp(m_j)
'   Next m_j
'   iLine = iLine + 1
'End Sub
'
'Sub GetPleft2()
'PLeft(1) = 500
'PLeft(2) = 1800
'PLeft(3) = 3200
'PLeft(4) = 5000
''PLeft(5) = 6500
''PLeft(6) = 12000
'End Sub
'
'Sub PrintTitle2()
'If m_PrintRpt = False Then
'   Printer.EndDoc
'   Printer.Orientation = 1 '1.ª½¦L 2.¾î¦L
'   m_PrintRpt = True
'End If
'
'GetPleft2
'iLine2 = 1
'
'Printer.Font.Size = 16
'Printer.Font.Underline = False
'Printer.FontBold = False
'
'Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("°Ó¼Ð¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¸ê®ÆÀË®Öªí¤@") / 2)
'Printer.CurrentY = iLine2 * 300
'Printer.Print "°Ó¼Ð¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "«D¥»©Ò®×¥ó¸ê®ÆÀË®Öªí¤G"
'
'Printer.Font.Size = 12
'Printer.Font.Underline = False
'Printer.FontBold = False
'
'iLine2 = iLine2 + 1
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = 900
'Printer.Print "¦C¦L¤H­û¡G" & strUserName
'Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(strSrvDate(2))) - 500
'Printer.CurrentY = 900
'Printer.Print "¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(strSrvDate(2))
'iLine2 = iLine2 + 1
'Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(strSrvDate(2))) - 500
'Printer.CurrentY = 1200
'Printer.Print "­¶¡@¡@¦¸¡G" & Printer.Page
'
'iLine2 = 5
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = iLine2 * 300
'Printer.Print "¼f©w¸¹¼Æ"
'Printer.CurrentX = PLeft(2)
'Printer.CurrentY = iLine2 * 300
'Printer.Print "¥Ó½Ð®×¸¹"
'Printer.CurrentX = PLeft(3)
'Printer.CurrentY = iLine2 * 300
'Printer.Print "°Ó¼ÐºØÃþ"
'Printer.CurrentX = PLeft(4)
'Printer.CurrentY = iLine2 * 300
'Printer.Print "ÀË®Ö¦³»~(?)ªº¸ê®Æ"
'
'iLine2 = iLine2 + 1
'Printer.CurrentX = PLeft(1)
'Printer.CurrentY = iLine2 * 300
'Printer.Print String(205, "-")
'iLine2 = iLine2 + 1
'End Sub
'
'Sub PrintDetail2()
'Dim m_j As Integer
'   For m_j = 1 To 4
'      Printer.CurrentX = PLeft(m_j)
'      Printer.CurrentY = iLine2 * 300
'      Printer.Print strTemp(m_j)
'   Next m_j
'   iLine2 = iLine2 + 1
'End Sub
