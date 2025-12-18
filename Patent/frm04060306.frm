VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060306 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "±M§Q¤½¶}¤½³øÂàÀÉ§@·~"
   ClientHeight    =   5640
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   5940
   Begin VB.CommandButton cmdPath 
      Height          =   330
      Left            =   5490
      Picture         =   "frm04060306.frx":0000
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   20
      Top             =   780
      Width           =   350
   End
   Begin VB.CommandButton cmdIPC 
      Caption         =   "¸ÉÂà¥¼¤ÀÃþªºIPC¤ÀÃþ"
      Height          =   345
      Left            =   3747
      TabIndex        =   19
      Top             =   1530
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.CommandButton cmdPA160 
      Caption         =   "¸ÉÂà®×¥óÄÝ©Ê"
      Height          =   400
      Left            =   4257
      TabIndex        =   18
      Top             =   2610
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1515
      Left            =   60
      TabIndex        =   17
      Top             =   4050
      Width           =   5745
      _ExtentX        =   10139
      _ExtentY        =   2667
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "·s²Ó©úÅé-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox text03 
      Height          =   264
      Left            =   1050
      MaxLength       =   7
      TabIndex        =   1
      Top             =   2370
      Width           =   1092
   End
   Begin VB.CommandButton cmdTransFile 
      Caption         =   "ÂàÀÉ(&T)"
      Height          =   400
      Left            =   3450
      TabIndex        =   5
      Top             =   1950
      Width           =   912
   End
   Begin VB.TextBox txtTMBM07 
      Height          =   264
      Left            =   1050
      MaxLength       =   4
      TabIndex        =   0
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   60
      TabIndex        =   11
      Top             =   3450
      Width           =   5805
      Begin VB.TextBox Text2 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H00FF0000&
         Height          =   300
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   5730
      End
   End
   Begin VB.FileListBox File2 
      Height          =   180
      Left            =   1560
      TabIndex        =   10
      Top             =   210
      Visible         =   0   'False
      Width           =   525
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   405
      Left            =   960
      TabIndex        =   9
      Top             =   210
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   720
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frm04060306.frx":0102
   End
   Begin VB.TextBox txtPath2 
      Height          =   315
      Left            =   1410
      TabIndex        =   3
      Text            =   "C:\GAZETTE\PGXml"
      Top             =   1140
      Width           =   4065
   End
   Begin VB.TextBox txtPath1 
      Height          =   315
      Left            =   1410
      TabIndex        =   2
      Text            =   "E:"
      Top             =   810
      Width           =   4065
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "«þ¨©¥úºÐ¸ê®Æ(&C)"
      Height          =   400
      Left            =   3300
      TabIndex        =   4
      Top             =   180
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4920
      TabIndex        =   6
      Top             =   180
      Width           =   912
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   270
      Top             =   210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.TextBox txtChkWord 
      Height          =   300
      Left            =   0
      TabIndex        =   21
      Top             =   0
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "¤½¶}¤é¡G"
      Height          =   180
      Left            =   300
      TabIndex        =   16
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "¤½³ø¨÷´Á¡G"
      Height          =   210
      Left            =   120
      TabIndex        =   15
      Top             =   2070
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "(               µ§)"
      Height          =   210
      Left            =   2190
      TabIndex        =   14
      Top             =   2070
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "ÂàÀÉ¤¤, ½Ðµy­Ô. . .(½Ð¤Å¥ô·NÃö³¬¦¹§@·~)"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   15.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   60
      TabIndex        =   12
      Top             =   3090
      Width           =   5835
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "«þ¨©¥Øªº¸ô®|¡G"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ÀÉ®×¨Ó·½¸ô®|¡G"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1260
   End
End
Attribute VB_Name = "frm04060306"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/3/3 Form2.0¤w­×§ï
'Memo by Morgan 2022/1/3 §ï¦¨Form2.0 (MSHFlexGrid1,Printer¦C¦L¥¼§ï)
'Memo By Morgan 2012/12/11 ´¼Åv¤H­ûÄæ¤w­×§ï
Option Explicit

Dim m_bolCharQ  As Boolean, m_strCharQNote As String
Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 8) As String
Dim iLine2 As Integer
Dim m_PrintRpt1 As Boolean, m_PrintRpt2 As Boolean
Dim ff1 As Integer
Dim m_strFileName1 As String, m_strFileName2 As String
Dim strErrTxt As String
Dim strTPG01 As String, strTPG02 As String, dblTPG03 As Double, strTPG04 As String
Dim strTPG05 As String, strTPG06 As String, strTPG07 As String, strTPG07_1 As String, strTPG07_temp1 As String
Dim strTPG08 As String, strTPG09 As String
Dim strAChinese As String, strAChinese1 As String, strAddress1 As String
Dim strOurAgentName As String
Dim pa() As String
Dim bolTaieCase As Boolean '¬O§_¬°¥»©Ò®×¥ó
Dim strTaieCaseNo As String
Dim strChkTPG04 As String, strChkTPG05 As String
Dim strTPG11 As String, strTPG12 As String, strTPG13 As String, strTPG14 As String
'Dim strTestTPG01 As String
Dim m_DefaultPrinter As String
Dim SeekPrint As Integer
'Add By Sindy 2012/1/16
Dim intPRow As Integer
Dim MaxHeight As Integer, MinHeight As Integer
'2012/1/16 End
'Add By Sindy 2013/8/27
Dim strTPG15 As String, strTPG16 As String, m_PI02 As String, strTPG17 As String
'2013/8/27 END
Dim strTPG18 As String 'Add By Sindy 2016/3/2
'Add By Sindy 2015/6/9 ¤ñ¹ï¹q¤lÀÉ¤º®e»P¥»©Ò®×¥ó©Ò«Ø¸ê®Æ¬O§_¤@­P
Dim strCaseChNm As String, strCaseEnNm As String 'µo©ú¤¤­^¤å¦WºÙ
Dim strApplDate As String '¥Ó½Ð¤é
Dim strAEng As String '¥Ó½Ð¤H­^¤å¦WºÙ
Dim strAEnCountry As String '¥Ó½Ð¤H°êÄy
Dim strApplName As String '¥Ó½Ð¤H
Dim strInventor As String 'µo©ú¤H
Dim strAgent As String '¥N²z¤H
Dim strClaims As String 'Àu¥ýÅv
Dim strGetData1 As String, strGetData2 As String, strGetData3 As String
'2015/6/9 END
'Add By Sindy 2018/11/12
Dim strTPGcApp(10) As String
Dim strTPGeApp(10) As String
Dim dblTPG39 As Double, dblTPG40 As Double, strTPG41 As String, strTPG42 As String
'2018/11/12 END
Dim strTPG43 As String 'Add By Sindy 2019/9/4
Dim adoStream As ADODB.Stream 'Add By Sindy 2022/3/3
Dim m_strTextBox As String 'Add by Sindy 2022/3/3
Dim m_strText As String 'Add by Sindy 2024/5/17

Private Sub cmdCopy_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim fs As Object, strTime As String
Dim DeleteFilePathErr As Boolean
   
On Error GoTo ErrHnd
   
   strTime = time()
   DeleteFilePathErr = False
   
   If IsEmptyText(txtPath1) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      'strMsg = "½Ð¿é¤J¥úºÐ¨Ó·½¸ô®|¡I"
      strMsg = "½Ð¿é¤JÀÉ®×¨Ó·½¸ô®|¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      txtPath1.SetFocus
      Exit Sub
   End If
   If IsEmptyText(txtPath2) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      'strMsg = "½Ð¿é¤J¥úºÐ¥Øªº¸ô®|¡I"
      strMsg = "½Ð¿é¤J«þ¨©¥Øªº¸ô®|¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      txtPath2.SetFocus
      Exit Sub
   End If
   If IsEmptyText(txtTMBM07) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤J¤½³ø¨÷´Á¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      txtTMBM07.SetFocus
      Exit Sub
   End If
   If IsEmptyText(text03) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤J¤½¶}¤é¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text03.SetFocus
      Exit Sub
   End If
   Call GetNoticeNumber(DBDATE(text03)) '¨Ì¿é¤Jªº¤½¶}¤é¨ú±o¬Û¹ïªº¤½§i¨÷´Á
   If Val(Left(txtTMBM07, 2)) <> Val(strChkTPG04) Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "¤½³ø¨÷¼Æ»P¤½¶}¤é´Á¤£²Å¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text03.SetFocus
      Exit Sub
   End If
   If Val(Right(txtTMBM07, 2)) <> Val(strChkTPG05) Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "¤½³ø´Á¼Æ»P¤½¶}¤é´Á¤£²Å¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text03.SetFocus
      Exit Sub
   End If
      
   If Right(Trim(txtPath1), 1) = "\" Then txtPath1 = Left(txtPath1, Len(txtPath1) - 1)
   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   Set fs = CreateObject("Scripting.FileSystemObject")
   
   'Add By Sindy 2020/5/11 ¥ý²M°£¸ÑÀ£ÁY«á,ÂÂªº¸ê®Æ§¨,¥H¨¾ªÅ¶¡¤£¨¬
   If Dir(txtPath1 & "\pub*") <> "" Then
      fs.DeleteFolder txtPath1 & "\pub*", True
      Sleep 1000
   End If
   '2020/5/11 END
   
   'Added by Sindy 2020/5/5
   '109/5/11¶}©l¨ú®ø¥úºÐ¡A§ï¤U¸üÀ£ÁYÀÉ
   'ÀË¬d¸ê®Æ§¨¬O§_¦s¦b
   strExc(0) = txtPath1 & "\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000")
   If fs.FolderExists(strExc(0) & "\patent") = False Then
      'ÀË¬dÀ£ÁYÀÉ¬O§_¦s¦b Ex:Pub018009_Publish.zip
      strExc(1) = strExc(0) & "_Publish.zip"
      If fs.FileExists(strExc(1)) = True Then
         PUB_UnZipFile strExc(1), strExc(0)
      Else
         MsgBox "¤½³øÀ£ÁYÀÉ(" & strExc(1) & ")¤£¦s¦b¡I", vbCritical
         Exit Sub
      End If
   End If
   'end 2020/5/5
   
   'Modify By Sindy 2013/1/2
   'File2.path = txtPath1 & "\img_1\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000")
   File2.path = txtPath1 & "\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\patent"
   '2013/1/2 End
   File2.Refresh
   If File2.ListCount = 0 Then
      'Modified by Sindy 2020/5/5
      'MsgBox "¥úºÐ¨Ó·½¸ô®|¤¤µL" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¤½¶}¤½³ø¸ê®Æ¡I"
      MsgBox "¨Ó·½¸ô®|¤¤µL" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¤½¶}¤½³ø¸ê®Æ¡I"
      '2020/5/5 END
      txtPath1.SetFocus
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   'Set fs = CreateObject("Scripting.FileSystemObject") 'Removed by Sindy 2020/5/5 §ï¨ì¤W­±
   DeleteFilePathErr = True
   
   'Modify By Sindy 2012/6/6
   If fs.FolderExists(txtPath2) = True Then
      fs.DeleteFile txtPath2 & "\*.*", True '§R°£XMLÀÉ¤Î°O¿ýª©¥»¤å¦rÀÉ(ver*.txt)
      'ÀË¬d¬O§_¦³±ý«þ¨©·í´ÁªºPDF¸ê®Æ§¨
      If fs.FolderExists(txtPath2 & "\img_1\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000")) = True Then
         fs.DeleteFolder txtPath2 & "\img_1\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000"), True
      End If
      '©T©w§R°£¤W­Ó¤ë¸Ó´ÁPDF¸ê®Æ§¨
      strDate = DBDATE(ChangeWStringToTString(DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(DBDATE(text03))))))
      Call GetNoticeNumber(strDate)
      If fs.FolderExists(txtPath2 & "\img_1\pub" & Format(strChkTPG04, "000") & Format(strChkTPG05, "000")) = True Then
         fs.DeleteFolder txtPath2 & "\img_1\pub" & Format(strChkTPG04, "000") & Format(strChkTPG05, "000"), True
      End If
   End If
   '2012/6/6 End
   'fs.DeleteFolder txtPath2, True
NotFolder76:
   'Modify By Sindy 2012/6/6
   If fs.FolderExists(txtPath2) = False Then
      fs.CreateFolder txtPath2 '¦s©ñXMLÀÉ
      fs.CreateFolder txtPath2 & "\img_1"
   End If
   '2012/6/6 End
   '¦s©ñPDF
   fs.CreateFolder txtPath2 & "\img_1\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000")
   'Modify By Sindy 2013/1/2
   'fs.CopyFile txtPath1 & "\xml\*.*", txtPath2 & "\"
   'fs.CopyFile txtPath1 & "\img_1\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\*.*", txtPath2 & "\img_1\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\"
   fs.CopyFile txtPath1 & "\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\patent\*.*", txtPath2 & "\"
   fs.CopyFile txtPath1 & "\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\sundrydata\*.*", txtPath2 & "\"
   fs.CopyFile txtPath1 & "\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\patent\pdf\*.*", txtPath2 & "\img_1\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\"
   '2013/1/2 End
   'Add By Sindy 2012/6/6
   '²£¥Í°O¿ýXMLª©¥»¤å¦rÀÉ(ver*.txt)
   Dim a As Object
   Set a = fs.CreateTextFile(txtPath2 & "\ver" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000.txt"), True)
   '2012/6/6 End
   Screen.MousePointer = vbDefault
   MsgBox "«þ¨©§¹²¦¡I(«þ¨©ªá¶O®É¶¡¡G" & strTime & "  " & time() & ")"
   Set fs = Nothing
   Exit Sub
   
ErrHnd:
   If Err.NUMBER = 76 And DeleteFilePathErr = True Then
      GoTo NotFolder76
   ElseIf Err.NUMBER = 68 Or Err.NUMBER = 76 Then
      'MsgBox "¥úºÐ¨Ó·½¸ô®|¤¤µL" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¤½¶}¤½³ø¸ê®Æ¡I"
      MsgBox "ÀÉ®×¨Ó·½¸ô®|¤¤µL" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¤½¶}¤½³ø¸ê®Æ¡I"
      txtPath1.SetFocus
   Else
      MsgBox Err.Description
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdIPC_Click()
Dim strDate1 As String, StrDate2 As String
Dim rsTmp As New ADODB.Recordset

On Error GoTo ErrHand
   
   strDate1 = DBDATE(Trim(InputBox("½Ð¿é¤J±ý¸ÉÂàªº°_©l¤½¶}¤é´Á")))
   If Val(strDate1) = 0 Then
      MsgBox "½Ð¿é¤J°_©l¤½¶}¤é´Á!!"
      Exit Sub
   End If
   StrDate2 = DBDATE(Trim(InputBox("½Ð¿é¤J±ý¸ÉÂàªººI¤î¤½¶}¤é´Á")))
   If Val(StrDate2) = 0 Then
      MsgBox "½Ð¿é¤JºI¤î¤½¶}¤é´Á!!"
      Exit Sub
   End If
   
   strSql = "SELECT count(*) FROM TPGazette " & _
            "WHERE TPG03>=" & strDate1 & " AND TPG03<=" & StrDate2 & _
             " AND TPG16 is null"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If rsTmp.RecordCount > 0 Then
         If rsTmp.Fields(0) = 0 Then
            MsgBox "µL«Ý¤ÀÃþªº¸ê®Æ!!"
            Exit Sub
         End If
      End If
   End If
   
   Screen.MousePointer = vbHourglass
   cnnConnection.BeginTrans
   
   strSql = "SELECT * FROM TPGazette " & _
            "WHERE TPG03>=" & strDate1 & " AND TPG03<=" & StrDate2 & _
             " AND TPG16 is null"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         strTPG01 = "": strTPG16 = ""
         
         strTPG01 = rsTmp.Fields("TPG01")
         strTPG16 = GetPatentIPC("1", rsTmp.Fields("TPG15"), "I") 'IPC¤ÀÃþ
         
         If strTPG16 <> "" Then
            strSql = "update TPGazette " & _
                     "set TPG16='" & strTPG16 & "' " & _
                     "where TPG01='" & strTPG01 & "'"
            cnnConnection.Execute strSql
         End If
         
         rsTmp.MoveNext
      Loop
   End If
   
   cnnConnection.CommitTrans
   
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   
   MsgBox "ÂàÀÉ§¹²¦¡I"
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Sub

'Add By Sindy 2016/3/2
'¸ÉÂà®×¥óÄÝ©Ê
Private Sub cmdPA160_Click()
Dim strTime As String
Dim stSQL As String, intR As Integer
Dim rsQuery As ADODB.Recordset
   
On Error GoTo ErrHand
   
   strTime = time()
   
   stSQL = "SELECT TPG01,TPG15,TPG16,TPG18 FROM TPGazette WHERE TPG16 is not null and TPG15 is not null and TPG18 is null"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      Screen.MousePointer = vbHourglass
      With rsQuery
         .MoveFirst
         Do While Not .EOF
            cnnConnection.BeginTrans
            
            strTPG18 = GetPatentIPC("3", .Fields("TPG15"), "")
            
            strSql = "UPDATE TPGazette SET TPG18='" & strTPG18 & "'" & _
                     " WHERE TPG01='" & .Fields("TPG01") & "'"
            cnnConnection.Execute strSql
            
            cnnConnection.CommitTrans
            .MoveNext
         Loop
      End With
      Screen.MousePointer = vbDefault
   End If
   Set rsQuery = Nothing
   
   MsgBox "ÂàÀÉ§¹²¦¡I(ÂàÀÉªá¶O®É¶¡¡G" & strTime & "  " & time() & ")"
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.NUMBER & " " & Err.Description
   End If
End Sub

''Add By Sindy 2013/8/27
'Private Sub cmdPA160_Click()
'Dim strTit As String
'Dim strMsg As String
'Dim nResponse
'Dim dblFCnt As Double
'Dim dblStar As Double, dblEnd As Double
'Dim dblChar As Double, dblLastEnd As Double
'Dim strText As String, strTitNM As String
'Dim strChar As String, strData As String
'Dim strFreeAgentCode As String
'Dim dblMaxWidth As Double
'Dim strTime As String, strTotRow As String
'Dim i As Integer, j As Integer
'Dim fs As Object
'
'On Error GoTo ErrHand
'
'   strTime = time()
'
'   '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
'   If TxtValidate = False Then Exit Sub
'
'   If IsRecordExist = False Then
'      MsgBox "¤½¶}¤½³ø¨÷´Á" & txtTMBM07 & "¸ê®Æ¤£¦s¦b¡I"
'      txtTMBM07.SetFocus
'      Exit Sub
'   End If
'
'   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
'
'   'ÀË¬d¤½³ø¨÷´Á
'   Set fs = CreateObject("Scripting.FileSystemObject")
'   File2.path = txtPath2.Text
'   File2.Refresh
'   If File2.ListCount = 0 Or _
'      fs.FileExists(txtPath2 & "\ver" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000.txt")) = False Then
'      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "¡^¤ºµL¸Ó´Á¤½¶}¤½³ø¸ê®Æ¡A½Ð¥ý«þ¨©¥úºÐ¸ê®Æ¡I"
'      txtPath2.SetFocus
'      Exit Sub
'   End If
'   Set fs = Nothing
'
'   Screen.MousePointer = vbHourglass
'   cnnConnection.BeginTrans
'
'   Call ResetGrid: intPRow = 0
'   strOurAgentName = GetTOurAgentName()
'   m_PrintRpt1 = False: m_PrintRpt2 = False: iLine2 = 0
'   strTotRow = File2.ListCount
'   Me.Height = MaxHeight
'   dblMaxWidth = 5730
'   Text2.Width = 0
'   Label2.Caption = "ÂàÀÉ¤¤, ½Ðµy­Ô . . ."
'   For dblFCnt = 0 To File2.ListCount - 1
'      'ÀÉ¦W«e3½X¬°sudªÌ¤£¶·Âà¤J¸ê®Æ
'      If (Asc(Left(Trim(File2.List(dblFCnt)), 1)) >= 48 And Asc(Left(Trim(File2.List(dblFCnt)), 1)) <= 57) And _
'         UCase(Right(Trim(File2.List(dblFCnt)), 3)) = "XML" Then
'         RichTextBox1.LoadFile (txtPath2.Text & "\" & File2.List(dblFCnt))
''         RichTextBox1.LoadFile (txtPath2.Text & "\099218880.xml")
'
'         Text2.Width = dblMaxWidth / Val(strTotRow) * (dblFCnt + 1): DoEvents
'
'         If ReadXmlData = False Then GoTo ErrHand
'
'         'Modify By Sindy 2016/3/2 +TPG18
'         strSql = "update TPGazette " & _
'                  "set TPG15='" & strTPG15 & "',TPG16='" & strTPG16 & "',TPG17='" & strTPG17 & "',TPG18='" & strTPG18 & "' " & _
'                  "where TPG01='" & strTPG01 & "'"
'         cnnConnection.Execute strSql
'      End If
'   Next dblFCnt
'
'   cnnConnection.CommitTrans
'
'   Screen.MousePointer = vbDefault
'
''   Call GetSendMailIPC
'   Call IsRecordExist '²£¥Íµ§¼Æ
'   MsgBox "ÂàÀÉ§¹²¦¡I(ÂàÀÉªá¶O®É¶¡¡G" & strTime & "  " & time() & ")" & vbCrLf & strMsg
'   Me.Height = MinHeight
'
'   Exit Sub
'
'ErrHand:
'   Screen.MousePointer = vbDefault
'   If Err.NUMBER = 76 Then
'      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "\img_1\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "¡^¤ºµL¸Ó´Á¤½¶}¤½³ø¸ê®Æ¡I"
'      txtPath2.SetFocus
'   Else
'      cnnConnection.RollbackTrans
'      If Err.NUMBER = -2147217873 Then
'         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½¶}¤½³ø¥Ó½Ð®×¸¹¡]" & strTPG01 & "¡^" & vbCrLf & strErrTxt & ": ¹H¤Ï¥²¶·¬°°ß¤@ªº­­¨î±ø¥ó"
'      Else
'         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½¶}¤½³ø¥Ó½Ð®×¸¹¡]" & strTPG01 & "¡^" & vbCrLf & strErrTxt & Err.Description
'      End If
'   End If
'End Sub

''Add By Sindy 2013/8/27 IPC¤ÀÃþÂkÃþ¤£¨ì®É,³qª¾69009·¨·¶¯Â
''Modify By Sindy 2020/5/13 ·¨·¶¯Â(ºÊ¹î¤H):¤w»P·¨¸g²z°Q½×¹L,¤é«á­Y¤½³øIPC¤ÀÃþ¦³°ÝÃD®É,½Ð¥Ñ¨t²Îª½±µÂàµ¹99033·¨¶²ªÚ¸g²z
'Private Sub GetSendMailIPC()
'   If m_PI02 <> "" Then
'      m_PI02 = Replace(m_PI02, "¡F", vbCrLf)
'      PUB_SendMail strUserNum, "99033;97038", "", "±M§Q¤½¶}¤½³ø" & txtTMBM07 & "´Á¦³°ê»Ú¤ÀÃþ¸¹¡A©|¥¼°µIPC¤ÀÃþ", "Dear Sirs," & vbCrLf & vbCrLf & _
'      "±M§Q¤½¶}¤½³ø" & txtTMBM07 & "´Á¦³°ê»Ú¤ÀÃþ¸¹¡A©|¥¼°µIPC¤ÀÃþ¡A¦p¤U¡G" & vbCrLf & vbCrLf & m_PI02 & vbCrLf & vbCrLf & _
'      "·Ð½Ð¦A³qª¾¹q¸£¤¤¤ßÀ³¦p¦ó¤ÀÃþ¡C" & vbCrLf & vbCrLf & vbCrLf & _
'      "                                                        ¹q¸£¤¤¤ß"
'   End If
'End Sub

'Added by Sindy 2020/5/5
Private Sub cmdPath_Click()
   Dim fName As String, strStartFolder As String
   
   If Dir(txtPath1 & "\", vbDirectory) <> "" Then strStartFolder = txtPath1
   
   fName = PUB_GetFolder(Me.hWnd, strStartFolder, "½Ð¿ï¨ú¸ê®Æ§¨:")
   If fName <> "" Then 'they did not hit cancel
      txtPath1 = fName
   End If
   
End Sub

Private Sub cmdTransFile_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim dblFCnt As Double
Dim dblStar As Double, dblEnd As Double
Dim dblChar As Double, dblLastEnd As Double
Dim strText As String, strTitNM As String
Dim strChar As String, strData As String
Dim rsTmp As New ADODB.Recordset
Dim strFreeAgentCode As String
Dim dblMaxWidth As Double
Dim strTime As String, strTotRow As String
Dim i As Integer, j As Integer
Dim fs As Object
Dim stCP12 As String, stCP13 As String, stCP09 As String, strFileName As String, strCP10 As String
Dim f
Dim bolTa04IsNull As Boolean 'Add By Sindy 2014/9/3
Dim TempFileName As String, strSys As String, strTo As String, ff As Integer
Dim arrData As Variant, arrData_1 As Variant
Dim strCP14, strCP48 As String 'Added by Lydia 2019/05/31 ¹w³]©Ó¿ì¤H©M©Ó¿ì´Á­­
   
On Error GoTo ErrHand
   
   strTime = time()
   
   '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
   If TxtValidate = False Then Exit Sub
   
   If IsRecordExist = True Then
      strTit = "¸ß°Ý"
      strMsg = "¤½¶}¤½³ø¨÷´Á" & txtTMBM07 & "¤w¦³¸ê®Æ¦s¦b¡A½T©w¬O§_­n­«·sÂàÀÉ¡H"
      nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
      If nResponse = vbNo Then Exit Sub
   End If
   
   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   
   'ÀË¬d¤½³ø¨÷´Á
   Set fs = CreateObject("Scripting.FileSystemObject")
   File2.path = txtPath2.Text
   File2.Refresh
   If File2.ListCount = 0 Or _
      fs.FileExists(txtPath2 & "\ver" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000.txt")) = False Then
      'MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "¡^¤ºµL¸Ó´Á¤½¶}¤½³ø¸ê®Æ¡A½Ð¥ý«þ¨©¥úºÐ¸ê®Æ¡I"
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "¡^¤ºµL¸Ó´Á¤½¶}¤½³ø¸ê®Æ¡A½Ð¥ý«þ¨©ÀÉ®×¸ê®Æ¡I"
      txtPath2.SetFocus
      Exit Sub
   End If
   'Set fs = Nothing
   
   Screen.MousePointer = vbHourglass
   
   'Add By Sindy 2015/6/11
   strSql = "delete FROM R04060306"
   cnnConnection.Execute strSql
   '2015/6/11 END
   strSql = "delete FROM TPGazette WHERE TPG04=" & CNULL(Left(txtTMBM07, 2)) & " and TPG05=" & CNULL(Right(txtTMBM07, 2))
   cnnConnection.Execute strSql
   
   Call ResetGrid: intPRow = 0 'Add By Sindy 2012/1/16
   strOurAgentName = GetTOurAgentName()
   m_PrintRpt1 = False: m_PrintRpt2 = False: iLine2 = 0
   strTotRow = File2.ListCount
   Me.Height = MaxHeight
   dblMaxWidth = 5730
   Text2.Width = 0
   Label2.Caption = "ÂàÀÉ¤¤, ½Ðµy­Ô . . ."
   For dblFCnt = 0 To File2.ListCount - 1
      'ÀÉ¦W«e3½X¬°sudªÌ¤£¶·Âà¤J¸ê®Æ
      If (Asc(Left(Trim(File2.List(dblFCnt)), 1)) >= 48 And Asc(Left(Trim(File2.List(dblFCnt)), 1)) <= 57) And _
         UCase(Right(Trim(File2.List(dblFCnt)), 3)) = "XML" Then
         
         'Add by Sindy 2022/3/3
         If strSrvDate(1) >= Form20¤W½u¤é Then
            adoStream.LoadFromFile (txtPath2.Text & "\" & File2.List(dblFCnt))
            m_strTextBox = adoStream.ReadText
         Else
         '2022/3/3 END
            RichTextBox1.LoadFile (txtPath2.Text & "\" & File2.List(dblFCnt))
            m_strTextBox = RichTextBox1.Text
         End If
         
         Text2.Width = dblMaxWidth / Val(strTotRow) * (dblFCnt + 1): DoEvents
         
         cnnConnection.BeginTrans
         
         If ReadXmlData = False Then GoTo ErrHand 'Modify By Sindy 2013/8/27 ²¾¦Ü¨ç¼Æ
         
'         If strTPG01 = "102141870" Then
'            MsgBox strTPG01
'         End If
         
         If ChkDataErr() = True Then GoTo ErrHand
         
         '¦a°Ï¦WºÙ¬°ªÅ¥Õ©Î020.¤¤°ê¤j³°,¥N²z¤H¦WºÙ¦³?®É,»Ý¦C¦L²M³æ (Or strTPG06 = "020")
         'Modify By Sindy 2015/9/23 +strTPG06 = "000"
         'Modify By Sindy 2019/9/4 + Or strTPG43 = "" Or strTPG43 = "¤¤µØ¥Á°ê" Or strTPG43 = "¥xÆW"
         txtChkWord = strTPG07 'Add By Sindy 2024/5/17
         If strTPG06 = "" Or strTPG06 = "000" Or _
            InStr(txtChkWord, "?") > 0 Or strTPG43 = "" Or strTPG43 = "¤¤µØ¥Á°ê" Or strTPG43 = "¥xÆW" Then
            Call ReadTxt1(strTPG01, strTPG02, strTPG06, strTPG07, strAChinese1, strAddress1)
            Call PrintPaper(strTPG01, strTPG02, strTPG06, strTPG07, strAddress1)
         End If
         
         'Add By Sindy 2018/11/12
         'ÀË¬d¥Ó½Ð¤H¦WºÙ¬O§_¦³?³y¦r
         For i = 1 To 10
            txtChkWord = strTPGcApp(i) 'Add By Sindy 2024/5/17
            If InStr(txtChkWord, "?") > 0 Then
               strMsg = "¥Ó½Ð®×¸¹" & strTPG01 & "¥Ó½Ð¤H¦WºÙ" & i & "¡u" & strTPGcApp(i) & "¡v¦³?¸¹"
               Call ReadTxt1(strTPG01, strTPG02, strMsg, "", "", "")
               Call PrintPaper(strTPG01, strTPG02, strMsg, "", "")
            End If
         Next i
         '2018/11/12 END
         
         '·s¼WTable
         strErrTxt = "·s¼W°ê¤º±M§Q¤½¶}¤½³øÀÉ.TPGazette"
         'Modify By Sindy 2016/3/2 +TPG18
         'Modify By Sindy 2019/9/4 +,TPG43
         strSql = "insert into TPGazette(TPG01,TPG02,TPG03,TPG04,TPG05,TPG06,TPG07,TPG08,TPG09,TPG15,TPG16,TPG17,TPG18" & _
                  ",TPG19,TPG20,TPG21,TPG22,TPG23,TPG24,TPG25,TPG26,TPG27,TPG28" & _
                  ",TPG29,TPG30,TPG31,TPG32,TPG33,TPG34,TPG35,TPG36,TPG37,TPG38" & _
                  ",TPG39,TPG40,TPG41,TPG42,TPG43" & _
                  ") values(" & CNULL(strTPG01) & "," & CNULL(strTPG02) & "," & dblTPG03 & "," & CNULL(strTPG04) & "," & CNULL(strTPG05) & _
                  "," & CNULL(strTPG06) & "," & CNULL(strTPG07_1) & "," & CNULL(strTPG08) & "," & CNULL(strTPG09) & _
                  "," & CNULL(strTPG15) & "," & CNULL(strTPG16) & "," & CNULL(strTPG17) & "," & CNULL(strTPG18) & _
                  "," & CNULL(strTPGcApp(1)) & "," & CNULL(strTPGcApp(2)) & "," & CNULL(strTPGcApp(3)) & "," & CNULL(strTPGcApp(4)) & "," & CNULL(strTPGcApp(5)) & _
                  "," & CNULL(strTPGcApp(6)) & "," & CNULL(strTPGcApp(7)) & "," & CNULL(strTPGcApp(8)) & "," & CNULL(strTPGcApp(9)) & "," & CNULL(strTPGcApp(10)) & _
                  "," & CNULL(strTPGeApp(1)) & "," & CNULL(strTPGeApp(2)) & "," & CNULL(strTPGeApp(3)) & "," & CNULL(strTPGeApp(4)) & "," & CNULL(strTPGeApp(5)) & _
                  "," & CNULL(strTPGeApp(6)) & "," & CNULL(strTPGeApp(7)) & "," & CNULL(strTPGeApp(8)) & "," & CNULL(strTPGeApp(9)) & "," & CNULL(strTPGeApp(10)) & _
                  "," & dblTPG39 & "," & dblTPG40 & "," & CNULL(strTPG41) & "," & CNULL(strTPG42) & "," & CNULL(strTPG43) & _
                  ")"
         cnnConnection.Execute strSql
         
         '¥»©Ò®×¥ó¤~§ó·s
         If bolTaieCase = True Then
            'Add By Sindy 2014/6/17 ·s¼W¶i«×
            'If pa(1) = "P" Then 'Modify By Sindy 2015/8/18 FCP¤]­n·s¼W¸Óµ§¶i«×
               strCP10 = "1229" '1229.¤½¶}¤½³ø
               'Modified by Lydia 2019/06/17 §ì¬O§_³¬¨÷¾P¨÷(closecase)
               'strSql = "SELECT cp09 FROM caseprogress " & _
                        "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' " & _
                         " AND CP10 = '" & strCP10 & "'"
               'Modified by Lydia 2019/07/01 debug
               'strSql = "SELECT cp09,pa57||pa108 as closecase FROM caseprogress,patent " & _
                        "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' " & _
                         " AND CP10 = '" & strCP10 & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
               strSql = "SELECT PA57||PA108 AS CLOSECASE,CP09 FROM PATENT," & _
                          "(SELECT CP09,CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10 = '" & strCP10 & "' ) X " & _
                          "WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' " & _
                          "AND PA01=CP01(+) AND PA02=CP02(+) AND PA03=CP03(+) AND PA04=CP04(+) "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               'Modified by Lydia 2019/07/01
               'If intI = 0 Then
               If intI = 1 Then
                  If "" & RsTemp.Fields("CP09") = "" Then
               'end 2019/07/01
                        stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
                        stCP12 = GetSalesArea(stCP13)
                        stCP09 = AutoNo("C", 6)
                        strExc(3) = "" 'Added by Lydia 2019/06/17
                        'Modified by Lydia 2019/05/31 ¥~±Mµ{§Ç¤u§@¤j¶µ¥ý¤£¤Wµo¤å¤é(¾ã§åµo¤å)
                        'strSql = "INSERT INTO CASEPROGRESS(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32)" & _
                                " VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & strSrvDate(1) & "','" & stCP09 & "'" & _
                                ",'" & strCP10 & "','" & stCP12 & "','" & stCP13 & "','" & strUserNum & "','N','N','" & strSrvDate(1) & "','N')"
                        If pa(1) = "FCP" Then
                              'Added by Lydia 2019/06/17 ¤w¤W³¬¨÷ªº®×¥ó¡A¦U¶µ¤j§å¶i«×ÀÉµo¤å¤é½Ð¥ý¤W111111
                              If "" & RsTemp.Fields("closecase") <> "" Then
                                  strExc(3) = "19221111"
                                  'Added by Morgan 2025/10/1 ÅÜ¼Æ­n­«³]¡A§_«h¤U­±·s¼W¶i«×·|¨S¤W°²µo¤å¤]±¾¿ùµ{§Ç¤H­û Ex:FCP-072920
                                  strCP14 = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
                                  strCP48 = 0
                                  'end 2025/10/1
                              Else
                              'end 2019/06/17
                                  'Added by Lydia 2024/10/07 §ï¦¨¦U°ÏFCPµ{§ÇºÞ¨î¤H---11/1¤W½u
                                  If strSrvDate(1) >= "20241101" Then
                                     strCP14 = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
                                  Else
                                  'end 2024/10/07
                                     strCP14 = Pub_GetSpecMan("¥~±Mµ{§Ç-¤½¶}¤½³ø")
                                  End If
                                  strCP48 = CompDate(2, 14, strSrvDate(1))
                              End If 'end 2019/06/17
                        Else
                            strCP14 = strUserNum
                            strCP48 = ""
                        End If
                        'Modified by Lydia 2019/06/17 ¤w¤W³¬¨÷ªº®×¥ó¡A¦U¶µ¤j§å¶i«×ÀÉµo¤å¤é½Ð¥ý¤W111111
                        'strSql = "INSERT INTO CASEPROGRESS(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP48)" & _
                                " VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & strSrvDate(1) & "','" & stCP09 & "'" & _
                                ",'" & strCP10 & "','" & stCP12 & "','" & stCP13 & "','" & strCP14 & "','N','N','" & IIf(strCP48 = "", strSrvDate(1), "") & "','N'," & CNULL(strCP48, True) & " )"
                        strSql = "INSERT INTO CASEPROGRESS(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP48)" & _
                                " VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & strSrvDate(1) & "','" & stCP09 & "'" & _
                                ",'" & strCP10 & "','" & stCP12 & "','" & stCP13 & "','" & strCP14 & "','N','N','" & IIf(strCP48 = "", IIf(strExc(3) <> "", strExc(3), strSrvDate(1)), "") & "','N'," & CNULL(strCP48, True) & " )"
                        'end 2019/05/31
                        cnnConnection.Execute strSql
                        '±Npdf file¦s¤JDB
                        strFileName = txtPath2.Text & "\img_1\pub0" & Left(txtTMBM07, 2) & "0" & Right(txtTMBM07, 2) & "\" & strTPG01 & ".pdf"
                        'Set fs = CreateObject("Scripting.FileSystemObject")
                        Set f = fs.GetFile(strFileName)
                        '¦sÀÉ
                        'Modify By Sindy 2022/5/6 CStr(Val(pa(2))) ==> pa(2)
                        If SaveAttFile_PDF(stCP09, strFileName, UCase(pa(1) & pa(2) & IIf(pa(3) <> "0" Or pa(4) <> "00", "-" & pa(3), "") & IIf(pa(4) <> "00", "-" & pa(4), "") & "." & strCP10 & ".pdf"), Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), True) = False Then
                           GoTo ErrHand
                        End If
                  End If 'Added by Lydia 2019/07/01 If "" & RsTemp.Fields("CP09") = "" Then
               End If
            'End If
            '2014/6/17 END
            
            ' §ó·s±M§Q°ò¥»ÀÉªº¤½¶}¤é¤Î¤½¶}¸¹
            strSql = "UPDATE Patent SET PA12 = " & dblTPG03 & ", " & _
                                       "PA13 = '" & strTPG02 & "' " & _
                     "WHERE PA11 = '" & strTPG01 & "'"
            cnnConnection.Execute strSql
         End If
         cnnConnection.CommitTrans
      End If
   Next dblFCnt
   
   '¸ÑªR¹ê¼f¤½¶}
   'Add by Sindy 2022/3/3
   If strSrvDate(1) >= Form20¤W½u¤é Then
      adoStream.LoadFromFile (txtPath2.Text & "\pubsud06.xml")
      m_strTextBox = adoStream.ReadText
   Else
   '2022/3/3 END
      RichTextBox1.LoadFile (txtPath2.Text & "\pubsud06.xml")
      m_strTextBox = RichTextBox1.Text
   End If
   
   strText = "PubSud06Dataset": strTitNM = "¹ê¼f¤½¶}"
   dblStar = InStr(m_strTextBox, "<" & strText)
   dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
   dblFCnt = 0
   'strTestTPG01 = ""
   If dblStar > 0 Then
      For dblChar = dblStar To dblLastEnd
         strTPG01 = ""
         strTPG11 = "": strTPG12 = "": strTPG13 = "": strTPG14 = ""
         For j = 1 To 5
            strData = ""
            If j = 1 Then
               strText = "aplno": strTitNM = "¥Ó½Ð®×¸¹"
            ElseIf j = 2 Then
               strText = "volno": strTitNM = "¹ê¼f¤½¶}¨÷¼Æ"
            ElseIf j = 3 Then
               strText = "isuno": strTitNM = "¹ê¼f¤½¶}´Á¼Æ"
            ElseIf j = 4 Then
               strText = "examdt": strTitNM = "¹ê¼f¥Ó½Ð¤é"
            ElseIf j = 5 Then
               strText = "checkyn": strTitNM = "¬O§_¥»¤H¥Ó½Ð"
            End If
            dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
            If dblStar < dblChar Then Exit For
            If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
            '***** ¸ÑªRXML *****
            If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd) = False Then
            '***** End
               Exit For
            End If
            If j = 1 Then
               strTPG01 = strData
               'strTestTPG01 = strTestTPG01 & ",'" & strTPG01 & "'"
               dblFCnt = dblFCnt + 1
               Label2.Caption = "ÂàÀÉ¤¤, ½Ðµy­Ô (¹ê¼f¤½¶} ²Ä" & dblFCnt & "µ§) . . ."
               DoEvents
            ElseIf j = 2 Then
               strTPG11 = Format(strData, "00")
            ElseIf j = 3 Then
               strTPG12 = Format(strData, "00")
            ElseIf j = 4 Then
               strTPG13 = DBDATE(strData)
            ElseIf j = 5 Then
               If strData = "¬O" Then
                  strTPG14 = "Y"
               ElseIf strData = "§_" Then
                  strTPG14 = "N"
               End If
               '§ó·s¸ê®Æ
               strErrTxt = "§ó·s°ê¤º±M§Q¤½¶}¤½³øÀÉ.TPGazette"
               strSql = "update TPGazette set " & _
                        "TPG10=" & DBDATE(text03) & _
                        ",TPG11='" & strTPG11 & "'" & _
                        ",TPG12='" & strTPG12 & "'" & _
                        ",TPG13=" & strTPG13 & _
                        ",TPG14='" & strTPG14 & "'" & _
                        "where TPG01='" & strTPG01 & "'"
               cnnConnection.Execute strSql
            End If
            dblChar = dblEnd
         Next j
      Next dblChar
      'If strTestTPG01 > "" Then strTestTPG01 = Mid(strTestTPG01, 2, Len(strTestTPG01)) '´ú¸Õ¥Î
   End If
   
   'Add By Sindy 2015/6/15
   strSql = "select pa01,pa01||'-'||pa02||'-'||pa03||'-'||pa04 caseno,pa11,r04060306.* from patent,r04060306" & _
            " where rcp01=pa01(+) and rcp02=pa02(+) and rcp03=pa03(+) and rcp04=pa04(+)" & _
            " order by rcp01,rcp02,rcp03,rcp04,rseqno"
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   With rsTmp
      If .RecordCount > 0 Then
         .MoveFirst
         TempFileName = ""
         Do While Not .EOF
            If TempFileName <> "" And strSys <> .Fields("pa01") Then
               Close ff
               If strSys = "P" Then
                  strTo = "79075" '³¢¶®®S
               Else
                  'modify by sonia 2016/7/15 ¨ú®ø73023¥[A4025¼B¤SµØ
                  'strTo = "73023;82045" '±iÀRªÚ;§d­Yªâ
                  'Modified by Morgan 2018/3/19
                  'strTo = "82045;A4025" '§d­Yªâ;¼B¤SµØ
                  'Modified by Lydia 2021/09/01 §ï¦¨¨t²Î³]©w
                  'strTo = "82045;A6019" '§d­Yªâ;¬x­§´P
                  'Added by Lydia 2024/10/07 §ï³qª¾FCPµ{§ÇºÞ¨î¤H(¥þ³¡)---11/1¤W½u
                  If strSrvDate(1) >= "20241101" Then
                     'Modified by Lydia 2024/11/04 ¥þ³¡µ{§Ç³£³qª¾----Sharon
                     'strExc(0) = "select na16 from nation,staff where na01 > '010' and nvl(na16,'N') <> 'N' and na16=st01(+) and st04='1' group by na16 "
                     'intI = 1
                     'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     'If intI = 1 Then
                     '   strTo = RsTemp.GetString(adClipString, , , ";")
                     '   If Right(strTo, 1) = ";" Then strTo = Mid(strTo, 1, Len(strTo) - 1)
                     'End If
                     strTo = "FCP_1"
                     'end 2024/11/04
                  Else
                  'end 2024/10/07
                     strTo = Pub_GetSpecMan("¥~±Mµ{§Ç-¤½¶}¤½³ø")
                  End If
               End If
               PUB_SendMail strUserNum, strTo, "", TempFileName, "Dear Sirs," & vbCrLf & vbCrLf & _
               "½Ð¬Ýªþ¥ó¡I" & vbCrLf & vbCrLf & vbCrLf & _
               "                                                        ¹q¸£¤¤¤ß", , txtPath2 & "\" & TempFileName & ".txt"
               TempFileName = ""
            End If
            If TempFileName = "" Then
               TempFileName = "°ê¤º±M§Q¤½¶}¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¸ê®Æ¤ñ¹ï©ú²Óªí-" & .Fields("pa01")
               ff = FreeFile
               If ff > 0 Then Close #ff
               ff = FreeFile
               Open txtPath2 & "\" & TempFileName & ".txt" For Output As ff
               Print #ff, "¥»©Ò®×¸¹     ¥Ó½Ð®×¸¹   ¶µ¥Ø         ¤º®e (¤W¡G¤½³ø¤º®e ¤U¡G¥»©Ò«ØÀÉ¤º®e)"
               Print #ff, "============ ========== ============ =================================================="
            End If
            For i = 1 To 8
               strTemp(i) = ""
            Next i
            strTemp(1) = convForm(CheckStr("" & .Fields("caseno")), 12)
            strTemp(2) = convForm(CheckStr("" & .Fields("PA11")), 10)
            strTemp(3) = convForm(CheckStr("" & .Fields("ritem")), 12)
            strTemp(4) = Replace(CheckStr("" & .Fields("rtext")), "!!", "!")
            strTemp(5) = Replace(CheckStr("" & .Fields("rdbtext")), "!!", "!")
            strSys = .Fields("pa01")
            '¤½³ø¤º®e
            'Modify By Sindy 2015/7/7 ±M§Q³B­n½ð°£´¼¼z§½µL«ØÀÉªº¸ê®Æ
            If strSys = "P" And strTemp(4) = "" Then GoTo ReadNext
            If strTemp(4) = "" Then
               Print #ff, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " ¤½³øµL¸ê®Æ"
            Else
               arrData = Split(strTemp(4), ";")
               For i = 0 To UBound(arrData)
                  arrData_1 = Split(arrData(i), "!")
                  For j = 0 To UBound(arrData_1)
                     If i = 0 And j = 0 Then
                        Print #ff, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & arrData_1(j)
                     Else
                        Print #ff, convForm(" ", 37) & arrData_1(j)
                     End If
                  Next j
               Next i
            End If
            '¥»©Ò«ØÀÉ¤º®e
            Print #ff, convForm(" ", 37) & "--------------------------------------------------"
            If strTemp(5) = "" Then
               Print #ff, convForm(" ", 37) & " ¸ê®Æ®wµL¸ê®Æ"
            Else
               arrData = Split(strTemp(5), ";")
               For i = 0 To UBound(arrData)
                  arrData_1 = Split(arrData(i), "!")
                  For j = 0 To UBound(arrData_1)
                     Print #ff, convForm(" ", 37) & arrData_1(j)
                  Next j
               Next i
            End If
            Print #ff, "---------------------------------------------------------------------------------------"
ReadNext:
            .MoveNext
         Loop
         If TempFileName <> "" Then Close ff
      End If
   End With
   rsTmp.Close
   If TempFileName <> "" Then
      If strSys = "P" Then
         strTo = "79075" '³¢¶®®S
      Else
         'modify by sonia 2016/7/15 ¨ú®ø73023¥[A4025¼B¤SµØ
         'strTo = "73023;82045" '±iÀRªÚ;§d­Yªâ
         'Modified by Morgan 2018/3/19
         'strTo = "82045;A4025" '§d­Yªâ;¼B¤SµØ
         'Modified by Lydia 2021/09/01 §ï¦¨¨t²Î³]©w
         'strTo = "82045;A6019" '§d­Yªâ;¬x­§´P
         strTo = Pub_GetSpecMan("¥~±Mµ{§Ç-¤½¶}¤½³ø")
      End If
      PUB_SendMail strUserNum, strTo, "", TempFileName, "Dear Sirs," & vbCrLf & vbCrLf & _
      "½Ð¬Ýªþ¥ó¡I" & vbCrLf & vbCrLf & vbCrLf & _
      "                                                        ¹q¸£¤¤¤ß", , txtPath2 & "\" & TempFileName & ".txt"
   End If
   '2015/6/15 END
   
   bolTa04IsNull = ReadTagentTa04IsNull(text03.Text) 'Add By Sindy 2014/9/3
   strMsg = ""
   'Modify By Sindy 2014/9/3
'   If m_PrintRpt1 = True Then
'      Close ff1
'      strMsg = "½Ð¦Ü¤U¦C¦ì¸m¦C¦LÀË®Öªí¡G" & PUB_Getdesktop & "\" & m_strFileName1
'   End If
'   'If m_PrintRpt2 = True Then
'   If intPRow > 0 Then
'      Call PrintRpt
'      Printer.EndDoc
'      strMsg = strMsg & "¡FÀË®Öªí¤w¦C¦L§¹¦¨"
'   End If
   If m_PrintRpt1 = True Or bolTa04IsNull = True Then
      If m_PrintRpt1 = True Then
         'Close ff1
         'Add By Sindy 2024/5/17
         If Dir(PUB_Getdesktop & "\" & m_strFileName1) <> "" Then
            Kill PUB_Getdesktop & "\" & m_strFileName1
            Sleep 100
         End If
         Call PUB_SaveTextAsUTF8(PUB_Getdesktop & "\" & m_strFileName1, m_strText)
         '2024/5/17 END
         If bolTa04IsNull = True Then m_strFileName1 = m_strFileName1 & "¡B" & "¤½³ø¥N²z¤H¨Æ°È©Ò¦WºÙÄæªÅ¥Õ²M³æ.txt"
      Else
         m_strFileName1 = "¤½³ø¥N²z¤H¨Æ°È©Ò¦WºÙÄæªÅ¥Õ²M³æ.txt"
      End If
      strMsg = "½Ð¦Ü¤U¦C¦ì¸m¦C¦LÀË®Öªí¡G" & PUB_Getdesktop & "\" & m_strFileName1
   End If
   'If m_PrintRpt2 = True Then
   If intPRow > 0 Or bolTa04IsNull = True Then
      If intPRow > 0 Then
         Call PrintRpt
         Printer.EndDoc
      End If
      strMsg = strMsg & "¡FÀË®Öªí¤w¦C¦L§¹¦¨"
   End If
   '2014/9/3 END
   
   Set fs = Nothing
   Set f = Nothing
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   
   'Modify By Sindy 2024/6/3 ·¨¶²ªÚ¸g²z«ü¥Ü,Á`¸g²z¤w®Ö¥Ü°±¤î¦¹¶µ¤ÀÃþ¤u§@¡A¦¹Ãþ³qª¾¤]¥i°±¤îµo°e
'   Call GetSendMailIPC
   Call IsRecordExist '²£¥Íµ§¼Æ
   MsgBox "ÂàÀÉ§¹²¦¡I(ÂàÀÉªá¶O®É¶¡¡G" & strTime & "  " & time() & ")" & vbCrLf & strMsg
   Me.Height = MinHeight
   
   Exit Sub
   
ErrHand:
   Set fs = Nothing
   Set f = Nothing
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   If Err.NUMBER = 76 Then
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "\img_1\pub" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "¡^¤ºµL¸Ó´Á¤½¶}¤½³ø¸ê®Æ¡I"
      txtPath2.SetFocus
   Else
      cnnConnection.RollbackTrans
      If Err.NUMBER = -2147217873 Then
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½¶}¤½³ø¥Ó½Ð®×¸¹¡]" & strTPG01 & "¡^" & vbCrLf & strErrTxt & ": ¹H¤Ï¥²¶·¬°°ß¤@ªº­­¨î±ø¥ó"
      Else
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½¶}¤½³ø¥Ó½Ð®×¸¹¡]" & strTPG01 & "¡^" & vbCrLf & strErrTxt & Err.Description
      End If
   End If
End Sub

'Modify By Sindy 2013/8/27
Private Function ReadXmlData() As Boolean
Dim dblStar As Double
Dim strMsg As String
Dim dblChar As Double, dblLastEnd As Double, dblEnd As Double
Dim strText As String, strTitNM As String
Dim strChar As String, strData As String
Dim rsTmp As New ADODB.Recordset
Dim strFreeAgentCode As String
Dim i As Integer, j As Integer
Dim dblRunStar As Double
Dim strChineseNM As String, strEnglishNM As String, intApp As Integer 'Add By Sindy 2018/11/12
Dim strUpdNewTA02 As String 'Add By Sindy 2020/1/9
   
   ReadXmlData = True
   
   strTPG01 = "": strTPG02 = "": dblTPG03 = Empty: strTPG04 = ""
   strTPG05 = "": strTPG06 = "": strTPG07 = "": strTPG07_1 = "": strTPG07_temp1 = "": strUpdNewTA02 = ""
   strTPG08 = "": strTPG09 = ""
   strAChinese = "": strAChinese1 = "": strAddress1 = ""
   bolTaieCase = False
   strTaieCaseNo = ""
   'Add By Sindy 2013/8/27
   'Modify By Sindy 2016/3/2 +: strTPG18 = ""
   strTPG15 = "": strTPG16 = "": strTPG17 = "": strTPG18 = ""
   '2013/8/27 END
   strTPG43 = "" 'Add By Sindy 2019/9/4
   'Add By Sindy 2015/6/10
   strCaseChNm = "": strCaseEnNm = "" 'µo©ú¤¤­^¤å¦WºÙ
   strApplDate = "" '¥Ó½Ð¤é
   strAEng = "" '¥Ó½Ð¤H­^¤å¦WºÙ
   strAEnCountry = "" '¥Ó½Ð¤H°êÄy
   strApplName = "" '¥Ó½Ð¤H
   strInventor = "" 'µo©ú¤H
   strAgent = "" '¥N²z¤H
   strClaims = "" 'Àu¥ýÅv
   strGetData1 = "": strGetData2 = "": strGetData3 = ""
   '2015/6/10 END
   'Add By Sindy 2018/11/12
   For i = 1 To 10
      strTPGcApp(i) = ""
      strTPGeApp(i) = ""
   Next i
   dblTPG39 = Empty: dblTPG40 = Empty: strTPG41 = "": strTPG42 = ""
   '2018/11/12 End
   
   If GetXmlData(1, "volno", "¨÷¼Æ", strData, dblEnd) = True Then
      strTPG04 = Format(strData, "00")
   End If
   If GetXmlData(1, "isuno", "´Á¼Æ", strData, dblEnd) = True Then
      strTPG05 = Format(strData, "00")
   End If
   dblStar = InStr(m_strTextBox, "<publication-reference>")
   If GetXmlData(dblStar, "doc-number", "¤½¶}¸¹", strData, dblEnd) = True Then
      strTPG02 = strData
   End If
   If GetXmlData(dblStar, "date", "¤½¶}¤é", strData, dblEnd) = True Then
      dblTPG03 = DBDATE(strData)
   End If
   dblStar = InStr(m_strTextBox, "<application-reference")
   If GetXmlData(dblStar, "doc-number", "¥Ó½Ð®×¸¹", strData, dblEnd) = True Then
      strTPG01 = strData
      '¥Ó½Ð®×¤~­n±a
      Erase pa
      ReDim pa(1 To TF_PA) As String
      strSql = "SELECT * FROM Patent " & _
               "WHERE PA11 = '" & strTPG01 & "' AND " & _
                     "PA09 = '000' and pa23='1'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            bolTaieCase = True
            strTaieCaseNo = RsTemp.Fields("PA01") & "-" & RsTemp.Fields("PA02") & "-" & RsTemp.Fields("PA03") & "-" & RsTemp.Fields("PA04")
            pa(1) = RsTemp.Fields("PA01")
            pa(2) = RsTemp.Fields("PA02")
            pa(3) = RsTemp.Fields("PA03")
            pa(4) = RsTemp.Fields("PA04")
            pa(14) = "" & RsTemp.Fields("PA14")
            pa(22) = "" & RsTemp.Fields("PA22")
            pa(72) = "" & RsTemp.Fields("PA72")
            pa(21) = "" & RsTemp.Fields("PA21")
            Call ClsPDReadPatentDatabase(pa(), °ê¤º, False) 'Add By Sindy 2015/6/10
         End If
      End If
   End If
   'Add By Sindy 2015/6/10
   If GetXmlData(dblStar, "date", "¥Ó½Ð¤é", strData, dblEnd) = True Then
      strApplDate = DBDATE(strData)
      dblTPG39 = strApplDate 'Add By Sindy 2018/11/12
   End If
   '2015/6/10 END
   
   If GetXmlData(1, "physical-examination", "¥Ó½Ð¹êÅé¼f¬d", strData, dblEnd) = True Then
      If strData = "µL" Then
         strTPG09 = "N"
      ElseIf strData = "¦³" Then
         strTPG09 = "Y"
      End If
   End If
   
   '°ê»Ú¤ÀÃþ
   dblStar = InStr(m_strTextBox, "<classification-")
   If dblStar > 0 Then
      If GetXmlData2(dblStar, "main-classification", "°ê»Ú¤ÀÃþ", strData, dblEnd) = True Then
         If Trim(strData) <> "" Then
            strTPG15 = strData '°ê»Ú¤ÀÃþ¸¹
            strTPG16 = GetPatentIPC("1", strTPG15, "I") 'IPC¤ÀÃþ
            strTPG17 = GetPatentIPC("2", strTPG15, "") '²£·~§O¤ÀÃþ
            strTPG18 = GetPatentIPC("3", strTPG15, "") '®×¥óÄÝ©Ê 'Add By Sindy 2016/3/2
            
            If strTPG17 = "" Then
               strErrTxt = "²£·~§O¤ÀÃþ¤£¥iªÅ¥Õ¡I"
               ReadXmlData = False
            End If
            'Add By Sindy 2016/3/2
            If strTPG18 = "" Then
               strErrTxt = "®×¥óÄÝ©Ê¤£¥iªÅ¥Õ¡I"
               ReadXmlData = False
            End If
            '2016/3/2 END
            
            'IPC¤ÀÃþÂkÃþ¤£¨ì®É,°O¿ý°ê»Ú¤ÀÃþ¸¹
            If strTPG16 = "" Then
               If InStr(m_PI02, strTPG15) = 0 Then
                  m_PI02 = m_PI02 & strTPG15 & " ¥Ó½Ð®×¸¹¬° " & strTPG01 & vbCrLf
               End If
            End If
         End If
      End If
   End If
   
   'Add By Sindy 2015/6/9 'µo©ú¤¤­^¤å¦WºÙ
   dblStar = InStr(m_strTextBox, "<invention-title")
   If GetXmlData(dblStar, "chinese-title", "µo©ú¤¤¤å¦WºÙ", strData, dblEnd) = True Then
      strCaseChNm = strData
   End If
   If GetXmlData(dblStar, "english-title", "µo©ú­^¤å¦WºÙ", strData, dblEnd) = True Then
      strCaseEnNm = strData
   End If
'   strText = "invention-title": strTitNM = "µo©ú¦WºÙ"
'   dblStar = InStr(m_strTextBox, "<" & strText & ">")
'   dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
'   If dblStar > 0 Then
'      For dblChar = dblStar To dblLastEnd
'         strData = ""
'         strText = "chinese-title": strTitNM = "¤¤¤å¦WºÙ"
'         dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
'         If dblStar < dblChar Then Exit For
'         If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
'         '***** ¸ÑªRXML *****
'         If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd) = False Then
'         '***** End
'            Exit For
'         Else
'            strCaseChNm = strData
'         End If
'         dblChar = dblEnd
'         strData = ""
'         strText = "english-title": strTitNM = "­^¤å¦WºÙ"
'         dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
'         If dblStar < dblChar Then Exit For
'         If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
'         '***** ¸ÑªRXML *****
'         If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd) = False Then
'         '***** End
'            Exit For
'         Else
'            strCaseEnNm = strData
'         End If
'         dblChar = dblEnd
'      Next dblChar
'   End If
'2015/6/9 END
   
   strText = "agents": strTitNM = "¥N²z¤H"
   dblStar = InStr(m_strTextBox, "<" & strText & ">")
   If dblStar > 0 And InStr(m_strTextBox, "<" & strText & " />") = 0 Then
      dblRunStar = InStr(m_strTextBox, "<" & strText & ">")
      dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
      For dblChar = dblStar To dblLastEnd
         strData = ""
         strText = "last-name": strTitNM = "¥N²z¤H¦WºÙ"
         dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
         If dblStar < dblChar Then Exit For
         If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
         '***** ¸ÑªRXML *****
         If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd) = False Then
         '***** End
            Exit For
         Else
            '©T©wªº¥N²z¤H¹ï·Óªí
'            If strData = "?ªF§÷" Then
'               strData = "üÚªF§÷"
'            ElseIf strData = "°ª?¼ü" Then strData = "°ªû^¼ü"
'            ElseIf strData = "¶À·Ó?" Then strData = "¶À·Óúh"
'            ElseIf strData = "¶À?¹a" Then strData = "¶ÀúE¹a"
'            ElseIf strData = "·¨ªø?" Then strData = "·¨ªøúh"
'            ElseIf strData = "±i·×?" Then strData = "±i·× h"
'            End If
            'Add By Sindy 2017/12/1 ¼W¥[¤ñ¹ï¥N²z¤H
            'Modify By Sindy 2023/8/2
'            strData = ReplaceMadeWord(strData, "?") 'Modify By Sindy 2018/5/21 ÀË¬d³y¦r
'            strData = PUB_FilterBulletinSpecWord("2", strData, "")
            '2023/8/2 END
            '2017/12/1 END
            'Modify By Sindy 2018/7/23 ±q¤U­±if²¾¥X¨Ó§PÂ_
'            If strData = "ÀF±Ò®õ" Then strData = "ÀF•K®õ"
            If bolTaieCase = True And strData <> "" Then
               If InStr(1, strOurAgentName, strData) > 0 Then
                  strTPG07 = GetTAgentName("01", "TA03")
                  strTPG07_1 = "01"
                  strTPG08 = GetTAgentName("01", "TA04")
               End If
            End If
            '2018/7/23 END
            If strTPG07_temp1 = "" Then strTPG07_temp1 = strData '°O¿ý²Ä¤@¦ì¥X¦W¥N²z¤H
            '©|¥¼Åª¨ú¨ì¥N²z¤H¦WºÙ®É
            'Modify By Sindy 2020/1/9
            'If Trim(strTPG07) = "" And strData <> "" Then
            If strData <> "" Then
            '2020/1/9 END
               'ÀË¬d¬O§_¬°¥»©Ò¥N²zªº®×¥ó
'                     strSql = "select cp09 from caseprogress,(SELECT PA01,PA02,PA03,PA04 FROM Patent WHERE PA11='" & strTPG01 & "' AND PA09='000' and pa23='1') " & _
'                              "Where CP01=pa01 And cp02=pa02 And cp03=pa03 And cp04=pa04 " & _
'                              "and instr('" & NewCasePtyList & "',cp10)>0 and cp27 is not null "
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                     If intI = 1 And InStr(1, strOurAgentName, strData) > 0 Then
'                        strTPG07 = GetTAgentName("01", "TA03")
'                        strTPG07_1 = "01"
'                        strTPG08 = GetTAgentName("01", "TA04")
'                        Exit For
'                     End If
'               If bolTaieCase = True Then
'                  If InStr(1, strOurAgentName, strData) > 0 Then
'                     strTPG07 = GetTAgentName("01", "TA03")
'                     strTPG07_1 = "01"
'                     strTPG08 = GetTAgentName("01", "TA04")
'                     Exit For
''                        Else
''                           strMsg = strTaieCaseNo & "¬°¥»©Ò®×¥ó¦ý¥N²z¤H¨Ã«D¥»©Ò"
''                           Call ReadTxt1(strTPG01, strTPG02, strMsg, "", "", "")
'                  End If
'               End If
               
               '¨ú±o¤w¦³½s¦Cªº¥N²z¤H¦WºÙ
               strSql = "SELECT * FROM TAGENT " & _
                         "WHERE TA01 = 'P' AND " & _
                                "replace(replace(TA03,'¡@',''),' ','')='" & Trim(strData) & "' "
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount > 0 Then
                  rsTmp.MoveFirst
                  'Modify By Sindy 2020/1/9
                  If strTPG08 = "" Then
                  '2020/1/9 END
                     If IsNull(rsTmp.Fields("TA02")) = False Then
                        strTPG07_1 = rsTmp.Fields("TA02")
                     End If
                     If IsNull(rsTmp.Fields("TA03")) = False Then
                        strTPG07 = rsTmp.Fields("TA03")
                     End If
                     If IsNull(rsTmp.Fields("TA04")) = False Then
                        strTPG08 = rsTmp.Fields("TA04")
                     End If
                  End If
                  'Modify By Sindy 2020/1/9 °j°é­n¶]§¹,Åª¨ú¥þ³¡¥X¦W¥N²z¤H¸ê®Æ
                  'rsTmp.Close: Exit For
               Else
                  'Modify By Sindy 2020/1/9
                  '·s¼W°ê¤º¤½³ø¥N²z¤HÀÉ
                  strFreeAgentCode = PUB_GetFreeAgentCode("P")
                  If strTPG07_1 = "" Then strTPG07_1 = strFreeAgentCode '°O¿ý²Ä¤@¦ì¥X¦W¥N²z¤HID
                  strUpdNewTA02 = strUpdNewTA02 & ",'" & strFreeAgentCode & "'" 'Add By Sindy 2020/1/9
                  strSql = "INSERT INTO TAgent (TA01,TA02,TA03,TA04,TA05) " & _
                           "VALUES ('P','" & strFreeAgentCode & "','" & Trim(strData) & "',null," & dblTPG03 & ")"
                  cnnConnection.Execute strSql
                  '2020/1/9 END
               End If
               rsTmp.Close
            End If
         End If
         dblChar = dblEnd
      Next dblChar
      '©|¥¼Åª¨ú¨ì¥N²z¤H¦WºÙ®É,«h§ó·s²Ä¤@¦ì¥X¦W¥N²z¤H¸ê®Æ
      If Trim(strTPG07) = "" And strTPG07_temp1 <> "" Then
         strTPG07 = strTPG07_temp1
         strTPG08 = strTPG07_temp1
         'Modify By Sindy 2020/1/9 Mark,§ï«e­±³vµ§µL¸ê®Æ,«hinsert
'         If InStr(strTPG07_temp1, "?") = 0 Then
'            '·s¼W°ê¤º¤½³ø¥N²z¤HÀÉ
'            strFreeAgentCode = PUB_GetFreeAgentCode("P")
'            strTPG07_1 = strFreeAgentCode
'            'Modify By Sindy 2014/9/2 ·s¥N²z¤Hªº¨Æ°È©Ò¦WºÙÄæ©ñNull
''            strSql = "INSERT INTO TAgent (TA01,TA02,TA03,TA04,TA05) " & _
''                     "VALUES ('P','" & strTPG07_1 & "','" & Trim(strTPG07) & "','" & Trim(strTPG08) & "'," & dblTPG03 & ")"
'            strSql = "INSERT INTO TAgent (TA01,TA02,TA03,TA04,TA05) " & _
'                     "VALUES ('P','" & strTPG07_1 & "','" & Trim(strTPG07) & "',Null," & dblTPG03 & ")"
'            cnnConnection.Execute strSql
'         End If
      'Modify By Sindy 2020/1/9 §ó·s,·s¥N²z¤Hªº¨Æ°È©Ò¦WºÙ
      ElseIf strTPG08 <> "" And strUpdNewTA02 <> "" Then
         strUpdNewTA02 = Mid(strUpdNewTA02, 2)
         strSql = "UPDATE TAgent SET TA04='" & strTPG08 & "'" & _
                  " WHERE TA01='P' AND TA02 in(" & strUpdNewTA02 & ")"
         cnnConnection.Execute strSql
         '2020/1/9 END
      End If
      '¬°¥»©Ò®×¥ó¦ý¥N²z¤H¨Ã«D¥»©Ò
      If bolTaieCase = True And strTPG07_1 <> "01" Then
         strMsg = strTaieCaseNo & "¬°¥»©Ò®×¥ó¦ý¥N²z¤H¨Ã«D¥»©Ò¡A¬°¡e" & strTPG07_1 & " " & strTPG07 & " " & strTPG08 & "¡f"
         Call ReadTxt1(strTPG01, strTPG02, strMsg, "", "", "")
         Call PrintPaper(strTPG01, strTPG02, strMsg, "", "")
      End If
   End If
   'Add By Sindy 2015/6/10
   strText = "agents": strTitNM = "¥N²z¤H"
   dblStar = InStr(m_strTextBox, "<" & strText & ">")
   If dblStar > 0 And InStr(m_strTextBox, "<" & strText & " />") = 0 Then
      dblRunStar = InStr(m_strTextBox, "<" & strText & ">")
      dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
      For dblChar = dblStar To dblLastEnd
         strData = ""
         strText = "last-name": strTitNM = "¥N²z¤H¦WºÙ"
         dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
         If dblStar < dblChar Then Exit For
         If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
         '***** ¸ÑªRXML *****
         If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd) = False Then
         '***** End
            Exit For
         End If
         strAgent = strAgent & ";" & strData
         dblChar = dblEnd
      Next dblChar
   End If
   '2015/6/10 END
   
   strText = "applicants": strTitNM = "¥Ó½Ð¤H"
   dblStar = InStr(m_strTextBox, "<" & strText & ">")
   If dblStar > 0 Then
      dblRunStar = InStr(m_strTextBox, "<" & strText & ">")
      dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
      For dblChar = dblStar To dblLastEnd
         For j = 1 To 4 '2
            strData = ""
            If j = 1 Then
               strText = "last-name": strTitNM = "¥Ó½Ð¤H¤¤¤å¦WºÙ"
            'Add By Sindy 2015/6/10
            ElseIf j = 2 Then
               strText = "last-name": strTitNM = "¥Ó½Ð¤H­^¤å¦WºÙ"
            '2015/6/10 END
            ElseIf j = 3 Then
               strText = "address": strTitNM = "¥Ó½Ð¤H¦a§}"
            ElseIf j = 4 Then
               strText = "english-country": strTitNM = "¥Ó½Ð¤H°êÄy"
            End If
            'Modify By Sindy 2015/6/12
            dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
            If dblStar > dblLastEnd Then
               dblStar = InStr(dblChar, m_strTextBox, "<" & strText & " />") + Len("<" & strText & " />") - 1
               If dblStar < dblRunStar Then
                  dblChar = dblLastEnd
                  Exit For
               Else
                  If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
                  dblEnd = dblStar
                  strData = ""
                  GoTo Step_Appl
               End If
            End If
            '2015/6/12 END
            If dblStar < dblChar Then Exit For
            If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
            '***** ¸ÑªRXML *****
            If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd) = False Then
            '***** End
               Exit For
            End If
Step_Appl:
            If dblEnd > dblLastEnd Then strData = "": dblChar = dblStar
            If j = 1 Then '¥Ó½Ð¤H¤¤¤å¦WºÙ
               strAChinese = strData
               If strAChinese1 = "" Then strAChinese1 = strData
            'Add By Sindy 2015/6/10
            ElseIf j = 2 Then '¥Ó½Ð¤H­^¤å¦WºÙ
               strAEng = strData
            '2015/6/10 END
            ElseIf j = 3 Then '¥Ó½Ð¤H¦a§}
               If strAddress1 = "" Then strAddress1 = strData
               If strData <> "" Then
                  If strTPG06 = "" Then
                     '¥ý¥Î¥þ¦W¤ñ¹ï¦a°Ï
                     'Modify By Sindy 2019/9/4 + , strTPG43
                     If GetNationNo(strData, strTPG43) <> "" Then
                        strTPG06 = strData
                        'Exit For
                     End If
                     '³v¦r¤ñ¹ï
                     For i = 1 To Len(strData)
                        strChar = Left(strData, i)
                        strChar = Replace(strChar, "»O", "¥x")
                        'Modify By Sindy 2019/9/4 + , strTPG43
                        If GetNationNo(strChar, strTPG43) <> "" Then
                           strTPG06 = strChar
                           Exit For
                        End If
                        '[¯S¨Ò]³B²z¥xÆW¦a°Ï¦WºÙ
                        If Len(strChar) = 3 Then
                           strChar = Left(strChar, 2) & "¿¤"
                           'Modify By Sindy 2019/9/4 + , strTPG43
                           If GetNationNo(strChar, strTPG43) <> "" Then
                              strTPG06 = strChar
                              Exit For
                           End If
                        End If
                     Next i
                     '¼Ò½k¤ñ¹ï¦a°Ï¦WºÙ
                     If strTPG06 = "" Or strTPG06 = "020" Then '020.¤¤°ê¤j³°
                        If strAChinese <> "" Then
                           'Modify By Sindy 2019/9/4 + , strTPG43
                           strChar = GetNationLike(strAChinese, strTPG43)
                           If strChar <> "" Then
                              strTPG06 = strChar
                              'Exit For
                           End If
                        End If
                     ElseIf strTPG06 <> "" Then
                        'Exit For
                     End If
                  End If
               End If
            'Add By Sindy 2015/6/10
            ElseIf j = 4 Then '¥Ó½Ð¤H°êÄy
               strAEnCountry = strData
               strApplName = strApplName & ";" & strAChinese & "!" & strAEng & "!" & strAEnCountry
            '2015/6/10 END
            End If
            dblChar = dblEnd
         Next j
         'Add By Sindy 2017/12/1
         'Modify By Sindy 2023/8/2
'         strAChinese1 = ReplaceMadeWord(strAChinese1, "?") 'Modify By Sindy 2018/5/21 ÀË¬d³y¦r
'         strAChinese1 = PUB_FilterBulletinSpecWord("1", strAChinese1, GetPrjNationName(strTPG06))
         '2023/8/2 END
         '2017/12/1 END
      Next dblChar
   End If
   
   'Add By Sindy 2015/6/10
   strText = "inventors": strTitNM = "µo©ú¤H"
   dblStar = InStr(m_strTextBox, "<" & strText & ">")
   If dblStar > 0 Then
      dblRunStar = InStr(m_strTextBox, "<" & strText & ">")
      dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
      For dblChar = dblStar To dblLastEnd
         For j = 1 To 3
            strData = ""
            If j = 1 Then
               strText = "last-name": strTitNM = "µo©ú¤H¤¤¤å¦WºÙ"
            ElseIf j = 2 Then
               strText = "last-name": strTitNM = "µo©ú¤H­^¤å¦WºÙ"
            ElseIf j = 3 Then
               strText = "english-country": strTitNM = "µo©ú¤H°êÄy"
            End If
            'Modify By Sindy 2015/6/12
            dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
            If dblStar > dblLastEnd Then
               dblStar = InStr(dblChar, m_strTextBox, "<" & strText & " />") + Len("<" & strText & " />") - 1
               If dblStar < dblRunStar Then
                  dblChar = dblLastEnd
                  Exit For
               Else
                  If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
                  dblEnd = dblStar
                  strData = ""
                  GoTo Step_Inventor
               End If
            End If
            '2015/6/12 END
            If dblStar < dblChar Then dblChar = dblLastEnd: Exit For
            If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
            '***** ¸ÑªRXML *****
            If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd) = False Then
            '***** End
               Exit For
            End If
Step_Inventor:
            If j = 1 Then 'µo©ú¤H¤¤¤å¦WºÙ
               strGetData1 = strData
            ElseIf j = 2 Then 'µo©ú¤H­^¤å¦WºÙ
               strGetData2 = strData
            ElseIf j = 3 Then 'µo©ú¤H°êÄy
               strGetData3 = strData
               strInventor = strInventor & ";" & strGetData1 & "!" & strGetData2 & "!" & strGetData3
            End If
            dblChar = dblEnd
         Next j
      Next dblChar
   End If
   '2015/6/10 END
   
   'Add By Sindy 2018/11/12 °ê¥~³¡·~°È©Ý®i³B­n¥Ó½Ð¤H¸ê®Æ°µ²Î­p¥Î
   strText = "applicants": strTitNM = "¥Ó½Ð¤H"
   dblStar = InStr(m_strTextBox, "<" & strText & ">")
   dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
   intApp = 0
   If dblStar > 0 Then
      For dblChar = dblStar To dblLastEnd
         strChineseNM = "": strEnglishNM = ""
         For j = 1 To 2
            strData = ""
            If j = 1 Then
               dblChar = InStr(dblChar, m_strTextBox, "<chinese-name")
               strText = "last-name": strTitNM = "¥Ó½Ð¤H¤¤¤å¦WºÙ"
            ElseIf j = 2 Then
               dblChar = InStr(dblChar, m_strTextBox, "<english-name")
               strText = "last-name": strTitNM = "¥Ó½Ð¤H­^¤å¦WºÙ"
            End If
            dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
            If dblStar < dblChar Then Exit For
            If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
            '***** ¸ÑªRXML *****
            'If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd, "<") = False Then
            If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd) = False Then
            '***** End
               Exit For
            End If
            If j = 1 Then '¥Ó½Ð¤H¤¤¤å¦WºÙ
               '©m¦W¦³³y¦r¦³¹Ï¤ù
               'strData=¸âµú<img align="absmiddle" height="18px" width="27px" file="106203003/106203003-009.TIF" alt="¨ä¥L«D¹Ï¦¡ ed10999.png" img-content="tif" orientation="portrait" inline="yes" giffile="106203003/106203003-009.png"></img>
               If InStr(strData, "<") > 0 Then
                  strData = Left(strData, InStr(strData, "<") - 1)
               End If
               'Modify By Sindy 2023/8/2
'               strData = ReplaceMadeWord(strData, "?") 'Modify By Sindy 2018/5/21 ÀË¬d³y¦r
'               strChineseNM = PUB_FilterBulletinSpecWord("1", strData, GetPrjNationName(strTPG06))
               strChineseNM = strData
               '2023/8/2 END
            ElseIf j = 2 Then '¥Ó½Ð¤H­^¤å¦WºÙ
               strEnglishNM = strData
            End If
            dblChar = dblEnd
         Next j
         intApp = intApp + 1
         '¸ê®Æ®w¥u¦s10¦ì¥Ó½Ð¤H
         If intApp >= 11 Then
            Exit For
         End If
         If strChineseNM <> "" Then
            strTPGcApp(intApp) = strChineseNM
         End If
         If strEnglishNM <> "" Then
            strTPGeApp(intApp) = strEnglishNM
         End If
      Next dblChar
   End If
   '2018/11/12 END
   
   'Add By Sindy 2015/6/10
   strText = "priority-claims": strTitNM = "Àu¥ýÅv"
   dblStar = InStr(m_strTextBox, "<" & strText & ">")
   If dblStar > 0 And InStr(m_strTextBox, "<" & strText & " />") = 0 Then
      dblRunStar = InStr(m_strTextBox, "<" & strText & ">")
      dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
      For dblChar = dblStar To dblLastEnd
         For j = 1 To 3
            strData = ""
            If j = 1 Then
               strText = "country": strTitNM = "Àu¥ýÅv°ê®a"
            ElseIf j = 2 Then
               strText = "doc-number": strTitNM = "Àu¥ýÅv¸¹¼Æ"
            ElseIf j = 3 Then
               strText = "date": strTitNM = "Àu¥ýÅv¤é´Á"
            End If
            'Modify By Sindy 2015/6/12
            dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
            If dblStar > dblLastEnd Then
               dblStar = InStr(dblChar, m_strTextBox, "<" & strText & " />") + Len("<" & strText & " />") - 1
               If dblStar < dblRunStar Then
                  dblChar = dblLastEnd
                  Exit For
               Else
                  If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
                  dblEnd = dblStar
                  strData = ""
                  GoTo Step_Claims
               End If
            End If
            '2015/6/12 END
            If dblStar < dblChar Then dblChar = dblLastEnd: Exit For
            If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
            '***** ¸ÑªRXML *****
            If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd) = False Then
            '***** End
               Exit For
            End If
Step_Claims:
            If j = 1 Then 'Àu¥ýÅv°ê®a
               strGetData1 = strData
            ElseIf j = 2 Then 'Àu¥ýÅv¸¹¼Æ
               strGetData2 = strData
            ElseIf j = 3 Then 'Àu¥ýÅv¤é´Á
               strGetData3 = strData
               strClaims = strClaims & ";" & strGetData1 & "!" & strGetData2 & "!" & strGetData3
               
               If dblTPG40 = 0 Then dblTPG40 = strGetData3 'Àu¥ýÅv¤é´Á Add By Sindy 2018/11/12
               strTPG41 = strTPG41 & ";" & strGetData2 'Àu¥ýÅv¸¹¼Æ Add By Sindy 2018/11/12
               strTPG42 = strTPG42 & ";" & strGetData1 'Àu¥ýÅv°ê®a Add By Sindy 2018/11/12
            End If
            dblChar = dblEnd
         Next j
      Next dblChar
   End If
   '2015/6/10 END
   If strTPG41 <> "" Then strTPG41 = Mid(strTPG41, 2) 'Add By Sindy 2018/11/12
   If strTPG42 <> "" Then strTPG42 = Mid(strTPG42, 2) 'Add By Sindy 2018/11/12
End Function

'Add By Sindy 2013/8/27
'ºI¨úXML¸ê®Æ¤G
Private Function GetXmlData2(dblChar As Double, strText As String, strTitNM As String, ByRef strData As String, ByRef dblEnd As Double) As Boolean
Dim dblStar As Double
   
   GetXmlData2 = False
   strData = "": dblEnd = 0
   dblStar = InStr(dblChar, m_strTextBox, "<" & strText)
   dblStar = InStr(dblStar, m_strTextBox, ">")
   If dblStar <= dblChar Then
      Exit Function
   End If
   dblEnd = InStr(dblStar, m_strTextBox, "</" & strText & ">") - 1
   If dblStar >= dblEnd Or dblEnd <= 0 Then
      Exit Function
   End If
   strData = Trim(Mid(m_strTextBox, dblStar + 1, (dblEnd - dblStar)))
   strData = Trim(Replace(ChgSQL(strData), "amp;", ""))
   GetXmlData2 = True
End Function

'ºI¨úXML¸ê®Æ
Private Function GetXmlData(dblChar As Double, strText As String, strTitNM As String, ByRef strData As String, ByRef dblEnd As Double) As Boolean
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
   If Asc(strData) = 13 Then strData = "" 'Add By Sindy 2015/6/11
   GetXmlData = True
End Function

Private Function IsTPG02Exist(ByVal strTPG02 As String, ByRef strErr As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   IsTPG02Exist = False
   '"substr(TPG01,1,9)<>'" & Left(strTPG01, 9) & "' "
   strSql = "SELECT * FROM TPGazette " & _
            "WHERE TPG02='" & strTPG02 & "' AND " & _
                  "TPG01<>'" & strTPG01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsTPG02Exist = True
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         If strErr <> "" Then strErr = strErr & ","
         strErr = strErr & rsTmp.Fields("TPG01")
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function ChkDataErr() As Boolean
Dim rsA As New ADODB.Recordset
Dim rsTemp1 As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim strMsg As String, strErr As String
Dim arrData As Variant, arrData_1 As Variant
Dim strTmpData1 As String, strTmpData2 As String, strTmpData3 As String
Dim strDBData1 As String, strDBData2 As String, strDBData3 As String
Dim bolFind As Boolean
Dim strDBText As String
   
   ChkDataErr = False
   
   Call GetNoticeNumber(CStr(dblTPG03)) '¨ÌÂàÀÉ¤¤ªº¤½¶}¤é¨ú±o¬Û¹ïªº¤½§i¨÷´Á
   If Val(Left(txtTMBM07, 2)) <> Val(strChkTPG04) Then
      strErrTxt = "¤½¶}¤é´Á¡]" & dblTPG03 & "¡^»Pµe­±¤W¿é¤Jªº¤½³ø¨÷¼Æ¡]" & Left(txtTMBM07, 2) & "¡^¤£²Å¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If
   If Val(strTPG04) <> Val(strChkTPG04) Then
      strErrTxt = "¤½¶}¤é´Á¡]" & dblTPG03 & "¡^»P¤½³ø¨÷¼Æ¡]" & strTPG04 & "¡^¤£²Å¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If
   If Val(Right(txtTMBM07, 2)) <> Val(strChkTPG05) Then
      MsgBox "¤½¶}¤é´Á¡]" & dblTPG03 & "¡^»Pµe­±¤W¿é¤Jªº¤½³ø´Á¼Æ¡]" & Right(txtTMBM07, 2) & "¡^¤£²Å¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If
   If Val(strTPG05) <> Val(strChkTPG05) Then
      MsgBox "¤½¶}¤é´Á¡]" & dblTPG03 & "¡^»P¤½³ø´Á¼Æ¡]" & strTPG05 & "¡^¤£²Å¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If
   
   If IsTPG02Exist(strTPG02, strErr) = True Then
      strErrTxt = "¤½¶}¸¹¡]" & strTPG02 & "¡^¤w¦s¦b¡]­«ÂÐªº¥Ó½Ð®×¸¹¡G" & strErr & "¡^¡A¤£¥i¦sÀÉ¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If
   
   '­Y¬°¥»©Ò®×¥ó
   If bolTaieCase = True Then
      strSql = "Select cp09 From CaseProgress Where CP01='" & pa(1) & "' And CP02='" & pa(2) & "' " & _
                                             "And CP03='" & pa(3) & "' And CP04='" & pa(4) & "' " & _
                                             "And CP10='416' And CP27 Is Not Null And CP57 Is Null"
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      '­Y¹êÅé¼f¬d¤wµo¤å¥¼¨ú®ø¦¬¤å
      If rsA.RecordCount > 0 And strTPG09 = "N" Then
         strMsg = strTaieCaseNo & "¦¹®×¥ó¤w´£¹ê¼f"
         Call SaveR04060306("¥Ó½Ð¹êÅé¼f¬d", "µL¡F " & strMsg, "¦³") 'Add By Sindy 2015/6/10
         Call ReadTxt1(strTPG01, strTPG02, strMsg, "", "", "")
         Call PrintPaper(strTPG01, strTPG02, strMsg, "", "")
      '­YµL¹êÅé¼f¬d©Î¹êÅé¼f¬d¥¼µo¤å
      ElseIf rsA.RecordCount <= 0 And strTPG09 = "Y" Then
         strMsg = strTaieCaseNo & "¦¹®×¥ó¥¼´£¹ê¼f¡A½Ð³qª¾±M·~³¡½T»{¸ê®Æ¬O§_¥¿½T"
         Call SaveR04060306("¥Ó½Ð¹êÅé¼f¬d", "¦³¡F " & "¦¹®×¥ó¥¼´£¹ê¼f", "µL") 'Add By Sindy 2015/6/10
         Call ReadTxt1(strTPG01, strTPG02, strMsg, "", "", "")
         Call PrintPaper(strTPG01, strTPG02, strMsg, "", "")
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      
      'Add By Sindy 2015/6/10 ¤ñ¹ï¸ê®Æ¬O§_¤£¤@­P
      '¤½¶}½s¸¹
      'Modify By Sindy 2015/7/6 +And pa(13) <> ""
      If Trim(pa(13)) <> Trim(strTPG02) And pa(13) <> "" Then
         Call SaveR04060306("¤½¶}¸¹", Trim(strTPG02), Trim(pa(13)))
      End If
      '¤½¶}¤é
      'Modify By Sindy 2015/7/6 +And Val(DBDATE(pa(12))) > 0
      If Val(DBDATE(pa(12))) <> Val(dblTPG03) And Val(DBDATE(pa(12))) > 0 Then
         Call SaveR04060306("¤½¶}¤é", CStr(dblTPG03), DBDATE(pa(12)))
      End If
      'µo©ú¤¤¤å¦WºÙ
      If Trim(pa(5)) <> Trim(strCaseChNm) Then
         Call SaveR04060306("µo©ú¤¤¤å¦WºÙ", strCaseChNm, pa(5))
      End If
      'µo©ú­^¤å¦WºÙ
      If Trim(UCase(Replace(pa(6), " ", ""))) <> Trim(UCase(Replace(strCaseEnNm, " ", ""))) Then
         Call SaveR04060306("µo©ú­^¤å¦WºÙ", strCaseEnNm, pa(6))
      End If
      '¥Ó½Ð®×¸¹
      If Trim(pa(11)) <> Trim(strTPG01) Then
         Call SaveR04060306("¥Ó½Ð®×¸¹", strTPG01, pa(11))
      End If
      '¥Ó½Ð¤é
      If Val(DBDATE(pa(10))) <> Val(strApplDate) Then
         Call SaveR04060306("¥Ó½Ð¤é", strApplDate, DBDATE(pa(10)))
      End If
      'Àu¥ýÅv
      If strClaims <> "" Then strClaims = Mid(strClaims, 2)
      strSql = "Select PD05,PD06,PD07,na03||','||na70 na03 From PriDate,nation Where PD01='" & pa(1) & "' And PD02='" & pa(2) & "' " & _
                                             "And PD03='" & pa(3) & "' And PD04='" & pa(4) & "' AND PD07=na01(+) "
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         '¥Ø«e¸ê®Æ®w¤º®e
         rsA.MoveFirst
         strDBText = ""
         Do While Not rsA.EOF
            strDBText = strDBText & ";" & "" & rsA.Fields("na03") & "!" & "" & rsA.Fields("PD06") & "!" & "" & rsA.Fields("PD05")
            rsA.MoveNext
         Loop
         If strDBText <> "" Then strDBText = Mid(strDBText, 2)
         'END
         arrData = Split(strClaims, ";")
         If strClaims = "" Or UBound(arrData) < 0 Or UBound(arrData) + 1 <> rsA.RecordCount Then
            Call SaveR04060306("Àu¥ýÅv", strClaims, strDBText)
         Else
            For i = 0 To UBound(arrData)
               arrData_1 = Split(arrData(i), "!")
               For j = 0 To 2
                  If j = 0 Then strTmpData1 = arrData_1(j) 'Àu¥ýÅv°ê®a
                  If j = 1 Then strTmpData2 = arrData_1(j) 'Àu¥ýÅv¸¹¼Æ
                  If j = 2 Then strTmpData3 = arrData_1(j) 'Àu¥ýÅv¤é´Á
               Next j
               rsA.MoveFirst
               bolFind = False
               Do While Not rsA.EOF
                  If rsA.Fields("PD06") = strTmpData2 Then
                     bolFind = True
                     strDBData1 = "" & rsA.Fields("na03") 'Àu¥ýÅv°ê®a
                     strDBData2 = "" & rsA.Fields("PD06") 'Àu¥ýÅv¸¹¼Æ
                     strDBData3 = "" & rsA.Fields("PD05") 'Àu¥ýÅv¤é´Á
                     If InStr(strDBData1, strTmpData1) = 0 Or _
                        strDBData2 <> strTmpData2 Or _
                        strDBData3 <> strTmpData3 Then
                        Call SaveR04060306("Àu¥ýÅv", strClaims, strDBText)
                        Exit For
                     End If
                     Exit Do
                  End If
                  rsA.MoveNext
               Loop
               If bolFind = False Then
                  Call SaveR04060306("Àu¥ýÅv", strClaims, strDBText)
                  Exit For
               End If
            Next i
         End If
      Else
         If Trim(strClaims) <> "" Then
            Call SaveR04060306("Àu¥ýÅv", strClaims, "")
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      'µo©ú¤H
      If strInventor <> "" Then strInventor = Mid(strInventor, 2)
      strSql = "Select IN04,IN05,substr(NA72,1,2) NA72 From PatentInventor,Inventor,Nation " & _
               "Where PI01='" & pa(1) & "' And PI02='" & pa(2) & "' And PI03='" & pa(3) & "' And PI04='" & pa(4) & "' " & _
               "AND substr(PI06,1,8)=IN01(+) AND substr(PI06,9,2)=IN02(+) " & _
               "AND IN11=na01(+) " & _
               "order by pi05 asc "
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         '¥Ø«e¸ê®Æ®w¤º®e
         rsA.MoveFirst
         strDBText = ""
         Do While Not rsA.EOF
            strDBText = strDBText & ";" & "" & rsA.Fields("IN04") & "!" & "" & rsA.Fields("IN05") & "!" & "" & rsA.Fields("NA72")
            rsA.MoveNext
         Loop
         If strDBText <> "" Then strDBText = Mid(strDBText, 2)
         'END
         arrData = Split(strInventor, ";")
         If strInventor = "" Or UBound(arrData) < 0 Or UBound(arrData) + 1 <> rsA.RecordCount Then
            Call SaveR04060306("µo©ú¤H", strInventor, strDBText)
         Else
            For i = 0 To UBound(arrData)
               arrData_1 = Split(arrData(i), "!")
               For j = 0 To 2
                  If j = 0 Then strTmpData1 = arrData_1(j) 'µo©ú¤H¤¤¤å¦WºÙ
                  If j = 1 Then
                     strTmpData2 = arrData_1(j) 'µo©ú¤H­^¤å¦WºÙ
                     If pa(1) = "P" Then strTmpData2 = "" 'Add By Sindy 2015/7/7 ±M§Q³B¤£¤ñ¹ïµo©ú¤H­^¤å
                  End If
                  If j = 2 Then strTmpData3 = arrData_1(j) 'µo©ú¤H°êÄy
               Next j
               rsA.MoveFirst
               bolFind = False
               Do While Not rsA.EOF
                  If rsA.Fields("IN04") = strTmpData1 Then
                     bolFind = True
                     strDBData1 = "" & rsA.Fields("IN04") 'µo©ú¤H¤¤¤å¦WºÙ
                     strDBData2 = "" & rsA.Fields("IN05") 'µo©ú¤H­^¤å¦WºÙ
                     If pa(1) = "P" Then strDBData2 = "" 'Add By Sindy 2015/7/7 ±M§Q³B¤£¤ñ¹ïµo©ú¤H­^¤å
                     strDBData3 = "" & rsA.Fields("NA72") 'µo©ú¤H°êÄy
                     
                     If Trim(UCase(Replace(Replace(strDBData1, "¡@", ""), " ", ""))) <> Trim(UCase(Replace(Replace(strTmpData1, "¡@", ""), " ", ""))) Or _
                        Trim(UCase(Replace(strDBData2, " ", ""))) <> Trim(UCase(Replace(strTmpData2, " ", ""))) Or _
                        UCase(strDBData3) <> UCase(strTmpData3) Then
                        Call SaveR04060306("µo©ú¤H", strInventor, strDBText)
                        Exit For
                     End If
                     Exit Do
                  End If
                  rsA.MoveNext
               Loop
               If bolFind = False Then
                  Call SaveR04060306("µo©ú¤H", strInventor, strDBText)
                  Exit For
               End If
            Next i
         End If
      Else
         If Trim(strInventor) <> "" Then
            Call SaveR04060306("µo©ú¤H", strInventor, "")
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      '¥Ó½Ð¤H
      If strApplName <> "" Then strApplName = Mid(strApplName, 2)
      'Modified by Morgan 2023/1/16 +Srt(¨Ì¥Ó½Ð¤H¶¶§Ç±Æ)
      strSql = "Select CU04,rtrim(ltrim(cu05||' '||cu88||' '||cu89||' '||cu90)) CU05,substr(NA72,1,2) NA72,1 Srt From Patent,customer,nation Where Pa01='" & pa(1) & "' And Pa02='" & pa(2) & "' And Pa03='" & pa(3) & "' And Pa04='" & pa(4) & "' AND Pa26 is not null AND substr(Pa26,1,8)=cu01(+) AND substr(Pa26,9,1)=cu02(+) AND substr(CU10,1,3)=na01(+)" & _
               " union Select CU04,rtrim(ltrim(cu05||' '||cu88||' '||cu89||' '||cu90)) CU05,substr(NA72,1,2) NA72,2 Srt From Patent,customer,nation Where Pa01='" & pa(1) & "' And Pa02='" & pa(2) & "' And Pa03='" & pa(3) & "' And Pa04='" & pa(4) & "' AND Pa27 is not null AND substr(Pa27,1,8)=cu01(+) AND substr(Pa27,9,1)=cu02(+) AND substr(CU10,1,3)=na01(+)" & _
               " union Select CU04,rtrim(ltrim(cu05||' '||cu88||' '||cu89||' '||cu90)) CU05,substr(NA72,1,2) NA72,3 Srt From Patent,customer,nation Where Pa01='" & pa(1) & "' And Pa02='" & pa(2) & "' And Pa03='" & pa(3) & "' And Pa04='" & pa(4) & "' AND Pa28 is not null AND substr(Pa28,1,8)=cu01(+) AND substr(Pa28,9,1)=cu02(+) AND substr(CU10,1,3)=na01(+)" & _
               " union Select CU04,rtrim(ltrim(cu05||' '||cu88||' '||cu89||' '||cu90)) CU05,substr(NA72,1,2) NA72,4 Srt From Patent,customer,nation Where Pa01='" & pa(1) & "' And Pa02='" & pa(2) & "' And Pa03='" & pa(3) & "' And Pa04='" & pa(4) & "' AND Pa29 is not null AND substr(Pa29,1,8)=cu01(+) AND substr(Pa29,9,1)=cu02(+) AND substr(CU10,1,3)=na01(+)" & _
               " union Select CU04,rtrim(ltrim(cu05||' '||cu88||' '||cu89||' '||cu90)) CU05,substr(NA72,1,2) NA72,5 Srt From Patent,customer,nation Where Pa01='" & pa(1) & "' And Pa02='" & pa(2) & "' And Pa03='" & pa(3) & "' And Pa04='" & pa(4) & "' AND Pa30 is not null AND substr(Pa30,1,8)=cu01(+) AND substr(Pa30,9,1)=cu02(+) AND substr(CU10,1,3)=na01(+) order by Srt"
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         '¥Ø«e¸ê®Æ®w¤º®e
         rsA.MoveFirst
         strDBText = ""
         Do While Not rsA.EOF
            strDBText = strDBText & ";" & "" & rsA.Fields("CU04") & "!" & "" & rsA.Fields("CU05") & "!" & "" & rsA.Fields("NA72")
            rsA.MoveNext
         Loop
         If strDBText <> "" Then strDBText = Mid(strDBText, 2)
         'END
         arrData = Split(strApplName, ";")
         If strApplName = "" Or UBound(arrData) < 0 Or UBound(arrData) + 1 <> rsA.RecordCount Then
            Call SaveR04060306("¥Ó½Ð¤H", strApplName, strDBText)
         Else
            For i = 0 To UBound(arrData)
               arrData_1 = Split(arrData(i), "!")
               For j = 0 To 2
                  If j = 0 Then strTmpData1 = arrData_1(j) '¥Ó½Ð¤H¤¤¤å¦WºÙ
                  If j = 1 Then strTmpData2 = arrData_1(j) '¥Ó½Ð¤H­^¤å¦WºÙ
                  If j = 2 Then strTmpData3 = arrData_1(j) '¥Ó½Ð¤H°êÄy
               Next j
               rsA.MoveFirst
               bolFind = False
               Do While Not rsA.EOF
                  If rsA.Fields("CU04") = strTmpData1 Then
                     bolFind = True
                     strDBData1 = "" & rsA.Fields("CU04") '¥Ó½Ð¤H¤¤¤å¦WºÙ
                     strDBData2 = "" & rsA.Fields("CU05") '¥Ó½Ð¤H­^¤å¦WºÙ
                     strDBData3 = "" & rsA.Fields("NA72") '¥Ó½Ð¤H°êÄy
                     
                     If Trim(UCase(Replace(Replace(strDBData1, "¡@", ""), " ", ""))) <> Trim(UCase(Replace(Replace(strTmpData1, "¡@", ""), " ", ""))) Or _
                        Trim(UCase(Replace(strDBData2, " ", ""))) <> Trim(UCase(Replace(strTmpData2, " ", ""))) Or _
                        UCase(strDBData3) <> UCase(strTmpData3) Then
                        Call SaveR04060306("¥Ó½Ð¤H", strApplName, strDBText)
                        Exit For
                     End If
                     Exit Do
                  End If
                  rsA.MoveNext
               Loop
               If bolFind = False Then
                  Call SaveR04060306("¥Ó½Ð¤H", strApplName, strDBText)
                  Exit For
               End If
            Next i
         End If
      Else
         Call SaveR04060306("¥Ó½Ð¤H", strApplName, "")
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      '¥N²z¤H
      If strAgent <> "" Then strAgent = Mid(strAgent, 2)
      strSql = "Select cp110 From CaseProgress Where CP01='" & pa(1) & "' And CP02='" & pa(2) & "' " & _
                                             "And CP03='" & pa(3) & "' And CP04='" & pa(4) & "' " & _
                                             "And instr('" & NewCasePtyList & "',CP10)>0 And CP27 Is Not Null And CP57 Is Null"
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         '¥Ø«e¸ê®Æ®w¤º®e
         strDBText = "" & rsA.Fields("cp110")
         'END
         arrData = Split(strAgent, ";")
         arrData_1 = Split("" & rsA.Fields("cp110"), ",")
         If (Trim(strAgent) = "" And Trim(strDBText) <> "") Or _
            (Trim(strAgent) <> "" And Trim(strDBText) = "") Or _
            UBound(arrData) + 1 <> UBound(arrData_1) + 1 Then
            Call SaveR04060306("¥N²z¤H", strAgent, strDBText)
         Else
            If strAgent <> "" And strDBText <> "" Then
               For i = 0 To UBound(arrData_1)
                  strDBData1 = ""
                  If arrData_1(i) = "81040" Then
                     strDBData1 = "ÀF±Ò®õ"
                  Else
                     strExc(0) = "SELECT st02 FROM staff WHERE ST01=" & CNULL(CStr(arrData_1(i)))
                     intI = 1
                     Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strDBData1 = rsTemp1.Fields("st02")
                     End If
                  End If
                  If InStr(strAgent, strDBData1) = 0 Then
                     Call SaveR04060306("¥N²z¤H", strAgent, strDBText)
                     Exit For
                  End If
               Next i
            End If
         End If
      Else
         If Trim(strAgent) <> "" Then
            Call SaveR04060306("¥N²z¤H", strAgent, "")
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      '2015/6/10 END
   End If
   Set rsA = Nothing
End Function

'Add By Sindy 2015/6/10
Private Sub SaveR04060306(strItem As String, strText As String, strDBText As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim intCnt As Integer
   
   strErrTxt = "·s¼W¤ñ¹ïÂàÀÉ¸ê®Æ¼È¦sÀÉ.R04060306"
   strSql = "SELECT nvl(max(Rseqno),0) FROM R04060306" & _
            " where RCP01='" & pa(1) & "' and RCP02='" & pa(2) & "' and RCP03='" & pa(3) & "' and RCP04='" & pa(4) & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      intCnt = rsTmp.Fields(0) + 1
   Else
      intCnt = 1
   End If
   rsTmp.Close
   
   strSql = "insert into R04060306(RCP01,RCP02,RCP03,RCP04,Rseqno,Ritem,Rtext,Rdbtext) " & _
            "values(" & CNULL(pa(1)) & "," & CNULL(pa(2)) & "," & CNULL(pa(3)) & "," & CNULL(pa(4)) & _
            "," & CStr(intCnt) & "," & CNULL(strItem) & "," & CNULL(ChgSQL(strText)) & _
            "," & CNULL(ChgSQL(strDBText)) & ")"
   cnnConnection.Execute strSql
   
   Set rsTmp = Nothing
End Sub

'¦a°Ï¦WºÙ¸ê®ÆÀË®Öªí
Private Sub ReadTxt1(strTPG01 As String, strTPG02 As String, strTPG06 As String, strTPG07 As String, strAChinese1 As String, strAddress1 As String)
Dim i As Integer
   
   If m_PrintRpt1 = False Then
      m_PrintRpt1 = True
'      If ff1 > 0 Then Close #ff1
'      ff1 = FreeFile
      m_strFileName1 = "°ê¤º±M§Q¤½¶}¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¸ê®ÆÀË®Öªí.txt"
'      Open PUB_Getdesktop & "\" & m_strFileName1 For Output As ff1
'      Print #ff1, "³Æµù¡G§ï¦r«¬Fixedsys¼Ð·Ç11¸¹¦r¥H¾î¦¡¤W¤U¥ª¥k¦U10MM¦C¦L"
'      Print #ff1, "¥Ó½Ð®×¸¹        ¤½¶}¸¹     ¦a°Ï¦WºÙ        ¥N²z¤H¦WºÙ   ¥Ó½Ð¤H¦a§}"
'      Print #ff1, "                           ©Î ´£¿ô³Æµù"
'      Print #ff1, "=============== ========== =============== ============ ============================================="
      
      m_strText = "³Æµù¡G§ï¦r«¬Fixedsys¼Ð·Ç11¸¹¦r¥H¾î¦¡¤W¤U¥ª¥k¦U10MM¦C¦L" & vbCrLf
      m_strText = m_strText & "¥Ó½Ð®×¸¹        ¤½¶}¸¹     ¦a°Ï¦WºÙ        ¥N²z¤H¦WºÙ   ¥Ó½Ð¤H¦a§}" & vbCrLf
      m_strText = m_strText & "                           ©Î ´£¿ô³Æµù" & vbCrLf
      m_strText = m_strText & "=============== ========== =============== ============ =============================================" & vbCrLf
   End If
   For i = 1 To 6
      strTemp(i) = ""
   Next i
   strTemp(1) = Trim(strTPG01)
   strTemp(2) = Trim(strTPG02)
   strTemp(3) = Trim(strTPG06)
   strTemp(4) = Trim(strTPG07)
   strTemp(5) = Trim(strAChinese1)
   strTemp(6) = Trim(strAddress1)
   
   If strTemp(3) = "" Then  '020.¤¤°ê¤j³° Or strTemp(3) = "020"
      strTemp(3) = "*" & strTemp(3) & GetPrjNationName(strTemp(3))
   Else
      strTemp(3) = strTemp(3) & GetPrjNationName(strTemp(3))
   End If
   txtChkWord = strTemp(4) 'Add By Sindy 2024/5/17
   If InStr(txtChkWord, "?") > 0 Then
      strTemp(4) = "*" & strTemp(4)
   End If
   
   strTemp(1) = convForm(CheckStr(strTemp(1)), 15)
   strTemp(2) = convForm(CheckStr(strTemp(2)), 10)
   If strTemp(5) <> "" Then '¥Nªí¶Ç¤Jªº¸ê®Æ¬°´£¿ô³Æµù¡A«hÅã¥Ü¥þ³¡¤º®e
      strTemp(3) = convForm(CheckStr(strTemp(3)), 15)
   End If
   strTemp(4) = convForm(CheckStr(strTemp(4)), 12)
   strTemp(5) = convForm(CheckStr(strTemp(5)), 45)
   strTemp(6) = convForm(CheckStr(strTemp(6)), 45)
   'Print #ff1, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(6)
   m_strText = m_strText & strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(6) & vbCrLf
End Sub

'­­©w¦r¦êªø«×
'Remove by Lydia 2018/08/24 »PbasQuery­«½Æ
'Private Function convForm(ByVal p_InStr As String, ByVal p_Num As Integer, Optional ByVal p_Char As String = " ") As String
'   convForm = StrConv(LeftB(StrConv(p_InStr & String(p_Num, p_Char), vbFromUnicode), p_Num), vbUnicode)
'End Function

Private Sub Form_Load()
Dim SeekPrintL As Integer
Dim i As Integer, j As Integer
   
   MoveFormToCenter Me
   
   'Modify By Sindy 2012/1/16
   MaxHeight = 4305
   MinHeight = 3450
   '2012/1/16 End
   
   Me.Height = MinHeight
   
   m_DefaultPrinter = Printer.DeviceName
'   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      'cmbPrinter2.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = m_DefaultPrinter Then
         SeekPrint = i
      End If
   Next i
   Set Printer = Printers(SeekPrint)
   'cmbPrinter2.Text = cmbPrinter2.List(SeekPrint)
   
   'Add By Sindy 2013/8/27
   If Pub_StrUserSt03 = "M51" Then
      cmdPA160.Visible = True
      cmdIPC.Visible = True
   Else
      cmdPA160.Visible = False
      cmdIPC.Visible = False
   End If
   '2013/8/27 END
   
   PUB_ReadPath txtPath1, Me.Name 'Added by Sindy 2020/5/5
   
   'Add By Sindy 2022/3/3
   Set adoStream = New ADODB.Stream
   adoStream.Charset = "UTF-8" '"UTF-8" Unicode
   adoStream.Open
   '2022/3/3 END

End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SavePath txtPath1, Me.Name 'Added by Sindy 2020/5/5
   
   'Add By Sindy 2022/3/3
   adoStream.Close
   Set adoStream = Nothing
   '2022/3/3 END
   
   Set frm04060306 = Nothing
End Sub

Private Sub text03_GotFocus()
   InverseTextBox text03
End Sub

Private Sub text03_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(text03) = False Then
      If CheckIsTaiwanDate(text03, False) = False Then
         Cancel = True
         strMsg = "½Ð¿é¤J¥¿½Tªº¤½¶}¤é"
         strTit = "¸ê®ÆÀË®Ö"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03_GotFocus
         GoTo EXITSUB
      End If
      
      '¤½¶}¤é¤£¯à¤j©ó¨t²Î¤é
      If DBDATE(text03) > strSrvDate(1) Then
         Cancel = True
         strMsg = "¤½¶}¤é¤£¯à¤j©ó¨t²Î¤é"
         strTit = "¸ê®ÆÀË®Ö"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03_GotFocus
      End If
   End If
EXITSUB:
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
         strMsg = "¤½³ø¨÷´Á¥u¥i¿é¤J¼Æ­È¸ê®Æ¡I"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTMBM07_GotFocus
         Exit Sub
      End If
      If Len(txtTMBM07) <> 4 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¤½³ø¨÷´Á¬°4½X¡I"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTMBM07_GotFocus
         Exit Sub
      End If
      Call IsRecordExist
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

If IsEmptyText(text03) = True Then
   strTit = "ÀË®Ö¸ê®Æ"
   strMsg = "½Ð¿é¤J¤½¶}¤é¡I"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   text03.SetFocus
   Exit Function
End If

If Me.txtTMBM07.Enabled = True Then
   Cancel = False
   txtTMBM07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.text03.Enabled = True Then
   Cancel = False
   text03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

Call GetNoticeNumber(DBDATE(text03)) '¨Ì¿é¤Jªº¤½¶}¤é¨ú±o¬Û¹ïªº¤½§i¨÷´Á
If Val(Left(txtTMBM07, 2)) <> Val(strChkTPG04) Then
   strTit = "ÀË®Ö¸ê®Æ"
   strMsg = "¤½³ø¨÷¼Æ»P¤½¶}¤é´Á¤£²Å¡I"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   text03.SetFocus
   Exit Function
End If
If Val(Right(txtTMBM07, 2)) <> Val(strChkTPG05) Then
   strTit = "ÀË®Ö¸ê®Æ"
   strMsg = "¤½³ø´Á¼Æ»P¤½¶}¤é´Á¤£²Å¡I"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   text03.SetFocus
   Exit Function
End If

If IsEmptyText(txtPath2) = True Then
   strTit = "ÀË®Ö¸ê®Æ"
   'strMsg = "½Ð¿é¤J¥úºÐ¥Øªº¸ô®|¡I"
   strMsg = "½Ð¿é¤J«þ¨©¥Øªº¸ô®|¡I"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   txtPath2.SetFocus
   Exit Function
End If

TxtValidate = True
End Function

' ÀË¬d°O¿ý¬O§_¤w¸g¦s¦b
Private Function IsRecordExist() As Boolean
   Dim rsTmp2 As New ADODB.Recordset
   Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   IsRecordExist = False
   
   strSql = "SELECT count(TPG01) FROM TPGazette WHERE TPG04=" & CNULL(Left(txtTMBM07, 2)) & " and TPG05=" & CNULL(Right(txtTMBM07, 2))
   
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
Private Function GetTAgentName(ByVal strData As String, ByVal strCol As String) As String
Dim strSql As String
Dim rsTmp2 As New ADODB.Recordset
   
   GetTAgentName = Empty
   strSql = "SELECT * FROM TAGENT " & _
            "WHERE TA01 = 'P' AND " & _
                  "TA02 = '" & strData & "' "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      If UCase(strCol) = "TA03" Then
         If IsNull(rsTmp2.Fields("TA03")) = False Then
            GetTAgentName = rsTmp2.Fields("TA03")
         End If
      ElseIf UCase(strCol) = "TA04" Then
         If IsNull(rsTmp2.Fields("TA04")) = False Then
            GetTAgentName = rsTmp2.Fields("TA04")
         End If
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
            "where OA01 in('P','FCP') " & _
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

' ¨ú±o°ê®aªº¥N½X
'Modify By Sindy 2019/9/4 + , ByRef strData_Nm As String °ê®a¦WºÙ
Private Function GetNationNo(ByRef strData As String, ByRef strData_Nm As String) As String
Dim strSql As String
Dim rsTmp2 As New ADODB.Recordset
Dim arrData, i As Integer 'Add By Sindy 2013/3/19
   
   GetNationNo = Empty
   
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA03 = '" & strData & "' AND length(na01)=3 "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      If IsNull(rsTmp2.Fields("NA71")) = False Then
         GetNationNo = rsTmp2.Fields("NA71")
         strData = rsTmp2.Fields("NA71")
         strData_Nm = rsTmp2.Fields("NA03") 'Add By Sindy 2019/9/4
         rsTmp2.Close: Set rsTmp2 = Nothing: Exit Function
      End If
   End If
   rsTmp2.Close
   
   If GetNationNo = "" Then
      'Modify By Sindy 2013/3/5 NA70·|¦s©ñ¦h­Ó¤½³ø¦a°Ï¦WºÙ
'      strSql = "SELECT * FROM NATION " & _
'               "WHERE NA70 = '" & strData & "' "
      strSql = "SELECT * FROM NATION " & _
               "WHERE instr(NA70,'" & strData & "')>0 AND length(na01)=3 "
      rsTmp2.CursorLocation = adUseClient
      rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp2.RecordCount > 0 Then
         'Modify By Sindy 2013/3/19
'         If IsNull(rsTmp2.Fields("NA71")) = False Then
'            GetNationNo = rsTmp2.Fields("NA71")
'            strData = rsTmp2.Fields("NA71")
'            rsTmp2.Close: Set rsTmp2 = Nothing: Exit Function
'         End If
         rsTmp2.MoveFirst
         Do While Not rsTmp2.EOF
            arrData = Split(rsTmp2.Fields("NA70"), ",")
            For i = 0 To UBound(arrData)
               If arrData(i) = strData Then
                  If IsNull(rsTmp2.Fields("NA71")) = False Then
                     GetNationNo = rsTmp2.Fields("NA71")
                     strData = rsTmp2.Fields("NA71")
                     strData_Nm = rsTmp2.Fields("NA03") 'Add By Sindy 2019/9/4
                     rsTmp2.Close: Set rsTmp2 = Nothing: Exit Function
                  End If
               End If
            Next i
            rsTmp2.MoveNext
         Loop
         '2013/3/19 End
      End If
      rsTmp2.Close
   End If
      
   Set rsTmp2 = Nothing
End Function

' ¼Ò½k¤ñ¹ï¯S®í¦a°Ï¦WºÙ
'Modify By Sindy 2019/9/4 + , ByRef strData_Nm As String °ê®a¦WºÙ
Private Function GetNationLike(ByVal strData As String, ByRef strData_Nm As String) As String
Dim strSql As String
Dim rsTmp2 As New ADODB.Recordset
Dim arrData, i As Integer 'Add By Sindy 2013/3/19
   
   GetNationLike = Empty
   
   strSql = "SELECT * FROM NATION WHERE instr('" & strData & "',na03)>0 AND length(na01)=3 order by length(na03) desc "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      rsTmp2.MoveFirst
      If IsNull(rsTmp2.Fields("NA71")) = False Then
         GetNationLike = rsTmp2.Fields("NA71")
         strData_Nm = rsTmp2.Fields("NA03") 'Add By Sindy 2019/9/4
         rsTmp2.Close
         Set rsTmp2 = Nothing
         Exit Function
      End If
   End If
   rsTmp2.Close
   
   'Modify By Sindy 2013/3/5 NA70·|¦s©ñ¦h­Ó¤½³ø¦a°Ï¦WºÙ
   'strSql = "SELECT * FROM NATION WHERE instr('" & strData & "',na70)>0 order by length(na70) desc "
   strSql = "SELECT * FROM NATION WHERE instr('" & strData & "',na70)>0 and instr(na70,',')=0 AND length(na01)=3 order by length(na70) desc" 'Modify By Sindy 2013/3/19
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      rsTmp2.MoveFirst
      If IsNull(rsTmp2.Fields("NA71")) = False Then
         GetNationLike = rsTmp2.Fields("NA71")
         strData_Nm = rsTmp2.Fields("NA03") 'Add By Sindy 2019/9/4
         rsTmp2.Close
         Set rsTmp2 = Nothing
         Exit Function
      End If
   End If
   rsTmp2.Close
   
   'Add By Sindy 2013/3/19
   strSql = "SELECT * FROM NATION WHERE instr(na70,',')>0 AND length(na01)=3 "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      rsTmp2.MoveFirst
      Do While Not rsTmp2.EOF
         arrData = Split(rsTmp2.Fields("NA70"), ",")
         For i = 0 To UBound(arrData)
            If InStr(strData, arrData(i)) > 0 Then
               If IsNull(rsTmp2.Fields("NA71")) = False Then
                  GetNationLike = rsTmp2.Fields("NA71")
                  strData_Nm = rsTmp2.Fields("NA03") 'Add By Sindy 2019/9/4
                  rsTmp2.Close
                  Set rsTmp2 = Nothing
                  Exit Function
               End If
            End If
         Next i
         rsTmp2.MoveNext
      Loop
   End If
   rsTmp2.Close
   '2013/3/19 End
   
   '°w¹ï¤j³°¦a°Ï
   strSql = "SELECT * FROM NATION WHERE na02='B00' and na03 like '%¥«' and instr('" & strData & "',replace(na03,'¥«',''))>0 AND length(na01)=3 "
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp2.RecordCount > 0 Then
      rsTmp2.MoveFirst
      If IsNull(rsTmp2.Fields("NA71")) = False Then
         GetNationLike = rsTmp2.Fields("NA71")
         strData_Nm = rsTmp2.Fields("NA03") 'Add By Sindy 2019/9/4
         rsTmp2.Close
         Set rsTmp2 = Nothing
         Exit Function
      End If
   End If
   rsTmp2.Close
   
   Set rsTmp2 = Nothing
End Function

Private Sub GetNoticeNumber(strDate As String)
Dim i As Integer, j As Integer
   
   strChkTPG04 = Format(Val(Val(Left(strDate, 4)) - 1911) - 91, "00")
   
   j = Val(Mid(strDate, 5, 2))
   i = (j - 1) * 2
   j = Val(Right(strDate, 2))
   If j >= 1 And j < 11 Then
      i = i + 1
   ElseIf j >= 11 And j < 21 Then
      i = i + 2
   End If
   '92¦~¤½³ø±q5¤ë¶}©l
   If Val(strDate) < 20040000 Then i = i - 8
   strChkTPG05 = Format(i, "00")
End Sub

Private Sub PrintPaper(strTPG01 As String, strTPG02 As String, strTPG06 As String, strTPG07 As String, strAddress1 As String)
   intPRow = intPRow + 1
   MSHFlexGrid1.Rows = intPRow + 1
   
   MSHFlexGrid1.TextMatrix(intPRow, 0) = strTPG01
   MSHFlexGrid1.TextMatrix(intPRow, 1) = strTPG02
   
   If strTPG06 = "" Then
      MSHFlexGrid1.TextMatrix(intPRow, 2) = "*"
   Else
      MSHFlexGrid1.TextMatrix(intPRow, 2) = strTPG06 & GetPrjNationName(strTPG06)
   End If
   
   txtChkWord = strTPG07 'Add By Sindy 2024/5/17
   If InStr(txtChkWord, "?") > 0 Then
      MSHFlexGrid1.TextMatrix(intPRow, 3) = "*" & strTPG07
   Else
      MSHFlexGrid1.TextMatrix(intPRow, 3) = strTPG07
   End If
   
   MSHFlexGrid1.TextMatrix(intPRow, 4) = strAddress1
End Sub

Private Sub PrintRpt()
Dim i As Integer, j As Integer
   
   For j = 1 To MSHFlexGrid1.Rows - 1
      For i = 1 To 5
         strTemp(i) = ""
      Next i
      
      strTemp(1) = MSHFlexGrid1.TextMatrix(j, 0)
      strTemp(2) = MSHFlexGrid1.TextMatrix(j, 1)
      strTemp(3) = MSHFlexGrid1.TextMatrix(j, 2)
      strTemp(4) = MSHFlexGrid1.TextMatrix(j, 3)
      strTemp(5) = MSHFlexGrid1.TextMatrix(j, 4)
      If iLine2 > 34 Or iLine2 = 0 Then
         If iLine2 > 0 Then Printer.NewPage
         PrintTitle '¦C¦LªíÀY
      End If
      PrintDetail '¦C¦L©ú²Ó
   Next j
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 1800
PLeft(3) = 3200
PLeft(4) = 5000
PLeft(5) = 6500
End Sub

Sub PrintTitle()
If m_PrintRpt2 = False Then
'   Printer.EndDoc
   Printer.Orientation = 2 '1.ª½¦L 2.¾î¦L
   m_PrintRpt2 = True
End If

GetPleft
iLine2 = 1

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("°ê¤º±M§Q¤½¶}¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¸ê®ÆÀË®Öªí") / 2)
Printer.CurrentY = iLine2 * 300
Printer.Print "°ê¤º±M§Q¤½¶}¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¸ê®ÆÀË®Öªí"

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

iLine2 = iLine2 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "¦C¦L¤H­û¡G" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(strSrvDate(2))
iLine2 = iLine2 + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 1200
Printer.Print "­¶¡@¡@¦¸¡G" & Printer.Page

iLine2 = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine2 * 300
Printer.Print "¥Ó½Ð®×¸¹"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine2 * 300
Printer.Print "¤½¶}¸¹"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine2 * 300
Printer.Print "¦a°Ï¦WºÙ"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine2 * 300
Printer.Print "¥N²z¤H¦WºÙ"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine2 * 300
Printer.Print "¥Ó½Ð¤H¦a§}"
iLine2 = 6
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine2 * 300
Printer.Print "©Î ´£¿ô³Æµù"

iLine2 = iLine2 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine2 * 300
Printer.Print String(205, "-")
iLine2 = iLine2 + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
   For m_j = 1 To 5
      Printer.CurrentX = PLeft(m_j)
      Printer.CurrentY = iLine2 * 300
      Printer.Print strTemp(m_j)
   Next m_j
   iLine2 = iLine2 + 1
End Sub

Private Sub ResetGrid()
   With MSHFlexGrid1
      .Clear
      .Rows = 2
      .FixedRows = 1
      .FixedCols = 0
      .FormatString = "¥Ó½Ð®×¸¹|¤½¶}¸¹|¦a°Ï¦WºÙ|¥N²z¤H¦WºÙ|¥Ó½Ð¤H¦a§}"
   End With
End Sub
