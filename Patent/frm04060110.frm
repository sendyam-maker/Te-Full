VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060110 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "±M§Q¤½³øÂàÀÉ§@·~"
   ClientHeight    =   5640
   ClientLeft      =   40
   ClientTop       =   280
   ClientWidth     =   6190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6190
   Begin VB.CommandButton cmdPath 
      Height          =   330
      Left            =   5490
      Picture         =   "frm04060110.frx":0000
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   21
      Top             =   810
      Width           =   350
   End
   Begin VB.CommandButton cmdTPB12 
      Caption         =   "¸ÉÂà®×¥óÄÝ©Ê"
      Height          =   400
      Left            =   4560
      TabIndex        =   20
      Top             =   1560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdTemp 
      Caption         =   "Âà¼È¦sÀÉ-¥Ó½Ð¤H"
      Height          =   400
      Left            =   4560
      TabIndex        =   7
      Top             =   2610
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdPA160 
      Caption         =   "¸ÉÂà°ê»Ú¤ÀÃþ"
      Height          =   400
      Left            =   4560
      TabIndex        =   6
      Top             =   2070
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   3600
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
      Left            =   90
      TabIndex        =   13
      Top             =   3480
      Width           =   6015
      Begin VB.TextBox Text2 
         Alignment       =   2  '¸m¤¤¹ï»ô
         BackColor       =   &H00FF0000&
         Height          =   300
         Left            =   30
         TabIndex        =   15
         Top             =   120
         Width           =   5940
      End
   End
   Begin VB.FileListBox File2 
      Height          =   240
      Left            =   1560
      TabIndex        =   12
      Top             =   210
      Visible         =   0   'False
      Width           =   525
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   405
      Left            =   960
      TabIndex        =   11
      Top             =   210
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   864
      _ExtentY        =   723
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frm04060110.frx":0102
   End
   Begin VB.TextBox txtPath2 
      Height          =   315
      Left            =   1410
      TabIndex        =   3
      Text            =   "C:\GAZETTE\PXml"
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
      Caption         =   "«þ¨©¸ê®Æ(&C)"
      Height          =   400
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5100
      TabIndex        =   8
      Top             =   240
      Width           =   912
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   270
      Top             =   210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1515
      Left            =   60
      TabIndex        =   19
      Top             =   4050
      Width           =   6045
      _ExtentX        =   10672
      _ExtentY        =   2663
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSForms.TextBox txtChkWord 
      Height          =   300
      Left            =   0
      TabIndex        =   22
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
      Caption         =   "¤½§i¤é¡G"
      Height          =   180
      Left            =   300
      TabIndex        =   18
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "¤½³ø¨÷´Á¡G"
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   2070
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "(               µ§)"
      Height          =   210
      Left            =   2190
      TabIndex        =   16
      Top             =   2070
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "ÂàÀÉ¤¤, ½Ðµy­Ô. . .(½Ð¤Å¥ô·NÃö³¬¦¹§@·~)"
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
      Left            =   90
      TabIndex        =   14
      Top             =   3120
      Width           =   6015
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "«þ¨©¥Øªº¸ô®|¡G"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ÀÉ®×¨Ó·½¸ô®|¡G"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1260
   End
End
Attribute VB_Name = "frm04060110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/3/3 Form2.0¤w­×§ï
'Memo By Morgan 2012/12/11 ´¼Åv¤H­ûÄæ¤w­×§ï
Option Explicit

Dim m_bolCharQ  As Boolean, m_strCharQNote As String
Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iLine2 As Integer
Dim m_PrintRpt1 As Boolean, m_PrintRpt2 As Boolean
Dim ff1 As Integer, FF2 As Integer
Dim m_strFileName1 As String, m_strFileName2 As String
Dim strErrTxt As String
Dim strTPB01 As String, strTPB02 As String, dblTPB03 As Double, strTPB04 As String
Dim strTPB05 As String, strTPB06 As String, strTPB07 As String, strTPB07_1 As String, strTPB07_temp1 As String
Dim strTPB08 As String, strTPB09 As String
'Add By Sindy 2012/8/9
Dim strTPB10 As String, strTPB11 As String, m_PI02 As String, strTPB12 As String
'2012/8/9 End
Dim strTPB13 As String 'Add By Sindy 2016/3/2
Dim strTPB38 As String 'Add By Sindy 2019/9/4
Dim strTPBcApp(10) As String 'Add By Sindy 2013/4/15
'Add By Sindy 2018/11/12
Dim strTPBeApp(10) As String
Dim dblTPB34 As Double, dblTPB35 As Double, strTPB36 As String, strTPB37 As String
'2018/11/12 END
Dim strAChinese As String, strAChinese1 As String, strAddress1 As String
Dim strOurAgentName As String
Dim pa() As String
Dim m_strPA14 As String '¹w©w¤½§i¤é
Dim m_bol412 As Boolean '¬O§_¦³µo¤å©µ½w¤½§i
Dim bolTaieCase As Boolean '¬O§_¬°¥»©Ò®×¥ó
Dim strTaieCaseNo As String
Dim m_strNextDueDate As String  '¤U¦¸Ãº¶O¤éªk©w´Á­­
Dim m_strNextFeeDate As String  '¤U¦¸Ãº¶O¤é¥»©Ò´Á­­
Dim m_strAgreeOnDate As String 'Add By Sindy 2021/8/17 ¤U¦¸Ãº¶O¤é¬ù©w´Á­­
Dim m_str421CP09 As String '§Þ³N³ø§iÁ`¦¬¤å¸¹
Dim m_str421CP14 As String '§Þ³N³ø§i©Ó¿ì¤H
Dim m_str421EP06 As String '§Þ³N³ø§i¤å¥ó»ô³Æ¤é
Dim m_str421CP48 As String '§Þ³N³ø§i©Ó¿ì´Á­­
Dim strChkTPB04 As String, strChkTPB05 As String
Dim m_DefaultPrinter As String
Dim SeekPrint As Integer
'Add By Sindy 2012/1/16
Dim intPRow As Integer
Dim MaxHeight As Integer, MinHeight As Integer
'2012/1/16 End
'Add By Sindy 2012/3/3
Dim strPA160 As String
Dim strMsg As String
Dim i As Integer, j As Integer
'2012/3/3 End
Dim adoStream As ADODB.Stream 'Add By Sindy 2022/3/3
Dim m_strTextBox As String 'Add by Sindy 2022/3/3
Dim m_strText As String 'Add by Sindy 2024/5/17


Private Sub cmdCopy_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim fs As Object, strTime As String
Dim DeleteFilePathErr As Boolean
Dim strPath As String, oFolder As Folder, oFile As File, strPathTemp As String
   
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
      strMsg = "½Ð¿é¤J¤½§i¤é¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text03.SetFocus
      Exit Sub
   End If
   Call GetNoticeNumber(DBDATE(text03)) '¨Ì¿é¤Jªº¤½§i¤é¨ú±o¬Û¹ïªº¤½§i¨÷´Á
   If Val(Left(txtTMBM07, 2)) <> Val(strChkTPB04) Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "¤½³ø¨÷¼Æ»P¤½§i¤é´Á¤£²Å¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text03.SetFocus
      Exit Sub
   End If
   If Val(Right(txtTMBM07, 2)) <> Val(strChkTPB05) Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "¤½³ø´Á¼Æ»P¤½§i¤é´Á¤£²Å¡I"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text03.SetFocus
      Exit Sub
   End If
   
   If Right(Trim(txtPath1), 1) = "\" Then txtPath1 = Left(txtPath1, Len(txtPath1) - 1)
   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   Set fs = CreateObject("Scripting.FileSystemObject")
   
   'Add By Sindy 2020/5/11 ¥ý²M°£¸ÑÀ£ÁY«á,ÂÂªº¸ê®Æ§¨,¥H¨¾ªÅ¶¡¤£¨¬
   If Dir(txtPath1 & "\isu*") <> "" Then
      fs.DeleteFolder txtPath1 & "\isu*", True
      Sleep 1000
   End If
   '2020/5/11 END
   
   'Added by Morgan 2020/5/5
   '109/5/11¶}©l¨ú®ø¥úºÐ¡A§ï¤U¸üÀ£ÁYÀÉ
   'ÀË¬d¸ê®Æ§¨¬O§_¦s¦b
   strExc(0) = txtPath1 & "\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000")
   If fs.FolderExists(strExc(0) & "\patent") = False Then
      'ÀË¬dÀ£ÁYÀÉ¬O§_¦s¦b Ex:Isu047013_Publish.zip
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
   'File2.path = txtPath1 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000")
   File2.path = txtPath1 & "\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\patent"
   '2013/1/2 End
   File2.Refresh
   If File2.ListCount = 0 Then
      'Modified by Morgan 2020/5/5
      'MsgBox "¥úºÐ¨Ó·½¸ô®|¤¤µL" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¤½³ø¸ê®Æ¡I"
      MsgBox "¨Ó·½¸ô®|¤¤µL" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¤½³ø¸ê®Æ¡I"
      'end 2020/5/5
      txtPath1.SetFocus
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   'Set fs = CreateObject("Scripting.FileSystemObject") 'Removed by Morgan 2020/5/5 §ï¨ì¤W­±
   DeleteFilePathErr = True
   
   'Modify By Sindy 2012/6/6
   If fs.FolderExists(txtPath2) = True Then
      fs.DeleteFile txtPath2 & "\*.*", True '§R°£XMLÀÉ¤Î°O¿ýª©¥»¤å¦rÀÉ(ver*.txt)
      'ÀË¬d¬O§_¦³±ý«þ¨©·í´ÁªºPDF¸ê®Æ§¨
      If fs.FolderExists(txtPath2 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000")) = True Then
         fs.DeleteFolder txtPath2 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000"), True
      End If
      '©T©w§R°£¤W­Ó¤ë¸Ó´ÁPDF¸ê®Æ§¨
      strDate = DBDATE(ChangeWStringToTString(DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(DBDATE(text03))))))
      Call GetNoticeNumber(strDate)
      If fs.FolderExists(txtPath2 & "\img_1\isu" & Format(strChkTPB04, "000") & Format(strChkTPB05, "000")) = True Then
         fs.DeleteFolder txtPath2 & "\img_1\isu" & Format(strChkTPB04, "000") & Format(strChkTPB05, "000"), True
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
   fs.CreateFolder txtPath2 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000")
   'Modify By Sindy 2013/1/2
   'fs.CopyFile txtPath1 & "\xml\*.*", txtPath2 & "\"
   'fs.CopyFile txtPath1 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\*.*", txtPath2 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\"
   fs.CopyFile txtPath1 & "\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\patent\*.*", txtPath2 & "\"
   fs.CopyFile txtPath1 & "\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\patent\pdf\*.*", txtPath2 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\"
   '2013/1/2 End
   'Add By Sindy 2012/6/6
   '²£¥Í°O¿ýXMLª©¥»¤å¦rÀÉ(ver*.txt)
   Dim a As Object
   Set a = fs.CreateTextFile(txtPath2 & "\ver" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000.txt"), True)
   'DoEvents
   '2012/6/6 End
   
   'Add By Sindy 2017/11/24
   If Pub_StrUserSt03 = "M51" Then
      strPath = PUB_Getdesktop & "\" & text03
   Else
      'Modified by Lydia 2024/07/22 §ï¦¨ÅÜ¼Æ
      'strPath = "\\Pat1\¹q¤l±M§Q¤½³ø\" & text03
      strPath = "\\" & strPat1Path & "\¹q¤l±M§Q¤½³ø\" & text03
   End If
   '¼È¦s¸ê®Æ§¨,¬°¤F¦X¨ÖÀÉ®×¨Ï¥Î
   strPathTemp = txtPath2 & "\img_1\temp"
   If fs.FolderExists(strPathTemp) = False Then
      fs.CreateFolder strPathTemp
   Else
      fs.DeleteFile strPathTemp & "\*.*", True '§R°£ÀÉ®×­«CopyFile
   End If
   'ÀË¬d¦s©ñ³]­p®×PDFÀÉ¸ê®Æ§¨¬O§_¤w¦s¦b
   If fs.FolderExists(strPath) = False Then
      fs.CreateFolder strPath
      'DoEvents
   Else
      fs.DeleteFile strPath & "\*.*", True '§R°£ÀÉ®×­«CopyFile
   End If
   If fs.FolderExists(strPath & "\3¥Ó­Ó®×") = False Then
      fs.CreateFolder strPath & "\3¥Ó­Ó®×"
      'DoEvents
   Else
      fs.DeleteFile strPath & "\3¥Ó­Ó®×\*.*", True '§R°£ÀÉ®×­«CopyFile
   End If
   'Copy³]­p®×PDFÀÉ
   ChDir App.path 'Add By Sindy 2020/4/6 ÄÀ©ñ¸ê®Æ§¨Åv­­
   Set oFolder = fs.GetFolder(txtPath1 & "\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\patent\pdf")
   If oFolder.files.Count > 0 Then
      For Each oFile In oFolder.files
         'D:\Isu044033\patent\pdf\106300034.pdf
         If UCase(Right(Trim(oFile.Name), 4)) = UCase(".pdf") And Mid(Trim(oFile.Name), 4, 1) = "3" Then
            fs.CopyFile txtPath1 & "\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\patent\pdf\" & oFile.Name, strPath & "\3¥Ó­Ó®×\" & oFile.Name
            fs.CopyFile txtPath1 & "\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\patent\pdf\" & oFile.Name, strPathTemp & "\" & oFile.Name
         End If
      Next
      Sleep 1000 'Add By Sindy 2020/4/6
      '¦X¨ÖÀÉ®×
      If MergePDF(strPathTemp, strPathTemp & "\*.*", "merge.pdf") = True Then
         fs.CopyFile strPathTemp & "\merge.pdf", strPath & "\3¥Ó¦X¨Ö.pdf"
      End If
   End If
   'Copyª§Ä³®×PDFÀÉ
   Set oFolder = fs.GetFolder(txtPath1 & "\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\sundrydata\pdf")
   If oFolder.files.Count > 0 Then
      For Each oFile In oFolder.files
         'D:\Isu044033\sundrydata\pdf\sud07_1.pdf, sud07_2.pdf
         If UCase(Trim(oFile.Name)) = UCase("sud07_1.pdf") Or UCase(Trim(oFile.Name)) = UCase("sud07_2.pdf") Then
            'fs.CopyFile txtPath1 & "\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\sundrydata\pdf\" & oFile.Name, strPath & "\" & oFile.Name
            fs.CopyFile txtPath1 & "\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "\sundrydata\pdf\" & oFile.Name, strPathTemp & "\" & oFile.Name
         End If
      Next
      '¦X¨ÖÀÉ®×
      'Modify By Sindy 2020/4/6
      If Dir(strPathTemp & "\sud07_1.pdf") <> "" And Dir(strPathTemp & "\sud07_2.pdf") <> "" Then
      '2020/4/6 END
         If MergePDF(strPathTemp, strPathTemp & "\sud07_1.pdf " & strPathTemp & "\sud07_2.pdf", "merge2.pdf") = True Then
            fs.CopyFile strPathTemp & "\merge2.pdf", strPath & "\ª§Ä³®×.pdf"
         End If
      End If
   End If
   '2017/11/24 END
   ChDir App.path 'Add By Sindy 2020/4/6 ÄÀ©ñ¸ê®Æ§¨Åv­­
   
   Screen.MousePointer = vbDefault
   MsgBox "«þ¨©§¹²¦¡I(«þ¨©ªá¶O®É¶¡¡G" & strTime & "  " & time() & ")"
   Set fs = Nothing
   Exit Sub
   
ErrHnd:
   If Err.NUMBER = 76 And DeleteFilePathErr = True Then
      GoTo NotFolder76
   ElseIf Err.NUMBER = 68 Or Err.NUMBER = 76 Then
      'MsgBox "¥úºÐ¨Ó·½¸ô®|¤¤µL" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¤½³ø¸ê®Æ¡I"
      MsgBox "ÀÉ®×¨Ó·½¸ô®|¤¤µL" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á¤½³ø¸ê®Æ¡I"
      txtPath1.SetFocus
   Else
      MsgBox Err.Description
   End If
   Screen.MousePointer = vbDefault
End Sub

'Add By Sindy 2017/11/24 ¦X¨ÖÀÉ®×
Private Function MergePDF(strFilePath As String, strFiles As String, strMergeName As String) As Boolean
Dim strCmd As String
Dim process_id As Long
Dim process_handle As Long
   
   MergePDF = False
   'pdftk.exe C:\97038\zPDF\*.pdf cat output C:\97038\zPDF\merge.pdf
   strCmd = pub_PdftkEXE & " " & strFiles & " cat output " & strFilePath & "\" & strMergeName
   process_id = Shell(strCmd, vbHide)
   process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
   If process_handle <> 0 Then
      For intI = 1 To 10
         If PUB_CheckIsRunning(pub_PdftkName) = True Then
            Sleep 1000
         Else
            Exit For
         End If
      Next
      If intI > 10 And Dir(strFilePath & "\" & strMergeName) = "" Then
         TerminateProcess process_handle, 0&
         CloseHandle process_handle
         MsgBox "¦X¨ÖPDF¥¢±Ñ¡I"
         Exit Function
      Else
         CloseHandle process_handle
      End If
   Else
      MsgBox "¦X¨ÖPDF¥¢±Ñ¡I"
      Exit Function
   End If
   MergePDF = True
End Function

Private Sub cmdExit_Click()
   Unload Me
End Sub

'Add By Sindy 2012/3/3
Private Sub cmdPA160_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim dblFCnt As Double
Dim dblMaxWidth As Double
Dim strTime As String, strTotRow As String
Dim fs As Object
   
On Error GoTo ErrHand
   
   strTime = time()
   
   '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
   If TxtValidate = False Then Exit Sub
   
   If IsRecordExist = False Then
      MsgBox "¤½³ø¨÷´Á" & txtTMBM07 & "¸ê®Æ¤£¦s¦b¡I"
      txtTMBM07.SetFocus
      Exit Sub
   End If
   
   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   
   'ÀË¬d¤½³ø¨÷´Á
   Set fs = CreateObject("Scripting.FileSystemObject")
   File2.path = txtPath2.Text
   File2.Refresh
   If File2.ListCount = 0 Or _
      fs.FileExists(txtPath2 & "\ver" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000.txt")) = False Then
      'MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡A½Ð¥ý«þ¨©¥úºÐ¸ê®Æ¡I"
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡A½Ð¥ý«þ¨©ÀÉ®×¸ê®Æ¡I"
      txtPath2.SetFocus
      Exit Sub
   End If
   Set fs = Nothing
   
   Screen.MousePointer = vbHourglass
   cnnConnection.BeginTrans
   
   Call ResetGrid: intPRow = 0 'Add By Sindy 2012/1/16
   strOurAgentName = GetTOurAgentName()
   m_PrintRpt1 = False: m_PrintRpt2 = False: iLine2 = 0
   strTotRow = File2.ListCount
   Me.Height = MaxHeight
   dblMaxWidth = 5940
   Text2.Width = 0
   m_PI02 = "" 'Add By Sindy 2012/8/16
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
         
         If ReadXmlData = False Then GoTo ErrHand
         
         '¥»©Ò¥Ó½Ð®×¤~§ó·s
         If bolTaieCase = True Then
            strSql = "UPDATE Patent SET PA160='" & strPA160 & "' " & _
                  " WHERE PA11 = '" & strTPB01 & "'"
            cnnConnection.Execute strSql
         End If
         'Add By Sindy 2012/8/9 °ê¤º±M§Q¤½³øÀÉ¼W¥[°ê»Ú¤ÀÃþ¸¹,IPC¤ÀÃþ
         'Modify By Sindy 2016/3/2 +,TPB13='" & strTPB13 & "'
         strSql = "UPDATE TPBulletin SET TPB10='" & strTPB10 & "',TPB11='" & strTPB11 & "',TPB12='" & strTPB12 & "',TPB13='" & strTPB13 & "'" & _
                  " WHERE TPB01='" & strTPB01 & "'"
         cnnConnection.Execute strSql
         '2012/8/9 End
      End If
   Next dblFCnt
   
   cnnConnection.CommitTrans
   
   Screen.MousePointer = vbDefault
   'Modify By Sindy 2024/6/3 ·¨¶²ªÚ¸g²z«ü¥Ü,Á`¸g²z¤w®Ö¥Ü°±¤î¦¹¶µ¤ÀÃþ¤u§@¡A¦¹Ãþ³qª¾¤]¥i°±¤îµo°e
'   Call GetSendMailIPC 'Add By Sindy 2012/8/16
   Call IsRecordExist '²£¥Íµ§¼Æ
   MsgBox "ÂàÀÉ§¹²¦¡I(ÂàÀÉªá¶O®É¶¡¡G" & strTime & "  " & time() & ")" & vbCrLf & strMsg
   Me.Height = MinHeight
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   If Err.NUMBER = 76 Then
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡I"
      txtPath2.SetFocus
   Else
      cnnConnection.RollbackTrans
      If Err.NUMBER = -2147217873 Then
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¥Ó½Ð®×¸¹¡]" & strTPB01 & "¡^" & vbCrLf & strErrTxt & ": ¹H¤Ï¥²¶·¬°°ß¤@ªº­­¨î±ø¥ó" & vbCrLf & strSql
      Else
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¥Ó½Ð®×¸¹¡]" & strTPB01 & "¡^" & vbCrLf & strErrTxt & Err.Description & vbCrLf & strSql
      End If
   End If
End Sub

''Add By Sindy 2012/8/16 IPC¤ÀÃþÂkÃþ¤£¨ì®É,³qª¾69009·¨·¶¯Â
''Modify By Sindy 2020/5/13 ·¨·¶¯Â(ºÊ¹î¤H):¤w»P·¨¸g²z°Q½×¹L,¤é«á­Y¤½³øIPC¤ÀÃþ¦³°ÝÃD®É,½Ð¥Ñ¨t²Îª½±µÂàµ¹99033·¨¶²ªÚ¸g²z
'Private Sub GetSendMailIPC()
'   If m_PI02 <> "" Then
'      'm_PI02 = Mid(m_PI02, 2, Len(m_PI02))
'      m_PI02 = Replace(m_PI02, "¡F", vbCrLf)
'      PUB_SendMail strUserNum, "99033;97038", "", "±M§Q¤½³ø" & txtTMBM07 & "´Á¦³°ê»Ú¤ÀÃþ¸¹¡A©|¥¼°µIPC¤ÀÃþ", "Dear Sirs," & vbCrLf & vbCrLf & _
'      "±M§Q¤½³ø" & txtTMBM07 & "´Á¦³°ê»Ú¤ÀÃþ¸¹¡A©|¥¼°µIPC¤ÀÃþ¡A¦p¤U¡G" & vbCrLf & vbCrLf & m_PI02 & vbCrLf & vbCrLf & _
'      "·Ð½Ð¦A³qª¾¹q¸£¤¤¤ßÀ³¦p¦ó¤ÀÃþ¡C" & vbCrLf & vbCrLf & vbCrLf & _
'      "                                                        ¹q¸£¤¤¤ß"
'   End If
'End Sub

'Added by Morgan 2020/5/5
Private Sub cmdPath_Click()
   Dim fName As String, strStartFolder As String
   
   If Dir(txtPath1 & "\", vbDirectory) <> "" Then strStartFolder = txtPath1
   
   fName = PUB_GetFolder(Me.hWnd, strStartFolder, "½Ð¿ï¨ú¸ê®Æ§¨:")
   If fName <> "" Then 'they did not hit cancel
      txtPath1 = fName
   End If
   
End Sub

'Add By Sindy 2013/4/15
Private Sub cmdTemp_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim dblFCnt As Double
Dim dblMaxWidth As Double
Dim strTime As String, strTotRow As String
Dim fs As Object
   
On Error GoTo ErrHand
   
   strTime = time()
   
   '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
   If TxtValidate = False Then Exit Sub
   
   If IsRecordExist_Temp = True Then
      strTit = "¸ß°Ý"
      strMsg = "¤½³ø¨÷´Á" & txtTMBM07 & "¤w¦³¸ê®Æ¦s¦b¡A½T©w¬O§_­n­«·sÂàÀÉ¡H"
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
      'MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡A½Ð¥ý«þ¨©¥úºÐ¸ê®Æ¡I"
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡A½Ð¥ý«þ¨©ÀÉ®×¸ê®Æ¡I"
      txtPath2.SetFocus
      Exit Sub
   End If
   Set fs = Nothing
   
   Screen.MousePointer = vbHourglass
   cnnConnection.BeginTrans
   
   strSql = "delete FROM TPBulletin_sonia WHERE TPB04=" & CNULL(Left(txtTMBM07, 2)) & " and TPB05=" & CNULL(Right(txtTMBM07, 2))
   cnnConnection.Execute strSql
   
   Call ResetGrid: intPRow = 0
   strOurAgentName = GetTOurAgentName()
   m_PrintRpt1 = False: m_PrintRpt2 = False: iLine2 = 0
   strTotRow = File2.ListCount
   Me.Height = MaxHeight
   dblMaxWidth = 5940
   Text2.Width = 0
   m_PI02 = "" 'Add By Sindy 2012/8/16
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
         
         If ReadXmlData = False Then GoTo ErrHand
         'If ChkDataErr() = True Then GoTo ErrHand
         
'         '¦a°Ï¦WºÙ¬°ªÅ¥Õ©Î020.¤¤°ê¤j³°,¥N²z¤H¦WºÙ¦³?®É,»Ý¦C¦L²M³æ (Or strTPB06 = "020")
'         If strTPB06 = "" Or _
'            InStr(strTPB07, "?") > 0 Then
'            Call ReadTxt1(strTPB01, strTPB02, strTPB06, strTPB07, strAChinese1, strAddress1)
'            Call PrintPaper(strTPB01, strTPB02, strTPB06, strTPB07, strAddress1)
'         End If
         
         '·s¼WTable
         strErrTxt = "°ê¤º±M§Q¤½³øÀÉ.TPBulletin_sonia"
         strSql = "insert into TPBulletin_sonia (TPB01,TPB02,TPB03,TPB04,TPB05,TPB06,TPB07,TPB08,TPB09,TPB10,TPB11," & _
                  "TPB12,TPB13,TPB14,TPB15,TPB16,TPB17,TPB18,TPB19,TPB20,TPB21,TPB22) " & _
                  "values(" & CNULL(strTPB01) & "," & CNULL(strTPB02) & "," & dblTPB03 & "," & CNULL(strTPB04) & "," & CNULL(strTPB05) & "," & CNULL(strTPB06) & "," & CNULL(strTPB07_1) & "," & CNULL(strTPB08) & "," & CNULL(strTPB09) & "," & CNULL(strTPB10) & "," & CNULL(strTPB11) & _
                  "," & CNULL(strTPBcApp(1)) & "," & CNULL(strTPBcApp(2)) & "," & CNULL(strTPBcApp(3)) & "," & CNULL(strTPBcApp(4)) & "," & CNULL(strTPBcApp(5)) & _
                  "," & CNULL(strTPBcApp(6)) & "," & CNULL(strTPBcApp(7)) & "," & CNULL(strTPBcApp(8)) & "," & CNULL(strTPBcApp(9)) & "," & CNULL(strTPBcApp(10)) & "," & CNULL(strTPB12) & ")"
         cnnConnection.Execute strSql
      End If
   Next dblFCnt
   
   cnnConnection.CommitTrans
   
'   strMsg = ""
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
   
   Screen.MousePointer = vbDefault
   
   Call IsRecordExist_Temp '²£¥Íµ§¼Æ
   MsgBox "ÂàÀÉ§¹²¦¡I(ÂàÀÉªá¶O®É¶¡¡G" & strTime & "  " & time() & ")" & vbCrLf & strMsg
   Me.Height = MinHeight
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
'   Set rsTmp = Nothing
   If Err.NUMBER = 76 Then
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡I"
      txtPath2.SetFocus
   Else
      cnnConnection.RollbackTrans
      If Err.NUMBER = -2147217873 Then
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¥Ó½Ð®×¸¹¡]" & strTPB01 & "¡^" & vbCrLf & strErrTxt & ": ¹H¤Ï¥²¶·¬°°ß¤@ªº­­¨î±ø¥ó" & vbCrLf & strSql
      Else
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¥Ó½Ð®×¸¹¡]" & strTPB01 & "¡^" & vbCrLf & strErrTxt & Err.Description & vbCrLf & strSql
      End If
   End If
End Sub

'Add By Sindy 2016/3/2
'¸ÉÂà®×¥óÄÝ©Ê
Private Sub cmdTPB12_Click()
Dim strTime As String
Dim stSQL As String, intR As Integer
Dim rsQuery As ADODB.Recordset
   
On Error GoTo ErrHand
   
   strTime = time()
   
   stSQL = "SELECT TPB01,TPB02,TPB10,TPB11,TPB13 FROM TPBulletin WHERE TPB11 is not null and TPB10 is not null and TPB13 is null"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      Screen.MousePointer = vbHourglass
      With rsQuery
         .MoveFirst
         Do While Not .EOF
            'cnnConnection.BeginTrans
            
            strTPB13 = GetPatentIPC("3", .Fields("TPB10"), .Fields("TPB02"))
            
            strSql = "UPDATE TPBulletin SET TPB13='" & strTPB13 & "'" & _
                     " WHERE TPB01='" & .Fields("TPB01") & "'"
            cnnConnection.Execute strSql
            
            'cnnConnection.CommitTrans
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
      'cnnConnection.RollbackTrans
      MsgBox Err.NUMBER & " " & Err.Description
   End If
End Sub

''Add By Sindy 2013/8/23
''¸ÉÂà²£·~§O¤ÀÃþ
'Private Sub cmdTPB12_Click()
'Dim strTit As String
'Dim strMsg As String
'Dim nResponse
'Dim dblFCnt As Double
'Dim dblMaxWidth As Double
'Dim strTime As String, strTotRow As String
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
'      MsgBox "¤½³ø¨÷´Á" & txtTMBM07 & "¸ê®Æ¤£¦s¦b¡I"
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
'      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡A½Ð¥ý«þ¨©¥úºÐ¸ê®Æ¡I"
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
'   dblMaxWidth = 5940
'   Text2.Width = 0
'   m_PI02 = ""
'   For dblFCnt = 0 To File2.ListCount - 1
'      'ÀÉ¦W«e3½X¬°sudªÌ¤£¶·Âà¤J¸ê®Æ
'      If (Asc(Left(Trim(File2.List(dblFCnt)), 1)) >= 48 And Asc(Left(Trim(File2.List(dblFCnt)), 1)) <= 57) And _
'         UCase(Right(Trim(File2.List(dblFCnt)), 3)) = "XML" Then
'         RichTextBox1.LoadFile (txtPath2.Text & "\" & File2.List(dblFCnt))
''         RichTextBox1.LoadFile (txtPath2.Text & "\097307080.xml")
'
'         Text2.Width = dblMaxWidth / Val(strTotRow) * (dblFCnt + 1): DoEvents
'
'         If ReadXmlData = False Then GoTo ErrHand
'
'         '°ê¤º±M§Q¤½³øÀÉ¼W¥[²£·~§O¤ÀÃþ
'         strSql = "UPDATE TPBulletin SET TPB12='" & strTPB12 & "'" & _
'                  " WHERE TPB01='" & strTPB01 & "'"
'         cnnConnection.Execute strSql
'      End If
'   Next dblFCnt
'
'   cnnConnection.CommitTrans
'
'   Screen.MousePointer = vbDefault
'
'   Call GetSendMailIPC
'   Call IsRecordExist '²£¥Íµ§¼Æ
'   MsgBox "ÂàÀÉ§¹²¦¡I(ÂàÀÉªá¶O®É¶¡¡G" & strTime & "  " & time() & ")" & vbCrLf & strMsg
'   Me.Height = MinHeight
'
'   Exit Sub
'
'ErrHand:
'   Screen.MousePointer = vbDefault
'   If Err.NUMBER = 76 Then
'      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡I"
'      txtPath2.SetFocus
'   Else
'      cnnConnection.RollbackTrans
'      If Err.NUMBER = -2147217873 Then
'         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¥Ó½Ð®×¸¹¡]" & strTPB01 & "¡^" & vbCrLf & strErrTxt & ": ¹H¤Ï¥²¶·¬°°ß¤@ªº­­¨î±ø¥ó" & vbCrLf & strSql
'      Else
'         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¥Ó½Ð®×¸¹¡]" & strTPB01 & "¡^" & vbCrLf & strErrTxt & Err.Description & vbCrLf & strSql
'      End If
'   End If
'End Sub

Private Sub cmdTransFile_Click()
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim dblFCnt As Double
'Dim dblStar As Double, dblEnd As Double
'Dim dblChar As Double, dblLastEnd As Double
'Dim strText As String, strTitNM As String
'Dim strChar As String, strData As String
'Dim rsTmp As New ADODB.Recordset
'Dim strFreeAgentCode As String
Dim dblMaxWidth As Double
Dim strTime As String, strTotRow As String
Dim fs As Object
Dim stCP12 As String, stCP13 As String, stCP09 As String, strFileName As String, strCP10 As String
Dim f
Dim bolTa04IsNull As Boolean 'Add By Sindy 2014/9/3
Dim intQ As Integer, rsQuery As New ADODB.Recordset   'Added by Lydia 2021/08/16
Dim strExSql As String 'Added by Lydia 2022/01/21

On Error GoTo ErrHand
   
   strTime = time()
   
   '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
   If TxtValidate = False Then Exit Sub
   
   If IsRecordExist = True Then
      strTit = "¸ß°Ý"
      strMsg = "¤½³ø¨÷´Á" & txtTMBM07 & "¤w¦³¸ê®Æ¦s¦b¡A½T©w¬O§_­n­«·sÂàÀÉ¡H"
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
      'MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡A½Ð¥ý«þ¨©¥úºÐ¸ê®Æ¡I"
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡A½Ð¥ý«þ¨©ÀÉ®×¸ê®Æ¡I"
      txtPath2.SetFocus
      Exit Sub
   End If
   'Set fs = Nothing
   
   Screen.MousePointer = vbHourglass
   'cnnConnection.BeginTrans
   
   strSql = "delete FROM TPBulletin WHERE TPB04=" & CNULL(Left(txtTMBM07, 2)) & " and TPB05=" & CNULL(Right(txtTMBM07, 2))
   cnnConnection.Execute strSql
   
   Call ResetGrid: intPRow = 0 'Add By Sindy 2012/1/16
   strOurAgentName = GetTOurAgentName()
   m_PrintRpt1 = False: m_PrintRpt2 = False: iLine2 = 0
   strTotRow = File2.ListCount
   Me.Height = MaxHeight
   dblMaxWidth = 5940
   Text2.Width = 0
   m_PI02 = "" 'Add By Sindy 2012/8/16
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
         
         cnnConnection.BeginTrans 'Modify By Sindy 2015/8/11
         
         If ReadXmlData = False Then GoTo ErrHand
         If ChkDataErr() = True Then GoTo ErrHand
         
         '¦a°Ï¦WºÙ¬°ªÅ¥Õ©Î020.¤¤°ê¤j³°,¥N²z¤H¦WºÙ¦³?®É,»Ý¦C¦L²M³æ (Or strTPB06 = "020")
         'Modify By Sindy 2015/9/23 +strTPB06 = "000"
         'Modify By Sindy 2019/9/4 + Or strTPB38 = "" Or strTPB38 = "¤¤µØ¥Á°ê" Or strTPB38 = "¥xÆW"
         txtChkWord = strTPB07 'Add By Sindy 2024/5/17
         If strTPB06 = "" Or strTPB06 = "000" Or _
            InStr(txtChkWord, "?") > 0 Or strTPB38 = "" Or strTPB38 = "¤¤µØ¥Á°ê" Or strTPB38 = "¥xÆW" Then
            Call ReadTxt1(strTPB01, strTPB02, strTPB06, strTPB07, strAChinese1, strAddress1)
            Call PrintPaper(strTPB01, strTPB02, strTPB06, strTPB07, strAddress1)
         End If
         
         'Add By Sindy 2017/2/21
         'ÀË¬d¥Ó½Ð¤H¦WºÙ¬O§_¦³?³y¦r
         For i = 1 To 10
            txtChkWord = strTPBcApp(i) 'Add By Sindy 2024/5/17
            If InStr(txtChkWord, "?") > 0 Then
               strMsg = "¥Ó½Ð®×¸¹" & strTPB01 & "¥Ó½Ð¤H¦WºÙ" & i & "¡u" & strTPBcApp(i) & "¡v¦³?¸¹"
               Call ReadTxt1(strTPB01, strTPB02, strMsg, "", "", "")
               Call PrintPaper(strTPB01, strTPB02, strMsg, "", "")
            End If
         Next i
         '2017/2/21 END
         
         '·s¼WTable
         strErrTxt = "°ê¤º±M§Q¤½³øÀÉ.TPBulletin"
         'Modify By Sindy 2012/8/9 +,TPB10,TPB11
         'Modify By Sindy 2016/3/2 +,TPB13
         'Modify By Sindy 2017/2/20 +,TPB14,TPB15,TPB16,TPB17,TPB18,TPB19,TPB20,TPB21,TPB22,TPB23
         'Modify By Sindy 2019/9/4 +,TPB38
         strSql = "insert into TPBulletin (TPB01,TPB02,TPB03,TPB04,TPB05,TPB06,TPB07,TPB08,TPB09,TPB10,TPB11,TPB12,TPB13" & _
                  ",TPB14,TPB15,TPB16,TPB17,TPB18,TPB19,TPB20,TPB21,TPB22,TPB23" & _
                  ",TPB24,TPB25,TPB26,TPB27,TPB28,TPB29,TPB30,TPB31,TPB32,TPB33" & _
                  ",TPB34,TPB35,TPB36,TPB37,TPB38" & _
                  ") values(" & CNULL(strTPB01) & "," & CNULL(strTPB02) & "," & dblTPB03 & "," & CNULL(strTPB04) & "," & CNULL(strTPB05) & "," & CNULL(strTPB06) & "," & CNULL(strTPB07_1) & "," & CNULL(strTPB08) & "," & CNULL(strTPB09) & "," & CNULL(strTPB10) & "," & CNULL(strTPB11) & "," & CNULL(strTPB12) & "," & CNULL(strTPB13) & _
                  "," & CNULL(strTPBcApp(1)) & "," & CNULL(strTPBcApp(2)) & "," & CNULL(strTPBcApp(3)) & "," & CNULL(strTPBcApp(4)) & "," & CNULL(strTPBcApp(5)) & _
                  "," & CNULL(strTPBcApp(6)) & "," & CNULL(strTPBcApp(7)) & "," & CNULL(strTPBcApp(8)) & "," & CNULL(strTPBcApp(9)) & "," & CNULL(strTPBcApp(10)) & _
                  "," & CNULL(strTPBeApp(1)) & "," & CNULL(strTPBeApp(2)) & "," & CNULL(strTPBeApp(3)) & "," & CNULL(strTPBeApp(4)) & "," & CNULL(strTPBeApp(5)) & _
                  "," & CNULL(strTPBeApp(6)) & "," & CNULL(strTPBeApp(7)) & "," & CNULL(strTPBeApp(8)) & "," & CNULL(strTPBeApp(9)) & "," & CNULL(strTPBeApp(10)) & _
                  "," & dblTPB34 & "," & dblTPB35 & "," & CNULL(strTPB36) & "," & CNULL(strTPB37) & "," & CNULL(strTPB38) & _
                  ")"
         cnnConnection.Execute strSql
         
         '¥»©Ò¥Ó½Ð®×¤~§ó·s
         If bolTaieCase = True Then
            'Add By Sindy 2014/6/17 ·s¼W¶i«×
            'If pa(1) = "P" Then 'Modify By Sindy 2015/8/18 FCP¤]­n·s¼W¸Óµ§¶i«×
               strCP10 = "1228" '1228.¤½§i¤½³ø
               strSql = "SELECT cp09 FROM caseprogress " & _
                        "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' " & _
                         " AND CP10 = '" & strCP10 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then
                  stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
                  stCP12 = GetSalesArea(stCP13)
                  stCP09 = AutoNo("C", 6)
                  strSql = "INSERT INTO CASEPROGRESS(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32)" & _
                          " VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & strSrvDate(1) & "','" & stCP09 & "'" & _
                          ",'" & strCP10 & "','" & stCP12 & "','" & stCP13 & "','" & strUserNum & "','N','N','" & strSrvDate(1) & "','N')"
                  cnnConnection.Execute strSql
                  '±Npdf file¦s¤JDB
                  strFileName = txtPath2.Text & "\img_1\isu0" & Left(txtTMBM07, 2) & "0" & Right(txtTMBM07, 2) & "\" & strTPB01 & ".pdf"
                  'Set fs = CreateObject("Scripting.FileSystemObject")
                  Set f = fs.GetFile(strFileName)
                  '¦sÀÉ
                  'Modify By Sindy 2022/5/6 CStr(Val(pa(2))) ==> pa(2)
                  If SaveAttFile_PDF(stCP09, strFileName, UCase(pa(1) & pa(2) & IIf(pa(3) <> "0" Or pa(4) <> "00", "-" & pa(3), "") & IIf(pa(4) <> "00", "-" & pa(4), "") & "." & strCP10 & ".pdf"), Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), True) = False Then
                     GoTo ErrHand
                  End If
               'Added by Lydia 2022/07/05
               Else
                  stCP09 = "" & RsTemp.Fields("cp09")
               'end 2022/07/05
               End If
            'End If
            '2014/6/17 END
            
            '93.8.1 ¥H«á¤½§i¸¹§ï¬°ÃÒ®Ñ¸¹
            '§ó·s±M§Q°ò¥»ÀÉªº¤½§i¤é,±M§Q¸¹¼Æ,¤½§i¸¹
            'Add By Sindy 2012/3/3 +°ê»Ú¤ÀÃþ
            strSql = "UPDATE Patent SET PA14=" & dblTPB03 & _
                  ",PA22='" & strTPB02 & "',PA15='" & strTPB02 & "',PA160='" & strPA160 & "' " & _
                  " WHERE PA11 = '" & strTPB01 & "'"
            cnnConnection.Execute strSql
            
            '§ó·s¤U¤@µ{§Ç¦~¶O´Á­­
            strExc(0) = Right(pa(72), 2)
            If Left(strExc(0), 1) = "," Then strExc(0) = Right(strExc(0), 1)
            m_strNextDueDate = CompDate(0, Val(strExc(0)), dblTPB03)
            m_strNextDueDate = CompDate(2, -1, m_strNextDueDate)
            m_strAgreeOnDate = "" 'Add By Sindy 2021/8/17
            'Added by Morgan 2014/10/28
            'Modified by Morgan 2014/11/20 ¥~±M§ï¦^ÂÂ³W«h
            If strSrvDate(1) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é And pa(1) <> "FCP" Then
               m_strNextFeeDate = PUB_GetOurDeadline(m_strNextDueDate)
            'Added by Morgan 2019/7/11 ¥~±M¥xÆW®×©Ò­­¥H§ï¤u§@¤Ñ­pºâ
            ElseIf strSrvDate(1) >= ¥~±M¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é And pa(1) = "FCP" Then
               'Modify By Sindy 2021/8/17 + , , m_strAgreeOnDate
               m_strNextFeeDate = PUB_GetFCPOurDeadline(m_strNextDueDate, 2, , m_strAgreeOnDate)
            'end 2019/7/11
            Else
            'end 2014/10/28
               m_strNextFeeDate = CompDate(2, -2, m_strNextDueDate)
            End If 'Added by Morgan 2014/10/28
            
            If pa(1) = "P" Then 'P®×¤~­n§ì¤u§@¤Ñ
               m_strNextFeeDate = PUB_GetWorkDay1(m_strNextFeeDate, True)
            End If
            'Modify By Sindy 2021/8/17 + ",NP23=" & CNULL(m_strAgreeOnDate)
            strSql = "update nextprogress set NP08=" & m_strNextFeeDate & ", np09=" & m_strNextDueDate & _
                     ",NP23=" & CNULL(m_strAgreeOnDate) & _
                     " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
                     " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
            cnnConnection.Execute strSql, intI
            
            '¤º±M­Y¦³¥¼µo¤å§Þ³N³ø§i®É§ó·s¤å¥ó»ô³Æ¤é(=¤½§i¤é)¤Î©Ó¿ì´Á­­
            If pa(1) = "P" Then
               If PUB_ChkCPExist(pa, "421", 1, m_str421CP09, m_str421CP14) = True Then
                  m_str421EP06 = dblTPB03
                  '§ó·s¤å¥ó»ô³Æ¤é
                  strSql = "Update EngineerProgress Set EP06=" & strSrvDate(1) & " Where EP02='" & m_str421CP09 & "' AND EP06 IS NULL"
                  cnnConnection.Execute strSql
                  If PUB_IfSetCP48(m_str421CP09) Then
                     '©Ó¿ì´Á­­§ï©I¥s¦@¥Î¨ç¼Æ­pºâ
                     m_str421CP48 = Pub_GetHandleDay(pa(1), "000", "421", m_str421EP06, , m_str421CP09)
                     If Val(m_str421CP48) > 0 Then
                        '§ó·s©Ó¿ì´Á­­
                        strSql = "Update CaseProgress Set CP48=" & m_str421CP48 & " Where CP09='" & m_str421CP09 & "' AND CP48 IS NULL"
                        cnnConnection.Execute strSql
                     End If
                  End If
                  
                  'Added by Morgan 2019/12/11 «DFMP®×§ó·s»ô³Æ¤é©Ó¿ì´Á­­¦b Trigger ³]©w
                  If Val(m_str421CP48) = 0 Then
                     strExc(0) = "select cp48 from caseprogress where cp09='" & m_str421CP09 & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        m_str421CP48 = "" & RsTemp(0)
                     End If
                  End If
                  'end 2019/12/11
               End If
            End If
            'Added by Lydia 2021/08/16 ¥~±M-ÃÄ«~±M§Q³sµ²¡G·í¡u±M§Q³sµ²³qª¾=Y¡v®É¶i«×ÀÉ¦Û°Ê·s¼W¤@±M§Q³sµ²³qª¾¦¬¤å(BÃþ¦¬¤å959)¡A¨Ã¥B¦Û°Ê¤Wµo¤å¤é
            If pa(1) = "FCP" Then
                strExc(0) = "select pa14,pa26,pa27,pa28,pa29,pa30,pa75,cp09,cp14,cp14t as cp14t from patent," & _
                                 "(select cp01,cp02,cp03,cp04,cp09,cp14,st04 as cp14t from caseprogress c1,staff " & _
                                 "where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp05||cp09 = (select max(cp05||cp09) maxno from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'  and cp10='959' and cp159=0 ) " & _
                                 "and cp14=st01(+)) vtb1 where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "' and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                   '1.¶i«×ÀÉ¦³¦¬¤å:¡u959ÃÄ«~±M§Q³sµ²§i¥N¡v¡B2.³]©w¤w¦³«ü¥Ü³qª¾°µ±M§Q³sµ²¤§«È¤á¡GY20412(Novo) ¤ÎY45493 (Lundbeck)¨âªÌ§tÃö«Y¥ø·~
                   If "" & RsTemp.Fields("cp09") <> "" Or InStr("Y20412,Y45493,", Left("" & RsTemp.Fields("pa75"), 6)) > 0 Then
                      '­Y¡u¬O§_®Ö¹ï¤w­ã±M§Q¡v¤§©Ê½è¬°Nªº®×¥ó¡A¨t²Î¦P®É¦Û°Ê¦¬¤å¡u§i¥N901¡v¡F
                      If PUB_CheckAuto926(pa) = False Then
                          strExc(6) = AutoNo("B", 6)
                          strExc(5) = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
                          strExc(4) = CompWorkDay(6, strSrvDate(1)) '©Ó¿ì´Á­­=¨t²Î¤é+5¤u§@¤Ñ
                          'Modified by Lydia 2022/07/05 ©Ó¿ì¤H±¾¤uµ{®v; Ex.FCP-62461
                          'strSql = "INSERT INTO CASEPROGRESS(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP48)" & _
                                  " VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & strSrvDate(1) & "','" & strExc(6) & "'" & _
                                  ",'901','" & GetST15(strExc(5)) & "','" & strExc(5) & "','" & strUserNum & "','N','N','N'," & strExc(4) & ")"
                          If "" & RsTemp.Fields("cp14") <> "" And "" & RsTemp.Fields("cp14t") <> "2" Then
                               strExc(2) = "" & RsTemp.Fields("cp14")
                          Else
                               strExc(2) = PUB_GetFCPPromoterNo(stCP09, "1228", "" & RsTemp.Fields("cp14"))
                          End If
                          strSql = "INSERT INTO CASEPROGRESS(CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP48)" & _
                                  " VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & strSrvDate(1) & "','" & strExc(6) & "'" & _
                                  ",'901','" & GetST15(strExc(5)) & "','" & strExc(5) & "','" & strExc(2) & "','N','N','N'," & strExc(4) & ")"
                          'end 2022/07/05
                          cnnConnection.Execute strSql
                      End If
                      
                      '±H³qª¾Email
                      '¦¬¥ó¤H­û: ©Ó¿ì¤uµ{®v¡Bµ{§ÇºÞ¨î¤H­û
                      strExc(1) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
                      strExc(2) = "" & RsTemp.Fields("cp14") & RsTemp.Fields("cp14t") '959ÃÄ«~±M§Q³sµ²§i¥N¤§©Ó¿ì¤uµ{®v
                      strExc(5) = ""
                      
                      '§PÂ_³Ì«á¤@¹D¦¬¤åªº¤uµ{®v»P959ÃÄ«~±M§Q³sµ²§i¥N¤§©Ó¿ì¤uµ{®v
                      strExc(0) = "select cp14,st04 as cp14t from caseprogress c1,staff " & _
                                 "where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp05||cp09 = (select max(cp05||cp09) maxno from caseprogress,staff " & _
                                 "where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp159=0 and cp14=st01(+) and st03='F21' and cp14 not like 'F%' ) " & _
                                 "and cp14=st01(+) "
                      intQ = 1
                      Set rsQuery = ClsLawReadRstMsg(intQ, strExc(0))
                      If intQ = 1 Then
                          If strExc(2) <> "" & rsQuery.Fields("cp14") & rsQuery.Fields("cp14t") Then
                             strExc(2) = "" & rsQuery.Fields("cp14") & rsQuery.Fields("cp14t")
                          End If
                      'Added by Lydia 2024/08/01 ¨S¦³¤uµ{®v,§ï³qª¾­t³d¤HbyªL§¡­§; Ex.FCP-071145
                      Else
                          strExc(0) = "select oman as cp14,st04 as cp14t from setspecman,staff where ocode='¥~±M¤uµ{®v­t³dÃÄ«~±M§Q³sµ²®×' and instr(oman,st01) > 0 and st04='1' order by st01 "
                          intQ = 1
                          Set rsQuery = ClsLawReadRstMsg(intQ, strExc(0))
                          If intQ = 1 Then
                             If strExc(2) <> "" & rsQuery.Fields("cp14") & rsQuery.Fields("cp14t") Then
                               strExc(2) = "" & rsQuery.Fields("cp14") & rsQuery.Fields("cp14t")
                             End If
                          Else
                             strExc(2) = "R" '¨S¦³¤uµ{®v+¨S¦³­t³d¤H=>³qª¾¥DºÞ
                          End If
                      'end 2024/08/01
                      End If
                      If Right(strExc(2), 1) <> "1" Then '¤H­û¤wÂ÷Â¾¡A§ï³qª¾¥DºÞ
                         'Added by Lydia 2024/08/01 ¨S¦³¤uµ{®v+¨S¦³­t³d¤H=>³qª¾¥DºÞ
                         If strExc(2) = "R" Then
                            strExc(2) = Pub_GetSpecMan("R")
                         Else
                         'end 2024/08/01
                            strExc(2) = Mid(strExc(2), 1, Len(strExc(2)) - 1)
                            strExc(2) = PUB_GetFCPEngSup(strExc(2))
                         End If 'Added by Lydia 2024/08/01
                      Else
                         strExc(2) = Mid(strExc(2), 1, Len(strExc(2)) - 1)
                         strExc(5) = PUB_GetFCPEngSup(strExc(2)) & ";" 'CC¥DºÞ
                      End If
                      
                      '°Æ¥»: ¤uµ{®v¥DºÞ¡Bµ{§Ç¥DºÞ¡B85033(©T©w®Ö¹ï¤½³øµ{§Ç¤H­û=¯S®í³]©w¤§¥~±Mµ{§Ç-³qª¾¦~¶O¹O´Á)
                      strExc(3) = Pub_GetSpecMan("¥~±Mµ{§Ç-³qª¾¦~¶O¹O´Á")
                      strExc(5) = strExc(5) & PUB_GetFCPProSup(strExc(1))
                      If InStr(strExc(5), strExc(3)) = 0 Then strExc(5) = strExc(5) & ";" & strExc(3)
                      
                      '¥D¦®: ¡iÃÄ«~±M§Q³sµ²®×¡jFCP-XXXXXX½ÐÀu¥ý³B²zÃÒ®Ñ¡B¤G¦¸®Ö¹ï¤w­ã¨Ã§iª¾«È¤á±M§Q¸ê°Tµn¿ý´Á­­¬°YY¦~YY¤ëYY¤é(¤½§i¤é«á¤§¦¸¤é°_45¤Ñ¡^
                      'Modified by Lydia 2021/09/29 debug§ï¬°¤é¾ä¤Ñ(©¹«e±À¤u§@¤Ñ); ex.FCP-057257
                      'strExc(9) = CompWorkDay(46, "" & dblTPB03)   '¸ê°Tµn¿ý´Á­­¡G¤½§i¤é«á¤§¦¸¤é°_45¤Ñ
                      'Modified by Lydia 2021/12/03 debug: ¤£¥Î­Ë±À¤u§@¤Ñ(9/29 email¦³´£¨ì)
                      'strExc(9) = CompWorkDay(1, CompDate(2, 45, "" & dblTPB03), 1)
                      strExc(9) = CompDate(2, 45, "" & dblTPB03)
                      strExc(0) = "¡iÃÄ«~±M§Q³sµ²®×¡j" & pa(1) & "-" & pa(2) & IIf(pa(3) = "0", "", "-" & pa(3)) & IIf(pa(4) = "00", "", "-" & pa(4)) & "½ÐÀu¥ý³B²zÃÒ®Ñ¡B¤G¦¸®Ö¹ï¤w­ã¨Ã§iª¾«È¤á±M§Q¸ê°Tµn¿ý´Á­­¬°" & ChangeWStringToTDateString(strExc(9))
                      '¤º¤å¡G°Ï¤À2¬q
                      '1-©Ó¿ì¤uµ{®vªº¤º¤å
                      strExc(10) = "TO¡G©Ó¿ì¤uµ{®v" & vbCrLf & _
                                         "¡@¡@" & pa(1) & "-" & pa(2) & IIf(pa(3) = "0", "", "-" & pa(3)) & IIf(pa(4) = "00", "", "-" & pa(4)) & "¦³¥iµn¿ý±M§Q³sµ²¤§¼Ðªº¨Ã¤w©ó" & ChangeWStringToTDateString("" & dblTPB03) & _
                                         "¤½§i¡Aµ{§Ç±H§¹ÃÒ®Ñ«á¡A½ÐÀu¥ý³B²z¤G¦¸®Ö¹ï¤w­ã¨Ã§iª¾«È¤á±M§Q¸ê°Tµn¿ý´Á­­¬°" & ChangeWStringToTDateString(strExc(9)) & _
                                         "¡A¨Ã½Ðª`·N¬O§_À³¤Ä¿ï¡u±M§Q³sµ²³qª¾¡v©Ê½è¡A­Y¬°¤£¤G¦¸®Ö¹ï¤w­ãªº®×¥ó¡A¹q¸£¨t²Î±N¦Û°Ê¦¬¤å§i¥N¥H¨Ñ¤uµ{®v³ø§i¸ê°Tµn¿ý´Á­­¡C"
                      '2-µ{§Ç¤H­ûªº¤º¤å
                      strExc(10) = strExc(10) & vbCrLf & vbCrLf & _
                                        "TO¡Gµ{§Ç¤H­û" & vbCrLf & _
                                        "¡@¡@½ÐÀu¥ý±HÃÒ®Ñ¡A" & GetStaffName(strExc(3)) & "½ÐÀu¥ý®Ö¹ï¤½³ø³t°h¤uµ{®v¶i¦æ¤G®Ö¡C"
                      'Modified by Lydia 2022/01/21 ¦¬¥ó¤H½Ð°²®É±H«H¼u°T®§·|¥d¦í§å¦¸
                      'Call PUB_SendMail(strUserNum, strExc(2) & ";" & strExc(1), "", strExc(0), strExc(10), , , , , , strExc(5))
                      strExSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                           " VALUES ( '" & strUserNum & "','" & strExc(2) & ";" & strExc(1) & "',to_char(sysdate,'yyyymmdd')" & _
                           ",to_char(sysdate,'hh24miss'),'" & strExc(0) & "','" & strExc(10) & "','" & strExc(5) & "')"
                      cnnConnection.Execute strExSql
                      'end 2022/01/21
                   End If
                End If
            End If
            'end 202/08/02
         End If
         
         '­Y¦³¥¼µo¤å§Þ³N³ø§i®Éµo Mail ³qª¾©Ó¿ì¤H
         If m_str421CP09 <> "" And m_str421CP14 <> "" Then
            Dim stPS As String
            stPS = "¡°ª`·N¡A¥»®×¤w¤½§i¤w¥i©Ó¿ì¥B©Ó¿ì´Á­­¬° " & ChangeTStringToTDateString(Format(Val(m_str421CP48) - 19110000)) & "¡I"
            'Modified by Lydia 2022/01/21 ¦¬¥ó¤H½Ð°²®É±H«H¼u°T®§·|¥d¦í§å¦¸
            'Call PUB_SendMail(strUserNum, m_str421CP14, m_str421CP09, "§Þ³N³ø§i¤å¥ó»ô³Æ³qª¾", "", stPS)
            strExSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                 " VALUES ( '" & strUserNum & "','" & m_str421CP14 & "',to_char(sysdate,'yyyymmdd')" & _
                 ",to_char(sysdate,'hh24miss'),'§Þ³N³ø§i¤å¥ó»ô³Æ³qª¾','" & stPS & "')"
            cnnConnection.Execute strExSql
            'end 2022/01/21
         End If
         
         cnnConnection.CommitTrans 'Modify By Sindy 2015/8/11
      End If
   Next dblFCnt
   
   'cnnConnection.CommitTrans
   
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
   
   Set rsQuery = Nothing 'Added by Lydia 2021/08/16
   Set fs = Nothing
   Set f = Nothing
   Screen.MousePointer = vbDefault
   
'   Set rsTmp = Nothing
   'Modify By Sindy 2024/6/3 ·¨¶²ªÚ¸g²z«ü¥Ü,Á`¸g²z¤w®Ö¥Ü°±¤î¦¹¶µ¤ÀÃþ¤u§@¡A¦¹Ãþ³qª¾¤]¥i°±¤îµo°e
'   Call GetSendMailIPC 'Add By Sindy 2012/8/16
   Call IsRecordExist '²£¥Íµ§¼Æ
   
   'Add By Sindy 2025/2/3 ¤º±M¤H­û¶×¤J¤½§i¤½³ø¡]1228¡^¡A¨t²Î¦Û°Êµo«Hª¾·|¥~±M¦Uµ{§Ç¤H­û
   PUB_SendMail strUserNum, "FCP_1@taie.com.tw", "", "¤½§i¤½³ø¤w¶×¤J¨÷©v°Ï¡A½Ð³B²z«áÄò¬yµ{¡C", "¦p¦®~"
   
   MsgBox "ÂàÀÉ§¹²¦¡I(ÂàÀÉªá¶O®É¶¡¡G" & strTime & "  " & time() & ")" & vbCrLf & strMsg
   Me.Height = MinHeight
   Call PUB_SendMailCache 'Added by Lydia 2022/01/21
   
   Exit Sub
   
ErrHand:
   Set fs = Nothing
   Set f = Nothing
   Screen.MousePointer = vbDefault
'   Set rsTmp = Nothing
   If Err.NUMBER = 76 Then
      MsgBox "ÂàÀÉ¸ê®Æ§¨¡]" & txtPath2 & "\img_1\isu" & Format(Left(txtTMBM07, 2), "000") & Format(Right(txtTMBM07, 2), "000") & "¡^¤ºµL¸Ó´Á¤½³ø¸ê®Æ¡I"
      txtPath2.SetFocus
   Else
      cnnConnection.RollbackTrans
      If Err.NUMBER = -2147217873 Then
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¥Ó½Ð®×¸¹¡]" & strTPB01 & "¡^" & vbCrLf & strErrTxt & ": ¹H¤Ï¥²¶·¬°°ß¤@ªº­­¨î±ø¥ó" & vbCrLf & strSql
      Else
         MsgBox "²Ä" & dblFCnt & "µ§¡AÂàÀÉ¥¢±Ñ¡I" & "¤½³ø¥Ó½Ð®×¸¹¡]" & strTPB01 & "¡^" & vbCrLf & strErrTxt & Err.Description & vbCrLf & strSql
      End If
   End If
End Sub

Private Function ReadXmlData() As Boolean
Dim strData As String, strText As String, strTitNM As String, strChar As String
Dim dblStar As Double, dblEnd As Double, dblLastEnd As Double, dblChar As Double
Dim rsTmp As New ADODB.Recordset
Dim strFreeAgentCode As String
Dim strChineseNM As String, strEnglishNM As String, intApp As Integer 'Add By Sindy 2013/4/15
Dim dblRunStar As Double 'Add By Sindy 2018/11/12
Dim strGetData1 As String, strGetData2 As String, strGetData3 As String 'Add By Sindy 2018/11/12
Dim strUpdNewTA02 As String 'Add By Sindy 2020/1/9
   
   ReadXmlData = True
   
   strTPB01 = "": strTPB02 = "": dblTPB03 = Empty: strTPB04 = ""
   strTPB05 = "": strTPB06 = "": strTPB07 = "": strTPB07_1 = "": strTPB07_temp1 = "": strUpdNewTA02 = ""
   strTPB08 = "": strTPB09 = ""
   strPA160 = "" 'Add By Sindy 2012/3/3
   'Add By Sindy 2012/8/9
   'Modify By Sindy 2016/3/2 +: strTPB13 = ""
   strTPB10 = "": strTPB11 = "": strTPB12 = "": strTPB13 = ""
   '2012/8/9 End
   strTPB38 = "" 'Add By Sindy 2019/9/4
   'Add By Sindy 2013/4/15
   For i = 1 To 10
      strTPBcApp(i) = ""
      strTPBeApp(i) = "" 'Add By Sindy 2018/11/12
   Next i
   '2013/4/15 End
   dblTPB34 = Empty: dblTPB35 = Empty: strTPB36 = "": strTPB37 = "" 'Add By Sindy 2018/11/12
   strAChinese = "": strAChinese1 = "": strAddress1 = ""
   m_strPA14 = Empty
   m_bol412 = False
   bolTaieCase = False: strTaieCaseNo = ""
   m_strNextDueDate = ""
   m_strNextFeeDate = ""
   m_strAgreeOnDate = "" 'Add By Sindy 2021/8/17
   m_str421CP09 = ""
   m_str421CP14 = ""
   m_str421EP06 = ""
   m_str421CP48 = ""
   strMsg = ""
   
   If GetXmlData(1, "volno", "¨÷¼Æ", strData, dblEnd) = True Then
      strTPB04 = Format(strData, "00")
   End If
   If GetXmlData(1, "isuno", "´Á¼Æ", strData, dblEnd) = True Then
      strTPB05 = Format(strData, "00")
   End If
   dblStar = InStr(m_strTextBox, "<publication-reference>")
   If GetXmlData(dblStar, "doc-number", "±M§Q¸¹¼Æ", strData, dblEnd) = True Then
      strTPB02 = strData
   End If
   If GetXmlData(dblStar, "date", "¤½§i¤é", strData, dblEnd) = True Then
      dblTPB03 = DBDATE(strData)
   End If
   dblStar = InStr(m_strTextBox, "<application-reference")
   If GetXmlData(dblStar, "doc-number", "¥Ó½Ð®×¸¹", strData, dblEnd) = True Then
      strTPB01 = strData
      '¥Ó½Ð®×¤~­n±a
      Erase pa
      ReDim pa(1 To TF_PA) As String
      strSql = "SELECT * FROM Patent " & _
               "WHERE PA11 = '" & strTPB01 & "' AND " & _
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
            '¥u±±¨î¤º±M´N¦n
            If "" & RsTemp.Fields("PA01") = "P" Then
               m_strPA14 = PUB_GetPrePA14(pa, m_bol412)
            End If
         End If
      End If
   End If
   If Mid(strTPB01, 4, 1) = "2" Then
      strTPB09 = "N"
   Else
      strTPB09 = ""
   End If
   'Add By Sindy 2018/11/12
   If GetXmlData(dblStar, "date", "¥Ó½Ð¤é", strData, dblEnd) = True Then
      dblTPB34 = DBDATE(strData)
   End If
   '2018/11/12 END
   
   'Add By Sindy 2012/3/3 +°ê»Ú¤ÀÃþ
   'dblStar = InStr(m_strTextBox, "<classification-locarno>") '³]­p : ³]­p¤ÀÃþ¸¹
   'dblStar = InStr(m_strTextBox, "<classification-ipc>") 'µo©ú/·s«¬ : °ê»Ú¤ÀÃþ¸¹
   dblStar = InStr(m_strTextBox, "<classification-")
   If dblStar > 0 Then
      If GetXmlData2(dblStar, "main-classification", "°ê»Ú¤ÀÃþ", strData, dblEnd) = True Then
         If Trim(strData) <> "" Then
            strPA160 = Left(strData, 4) '°ê»Ú¤ÀÃþ«e4½X
            'Add By Sindy 2012/8/9
            strTPB10 = strData '°ê»Ú¤ÀÃþ¸¹
            
            'Add By Sindy 2013/8/19 ²£·~§O¤ÀÃþ
            strTPB12 = GetPatentIPC("2", strTPB10, strTPB02)
            '2013/8/19 END
            'Add By Sindy 2016/3/2 ®×¥óÄÝ©Ê
            strTPB13 = GetPatentIPC("3", strTPB10, strTPB02)
            '2016/3/2 END
            
            'Åª¨úIPC¤ÀÃþ:
            '1.³]­p±M§Q§¡¬°11.³]­pÃþ
            If Left(strTPB02, 1) = "D" Then
               strPA160 = strData '³]­p¤ÀÃþ¸¹¥þ¼Æ¦s¤J
               strTPB11 = "11"
            Else
               'Modify By Sindy 2013/8/19 ¼g¦¨¦@¥Î¨ç¼Æ
               strTPB11 = GetPatentIPC("1", strTPB10, strTPB02)
               '2013/8/19 END
            End If
            '2012/8/9 End
            If strPA160 = "" Then
               strErrTxt = "°ê»Ú¤ÀÃþ¤£¥iªÅ¥Õ¡I"
               ReadXmlData = False
            End If
            
            'Add By Sindy 2013/8/19
            If strTPB12 = "" Then
               strErrTxt = "²£·~§O¤ÀÃþ¤£¥iªÅ¥Õ¡I"
               ReadXmlData = False
            End If
            '2013/8/19 END
            
            'Add By Sindy 2016/3/2
            If strTPB13 = "" Then
               strErrTxt = "®×¥óÄÝ©Ê¤£¥iªÅ¥Õ¡I"
               ReadXmlData = False
            End If
            '2016/3/2 END
            
            'Add By Sindy 2012/8/16 IPC¤ÀÃþÂkÃþ¤£¨ì®É,°O¿ý°ê»Ú¤ÀÃþ¸¹
            If strTPB11 = "" Then
               If InStr(m_PI02, strTPB10) = 0 Then
                  'Modify By Sindy 2013/2/18
                  'm_PI02 = m_PI02 & "¡F" & strTPB10
                  m_PI02 = m_PI02 & strTPB10 & " ¥Ó½Ð®×¸¹¬° " & strTPB01 & vbCrLf
                  '2013/2/18 End
               End If
            End If
            '2012/8/16 End
         End If
      End If
   End If
   '2012/3/3 End
   
   strText = "agents": strTitNM = "¥N²z¤H"
   dblStar = InStr(m_strTextBox, "<" & strText & ">")
   dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
   If dblStar > 0 Then
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
            'Add By Sindy 2017/10/30 ¼W¥[¤ñ¹ï¥N²z¤H
            'Modify By Sindy 2023/8/2
'            strData = ReplaceMadeWord(strData, "?") 'Modify By Sindy 2018/5/21 ÀË¬d³y¦r
'            strData = PUB_FilterBulletinSpecWord("2", strData, "")
            '2023/8/2 END
            '2017/10/30 END
            'Modify By Sindy 2018/7/23 ±q¤U­±if²¾¥X¨Ó§PÂ_
'            If strData = "ÀF±Ò®õ" Then strData = "ÀF•K®õ"
            If bolTaieCase = True And strData <> "" Then
               If InStr(1, strOurAgentName, strData) > 0 Then
                  strTPB07 = GetTAgentName("01", "TA03")
                  strTPB07_1 = "01"
                  strTPB08 = GetTAgentName("01", "TA04")
               End If
            End If
            '2018/7/23 END
            If strTPB07_temp1 = "" Then strTPB07_temp1 = strData '°O¿ý²Ä¤@¦ì¥X¦W¥N²z¤H
            '©|¥¼Åª¨ú¨ì¥N²z¤H¦WºÙ®É
            'Modify By Sindy 2020/1/9
            'If Trim(strTPB07) = "" And Trim(strData) <> "" Then
            If Trim(strData) <> "" Then
            '2020/1/9 END
               'ÀË¬d¬O§_¬°¥»©Ò¥N²zªº®×¥ó
'                     strSql = "select cp09 from caseprogress,(SELECT PA01,PA02,PA03,PA04 FROM Patent WHERE PA11='" & strTPB01 & "' AND PA09='000' and pa23='1') " & _
'                              "Where CP01=pa01 And cp02=pa02 And cp03=pa03 And cp04=pa04 " & _
'                              "and instr('" & NewCasePtyList & "',cp10)>0 and cp27 is not null "
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                     If intI = 1 And InStr(1, strOurAgentName, strData) > 0 Then
'                        strTPB07 = GetTAgentName("01", "TA03")
'                        strTPB07_1 = "01"
'                        strTPB08 = GetTAgentName("01", "TA04")
'                        Exit For
'                     End If
'               If bolTaieCase = True Then
'                  If InStr(1, strOurAgentName, strData) > 0 Then
'                     strTPB07 = GetTAgentName("01", "TA03")
'                     strTPB07_1 = "01"
'                     strTPB08 = GetTAgentName("01", "TA04")
'                     Exit For
''                        Else
''                           strMsg = strTaieCaseNo & "¬°¥»©Ò®×¥ó¦ý¥N²z¤H¨Ã«D¥»©Ò"
''                           Call ReadTxt1(strTPB01, strTPB02, strMsg, "", "", "")
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
                  If strTPB08 = "" Then
                  '2020/1/9 END
                     If IsNull(rsTmp.Fields("TA02")) = False Then
                        strTPB07_1 = rsTmp.Fields("TA02")
                     End If
                     If IsNull(rsTmp.Fields("TA03")) = False Then
                        strTPB07 = rsTmp.Fields("TA03")
                     End If
                     If IsNull(rsTmp.Fields("TA04")) = False Then
                        strTPB08 = rsTmp.Fields("TA04")
                     End If
                  End If
                  'Modify By Sindy 2020/1/9 °j°é­n¶]§¹,Åª¨ú¥þ³¡¥X¦W¥N²z¤H¸ê®Æ
                  'rsTmp.Close: Exit For
               Else
                  'Modify By Sindy 2020/1/9
                  '·s¼W°ê¤º¤½³ø¥N²z¤HÀÉ
                  strFreeAgentCode = PUB_GetFreeAgentCode("P")
                  If strTPB07_1 = "" Then strTPB07_1 = strFreeAgentCode '°O¿ý²Ä¤@¦ì¥X¦W¥N²z¤HID
                  strUpdNewTA02 = strUpdNewTA02 & ",'" & strFreeAgentCode & "'" 'Add By Sindy 2020/1/9
                  strSql = "INSERT INTO TAgent (TA01,TA02,TA03,TA04,TA05) " & _
                           "VALUES ('P','" & strFreeAgentCode & "','" & Trim(strData) & "',null," & dblTPB03 & ")"
                  cnnConnection.Execute strSql
                  '2020/1/9 END
               End If
               rsTmp.Close
            End If
         End If
         dblChar = dblEnd
      Next dblChar
      '©|¥¼Åª¨ú¨ì¥N²z¤H¦WºÙ®É,«h§ó·s²Ä¤@¦ì¥X¦W¥N²z¤H¸ê®Æ
      If Trim(strTPB07) = "" And strTPB07_temp1 <> "" Then
         strTPB07 = strTPB07_temp1
         strTPB08 = strTPB07_temp1
         'Modify By Sindy 2020/1/9 Mark,§ï«e­±³vµ§µL¸ê®Æ,«hinsert
'         If InStr(strTPB07_temp1, "?") = 0 Then
'            '·s¼W°ê¤º¤½³ø¥N²z¤HÀÉ
'            strFreeAgentCode = PUB_GetFreeAgentCode("P")
'            strTPB07_1 = strFreeAgentCode
'            'Modify By Sindy 2014/9/2 ·s¥N²z¤Hªº¨Æ°È©Ò¦WºÙÄæ©ñNull
''            strSql = "INSERT INTO TAgent (TA01,TA02,TA03,TA04,TA05) " & _
''                     "VALUES ('P','" & strTPB07_1 & "','" & Trim(strTPB07) & "','" & Trim(strTPB08) & "'," & dblTPB03 & ")"
'            strSql = "INSERT INTO TAgent (TA01,TA02,TA03,TA04,TA05) " & _
'                     "VALUES ('P','" & strTPB07_1 & "','" & Trim(strTPB07) & "',null," & dblTPB03 & ")"
'            cnnConnection.Execute strSql
'         End If
      'Modify By Sindy 2020/1/9 §ó·s,·s¥N²z¤Hªº¨Æ°È©Ò¦WºÙ
      ElseIf strTPB08 <> "" And strUpdNewTA02 <> "" Then
         strUpdNewTA02 = Mid(strUpdNewTA02, 2)
         strSql = "UPDATE TAgent SET TA04='" & strTPB08 & "'" & _
                  " WHERE TA01='P' AND TA02 in(" & strUpdNewTA02 & ")"
         cnnConnection.Execute strSql
         '2020/1/9 END
      End If
      '¬°¥»©Ò®×¥ó¦ý¥N²z¤H¨Ã«D¥»©Ò
      If bolTaieCase = True And strTPB07_1 <> "01" Then
         strMsg = strTaieCaseNo & "¬°¥»©Ò®×¥ó¦ý¥N²z¤H¨Ã«D¥»©Ò¡A¬°¡e" & strTPB07_1 & " " & strTPB07 & " " & strTPB08 & "¡f"
         Call ReadTxt1(strTPB01, strTPB02, strMsg, "", "", "")
         Call PrintPaper(strTPB01, strTPB02, strMsg, "", "")
      End If
   End If
   
   strText = "applicants": strTitNM = "¥Ó½Ð¤H"
   dblStar = InStr(m_strTextBox, "<" & strText & ">")
   dblLastEnd = InStr(m_strTextBox, "</" & strText & ">")
   If dblStar > 0 Then
      For dblChar = dblStar To dblLastEnd
         For j = 1 To 2
            strData = ""
            If j = 1 Then
               strText = "last-name": strTitNM = "¥Ó½Ð¤H¦WºÙ"
            ElseIf j = 2 Then
               strText = "address": strTitNM = "¥Ó½Ð¤H¦a§}"
            End If
            dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
            If dblStar < dblChar Then Exit For
            If dblStar > dblLastEnd Then dblChar = dblStar: Exit For
            '***** ¸ÑªRXML *****
            If GetXmlData(dblChar, strText, strTitNM, strData, dblEnd) = False Then
            '***** End
               Exit For
            End If
            If j = 1 Then '¥Ó½Ð¤H¦WºÙ
               'Modify By Sindy 2017/7/3
               '©m¦W¦³³y¦r¦³¹Ï¤ù
               'strData=¸âµú<img align="absmiddle" height="18px" width="27px" file="106203003/106203003-009.TIF" alt="¨ä¥L«D¹Ï¦¡ ed10999.png" img-content="tif" orientation="portrait" inline="yes" giffile="106203003/106203003-009.png"></img>
               If InStr(strData, "<") > 0 Then
                  strData = Left(strData, InStr(strData, "<") - 1)
               End If
               '2017/7/3 END
               strAChinese = strData
               If strAChinese1 = "" Then strAChinese1 = strData
            ElseIf j = 2 Then '¥Ó½Ð¤H¦a§}
               If strAddress1 = "" Then strAddress1 = strData
               If Trim(strData) <> "" Then
                  If strTPB06 = "" Then
                     '¥ý¥Î¥þ¦W¤ñ¹ï¦a°Ï
                     'Modify By Sindy 2019/9/4 + , strTPB38
                     If GetNationNo(strData, strTPB38) <> "" Then
                        strTPB06 = strData
                        Exit For
                     End If
                     '³v¦r¤ñ¹ï
                     For i = 1 To Len(strData)
                        strChar = Left(strData, i)
                        strChar = Replace(strChar, "»O", "¥x")
                        'Modify By Sindy 2019/9/4 + , strTPB38
                        If GetNationNo(strChar, strTPB38) <> "" Then
                           strTPB06 = strChar
                           Exit For
                        End If
                        '[¯S¨Ò]³B²z¥xÆW¦a°Ï¦WºÙ
                        If Len(strChar) = 3 Then
                           strChar = Left(strChar, 2) & "¿¤"
                           'Modify By Sindy 2019/9/4 + , strTPB38
                           If GetNationNo(strChar, strTPB38) <> "" Then
                              strTPB06 = strChar
                              Exit For
                           End If
                        End If
                     Next i
                     '¼Ò½k¤ñ¹ï¦a°Ï¦WºÙ
                     If strTPB06 = "" Or strTPB06 = "020" Then '020.¤¤°ê¤j³°
                        If strAChinese <> "" Then
                           'Modify By Sindy 2019/9/4 + , strTPB38
                           strChar = GetNationLike(strAChinese, strTPB38)
                           If strChar <> "" Then
                              strTPB06 = strChar
                              Exit For
                           End If
                        End If
                     ElseIf strTPB06 <> "" Then
                        Exit For
                     End If
                  End If
               End If
            End If
            dblChar = dblEnd
         Next j
         'Modify By Sindy 2023/8/2
'         strAChinese1 = ReplaceMadeWord(strAChinese1, "?") 'Modify By Sindy 2018/5/21 ÀË¬d³y¦r
'         strAChinese1 = PUB_FilterBulletinSpecWord("1", strAChinese1, GetPrjNationName(strTPB06))
         '2023/8/2 END
      Next dblChar
   End If
   
   'Add By Sindy 2013/4/15 ¤ý°ÆÁ`­n¥Ó½Ð¤H¸ê®Æ°µ²Î­p¥Î,¦s¼È¦sÀÉ
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
               'Modify By Sindy 2017/7/3
               '©m¦W¦³³y¦r¦³¹Ï¤ù
               'strData=¸âµú<img align="absmiddle" height="18px" width="27px" file="106203003/106203003-009.TIF" alt="¨ä¥L«D¹Ï¦¡ ed10999.png" img-content="tif" orientation="portrait" inline="yes" giffile="106203003/106203003-009.png"></img>
               If InStr(strData, "<") > 0 Then
                  strData = Left(strData, InStr(strData, "<") - 1)
               End If
               '2017/7/3 END
               'Modify By Sindy 2023/8/2
'               strData = ReplaceMadeWord(strData, "?") 'Modify By Sindy 2018/5/21 ÀË¬d³y¦r
'               strChineseNM = PUB_FilterBulletinSpecWord("1", strData, GetPrjNationName(strTPB06))
               strChineseNM = strData
               '2023/8/2 END
            ElseIf j = 2 Then '¥Ó½Ð¤H­^¤å¦WºÙ
               strEnglishNM = strData
            End If
            dblChar = dblEnd
         Next j
         intApp = intApp + 1
         'Add By Sindy 2015/12/11 ¸ê®Æ®w¥u¦s10¦ì¥Ó½Ð¤H
         If intApp >= 11 Then
            Exit For
         End If
         '2015/12/11 END
         'Add By Sindy 2018/11/12
'         If strChineseNM <> "" Then
'            strTPBcApp(intApp) = strChineseNM
'         Else
'            If strEnglishNM <> "" Then
'               strTPBcApp(intApp) = strEnglishNM
'            End If
'         End If
         If strChineseNM <> "" Then
            strTPBcApp(intApp) = strChineseNM
         End If
         If strEnglishNM <> "" Then
            strTPBeApp(intApp) = strEnglishNM
         End If
         '2018/11/12 END
      Next dblChar
   End If
   '2013/4/15 End
   
   'Add By Sindy 2018/11/12
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
               strGetData3 = DBDATE(strData)
               
               If dblTPB35 = 0 Then dblTPB35 = strGetData3 'Àu¥ýÅv¤é´Á
               strTPB36 = strTPB36 & ";" & strGetData2 'Àu¥ýÅv¸¹¼Æ
               strTPB37 = strTPB37 & ";" & strGetData1 'Àu¥ýÅv°ê®a
            End If
            dblChar = dblEnd
         Next j
      Next dblChar
   End If
   If strTPB36 <> "" Then strTPB36 = Mid(strTPB36, 2)
   If strTPB37 <> "" Then strTPB37 = Mid(strTPB37, 2)
   '2018/11/12 END
   
   Set rsTmp = Nothing
End Function

'ºI¨úXML¸ê®Æ¤@
'Modify By Sindy 2013/4/15 +strEndTag
Private Function GetXmlData(dblChar As Double, strText As String, strTitNM As String, ByRef strData As String, ByRef dblEnd As Double, Optional strEndTag As String = "") As Boolean
Dim dblStar As Double
   
   GetXmlData = False
   strData = "": dblEnd = 0
   dblStar = InStr(dblChar, m_strTextBox, "<" & strText & ">") + Len("<" & strText & ">") - 1
   If dblStar <= dblChar Then
      Exit Function
   End If
   'Modify By Sindy 2013/4/15
   If strEndTag <> "" Then
      dblEnd = InStr(dblStar, m_strTextBox, strEndTag) - 1
   Else
   '2013/4/15 End
      dblEnd = InStr(dblStar, m_strTextBox, "</" & strText & ">") - 1
   End If
   If dblStar >= dblEnd Or dblEnd <= 0 Then
      Exit Function
   End If
   strData = Trim(Mid(m_strTextBox, dblStar + 1, (dblEnd - dblStar)))
   strData = Trim(Replace(ChgSQL(strData), "amp;", ""))
   GetXmlData = True
End Function

'Add By Sindy 2012/3/3
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

Private Function IsPA22Ok(ByVal stPA11 As String, ByVal stPA22 As String, ByRef stMomPA22 As String) As Boolean

On Error GoTo ErrHnd
   
   IsPA22Ok = True
   
   '¥Ó½Ð®×¸¹§ï½X¼Æ
   strSql = "Select PA22 FROM PATENT where PA11='" & Left(stPA11, 9) & "' AND PA01='P' AND PA09='000' AND PA23='1' AND PA22 IS NOT NULL"
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      stMomPA22 = "" & adoRecordset.Fields("PA22")
      '­Y¥À®×ÃÒ®Ñ¸¹¬°¼Æ¦r«h¥u¤ñ¸û¼Æ¦r³¡¤À
      If IsNumeric(stMomPA22) Then stPA22 = Mid(stPA22, 2)
'      If stPA22 = stMomPA22 Then
'         IsPA22Ok = True
      If stPA22 <> stMomPA22 Then
         IsPA22Ok = False
      End If
   Else
      stMomPA22 = ""
   End If
   
   CheckOC
   Exit Function
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Function IsTPB02Exist(ByVal strTPB02 As String, ByRef strErr As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
'   If strTPB02 = "D144217" Or strTPB02 = "D144062" Or strTPB02 = "D144063" Then
'      MsgBox strTPB02
'   End If
   
   IsTPB02Exist = False
   If Len(strTPB01) > 9 Then
      strSql = "SELECT * FROM TPBulletin " & _
               "WHERE TPB02='" & strTPB02 & "' AND " & _
                  "substr(TPB01,1,9)<>'" & Left(strTPB01, 9) & "' " & _
                  "AND TPB04||TPB05<'" & strTPB04 & strTPB05 & "' "
   Else
      strSql = "SELECT * FROM TPBulletin " & _
               "WHERE TPB02='" & strTPB02 & "' AND " & _
                  "TPB01<>'" & strTPB01 & "' " & _
                  "AND TPB04||TPB05<'" & strTPB04 & strTPB05 & "' "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsTPB02Exist = True
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         If strErr <> "" Then strErr = strErr & ","
         strErr = strErr & rsTmp.Fields("TPB01")
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function ChkDataErr() As Boolean
Dim stPA22 As String '¥À®×ÃÒ®Ñ¸¹
Dim i As Integer, j As Integer
Dim strMsg As String, strErr As String
   
   ChkDataErr = False
   
   Call GetNoticeNumber(CStr(dblTPB03)) '¨ÌÂàÀÉ¤¤ªº¤½§i¤é¨ú±o¬Û¹ïªº¤½§i¨÷´Á
   If Val(Left(txtTMBM07, 2)) <> Val(strChkTPB04) Then
      strErrTxt = "¤½§i¤é´Á¡]" & dblTPB03 & "¡^»Pµe­±¤W¿é¤Jªº¤½³ø¨÷¼Æ¡]" & Left(txtTMBM07, 2) & "¡^¤£²Å¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If
   If Val(strTPB04) <> Val(strChkTPB04) Then
      strErrTxt = "¤½§i¤é´Á¡]" & dblTPB03 & "¡^»P¤½³ø¨÷¼Æ¡]" & strTPB04 & "¡^¤£²Å¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If
   If Val(Right(txtTMBM07, 2)) <> Val(strChkTPB05) Then
      MsgBox "¤½§i¤é´Á¡]" & dblTPB03 & "¡^»Pµe­±¤W¿é¤Jªº¤½³ø´Á¼Æ¡]" & Right(txtTMBM07, 2) & "¡^¤£²Å¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If
   If Val(strTPB05) <> Val(strChkTPB05) Then
      strErrTxt = "¤½§i¤é´Á¡]" & dblTPB03 & "¡^»P¤½³ø¨÷´Á¡]" & strTPB05 & "¡^¤£²Å¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If
   
   'Áp¦X®×
   If Len(strTPB01) > 9 Then
      '­Y¬°¥»©Ò®×¥ó»ÝÀË¬d»P¥À®×¬Û¦P
      If bolTaieCase = True Then
         If IsPA22Ok(strTPB01, strTPB02, stPA22) = False Then
            strMsg = "ÃÒ®Ñ¸¹»P¥À®×ÃÒ®Ñ¸¹¡i" & stPA22 & "¡j¤£¦P"
            Call ReadTxt1(strTPB01, strTPB02, strMsg, "", "", "")
            Call PrintPaper(strTPB01, strTPB02, strMsg, "", "")
         End If
      End If
   ElseIf IsTPB02Exist(strTPB02, strErr) = True Then
      strErrTxt = "ÃÒ®Ñ¸¹¡]" & strTPB02 & "¡^¤w¦s¦b¡]­«ÂÐªº¥Ó½Ð®×¸¹¡G" & strErr & "¡^¡A¤£¥i¦sÀÉ¡I" & vbCrLf
      ChkDataErr = True
      Exit Function
   End If
   
   If bolTaieCase = True Then
      If Val(pa(14)) > 0 Then
         If Val(dblTPB03) <> Val(pa(14)) Then
            strErrTxt = "¤½§i¤é¡]" & ChangeTStringToTDateString(Format(Val(dblTPB03) - 19110000)) & "¡^»P²Ä¤@¦¸¿é¤J¡i" & ChangeTStringToTDateString(Format(Val(pa(14)) - 19110000)) & "¡j¤£¦P¡A¤£¥i¦sÀÉ¡I" & vbCrLf
            ChkDataErr = True
            Exit Function
         End If
      Else
         '¤½§i¤é»P¥Ó½Ð©µ½w¤½§iªº¤é´Á¤£¦P®É´£¿ô
         If m_bol412 = True Then
            If Val(dblTPB03) <> Val(m_strPA14) Then
               strMsg = "¤½§i¤é¡]" & ChangeTStringToTDateString(Format(Val(dblTPB03) - 19110000)) & "¡^»P©µ½w¤½§i¤é¡i" & ChangeTStringToTDateString(Format(Val(m_strPA14) - 19110000)) & "¡j¤£¦P"
               Call ReadTxt1(strTPB01, strTPB02, strMsg, "", "", "")
               Call PrintPaper(strTPB01, strTPB02, strMsg, "", "")
            End If
         End If
      End If
      
      '¦³µoÃÒ¤é¤~ÀË¬d
      If pa(22) <> "" And pa(21) <> "" Then
         If strTPB02 <> pa(22) Then
            strErrTxt = "ÃÒ®Ñ¸¹¡]" & strTPB02 & "¡^»P²Ä¤@¦¸¿é¤J¡i" & pa(22) & "¡j¤£¦P¡A¤£¥i¦sÀÉ¡I" & vbCrLf
            ChkDataErr = True
            Exit Function
         End If
      End If
   
      If Check413 = True Then
         strErrTxt = "¥»®×¤w¥Ó½Ð¦ÛºM¡AÀ³¤£¤©¤½§i¡A½Ð¬d©ú¡I" & vbCrLf
         ChkDataErr = True
         Exit Function
      End If
   End If
End Function

'ÀË¬d¦³µo¤å¥Ó½Ðµ{§Çªº¦Û½ÐºM¦^
Private Function Check413() As Boolean
   strExc(0) = "select 1 from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='413' and cp27>0 and cp57 is null" & _
      " and exists(select * from caseprogress b where b.cp09=a.cp43 and instr('101,102,103,104,105,107,301,302,303,304,305,306,307',b.cp10)>0)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Check413 = True
   End If
End Function

'¦a°Ï¦WºÙ¸ê®ÆÀË®Öªí
Private Sub ReadTxt1(strTPB01 As String, strTPB02 As String, strTPB06 As String, strTPB07 As String, strAChinese1 As String, strAddress1 As String)
Dim i As Integer
   
   If m_PrintRpt1 = False Then
      m_PrintRpt1 = True
'      If ff1 > 0 Then Close #ff1
'      ff1 = FreeFile
      m_strFileName1 = "°ê¤º±M§Q¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¸ê®ÆÀË®Öªí.txt"
'      Open PUB_Getdesktop & "\" & m_strFileName1 For Output As ff1
'      Print #ff1, "³Æµù¡G§ï¦r«¬Fixedsys¼Ð·Ç11¸¹¦r¥H¾î¦¡¤W¤U¥ª¥k¦U10MM¦C¦L"
'      Print #ff1, "¥Ó½Ð®×¸¹        ±M§Q¸¹¼Æ   ¦a°Ï¦WºÙ        ¥N²z¤H¦WºÙ   ¥Ó½Ð¤H¦a§}"
'      Print #ff1, "                           ©Î ´£¿ô³Æµù"
'      Print #ff1, "=============== ========== =============== ============ ============================================="
      
      m_strText = "³Æµù¡G§ï¦r«¬Fixedsys¼Ð·Ç11¸¹¦r¥H¾î¦¡¤W¤U¥ª¥k¦U10MM¦C¦L" & vbCrLf
      m_strText = m_strText & "¥Ó½Ð®×¸¹        ±M§Q¸¹¼Æ   ¦a°Ï¦WºÙ        ¥N²z¤H¦WºÙ   ¥Ó½Ð¤H¦a§}" & vbCrLf
      m_strText = m_strText & "                           ©Î ´£¿ô³Æµù" & vbCrLf
      m_strText = m_strText & "=============== ========== =============== ============ =============================================" & vbCrLf
   End If
   For i = 1 To 6
      strTemp(i) = ""
   Next i
   strTemp(1) = Trim(strTPB01)
   strTemp(2) = Trim(strTPB02)
   strTemp(3) = Trim(strTPB06)
   strTemp(4) = Trim(strTPB07)
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
   
'   Print #ff1, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(6)
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
   
   'Add By Sindy 2012/3/3
   If Pub_StrUserSt03 = "M51" Then
      cmdPA160.Visible = True
      cmdTemp.Visible = True 'Add By Sindy 2013/4/15
      cmdTPB12.Visible = True 'Add By Sindy 2013/8/23
   Else
      cmdPA160.Visible = False
      cmdTemp.Visible = False 'Add By Sindy 2013/4/15
      cmdTPB12.Visible = False 'Add By Sindy 2013/8/23
   End If
   
   PUB_ReadPath txtPath1, Me.Name 'Added by Morgan 2020/5/5
   
   'Add By Sindy 2022/3/3
   Set adoStream = New ADODB.Stream
   adoStream.Charset = "UTF-8" '"UTF-8" Unicode
   adoStream.Open
   '2022/3/3 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SavePath txtPath1, Me.Name 'Added by Morgan 2020/5/5
   
   'Add By Sindy 2022/3/3
   adoStream.Close
   Set adoStream = Nothing
   '2022/3/3 END
   
   Set frm04060110 = Nothing
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
         strMsg = "½Ð¿é¤J¥¿½Tªº¤½§i¤é"
         strTit = "¸ê®ÆÀË®Ö"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03_GotFocus
         GoTo EXITSUB
      End If
      
      '¤½§i¤é¤£¯à¤j©ó¨t²Î¤é
      If DBDATE(text03) > strSrvDate(1) Then
         Cancel = True
         strMsg = "¤½§i¤é¤£¯à¤j©ó¨t²Î¤é"
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
   strMsg = "½Ð¿é¤J¤½§i¤é¡I"
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

Call GetNoticeNumber(DBDATE(text03)) '¨Ì¿é¤Jªº¤½§i¤é¨ú±o¬Û¹ïªº¤½§i¨÷´Á
If Val(Left(txtTMBM07, 2)) <> Val(strChkTPB04) Then
   strTit = "ÀË®Ö¸ê®Æ"
   strMsg = "¤½³ø¨÷¼Æ»P¤½§i¤é´Á¤£²Å¡I"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   text03.SetFocus
   Exit Function
End If
If Val(Right(txtTMBM07, 2)) <> Val(strChkTPB05) Then
   strTit = "ÀË®Ö¸ê®Æ"
   strMsg = "¤½³ø´Á¼Æ»P¤½§i¤é´Á¤£²Å¡I"
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
   
   strSql = "SELECT count(TPB01) FROM TPBulletin WHERE TPB04=" & CNULL(Left(txtTMBM07, 2)) & " and TPB05=" & CNULL(Right(txtTMBM07, 2))
   
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

'Add By Sindy 2013/4/15
' ÀË¬d°O¿ý¬O§_¤w¸g¦s¦b
Private Function IsRecordExist_Temp() As Boolean
   Dim rsTmp2 As New ADODB.Recordset
   Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   IsRecordExist_Temp = False
   
   strSql = "SELECT count(TPB01) FROM TPBulletin_sonia WHERE TPB04=" & CNULL(Left(txtTMBM07, 2)) & " and TPB05=" & CNULL(Right(txtTMBM07, 2))
   
   ' Åª¨ú¸ê®Æ®w
   rsTmp2.CursorLocation = adUseClient
   rsTmp2.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   IsRecordExist_Temp = False
   Label3.Caption = "(               µ§)"
   ' ÀË¬dÅª¨úªº¸ê®Æµ§¼Æ
   If rsTmp2.RecordCount > 0 Then
      If rsTmp2.Fields(0) > 0 Then
         IsRecordExist_Temp = True
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
   
   strChkTPB04 = Val(Val(Left(strDate, 4)) - 1911) - 62
   
   j = Val(Mid(strDate, 5, 2))
   i = (j - 1) * 3
   j = Val(Right(strDate, 2))
   If j >= 1 And j < 11 Then
      i = i + 1
   ElseIf j >= 11 And j < 21 Then
      i = i + 2
   ElseIf j >= 21 Then
      i = i + 3
   End If
   strChkTPB05 = i
End Sub

Private Sub PrintPaper(strTPB01 As String, strTPB02 As String, strTPB06 As String, strTPB07 As String, strAddress1 As String)
   intPRow = intPRow + 1
   MSHFlexGrid1.Rows = intPRow + 1
   
   MSHFlexGrid1.TextMatrix(intPRow, 0) = strTPB01
   MSHFlexGrid1.TextMatrix(intPRow, 1) = strTPB02
   
   If strTPB06 = "" Then
      MSHFlexGrid1.TextMatrix(intPRow, 2) = "*"
   Else
      MSHFlexGrid1.TextMatrix(intPRow, 2) = strTPB06 & GetPrjNationName(strTPB06)
   End If
   
   txtChkWord = strTPB07 'Add By Sindy 2024/5/17
   If InStr(txtChkWord, "?") > 0 Then
      MSHFlexGrid1.TextMatrix(intPRow, 3) = "*" & strTPB07
   Else
      MSHFlexGrid1.TextMatrix(intPRow, 3) = strTPB07
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

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("°ê¤º±M§Q¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¸ê®ÆÀË®Öªí") / 2)
Printer.CurrentY = iLine2 * 300
Printer.Print "°ê¤º±M§Q¤½³ø" & Left(txtTMBM07, 2) & "¨÷" & Right(txtTMBM07, 2) & "´Á" & "¸ê®ÆÀË®Öªí"

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
Printer.Print "±M§Q¸¹¼Æ"
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
      .FormatString = "¥Ó½Ð®×¸¹|±M§Q¸¹¼Æ|¦a°Ï¦WºÙ|¥N²z¤H¦WºÙ|¥Ó½Ð¤H¦a§}"
   End With
End Sub
