VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090905 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "¤uµ{®v¤W¶Ç§@·~"
   ClientHeight    =   6920
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6920
   ScaleWidth      =   8040
   Begin VB.TextBox txtPath 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   280
      Index           =   3
      Left            =   2595
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   " ¦s©ñ´£¨ÑÂ½Ä¶°Ñ¦Ò¥Î¤§»¡©ú®Ñ©M¬Û¦ü¤ñ¹ïµ²ªGÀÉ®×"
      Top             =   5400
      Width           =   4125
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "ÂsÄý"
      Height          =   300
      Index           =   3
      Left            =   6780
      TabIndex        =   17
      Top             =   5400
      Width           =   800
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "»¡©ú"
      Height          =   300
      Left            =   3000
      TabIndex        =   18
      Top             =   6000
      Width           =   800
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "ÂsÄý"
      Height          =   300
      Index           =   2
      Left            =   6780
      TabIndex        =   16
      Top             =   5040
      Width           =   800
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "ÂsÄý"
      Height          =   300
      Index           =   1
      Left            =   6780
      TabIndex        =   14
      Top             =   4680
      Width           =   800
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "ÂsÄý"
      Height          =   300
      Index           =   0
      Left            =   6780
      TabIndex        =   12
      Top             =   4320
      Width           =   800
   End
   Begin VB.TextBox txtFilePath 
      Height          =   270
      Left            =   6360
      TabIndex        =   31
      Text            =   "txtFilePath"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   1680
      Left            =   5640
      TabIndex        =   30
      Top             =   1920
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.TextBox txtPath 
      Height          =   280
      Index           =   2
      Left            =   2595
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "\\Typing2\English_Vers"
      Top             =   5040
      Width           =   4000
   End
   Begin VB.TextBox txtPath 
      Height          =   280
      Index           =   1
      Left            =   2595
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "\\Typing2\±M§Q®×¥ó\(®×¸¹«e3½X)"
      Top             =   4680
      Width           =   4000
   End
   Begin VB.TextBox txtPath 
      Height          =   280
      Index           =   0
      Left            =   2595
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "\\Typing2\¹q¤l°e¥ó¼È¦s°Ï"
      Top             =   4320
      Width           =   4000
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "´M§ä(&F)"
      Default         =   -1  'True
      Height          =   300
      Left            =   3300
      TabIndex        =   3
      Top             =   360
      Width           =   800
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   3
      Left            =   2760
      MaxLength       =   2
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   2
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   1
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   23
      Text            =   "FCP"
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   0
      Left            =   5850
      TabIndex        =   9
      Top             =   120
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6840
      TabIndex        =   10
      Top             =   120
      Width           =   930
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   2085
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   7755
      Begin VB.CommandButton cmdAddDir 
         Caption         =   "¶×¤J¸ê®Æ§¨"
         Height          =   345
         Left            =   0
         TabIndex        =   4
         Top             =   1680
         Width           =   1155
      End
      Begin VB.ListBox lstAtt 
         Height          =   1660
         ItemData        =   "frm090905.frx":0000
         Left            =   0
         List            =   "frm090905.frx":0007
         MultiSelect     =   2  '¶i¶¥¦h­«¿ï¨ú
         Sorted          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   7740
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "¶}±Ò"
         Height          =   345
         Left            =   2730
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1680
         Width           =   675
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "¥[¤J"
         Height          =   345
         Left            =   1230
         TabIndex        =   5
         Top             =   1680
         Width           =   675
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "²¾°£"
         Height          =   345
         Left            =   1980
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1680
         Width           =   675
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "¥þ¿ï"
         Height          =   345
         Left            =   3480
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1680
         Width           =   675
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   285
      Left            =   1080
      TabIndex        =   42
      Top             =   720
      Width           =   6480
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "11430;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   225
      Index           =   4
      Left            =   2100
      TabIndex        =   41
      Top             =   1400
      Width           =   5800
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Caption         =   "Lbl3_FM2"
      Size            =   "10231;397"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   225
      Index           =   3
      Left            =   1080
      TabIndex        =   40
      Top             =   1400
      Width           =   980
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Caption         =   "Lbl3_FM2"
      Size            =   "1729;397"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   225
      Index           =   2
      Left            =   2100
      TabIndex        =   39
      Top             =   1095
      Width           =   5800
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Caption         =   "Lbl3_FM2"
      Size            =   "10231;397"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Lbl3 
      Height          =   225
      Index           =   1
      Left            =   1080
      TabIndex        =   38
      Top             =   1095
      Width           =   980
      BackColor       =   -2147483637
      VariousPropertyBits=   27
      Caption         =   "Lbl3_FM2"
      Size            =   "1729;397"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "¡@¡@ ¡@ ­Y¥u­×¹Ï¦¡(¥HPDF¦^¦s)¡A½Ð¦Û¦æ©R¦W¬°FCP05XXXXX-°e¥ó¤é(¥Á°ê¦~).FIG.PDF"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   4
      Left            =   480
      TabIndex        =   36
      Top             =   6600
      Width           =   6750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "³Æµù3¡G³Ì²×ª©¥»¤¤»¡À³¬°¹º½uª©¡A½Ð¦Û¦æ©R¦W¬°FCP0XXXXX-°e¥ó¤é(¥Á°ê¦~).FIX_U"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   35
      Top             =   6360
      Width           =   6645
   End
   Begin VB.Label Label7 
      Caption         =   "¨ä¥LÀÉ®×-¦s©ñ¸ô®|¡G     "
      Height          =   255
      Left            =   600
      TabIndex        =   34
      Top             =   5400
      Width           =   1905
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "³Æµù1¡GÀÉ®×ªºÀÉ¦W¥²»Ý¬°FCPXXXX(6½X)¶}ÀY¡F¤W¶Ç«áÀÉ®×¦WºÙ¦Û°Ê¥h±¼«D­^¼Æ¦r¡C"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   33
      Top             =   5760
      Width           =   6660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "³Æµù2¡G¸Ô²ÓÂkÀÉ»¡©ú¡A½ÐÂI¿ï"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   32
      Top             =   6060
      Width           =   2430
   End
   Begin VB.Label Label6 
      Caption         =   "¨ä¥L¥~¤å¥»-¦s©ñ¸ô®|¡G"
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   5040
      Width           =   1995
   End
   Begin VB.Label Label5 
      Caption         =   "¤¤»¡-¦s©ñ¸ô®|¡G"
      Height          =   270
      Left            =   600
      TabIndex        =   28
      Top             =   4680
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "¥~¤å´£¥Ó¥»-¦s©ñ¸ô®|¡G"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó¦WºÙ¡G"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   772
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H1 : "
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   25
      Top             =   1095
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¥N²z¤H : "
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   24
      Top             =   1400
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "³Æµù1¡G«ö¤U½T»{¡AÀÉ®×±N½Æ»s¨ì\\Typing2ªº¨t²Î¦s©ñ¸ô®|¡A½Ð°Ñ¦Ò¤U¦C¸ô®|»¡©ú¡C"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   3960
      Width           =   6540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥»©Ò®×¸¹¡G"
      Height          =   180
      Left            =   120
      TabIndex        =   21
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm090905"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/24 §ï¦¨Form2.0 ; Combo1¡BLbl3(index)
'Create By Lydia 2018/03/12 ¤uµ{®v¤W¶Ç§@·~
Option Explicit

'Private Const UniPath_¤¤»¡ As String = "å°ˆåˆ©æ¡ˆä»¶" '±M§Q®×¥óªº¹ïÀ³Unicode­È 'Mark by Lydia 2024/02/16 ¤w§ï¨ì­ì©lÀÉ°Ï
Dim m_UniPathList As String '±M§Q®×¥ó©³¤UªºÀÉ®×²M³æ

Dim ii As Integer
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
'Modified by Lydia 2020/01/20 §ï¦¨°}¦C
'Dim m_PA01 As String, m_PA02 As String, m_PA03 As String, m_PA04 As String '¤w¬d¸ßªº¥»©Ò®×¸¹
Dim m_Pa(1 To 4) As String
Dim strCompName As String 'ÀÉ¦W¶}ÀY6½X¬y¤ô¸¹
Dim strRepName As String 'ÀÉ¦W¶}ÀY5½X¬y¤ô¸¹
Dim stFtpIP As String
Dim mStrPath3 As String 'Added by Lydia 2018/10/22 Â½Ä¶°Ñ¦Ò¥Î¤§wordª©»¡©ú®Ñªº¦s©ñ¸ô®|


Private Sub FormClear(Optional bolAll As Boolean = False)
Dim oLbl As Control

    If bolAll = True Then
        txtData(1) = ""
        txtData(2) = ""
        txtData(3) = ""
    End If
    
    Combo1.Clear
    For Each oLbl In Lbl3
         oLbl.Caption = ""
    Next
    lstAtt.Clear
    
    'ÁÙ­ì¹w³]­È
    'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
    'txtPath(0).Text = "\\Typing2\¹q¤l°e¥ó¼È¦s°Ï"
    'txtPath(1).Text = "\\Typing2\±M§Q®×¥ó"
    'txtPath(2).Text = "\\Typing2\English_Vers"
    txtPath(0).Text = "\\" & strTyping2Path & "\¹q¤l°e¥ó¼È¦s°Ï"
    'Modified by Lydia 2025/10/02 ¹w³]ªÅ¥Õ---Á§¸g²z
    'txtPath(1).Text = "\\" & strTyping2Path & "\±M§Q®×¥ó"
    'txtPath(2).Text = "\\" & strTyping2Path & "\English_Vers"
    'end 2024/07/22
    txtPath(1).Text = ""
    txtPath(2).Text = ""
    'end 2025/10/02
    txtFilePath.Text = ""
    'Added by Lydia 2020/01/20 ±M§Q®×¥ó©MEnglish_VersÀÉ®×¡G°O¿ý¦¬¤å¸¹
    txtPath(0).Tag = ""
    txtPath(1).Tag = ""
    txtPath(2).Tag = ""
    
    Call CmdEnabled(False)
End Sub

Private Sub cmdAddDir_Click()
   Dim fName As String, strStartFolder As String
   
   '¹w³]¤W¤@¦¸ªº¸ô®|
   strStartFolder = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
   If Trim(strStartFolder) = "" Then strStartFolder = PUB_Getdesktop
   
   fName = PUB_GetFolder(Me.hWnd, strStartFolder, "½Ð¿ï¨ú¸ê®Æ§¨:")
   If fName <> "" Then 'they did not hit cancel
       If InStrRev(fName, "\") = 0 Then
           Exit Sub
       End If
       'Added by Lydia 2018/05/03 ¸ô®|±Æ°£&
       If InStr(fName, "&") > 0 Then
            MsgBox fName & vbCrLf & vbCrLf & "¡i&¡j²Å¸¹¬°¨t²Î«O¯d¦r¡A¤£¥i¨Ï¥Î©ó¸ô®|¡I", vbExclamation
            Exit Sub
       End If
       'end 2018/05/03
       'Åª¨ú¸ê®Æ§¨ªºÀÉ®×
       txtFilePath.Text = fName
       SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", txtFilePath
       RefreshList
   End If
            
End Sub

Private Sub RefreshList()
Dim strTmp As String
Dim tmpArr As Variant
Dim fs, f

   strExc(1) = "": strExc(2) = "": strExc(3) = ""
   File1.path = txtFilePath.Text
   File1.Refresh
   If File1.ListCount > 0 Then
      For ii = 0 To File1.ListCount - 1
         strTmp = File1.List(ii)
         '­­PDF©MWordÀÉ
         'Modified by Lydia 2018/04/27 PDFÀÉ:­­¨îÀÉ¦W¬°FCP0XXXXX.ORI.PDF©ÎFCP0XXXXX.FIG.PDF
         'If UCase(Right(Trim(strTmp), 4)) = ".PDF" Or UCase(Right(Trim(strTmp), 4)) = ".DOC" Or UCase(Right(Trim(strTmp), 5)) = ".DOCX" Then
         'Modified by Lydia 2018/05/18 ÀË¬d°ÆÀÉ¦W
         'If UCase(Right(Trim(strTmp), 8)) = ".ORI.PDF" Or UCase(Right(Trim(strTmp), 8)) = ".FIG.PDF" _
                    Or UCase(Right(Trim(strTmp), 4)) = ".DOC" Or UCase(Right(Trim(strTmp), 5)) = ".DOCX" Then
         'Modify By Sindy 2025/10/27 ChkAttFileName§ï¬°¦@¥Î¨ç¼Æ +, m_Pa(1), m_Pa(2)
         If ChkAttFileName(strTmp, m_Pa(1), m_Pa(2)) = True Then
               'ÀÉ®×©R¦W¤£²Å³W©w¡A¦r­º¥²¶·¬°ex. FCP012345
               If Mid(UCase(strTmp), 1, Len(strCompName)) <> UCase(strCompName) And _
                    Mid(UCase(strTmp), 1, Len(strRepName)) <> UCase(strRepName) Then
                    strExc(2) = strExc(2) & strTmp & vbCrLf
                    GoTo JumpNextFile
               End If
               'ÀË¬dÀÉ®×¬O§_¥¿¦b¨Ï¥Î¤¤
               If PUB_ChkFileOpening(txtFilePath & "\" & strTmp) = True Then
                     MsgBox txtFilePath & "\" & strTmp & vbCrLf & "ÀÉ®×¥¿¦b¨Ï¥Î¤¤¡]½ÐÃö³¬¡^¡A¤è¥iÄ~Äò¾Þ§@¡C", vbExclamation
                     Exit Sub
               End If
               
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set f = fs.GetFile(txtFilePath & "\" & strTmp)
                'ÀÉ®×¤j¤p¬° 0 KB ¦³»~
                If f.Size = 0 Then
                   ShowMsg strTmp & MsgText(9221)
                   Exit Sub
                End If
               '¥þ³¡ÀË¬d«á¤~·s¼W
               strExc(1) = strExc(1) & strTmp & "&"
         Else
               strExc(3) = strExc(3) & vbCrLf & txtFilePath & "\" & strTmp
         End If
JumpNextFile:
      Next
   End If
   
   If strExc(1) = "" Then 'Added by Lydia 2020/06/17 ¸ê®Æ§¨¤º¦³²Å¦XªºÀÉ®×,´N¤£¼u°T®§¥u¦C¥X²Å¦XªºÀÉ®×(by ¨¦©s¾Ç: ¦P¤@¸ê®Æ§¨©ñ¤£¦P®×¥ó)
        '¥þ³¡ÀË¬d«á¤~·s¼W
        If strExc(2) <> "" Then
             MsgBox "¤U¦CÀÉ®×©R¦W¤£²Å³W©w¡A¦r­º¥²¶·¬°" & strCompName & vbCrLf & strExc(2), vbCritical
             Exit Sub
        End If
        If strExc(3) <> "" Then
             'Modified by Lydia 2018/04/27
             'MsgBox "¤U¦CÀÉ®×¤£¥i¥[¤J¡A½Ð¿ï¾ÜWordÀÉ(*.DOCX/*.DOC)©ÎPDFÀÉ®×(*.PDF) !" & strExc(3), vbInformation
             'Modified by Lydia 2018/05/18 +ZIPÀÉ
             'Modified by Lydia 2018/06/08
             'MsgBox "¤U¦CÀÉ®×¤£¥i¥[¤J¡A½Ð¿ï¾ÜWordÀÉ(*.DOCX/*.DOC)¡B*.ZIPÀÉ¡B*.ORI.PDFÀÉ©Î*.FIG.PDFÀÉ®× !" & strExc(3), vbInformation
             MsgBox "¤U¦CÀÉ®×°ÆÀÉ¦W¤£²Å³W«h¡A½Ð°Ñ¦Ò¸Ô²ÓÂkÀÉ»¡©ú !" & strExc(3), vbInformation
        End If
   End If 'Added by Lydia 2020/06/17
   
   If strExc(1) <> "" Then
        If lstAtt.ListCount > 0 Then
            If MsgBox("¬O§_²MªÅ¦Cªí¡H", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
                 lstAtt.Clear
            End If
        End If
        tmpArr = Empty
        tmpArr = Split(strExc(1), "&")
        For ii = 0 To UBound(tmpArr)
             If Trim(tmpArr(ii)) <> "" Then
                 lstAtt.AddItem Trim(txtFilePath & "\" & tmpArr(ii))
             End If
        Next
   End If
End Sub

Private Sub cmdFind_Click()
Dim Cancel As Boolean

    Txtdata_Validate 0, Cancel
    If Cancel = True Then
        Exit Sub
    End If
    If Len(txtData(1)) <> 6 Then
        MsgBox "¥»©Ò®×¸¹½Ð¿é¤J6½X!! '", vbCritical
        txtData(1).SetFocus
        Txtdata_GotFocus 1
        Exit Sub
    End If
    If Trim(txtData(2)) = "" Then txtData(2) = "0"
    If Trim(txtData(3)) = "" Then txtData(3) = "00"
    'Added by Lydia 2024/02/23 ¶}©ñP®×¥i¤W¶ÇÀÉ®×
    If txtData(0) = "P" Or txtData(0) = "FCP" Then
       If txtData(0) = "P" Then
         If PUB_ChkIsFMP(txtData(0), txtData(1), txtData(2), txtData(3)) = False Then
             MsgBox "¥u¥i¤W¶Ç¾ÈµØ®×¡þFMP®×¡I"
             txtData(0).SetFocus
             Txtdata_GotFocus 0
             Exit Sub
         End If
       End If
    End If
    'end 2024/02/23
    
    Call FormClear
    m_Pa(1) = txtData(0): m_Pa(2) = txtData(1)
    m_Pa(3) = txtData(2): m_Pa(4) = txtData(3)
    
   '«È¤á¦WºÙ:¤¤->­^->¤é ; ¥N²z¤H¦WºÙ: ­^->¤¤->¤é
   'Modified by Lydia 2019/11/01 +¥Ó½Ð¤H2~5 (PA27~PA30)
   strExc(0) = "SELECT PA01,PA02,PA03,PA04,PA05,PA06,PA07,PA26,PA75," & _
                     " NVL(CU04,NVL(CU05,CU06)) CNAME,NVL(FA05,NVL(FA04,FA06)) FNAME" & _
                     " ,PA27,PA28,PA29,PA30" & _
                     " FROM PATENT,CUSTOMER,FAGENT" & _
                     " WHERE " & ChgPatent(m_Pa(1) & m_Pa(2) & m_Pa(3) & m_Pa(4)) & _
                     " AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+)" & _
                     " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+)"
    intI = 0
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        'Added by Lydia 2019/11/01 §Q¯q½Ä¬ð®×¥ó¡G³v®×¸¹§PÂ_
        If strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤é And XY¯S®íÅv­­½d³ò <> "" Then
            If PUB_ChkCufaByCase(Me.Name, m_Pa(1), m_Pa(1) & m_Pa(2) & m_Pa(3) & m_Pa(4), "" & RsTemp.Fields("PA26") & "," & RsTemp.Fields("PA27") & "," & RsTemp.Fields("PA28") & "," & RsTemp.Fields("PA29") & "," & RsTemp.Fields("PA30"), "" & RsTemp.Fields("PA75")) = False Then
                MsgBox MsgText(1109), vbInformation, MsgText(1110)
                txtData(1).SetFocus
                Txtdata_GotFocus 1
                GoTo JumpToExit
            End If
        End If
        'end 2019/11/01
        
        Combo1.AddItem "¤¤:" & RsTemp.Fields("PA05")
        Combo1.AddItem "­^:" & RsTemp.Fields("PA06")
        'Modified by Lydia 2022/04/25 ¡u¤é¤å¦WºÙ¡v§ï¬°¡u¥~¤å¦WºÙ¡v
        Combo1.AddItem "¥~:" & RsTemp.Fields("PA07")
        Combo1.ListIndex = 0
        Lbl3(1).Caption = "" & RsTemp.Fields("PA26")
        Lbl3(2).Caption = "" & RsTemp.Fields("CNAME")
        Lbl3(3).Caption = "" & RsTemp.Fields("PA75")
        Lbl3(4).Caption = "" & RsTemp.Fields("FNAME")
        
        'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
        'txtPath(0).Text = "\\Typing2\¹q¤l°e¥ó¼È¦s°Ï\" & m_Pa(1) & m_Pa(2)
        'txtPath(1).Text = "\\Typing2\±M§Q®×¥ó\" & Left(Val(m_Pa(2)), 3)
        txtPath(0).Text = "\\" & strTyping2Path & "\¹q¤l°e¥ó¼È¦s°Ï\" & m_Pa(1) & m_Pa(2)
        txtPath(1).Text = "\\" & strTyping2Path & "\±M§Q®×¥ó\" & Left(Val(m_Pa(2)), 3)
        'end 2024/07/22
        
        'Modified by Lydia 2018/05/09 +¨t²Î§O
        'txtPath(2).Text = Pub_GetFCPcaseFilePath(m_Pa(2), , m_Pa(1))  'Remove by Lydia 2021/12/06 (109/4/6)¤w±N\\Typing2ªº"English_Vers"©M"±M§Q®×¥ó"ªº®×¥ó¸ê®Æ§¨¡A¥þ³¡·h¨ì­ì©lÀÉ°Ï
        strCompName = PUB_FCPCaseNo2FileName(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4)) '¸g²z¡G¤£¥Î¥[²Å¸¹
        strRepName = PUB_CaseNo2FileName(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4))  ' ¸g²z¡G5½X¤]¥i¥H¡A¤W¶Ç«á¦Û°Ê´«6½X

        'Added by Lydia 2020/01/20 ±M§Q®×¥ó©MEnglish_VersÀÉ®×¡G§PÂ_ÀÉ®×¤W¶Ç¥Øªº¦a
        ' ¤w©ñ¦b­ì©lÀÉ°Ï
        If PUB_ChkCPExist(m_Pa, cnt±M§Q®×¥ó, , strExc(1), , "D") = True Then '±M§Q®×¥ó991
            txtPath(1).Text = "¡e­ì©lÀÉ°Ï¡f\±M§Q®×¥ó(" & strExc(1) & ")"
            txtPath(1).Tag = strExc(1)
        ElseIf strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× Then
            txtPath(1).Text = "¡e­ì©lÀÉ°Ï¡f"
        End If
        If PUB_ChkCPExist(m_Pa, cntEnglish_Vers, , strExc(1), , "D") = True Then 'English_Vers992
            txtPath(2).Text = "¡e­ì©lÀÉ°Ï¡f\English_Vers(" & strExc(1) & ")"
            txtPath(2).Tag = strExc(1)
        ElseIf strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× Then
            txtPath(2).Text = "¡e­ì©lÀÉ°Ï¡f"
        End If
        'end 2020/01/20
        
        Call CmdEnabled(True)
    End If
    
JumpToExit: 'Added by Lydia 2019/11/01
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim stFileName As String
Dim sFile As Variant
Dim strFlist As String
Dim bolUp As Boolean
Dim inX As Integer
Dim tmpBol As Boolean
Dim strChg As String 'Added by Lydia 2018/04/27
'Added by Lydia 2020/01/20
Dim nCP09 As String
Dim nFileName As String '¤W¶Ç«áªºÀÉ¦W

   '½T©w
   Select Case Index
          Case 0 '½T©w
                If lstAtt.ListCount = 0 Then
                   MsgBox "½Ð¥[¤Jªþ¥ó¡I", vbExclamation
                   Exit Sub
                End If
                stFileName = ""
                'Åª¨ú±M§Q®×¥óªº¥Ø«eÀÉ®×
                'Mark by Lydia 2024/02/16 ¤w§ï¨ì­ì©lÀÉ°Ï
                'If stFtpIP = "" Then stFtpIP = Pub_GetSpecMan("FTP_TYPING2")
                'If PUB_ChkFtpDirectory(stFtpIP, "//" & UniPath_¤¤»¡ & "/" & Left(Val(m_Pa(2)), 3), "R", "*.*", m_UniPathList) = False Then
                '     m_UniPathList = ""
                'End If
                'end 2024/02/16
                
                For ii = 0 To lstAtt.ListCount - 1
                    strExc(3) = ""
                     'Remove by Lydia 2018/03/29 ¤£¥Î§PÂ_ÀÉ¦W+" ("ÀÉ®×¤j¤p+¤é´Á
                    'If InStrRev(lstAtt.List(ii), " (") > 0 Then
                    '    strExc(1) = Mid(lstAtt.List(ii), 1, InStrRev(lstAtt.List(ii), " (") - 1)
                    'Else
                        strExc(1) = lstAtt.List(ii)
                    'End If
                    'end 2018/03/29
                    'ÀË¬dÀÉ®×¬O§_¦s¦b
                    If Dir(strExc(1)) = "" Then
                        'Modified by Lydia 2024/08/15 «Ø¥ß¸ê®Æ§¨
                        Pub_ChkExcelPath strExc(1)
                        If Dir(strExc(1)) = "" Then
                        'end 2024/08/15
                           MsgBox strExc(1) & "ÀÉ®×¸ô®|¤£¦s¦b !!", vbCritical
                        End If
                    End If
                    'ÀË¬dÀÉ®×¬O§_¥¿¦b¨Ï¥Î¤¤
                    If PUB_ChkFileOpening(strExc(1)) = True Then
                         MsgBox strExc(1) & vbCrLf & "ÀÉ®×¥¿¦b¨Ï¥Î¤¤¡]½ÐÃö³¬¡^¡A¤è¥iÄ~Äò¾Þ§@¡C", vbExclamation
                         Exit Sub
                    End If
                     'ÀË¬dÀÉ®×¬O§_¦³­«ÂÐ¿ï¨ú
                    If InStrRev(strExc(1), "\") > 0 Then
                        strExc(1) = Mid(strExc(1), InStrRev(strExc(1), "\") + 1)
                    End If
                    If InStr(UCase(stFileName), UCase(strExc(1))) > 0 Then
                         MsgBox "ÀÉ®×¦³­«ÂÐ¿ï¨ú !!", vbCritical
                         Exit Sub
                    End If
                    'ÀË¬d¥Øªº¦aÀÉ®×¬O§_­«½Æ
                    'Remove by Lydia 2018/04/27 ¹J¨ì­«½ÆÀÉ®×¡A°£¹q¤l°e¥ó¼È¦s°Ï¥iª½±µÂÐ»\ÀÉ®×¡A¨ä¥L\\English_vers©M±M§Q®×¥ó¸ê®Æ§¨¤@«ß§ó¦W¬°FCP0XXXXX+"-"+¤é´Á+®É¶¡+­ìÀÉ¦W
                    'strExc(3) = "": tmpBol = False
                    'If ChkIsExists(strExc(1), strExc(3), tmpBol) = True Then
                    '     If tmpBol = True Then
                    '         Exit Sub
                    '     Else
                    '         MsgBox "¤U¦CÀÉ®×»P¥Øªº¦a¸ê®Æ§¨ªºÀÉ¦W­«½Æ¡A½Ð§ó¦W¡I" & vbCrLf & strExc(3), vbCritical
                    '         Exit Sub
                    '     End If
                    'End If
                    'end 2018/04/27
                    
                    strFlist = strFlist & "&" & strExc(1)
                    stFileName = stFileName & "&" & lstAtt.List(ii)
                Next ii
                
On Error GoTo ErrHandle
                '¤W¶ÇÀÉ®×
                bolUp = True
                sFile = Empty
                sFile = Split(stFileName, "&")
                inX = -1
                strChg = ""  'Added by Lydia 2018/04/27
                For ii = 0 To UBound(sFile)
                    If Trim(sFile(ii)) <> "" Then
                       strExc(1) = sFile(ii)
                       inX = inX + 1
                       'Remove by Lydia 2018/03/29 ¤£¥Î§PÂ_ÀÉ¦W+" ("ÀÉ®×¤j¤p+¤é´Á
                       'If InStrRev(strExc(1), " (") > 0 Then
                       '    strExc(1) = Mid(strExc(1), 1, InStrRev(strExc(1), " (") - 1)
                       'End If
                       'end 2018/03/29
                       strExc(2) = strExc(1)
                       If InStrRev(strExc(2), "\") > 0 Then
                           strExc(2) = Mid(strExc(2), InStrRev(strExc(2), "\") + 1)
                       End If
                       '¬y¤ô½X5½X¸É¨ì6½X
                       If Mid(UCase(strExc(2)), 1, Len(strRepName)) = UCase(strRepName) Then
                            strExc(2) = strCompName & Mid(strExc(2), Len(strRepName) + 1)
                       End If
                       
                       'Added by Lydia 2018/04/27 ¹J¨ì*.ORI.PDF:¼u°T®§¸ß°Ý¡u¬O§_¬°¹q¤l°e¥ó¡v¡A¿ï¡¨¬O¡¨«h¤W¶Ç¨ì¹q¤l°e¥ó¼È¦s°Ï¡A¿ï¡¨§_¡¨«h¤W¶Ç¨ì\\English_vers¡C
                       'Remove by Lydia 2018/05/18 ¨ú®ø¸ß°Ý,¹w³]¬°¹q¤l°e¥ó
                       'If strChg = "" And Right(UCase(strExc(2)), Len(".ORI.PDF")) = ".ORI.PDF" Then
                       '    If MsgBox("¬O§_¬°¹q¤l°e¥ó¡H" & vbCrLf & "¿ï¡u¬O¡v¤W¶Ç¨ì¹q¤l°e¥ó¼È¦s°Ï¡A" & vbCrLf & "¿ï¡u§_¡v¤W¶Ç¨ì\\English_vers¡C", vbInformation + vbYesNo + vbDefaultButton1, "½T»{ORI.PDF") = vbYes Then
                       '       strChg = "0"
                       '    Else
                       '       strChg = "2"
                       '    End If
                       'End If
                       'end 2018/04/27
                       'end 2018/05/18
                       
                       '*.PDF©ñ¦b¹q¤l°e¥ó¼È¦s°Ï (¤uµ{®v¦³Åv­­ª½±µ½Æ»s¶K¤W)
                       'Modified by Lydia 2018/04/27 ¹q¤l°e¥óªº.ORI.PDF
                       'If Right(UCase(strExc(2)), Len(".PDF")) = ".PDF" Then
                       'Modified by Lydia 2018/05/18 ¨ú®ø¸ß°Ý,¹w³]ZIPÀÉ¡BORI.PDF©MFIX_X.PDF¬°¹q¤l°e¥ó
                       'If strChg = "0" And Right(UCase(strExc(2)), Len(".ORI.PDF")) = ".ORI.PDF" Then
                       'Modified by Lydia 2019/01/16 ±Æ°£¬Û¦üµ²ªG¤ñ¹ïÀÉ®×(.RES.PDF)
                       'If (Right(UCase(strExc(2)), Len(".PDF")) = ".PDF" And Right(UCase(strExc(2)), Len(".FIG.PDF")) <> ".FIG.PDF") Or (Right(UCase(strExc(2)), Len(".ZIP")) = ".ZIP") Then
                       'Modified by Lydia 2020/07/09 +°Ñ¦Ò¥».SEP.PDF
                       If (Right(UCase(strExc(2)), Len(".PDF")) = ".PDF" And Right(UCase(strExc(2)), Len(".FIG.PDF")) <> ".FIG.PDF" And Right(UCase(strExc(2)), Len(".RES.PDF")) <> ".RES.PDF" And Right(UCase(strExc(2)), Len(".SEP.PDF")) <> ".SEP.PDF") _
                                Or (Right(UCase(strExc(2)), Len(".ZIP")) = ".ZIP") Then

                             If Dir(txtPath(0).Text, vbDirectory) = "" Then
                                  MkDir txtPath(0).Text
                             End If
                             
                             strExc(2) = GetFinalName(strExc(2), "0") 'Added by Lydia 2018/04/27 ¨ú±o¤W¶Ç«áÀÉ¦W
                             
                             'Added by Lydia 2018/05/03 §PÂ_ÀÉ¦W­«½Æ,¸ß°Ý¬O§_¤W¶Ç
                             If strExc(2) <> "" Then
                                    FileCopy strExc(1), txtPath(0).Text & "\" & strExc(2)
                                    lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¦¨¥\)" '¦b«á­±¥[µù
                                    inX = inX + 1
                                    lstAtt.AddItem "-->" & txtPath(0).Text & "\" & strExc(2), inX
                                    'Added by Lydia 2020/02/12 +*.FIX.ORI.PDF ¦P®É¤W¶Ç¨ìEnglish_Vers
                                    If Right(UCase(strExc(2)), Len(".FIX.ORI.PDF")) = ".FIX.ORI.PDF" Then
                                        'Modified by Lydia 2020/03/18 +­ì©lÀÉ°Ï
                                        'If strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× Then
                                        If strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× Or InStr(txtPath(2).Text, "­ì©lÀÉ") > 0 Then
                                             nCP09 = txtPath(2).Tag
                                             strExc(6) = ""
                                             'English_Vers992 : ¹w³]©Ó¿ì¤H2=¾Þ§@ªÌ,­Y¦³­«ÂÐÀÉ®×¤£§R°£(A)
                                             If PUB_UploadCPFfile("2", strExc(1), m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), cntEnglish_Vers, nCP09, , "A", False, strExc(6), nFileName) = False Then
                                                  lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¥¢±Ñ)"
                                                  If strExc(6) <> "" Then
                                                      inX = inX + 1
                                                      lstAtt.AddItem "-->" & strExc(6), inX
                                                  End If
                                                  GoTo ErrHandle
                                             Else
                                                  If txtPath(2).Tag <> nCP09 Then  'ÅÜ§óÂsÄý«ö¶s
                                                       txtPath(2).Text = "¡e­ì©lÀÉ°Ï¡f\English_Vers(" & nCP09 & ")"
                                                       txtPath(2).Tag = nCP09
                                                  End If
                                                  lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¦¨¥\)"
                                                  inX = inX + 1
                                                  lstAtt.AddItem "-->" & txtPath(2).Text & "\English_Vers(" & nCP09 & ")\" & nFileName, inX
                                             End If
                                        'Mark by Lydia 2024/02/16 ¤w§ï¨ì­ì©lÀÉ°Ï
                                        'Else '­ì¥»¤W¶Ç¨ì\\Typing2\English_Vers
                                         '    If Pub_FtpPutTyping2(strExc(1), txtPath(2).Text & "\" & strExc(2)) = False Then
                                         '        GoTo ErrHandle
                                         '    Else
                                         '         lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¦¨¥\)"
                                         '         inX = inX + 1
                                         '         lstAtt.AddItem "-->" & txtPath(2).Text & "\" & strExc(2), inX
                                         '    End If
                                         'end 2024/02/16
                                        End If
                                    End If
                                    'end 2020/02/12
                                    
                             'Added by Lydia 2018/05/03
                             Else
                                     lstAtt.List(inX) = strExc(1) & " (¨ú®ø¤W¶Ç)"
                             End If 'end 2018/05/03
                             
                       'WordÀÉ¤À§O©ñ¦b±M§Q®×¥ó©MEnglish_Vers (¤uµ{®v¨S¦³Åv­­ª½±µ½Æ»s¶K¤W,¨Ï¥ÎFTP¤W¶Ç)
                       'Modified by Lydia 2018/04/27
                       'ElseIf Right(UCase(strExc(2)), Len(".DOCX")) = ".DOCX" Or Right(UCase(strExc(2)), Len(".DOC")) = ".DOC" Then
                       Else
                             'ÀÉ¦W¥u¯à¬°­^¼Æ¦r
                             strExc(2) = PUB_GetSimpleName(strExc(2))
                             '¥~¤å¥»-> English_Vers
                             'Modified by Lydia 2018/04/27 +«D¹q¤l°e¥óªº.ORI.PDF
                             'Modified by Lydia 2018/05/07 .FIX.ORIÀÉ=.ORIÀÉ
                             'If Right(UCase(strExc(2)), Len(".ORI.DOCX")) = ".ORI.DOCX" Or Right(UCase(strExc(2)), Len(".ORI.DOC")) = ".ORI.DOC" Then
                             'Modified by Lydia 2018/05/18 ¤£§t «D¹q¤l°e¥óªº.ORI.PDF
                             'If (strChg = "2" And Right(UCase(strExc(2)), Len(".ORI.PDF")) = ".ORI.PDF") _
                                    Or Right(UCase(strExc(2)), Len(".ORI.DOCX")) = ".ORI.DOCX" Or Right(UCase(strExc(2)), Len(".ORI.DOC")) = ".ORI.DOC" Then
                             'Modified by Lydia 2020/01/16 +TXTÀÉ
                             'If Right(UCase(strExc(2)), Len(".ORI.DOCX")) = ".ORI.DOCX" Or Right(UCase(strExc(2)), Len(".ORI.DOC")) = ".ORI.DOC" Then
                             If InStr(".ORI.DOC;.ORI.TXT", Right(UCase(strExc(2)), 8)) > 0 _
                                    Or InStr(".ORI.DOCX", Right(UCase(strExc(2)), 9)) > 0 Then
                            
                                  strExc(2) = GetFinalName(strExc(2), "2") 'Added by Lydia 2018/04/27 ¨ú±o¤W¶Ç«áÀÉ¦W
                                                               
                                  'Added by Lydia 2018/05/03 §PÂ_ÀÉ¦W­«½Æ,¸ß°Ý¬O§_¤W¶Ç
                                  If strExc(2) <> "" Then
                                        'Added by Lydia 2020/01/20 ±M§Q®×¥ó©MEnglish_VersÀÉ®×¡G¤W¶Ç¨ì­ì©lÀÉ°Ï
                                        'Modified by Lydia 2020/03/18 +­ì©lÀÉ°Ï
                                        'If strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× Then
                                        If strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× Or InStr(txtPath(2).Text, "­ì©lÀÉ") > 0 Then
                                             nCP09 = txtPath(2).Tag
                                             strExc(6) = ""
                                             'English_Vers992 : ¹w³]©Ó¿ì¤H2=¾Þ§@ªÌ,­Y¦³­«ÂÐÀÉ®×ª½±µ§R°£(D)
                                             If PUB_UploadCPFfile("2", strExc(1), m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), cntEnglish_Vers, nCP09, , "D", False, strExc(6), nFileName) = False Then
                                                  lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¥¢±Ñ)"
                                                  If strExc(6) <> "" Then
                                                      inX = inX + 1
                                                      lstAtt.AddItem "-->" & strExc(6), inX
                                                  End If
                                                  GoTo ErrHandle
                                             Else
                                                  If txtPath(2).Tag <> nCP09 Then  'ÅÜ§óÂsÄý«ö¶s
                                                       txtPath(2).Text = "¡e­ì©lÀÉ°Ï¡f\English_Vers(" & nCP09 & ")"
                                                       txtPath(2).Tag = nCP09
                                                  End If
                                                  lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¦¨¥\)"
                                                  inX = inX + 1
                                                  lstAtt.AddItem "-->" & txtPath(2).Text & "\English_Vers(" & nCP09 & ")\" & nFileName, inX
                                             End If
                                        'Mark by Lydia 2024/02/16 ¤w§ï¨ì­ì©lÀÉ°Ï
                                        'Else '­ì¥»¤W¶Ç¨ì\\Typing2\English_Vers
                                        '     If Pub_FtpPutTyping2(strExc(1), txtPath(2).Text & "/" & strExc(2)) = False Then
                                        ''end 2020/01/20
                                        '         GoTo ErrHandle
                                        '     Else
                                        '          lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¦¨¥\)"
                                        '          inX = inX + 1
                                        '          lstAtt.AddItem "-->" & txtPath(2).Text & "\" & strExc(2), inX
                                        '     End If
                                        'end 2024/02/16
                                        End If 'end 20/01/15
                                        
                                  'Added by Lydia 2018/05/03
                                  Else
                                       lstAtt.List(inX) = strExc(1) & " (¨ú®ø¤W¶Ç)"
                                  End If 'end 2018/05/03
                             'Added by Lydia 2018/10/22 Â½Ä¶°Ñ¦Ò¥Î¤§wordª©»¡©ú®Ñ
                             'Modified by Lydia 2018/10/25 Â½Ä¶°Ñ¦Ò¥Î¤§»¡©ú®Ñ¤£­­ÀÉ®×®æ¦¡
                             'ElseIf Right(UCase(strExc(2)), 9) = ".SEP.DOCX" Or Right(UCase(strExc(2)), 8) = ".SEP.DOC" Then
                             'Modified by Lydia 2019/01/16 ¬Û¦üµ²ªG¤ñ¹ïÀÉ®×(*.RES)
                             'ElseIf InStr(UCase(strExc(2)), ".SEP.") > 0 Then
                             ElseIf InStr(UCase(strExc(2)), ".SEP.") > 0 Or InStr(UCase(strExc(2)), ".RES.DOC") > 0 Or InStr(UCase(strExc(2)), ".RES.PDF") > 0 Then
                                  strExc(2) = GetFinalName(strExc(2), "3")
                                  If strExc(2) <> "" Then
                                        'Memo by Lydia 2020/01/20 Â½Ä¶°Ñ¦Ò¥Î¤§»¡©ú®Ñ(*.SEP.)©M¬Û¦ü¤ñ¹ïµ²ªG(*.RES)¤´ÂÂ©ñ¦bPub_GetSpecMan("FCP¬Û¦ü¤ñ¹ïµ²ªG¼È¦s")
                                        If Pub_FtpPutTyping2(strExc(1), mStrPath3 & "/" & strExc(2)) = False Then
                                             GoTo ErrHandle
                                        Else
                                              lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¦¨¥\)"
                                              inX = inX + 1
                                              lstAtt.AddItem "-->" & mStrPath3 & "\" & strExc(2), inX
                                        End If
                                        
                                        'Added by Lydia 2019/01/16 ¬Û¦üµ²ªG¤ñ¹ïÀÉ®×(*.RES) =>§ó·sÂ½Ä¶-«Ý¤ñ¹ï
                                        If InStr(UCase(strExc(2)), ".RES.DOC") > 0 Or InStr(UCase(strExc(2)), ".RES.PDF") > 0 Then
                                            strSql = "UPDATE TRANSFEE SET TF29=NULL WHERE TF01 IN " & _
                                                         "(SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & m_Pa(1) & "' AND CP02='" & m_Pa(2) & "' AND CP03='" & m_Pa(3) & "' AND CP04='" & m_Pa(4) & "' AND CP10 IN (" & GetAddStr(FcpTctPtys) & ")) "
                                            cnnConnection.Execute strSql, intI
                                        End If
                                        'end 2019/01/16
                                  Else
                                       lstAtt.List(inX) = strExc(1) & " (¨ú®ø¤W¶Ç)"
                                  End If
                             'end 2018/10/22
                
                             '¤¤»¡->±M§Q®×¥ó
                             Else
                                    'Memo by Lydia 2018/05/18 §t´À´«¥»(*.FIX.DOC, *.COR.DOC, *.DES.DOC)©M­×¥¿¥»(*.FIX_U.DOC, *.COR_U.DOC)©M¹ÏÀÉ(*.FIX.PDF)
                                    'Memo by Lydia 2020/01/16 +TXTÀÉ(*.FIX.TXT, *.COR.TXT, *.FIX_U.TXT, *.COR_U.TXT)
                                    'Memo by Lydia 2022/06/28 ³Ì²×ª©§Ç¦Cªí(.FIX.SEQ.)¡G¤ä´©TXT¡BWORD¡BXML
                                    
                                    'ÀÉ¦W¥u¯à¬°­^¼Æ¦r
                                    strExc(2) = PUB_GetSimpleName(strExc(2))
                                    strExc(2) = GetFinalName(strExc(2), "1") 'Added by Lydia 2018/04/27 ¨ú±o¤W¶Ç«áÀÉ¦W
                                    'Added by Lydia 2018/05/03 §PÂ_ÀÉ¦W­«½Æ,¸ß°Ý¬O§_¤W¶Ç
                                    If strExc(2) <> "" Then
                                        'Added by Lydia 2020/01/20 ±M§Q®×¥ó©MEnglish_VersÀÉ®×¡G¤W¶Ç¨ì­ì©lÀÉ°Ï
                                        'Modified by Lydia 2020/03/18 +­ì©lÀÉ°Ï
                                        'If strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× Then
                                        If strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× Or InStr(txtPath(1).Text, "­ì©lÀÉ") > 0 Then
                                             nCP09 = txtPath(1).Tag
                                             strExc(6) = ""
                                             '±M§Q®×¥ó991 : ¹w³]©Ó¿ì¤H2=¾Þ§@ªÌ, ´À´«¥»*.FIX©M­×¥¿¥»*.COR­Y¦³­«ÂÐÀÉ®×ª½±µ§R°£(D)
                                             If PUB_UploadCPFfile("2", strExc(1), m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), cnt±M§Q®×¥ó, nCP09, , "D", False, strExc(6), nFileName) = False Then
                                                  lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¥¢±Ñ)"
                                                  If strExc(6) <> "" Then
                                                      inX = inX + 1
                                                      lstAtt.AddItem "-->" & strExc(6), inX
                                                  End If
                                                  GoTo ErrHandle
                                             Else
                                                  If txtPath(1).Tag <> nCP09 Then  'ÅÜ§óÂsÄý«ö¶s
                                                       txtPath(1).Text = "¡e­ì©lÀÉ°Ï¡f\±M§Q®×¥ó(" & nCP09 & ")"
                                                       txtPath(1).Tag = nCP09
                                                  End If
                                                  lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¦¨¥\)"
                                                  inX = inX + 1
                                                  lstAtt.AddItem "-->" & txtPath(1).Text & "\±M§Q®×¥ó(" & nCP09 & ")\" & nFileName, inX
                                                  
                                                  '¤W¶Ç³]­p®×*.des.doc®É¡A¦Û°Ê±q\\Typing2\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\±N¬Û¹ïÀ³ªº¸ê®ÆÀÉ¤@¨Ö§ì¤J±M§Q®×¥ó°Ï¡C
                                                  If Right(UCase(strExc(2)), Len(".DES.DOCX")) = ".DES.DOCX" Or Right(UCase(strExc(2)), Len(".DES.DOC")) = ".DES.DOC" Then
                                                       'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
                                                       'strExc(7) = Dir("\\Typing2\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\fcp*" & Val(txtData(1)) & ".*")
                                                       strExc(7) = Dir("\\" & strTyping2Path & "\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\fcp*" & Val(txtData(1)) & ".*")
                                                       If strExc(7) <> "" Then
                                                            Do While strExc(7) <> ""
                                                                strExc(2) = PUB_GetSimpleName(strExc(7))
                                                                strExc(2) = GetFinalName(strExc(2), "1", False)
                                                                '¥ý¤U¸ü¨ì¥»¾÷ºÝ
                                                                'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
                                                                'FileCopy "\\Typing2\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\" & strExc(7), App.path & "\" & strUserNum & "\" & strExc(2)
                                                                FileCopy "\\" & strTyping2Path & "\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\" & strExc(7), App.path & "\" & strUserNum & "\" & strExc(2)
                                                                DoEvents
                                                                If PUB_UploadCPFfile("2", App.path & "\" & strUserNum & "\" & strExc(2), m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), cnt±M§Q®×¥ó, nCP09, , , False, strExc(6)) = False Then
                                                                     inX = inX + 1
                                                                     lstAtt.AddItem "-->" & " (¥Ó½Ð¤H¸ê®Æ-¤W¶Ç¥¢±Ñ)", inX
                                                                     If strExc(6) <> "" Then
                                                                        inX = inX + 1
                                                                        lstAtt.AddItem "-->" & strExc(6), inX
                                                                     End If
                                                                     GoTo ErrHandle
                                                                Else
                                                                     DoEvents
                                                                     Kill App.path & "\" & strUserNum & "\" & strExc(2)
                                                                     DoEvents
                                                                     'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
                                                                     'Kill "\\Typing2\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\" & strExc(7)
                                                                     Kill "\\" & strTyping2Path & "\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\" & strExc(7)
                                                                End If
                                                                'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
                                                                'strExc(7) = Dir("\\Typing2\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\fcp*" & Val(txtData(1)) & ".*")
                                                                strExc(7) = Dir("\\" & strTyping2Path & "\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\fcp*" & Val(txtData(1)) & ".*")
                                                            Loop
                                                       End If
                                                  End If
                                             End If
                                        'Mark by Lydia 2024/02/16 ¤w§ï¨ì­ì©lÀÉ°Ï
                                        'Else '­ì¥»¤W¶Ç¨ì\\Typing2\±M§Q®×¥ó
                                        ''end 2020/01/20
                                        '    '¥Ø«e¥u¦³"±M§Q®×¥ó"¯àÂà¦¨¹ïÀ³Unicode
                                        '    If Pub_FtpPutTyping2(strExc(1), "//Typing2/" & UniPath_¤¤»¡ & "/" & Left(Val(m_Pa(2)), 3) & "/" & strExc(2)) = False Then
                                        '         GoTo ErrHandle
                                        '    Else
                                         '        lstAtt.List(inX) = strExc(1) & " (¤W¶Ç¦¨¥\)"
                                         '        inX = inX + 1
                                         '        lstAtt.AddItem "-->" & txtPath(1).Text & "\" & strExc(2), inX
                                         '        'Added by Lydia 2018/05/18 ¤W¶Ç³]­p®×*.des.doc®É¡A¦Û°Ê±q\\Typing2\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\±N¬Û¹ïÀ³ªº¸ê®ÆÀÉ¤@¨Ö§ì¤J±M§Q®×¥ó°Ï¡C
                                         '        If Right(UCase(strExc(2)), Len(".DES.DOCX")) = ".DES.DOCX" Or Right(UCase(strExc(2)), Len(".DES.DOC")) = ".DES.DOC" Then
                                         '              strExc(7) = Dir("\\Typing2\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\fcp*" & Val(txtData(1)) & ".*")
                                         '              If strExc(7) <> "" Then
                                         '                   Do While strExc(7) <> ""
                                         '                       strExc(2) = PUB_GetSimpleName(strExc(7))
                                         '                       strExc(2) = GetFinalName(strExc(2), "1", False)
                                         '                       FileCopy "\\Typing2\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\" & strExc(7), App.path & "\" & strUserNum & "\" & strExc(2)
                                         '                       DoEvents
                                         '                       If Pub_FtpPutTyping2(App.path & "\" & strUserNum & "\" & strExc(2), "//Typing2/" & UniPath_¤¤»¡ & "/" & Left(Val(m_Pa(2)), 3) & "/" & strExc(2)) = False Then
                                         '                            GoTo ErrHandle
                                         '                       Else
                                         '                            DoEvents
                                         '                            Kill App.path & "\" & strUserNum & "\" & strExc(2)
                                         '                            DoEvents
                                         '                            Kill "\\Typing2\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\" & strExc(7)
                                         '                       End If
                                         '                       strExc(7) = Dir("\\Typing2\¥~±M°e¥ó\¥Ó½Ð¤H¸ê®Æ\fcp*" & Val(txtData(1)) & ".*")
                                         '                   Loop
                                         '              End If
                                         '        End If
                                          '       'end 2018/05/18
                                          '  End If
                                        End If 'Added by Lydia 2020/01/20
                                    'Added by Lydia 2018/05/03
                                    Else
                                        lstAtt.List(inX) = strExc(1) & " (¨ú®ø¤W¶Ç)"
                                    End If 'end 2018/05/03
                             End If

                       End If
                    End If
                Next
                
                MsgBox "¤W¶Ç§@·~§¹²¦¡I", vbInformation
                Call CmdEnabled(False)
          Case 1 'µ²§ô
                Unload Me
   End Select
   
   Exit Sub
   
ErrHandle:
   If Err.Number = 0 Then Exit Sub
   If bolUp = True Then
       If Right(UCase(strExc(1)), Len(".PDF")) = ".PDF" Then
             strExc(5) = txtPath(0).Text
       'Modified by Lydia 2018/04/27
       'ElseIf Right(UCase(strExc(1)), Len(".ORI.DOCX")) = ".ORI.DOCX" Or Right(UCase(strExc(1)), Len(".ORI.DOC")) = ".ORI.DOC" Then
       'Modified by Lydia 2018/10/22
       'ElseIf Right(UCase(strExc(1)), Len(".ORI.DOCX")) = ".ORI.DOCX" Or Right(UCase(strExc(1)), Len(".ORI.DOC")) = ".ORI.DOC" _
                Or Right(UCase(strExc(1)), Len(".FIX.DOCX")) = ".FIX.DOCX" Or Right(UCase(strExc(1)), Len(".FIX.DOC")) = ".FIX.DOC" Then
       'Modified by Lydia 2020/01/16 +TXTÀÉ
       'ElseIf InStr(".FIX.DOC;.COR.DOC;.DES.DOC", Right(strExc(1), 8)) > 0 _
                  Or InStr(".FIX.DOCX;.COR.DOCX;.DES.DOCX", Right(strExc(1), 9)) > 0 _
                  Or InStr(strExc(1), ".FIX_U.DOC") > 0 Or InStr(strExc(1), ".COR_U.DOC") > 0 _
                  Or InStr(strExc(1), ".FIX_U.DOCX") > 0 Or InStr(strExc(1), ".COR_U.DOCX") > 0 Then
       ElseIf InStr(".FIX.DOC;.COR.DOC;.DES.DOC;.FIX.TXT;.COR.TXT", Right(strExc(1), 8)) > 0 _
                  Or InStr(".FIX.DOCX;.COR.DOCX;.DES.DOCX", Right(strExc(1), 9)) > 0 _
                  Or InStr(".FIX_U.DOC;.COR_U.DOC", Right(strExc(1), 10)) > 0 _
                  Or InStr(".FIX_U.DOCX;.COR_U.DOCX", Right(strExc(1), 11)) > 0 Then
             
             'Added by Lydia 2020/01/20 §PÂ_¬O§_¤W¶Ç­ì©lÀÉ°Ï
             If InStr(txtPath(1), "­ì©lÀÉ") > 0 Or strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× Then
                  strExc(5) = "­ì©lÀÉ°Ï"
             Else
             'end 2020/01/20
                  strExc(5) = txtPath(1).Text
             End If 'Added by Lydia 2020/01/20
             
       'Added by Lydia 2018/10/22 Â½Ä¶°Ñ¦Ò¥Î¤§wordª©»¡©ú®Ñªº¦s©ñ¸ô®|
       'Modified by Lydia 2018/10/25
       'ElseIf Right(strExc(1), 8) = ".SEP.DOC" Or Right(strExc(1), 9) = ".SEP.DOCX" Then
       'Modified by Lydia 2019/01/16 +¬Û¦üµ²ªG¤ñ¹ïÀÉ®×(*.RES)
       ElseIf InStr(strExc(1), ".SEP.") > 0 Or InStr(strExc(1), ".RES.DOC") > 0 Or InStr(strExc(1), ".RES.PDF") > 0 Then
             strExc(5) = mStrPath3
       'end 2018/10/22
       
       'Modified by Lydia 2020/01/16 +TXTÀÉ
       ElseIf Right(UCase(strExc(1)), Len(".DOCX")) = ".DOCX" Or Right(UCase(strExc(1)), Len(".DOC")) = ".DOC" Or Right(UCase(strExc(1)), Len(".TXT")) = ".TXT" Then
             'Added by Lydia 2020/01/20 §PÂ_¬O§_¤W¶Ç­ì©lÀÉ°Ï
             If InStr(txtPath(2), "­ì©lÀÉ") > 0 Or strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× Then
                  strExc(5) = "­ì©lÀÉ°Ï"
             Else
             'end 2020/01/20
                  strExc(5) = txtPath(2).Text
             End If 'Added by Lydia 2020/01/20
       End If
       MsgBox "µLªk¼g¤J" & strExc(5) & "¡A½Ð³qª¾¹q¸£¤¤¤ß¡I", vbCritical
       inX = inX + 1
       lstAtt.AddItem "-->¤W¶Ç¥¢±Ñ:" & Err.Description, inX
       Call CmdEnabled(False)
   Else
       MsgBox Err.Description
   End If
End Sub

Private Sub CmdEnabled(ByVal bolE As Boolean)
     cmdAddDir.Enabled = bolE
     cmdAddAtt.Enabled = bolE
     cmdOpenAtt.Enabled = bolE
     cmdRemAtt.Enabled = bolE
     cmdSelect.Enabled = bolE
     cmdOK(0).Enabled = bolE
     If m_Pa(1) & m_Pa(2) & m_Pa(3) & m_Pa(4) = txtData(0) & txtData(1) & txtData(2) & txtData(3) Then
          CmdOpen(0).Enabled = True
          CmdOpen(1).Enabled = True
          CmdOpen(2).Enabled = True
     Else
          CmdOpen(0).Enabled = False
          CmdOpen(1).Enabled = False
          CmdOpen(2).Enabled = False
     End If
End Sub

Private Sub cmdOpen_Click(Index As Integer)
Dim hLocalFile As Long 'Added by Lydia 2018/06/21

On Error GoTo ErrHand01
    
    
    'Added by Lydia 2020/01/20 ¶}±Ò[­ì©lÀÉ°Ï]
    If (Index = 1 Or Index = 2) And InStr(txtPath(Index), "­ì©lÀÉ") > 0 Then
        'Added by Lydia 2020/02/26 ¥ýÀË¬d
        If Index = 1 Or Index = 2 Then
            If PUB_CheckFormExist("frm100101_M") Then
                MsgBox "½Ð¥ýÃö³¬¦@¦P¬d¸ß¡e­ì©lÀÉ°Ï¡fµe­±¡I"
                Exit Sub
            End If
        End If
        'end 2020/02/26
        If txtPath(Index).Tag = "" Then
            MsgBox m_Pa(1) & "-" & m_Pa(2) & "¦b¡e­ì©lÀÉ°Ï¡fªº" & IIf(Index = 1, "±M§Q®×¥ó", "English_Vers") & "¦¬¤å¸¹¤£¦s¦b!", vbInformation
        Else
            frm100101_M.m_strKey = txtPath(Index).Tag '¦hµ§Á`¦¬¤å¸¹
            frm100101_M.SetParent Me
            If frm100101_M.QueryData = True Then
               frm100101_M.Show
               Me.Hide
            End If
        End If
    Else
    'end 2020/01/20
        strExc(1) = "" 'Added by Lydia 2018/11/14
        If Index >= 0 And Index <= 2 Then 'Added by Lydia 2018/11/14 +¨ä¥L¸ô®|
            strExc(1) = txtPath(Index).Text
        'Added by Lydia 2018/11/14
        Else
            'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
            'strExc(1) = "\\Typing2\FCP_workflow\SIMILAR_RESULT"
            'Modified by Lydia 2024/12/31  Â½Ä¶°Ñ¦Ò¥Î¤§»¡©ú®Ñ(*.SEP.)©M¬Û¦ü¤ñ¹ïµ²ªG(*.RES)
            'strExc(1) = "\\" & strTyping2Path & "\FCP_workflow\SIMILAR_RESULT"
            strExc(1) = mStrPath3
        End If
        If strExc(1) = "" Then
            Exit Sub
        End If
        'end 2018/11/14
        'Modified by Lydia 2024/12/31 ¬d¥»¨­¥Ø¿ý + "\."
        If Dir(strExc(1) & "\.", vbDirectory) <> "" Then
             'Modified by Lydia 2018/06/21 ¥ÎÀÉ®×Á`ºÞ¶}±Ò©ñ¸m1~2¤ÀÄÁ«á,ÀÉ®×Á`ºÞ·|¥X¿ù(ex. A2037, A4041)
             'SHELL "Explorer.exe " & strExc(1), vbNormalFocus  '¶}±Ò¸ê®Æ§¨
             ShellExecute hLocalFile, "explore", strExc(1), vbNullString, vbNullString, 1
        Else
             MsgBox strExc(1) & " ¸ê®Æ§¨¤£¦s¦b ¡I", vbInformation
        End If
    End If
    Exit Sub
    
ErrHand01:
    If Err.Number <> 0 Then
         'Modified by Lydia 2018/11/14
         'MsgBox "µLªkÅª¨ú" & txtPath(Index).Text & "¡A½Ð³qª¾¹q¸£¤¤¤ß¡I", vbCritical
         MsgBox "µLªkÅª¨ú" & strExc(1) & "¡A½Ð³qª¾¹q¸£¤¤¤ß¡I", vbCritical
         Resume Next
    End If
End Sub

Private Sub Form_Load()
   txtPath(0).BackColor = &H8000000F
   txtPath(1).BackColor = &H8000000F
   txtPath(2).BackColor = &H8000000F
   
   MoveFormToCenter Me
   Call FormClear(True)
   
   mStrPath3 = Pub_GetSpecMan("FCP¬Û¦ü¤ñ¹ïµ²ªG¼È¦s") 'Added by Lydia 2018/10/22 Â½Ä¶°Ñ¦Ò¥Î¤§wordª©»¡©ú®Ñ(*.SEP)©M¬Û¦ü¤ñ¹ïµ²ªG(*.RES)©ñ¦b¤@°_
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090905 = Nothing
End Sub

'¶}±Òªþ¥ó
Private Sub cmdOpenAtt_Click()
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String
   Dim bolIsSelect As Boolean
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   
   strAtt = lstAtt.Text
   
   If strAtt = "" Then
      MsgBox "½Ð¿ï¾Ü±ý¶}±Òªºªþ¥ó¡I", vbExclamation
   Else
      For ii = 0 To lstAtt.ListCount - 1
         If lstAtt.Selected(ii) Then
            bolIsSelect = True
            stFileName = lstAtt.List(ii)
            'Remove by Lydia 2018/03/29 ¤£¥Î§PÂ_ÀÉ¦W+" ("ÀÉ®×¤j¤p+¤é´Á
            'If InStrRev(stFileName, " (") > 0 Then
            '   stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            'End If
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      Next ii
      If bolIsSelect = False Then
         MsgBox "½Ð¿ï¾Ü±ý¶}±Òªºªþ¥ó¡I", vbExclamation
      End If
   End If
   
   Screen.MousePointer = vbDefault
End Sub

'¥þ¿ï
Private Sub cmdSelect_Click()
   Dim oList As ListBox
   
   Set oList = lstAtt
   For ii = 0 To oList.ListCount - 1
      lstAtt.Selected(ii) = True
   Next
End Sub

'¥[¤J
Private Sub cmdAddAtt_Click()
   Dim stFileName As String
   Dim sFile
   Dim fs, f
   Dim strFile As String
   Dim bolMsg As Boolean, strExcept As String
   
On Error GoTo ErrHnd

   With CommonDialog1
      .CancelError = True
      .FileName = "*.*"
      .Filter = "All Files *.*|(*.*)"
      strExcept = ""
      
      '¹w³]¤W¤@¦¸ªº¸ô®|
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      bolMsg = False
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            '¦h¿ï
            sFile = Split(.FileName, ChrW$(0))
            '°O¿ý¸ô®|
             SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
             'Added by Lydia 2018/05/03 ¸ô®|±Æ°£&
             If InStr(CStr(sFile(0)), "&") > 0 Then
                  MsgBox CStr(sFile(0)) & vbCrLf & vbCrLf & "¡i&¡j²Å¸¹¬°¨t²Î«O¯d¦r¡A¤£¥i¨Ï¥Î©ó¸ô®|¡I", vbExclamation
                  Exit Sub
             End If
             'end 2018/05/03
            For ii = 1 To UBound(sFile)
               If InStr(CStr(sFile(ii)), "#") > 0 Or InStr(CStr(sFile(ii)), "&") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "¡i#©M&¡j²Å¸¹¬°¨t²Î«O¯d¦r¡A¤£¥i¨Ï¥Î©óÀÉ®×©R¦W¡I", vbExclamation
                  Exit Sub
               End If
               
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
              
               '¥²¶·¬°WordÀÉ©ÎPDFÀÉ®×
               'Modified by Lydia 2018/04/27 PDFÀÉ:­­¨îÀÉ¦W¬°FCP0XXXXX.ORI.PDF©ÎFCP0XXXXX.FIG.PDF
               'If Right(UCase(stFileName), Len(".PDF")) = ".PDF" Or Right(UCase(stFileName), Len(".DOCX")) = ".DOCX" Or Right(UCase(stFileName), Len(".DOC")) = ".DOC" Then
               'Modified by Lydia 2018/05/18 ÀË¬d°ÆÀÉ¦W
               'If Right(UCase(stFileName), Len(".ORI.PDF")) = ".ORI.PDF" Or Right(UCase(stFileName), Len(".FIG.PDF")) = ".FIG.PDF" _
                         Or Right(UCase(stFileName), Len(".DOCX")) = ".DOCX" Or Right(UCase(stFileName), Len(".DOC")) = ".DOC" Then
               'Modify By Sindy 2025/10/27 ChkAttFileName§ï¬°¦@¥Î¨ç¼Æ +, m_Pa(1), m_Pa(2)
               If ChkAttFileName(stFileName, m_Pa(1), m_Pa(2)) = True Then
                   'ÀË¬dÀÉ¦W³W«h
                   If InStrRev(stFileName, "\") > 0 Then
                       strExc(1) = Mid(stFileName, InStrRev(stFileName, "\") + 1)
                   Else
                       strExc(1) = stFileName
                   End If
                   If Mid(UCase(strExc(1)), 1, Len(strCompName)) <> UCase(strCompName) And _
                         Mid(UCase(strExc(1)), 1, Len(strRepName)) <> UCase(strRepName) Then
                       MsgBox "ÀÉ®×©R¦W¤£²Å³W©w¡A¦r­º¥²¶·¬°" & strCompName
                       Exit Sub
                   End If
                   'ÀË¬dÀÉ®×¬O§_¥¿¦b¨Ï¥Î¤¤
                   If PUB_ChkFileOpening(stFileName) = True Then
                        MsgBox stFileName & vbCrLf & "ÀÉ®×¥¿¦b¨Ï¥Î¤¤¡]½ÐÃö³¬¡^¡A¤è¥iÄ~Äò¾Þ§@¡C", vbExclamation
                        Exit Sub
                   End If
                   
                   Set fs = CreateObject("Scripting.FileSystemObject")
                   Set f = fs.GetFile(stFileName)
                   'ÀÉ®×¤j¤p¬° 0 KB ¦³»~
                   If f.Size = 0 Then
                      ShowMsg sFile(ii) & MsgText(9221)
                      Exit Sub
                   End If
                   
                   '°t¦X¥[¤J¥Ø¿ý¥u¦³Åã¥ÜÀÉ®×¦WºÙ
                   PUB_AddListX lstAtt, stFileName
               Else
                   bolMsg = True
                   strExcept = strExcept & vbCrLf & stFileName
               End If
            Next ii
            If bolMsg = True Then
                'Modified by Lydia 2018/04/27
                'MsgBox "¤U¦CÀÉ®×¤£¥i¥[¤J¡A½Ð¿ï¾ÜWordÀÉ(*.DOCX/*.DOC)©ÎPDFÀÉ®×(*.PDF) !" & strExcept, vbInformation
                'Modified by Lydia 2018/05/18 +ZIPÀÉ
                'Modified by Lydia 2018/06/08
                'MsgBox "¤U¦CÀÉ®×¤£¥i¥[¤J¡A½Ð¿ï¾ÜWordÀÉ(*.DOCX/*.DOC)¡B*.ZIPÀÉ¡B*.ORI.PDFÀÉ©Î*.FIG.PDFÀÉ®× !" & strExcept, vbInformation
                MsgBox "¤U¦CÀÉ®×°ÆÀÉ¦W¤£²Å³W«h¡A½Ð°Ñ¦Ò¸Ô²ÓÂkÀÉ»¡©ú !" & strExcept, vbInformation
            End If
         Else '³æ¿ï
             'Added by Lydia 2018/05/03 ¸ô®|±Æ°£&
             strExc(1) = Mid(.FileName, 1, InStrRev(.FileName, "\") - 1)
             If InStr(strExc(1), "&") > 0 Then
                  MsgBox strExc(1) & vbCrLf & vbCrLf & "¡i&¡j²Å¸¹¬°¨t²Î«O¯d¦r¡A¤£¥i¨Ï¥Î©ó¸ô®|¡I", vbExclamation
                  Exit Sub
             End If
             'end 2018/05/03
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Or InStr(strFile, "&") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "¡i#©M&¡j²Å¸¹¬°¨t²Î«O¯d¦r¡A¤£¥i¨Ï¥Î©óÀÉ®×©R¦W¡I", vbExclamation
               Exit Sub
            End If
            
            '°O¿ý¸ô®|
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     Exit For
                  End If
               Next ii
            End If
            
            stFileName = .FileName
            '¥²¶·¬°WordÀÉ©ÎPDFÀÉ®×
            'Modified by Lydia 2018/04/27 PDFÀÉ:­­¨îÀÉ¦W¬°FCP0XXXXX.ORI.PDF©ÎFCP0XXXXX.FIG.PDF
            'If Right(UCase(stFileName), Len(".PDF")) = ".PDF" Or Right(UCase(stFileName), Len(".DOCX")) = ".DOCX" Or Right(UCase(stFileName), Len(".DOC")) = ".DOC" Then
            'Modified by Lydia 2018/05/18 ÀË¬d°ÆÀÉ¦W
            'If Right(UCase(stFileName), Len(".ORI.PDF")) = ".ORI.PDF" Or Right(UCase(stFileName), Len(".FIG.PDF")) = ".FIG.PDF" _
                     Or Right(UCase(stFileName), Len(".DOCX")) = ".DOCX" Or Right(UCase(stFileName), Len(".DOC")) = ".DOC" Then
            'Modify By Sindy 2025/10/27 ChkAttFileName§ï¬°¦@¥Î¨ç¼Æ +, m_Pa(1), m_Pa(2)
            If ChkAttFileName(stFileName, m_Pa(1), m_Pa(2)) = True Then
                'ÀË¬dÀÉ¦W³W«h
                If InStrRev(stFileName, "\") > 0 Then
                    strExc(1) = Mid(stFileName, InStrRev(stFileName, "\") + 1)
                Else
                    strExc(1) = stFileName
                End If
                If Mid(UCase(strExc(1)), 1, Len(strCompName)) <> UCase(strCompName) And _
                      Mid(UCase(strExc(1)), 1, Len(strRepName)) <> UCase(strRepName) Then
                    MsgBox "ÀÉ®×©R¦W¤£²Å³W©w¡A¦r­º¥²¶·¬°" & strCompName
                    Exit Sub
                End If
                'ÀË¬dÀÉ®×¬O§_¥¿¦b¨Ï¥Î¤¤
                If PUB_ChkFileOpening(stFileName) = True Then
                     MsgBox stFileName & vbCrLf & "ÀÉ®×¥¿¦b¨Ï¥Î¤¤¡]½ÐÃö³¬¡^¡A¤è¥iÄ~Äò¾Þ§@¡C", vbExclamation
                     Exit Sub
                End If
                   
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set f = fs.GetFile(stFileName)
                'ÀÉ®×¤j¤p¬° 0 KB ¦³»~
                If f.Size = 0 Then
                   ShowMsg strFile & MsgText(9221)
                   Exit Sub
                End If
                '°t¦X¥[¤J¥Ø¿ý¥u¦³Åã¥ÜÀÉ®×¦WºÙ
                PUB_AddListX lstAtt, stFileName
            Else
                bolMsg = True
                'Modified by Lydia 2018/04/27
                'MsgBox "¤U¦CÀÉ®×¤£¥i¥[¤J¡A½Ð¿ï¾ÜWordÀÉ(*.DOCX/*.DOC)©ÎPDFÀÉ®×(*.PDF) !" & vbCrLf & stFileName, vbInformation
                'Modified by Lydia 2018/05/18 +ZIPÀÉ
                'Modified by Lydia 2018/06/08
                'MsgBox "¤U¦CÀÉ®×¤£¥i¥[¤J¡A½Ð¿ï¾ÜWordÀÉ(*.DOCX/*.DOC)¡B*.ZIPÀÉ¡B*.ORI.PDFÀÉ©Î*.FIG.PDFÀÉ®× !" & vbCrLf & stFileName, vbInformation
                MsgBox "¤U¦CÀÉ®×°ÆÀÉ¦W¤£²Å³W«h¡A½Ð°Ñ¦Ò¸Ô²ÓÂkÀÉ»¡©ú !" & vbCrLf & stFileName, vbInformation
            End If
         End If
      End If
   End With
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

'ÀË¬dÀÉ®×¦b¥Øªº¦a¬O§_¦s¦b
'Mark  by Lydia 2018/04/27
'Private Function ChkIsExists(ByVal pName As String, Optional ByRef pEList As String, Optional ByRef bError As Boolean) As Boolean
'Dim strA1 As String, strA2 As String
'
'On Error GoTo ErrHand01 'µLÅv­­ªº¿ù»~­n§ï°T®§
'
'    ChkIsExists = False
'    bError = False
'
'    If InStrRev(pName, "\") > 0 Then
'         strA2 = Mid(pName, InStrRev(pName, "\") + 1)
'    Else
'         strA2 = pName
'    End If
'    '¬y¤ô½X5½X¸É¨ì6½X
'    If Mid(UCase(strA2), 1, Len(strRepName)) = UCase(strRepName) Then
'        strA2 = strCompName & Mid(strA2, Len(strRepName) + 1)
'    End If
'    strA1 = PUB_GetSimpleName(strA2)  '¥h±¼«D­^¼Æ¦rªºÀÉ¦W
'
'    If Right(UCase(strA1), Len(".PDF")) = ".PDF" Then
'        '¹q¤l°e¥ó¼È¦s°Ï(¦³Åª¨úÅv­­)
'        strA1 = strA2  '¥i¤W¶Ç¤¤¤åÀÉ¦W
'        If Dir(txtPath(0).Text & "\" & strA1) <> "" Then
'             pEList = pEList & IIf(strA1 <> strA2, "­ìÀÉ¦W¡G" & strA2 & "-->", "") & "¤W¶Ç«á¡G" & strA1 & vbCrLf
'             ChkIsExists = True
'        End If
'    ElseIf Right(UCase(strA1), Len(".ORI.DOCX")) = ".ORI.DOCX" Or Right(UCase(strA1), Len(".ORI.DOC")) = ".ORI.DOC" Then
'        '¥~¤å¥»(¦³Åª¨úÅv­­)
'        If Dir(txtPath(2).Text & "\" & strA1) <> "" Then
'             pEList = pEList & IIf(strA1 <> strA2, "­ìÀÉ¦W¡G" & strA2 & "-->", "") & "¤W¶Ç«á¡G" & strA1 & vbCrLf
'             ChkIsExists = True
'        End If
'    ElseIf Right(UCase(strA1), Len(".DOCX")) = ".DOCX" Or Right(UCase(strA1), Len(".DOC")) = ".DOC" Then
'        '±M§Q®×¥ó(¨S¦³Åv­­)¡A¹w¥ýÅª¤JÀÉ®×
'        If InStr(UCase(m_UniPathList), UCase(strA1)) > 0 Then
'             pEList = pEList & IIf(strA1 <> strA2, "­ìÀÉ¦W¡G" & strA2 & "-->", "") & "¤W¶Ç«á¡G" & strA1 & vbCrLf
'             ChkIsExists = True
'        End If
'    End If
'
'    Exit Function
'
'ErrHand01:
'    If Err.Number <> 0 Then
'         bError = True
'         strExc(1) = Err.Description
'         If Right(UCase(strA1), Len(".PDF")) = ".PDF" Then
'             strExc(2) = txtPath(0).Text
'         ElseIf Right(UCase(strA1), Len(".ORI.DOCX")) = ".ORI.DOCX" Or Right(UCase(strA1), Len(".ORI.DOC")) = ".ORI.DOC" Then
'             strExc(2) = txtPath(2).Text
'         ElseIf Right(UCase(strA1), Len(".DOCX")) = ".DOCX" Or Right(UCase(strA1), Len(".DOC")) = ".DOC" Then
'             strExc(2) = txtPath(1).Text
'         End If
'         '¦]¬°³s½u¤è¦¡¦³2ºØ,©Ò¥H¥þ³¡¿ù»~°T®§²Î¤@
'         MsgBox "µLªk¼g¤J" & strExc(2) & "¡A½Ð³qª¾¹q¸£¤¤¤ß¡I", vbCritical
'         Resume Next
'    End If
'End Function

'²¾°£
Private Sub cmdRemAtt_Click()
   Call PUB_RemoveList(lstAtt)
End Sub

Private Sub SetListScroll(oList As ListBox)
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
    TextInverse txtData(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 1 Or Index = 2 Or Index = 3 Then
      KeyAscii = Pub_NumAscii(KeyAscii)
   'Added by Lydia 2024/12/20
   Else
      KeyAscii = UpperCase(KeyAscii)
   'end 2024/12/20
   End If
End Sub

Private Sub Txtdata_LostFocus(Index As Integer)
    If Index = 1 Then
        If txtData(Index).Text <> "" Then
           If Len(txtData(Index)) <> 6 Then
                 MsgBox "¥»©Ò®×¸¹½Ð¿é¤J6½X!! '"
                 txtData(Index).SetFocus
                 Txtdata_GotFocus Index
           End If
        End If
    End If
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
    If Index = 0 Then
        'Modified by Lydia 2024/12/20 +P®×
        If txtData(Index) <> "FCP" And txtData(Index) <> "P" Then
            MsgBox "¨t²Î§O½Ð¿é¤JFCP©ÎP®× !! '"
            txtData(Index).SetFocus
            Txtdata_GotFocus Index
            Cancel = True
        End If
    End If
End Sub

'Added by Lydia 2018/04/27 ÀË¬dÀÉ®×¦b¥Øªº¦a¬O§_¦s¦b,­Y¦s¦bµø±¡ªp­n§ó¦W
'Modified by Lydia 2018/05/18 +¬O§_§ó¦WbUpdate
Private Function GetFinalName(ByVal pName As String, Optional ByVal pPath As String = "", Optional ByVal bUpdate As Boolean = True) As String
Dim strA1 As String, strA2 As String, strA3 As String
Dim strMid As String
Dim intN As Integer
'Added by Lydia 2020/01/20
Dim intQ As Integer, strR1 As String
Dim rsQuery As New ADODB.Recordset

On Error GoTo ErrHand01 'µLÅv­­ªº¿ù»~­n§ï°T®§

    GetFinalName = ""
    
    If InStrRev(pName, "\") > 0 Then
         strA2 = Mid(pName, InStrRev(pName, "\") + 1)
    Else
         strA2 = pName
    End If
    '¬y¤ô½X5½X¸É¨ì6½X
    If Mid(UCase(strA2), 1, Len(strRepName)) = UCase(strRepName) Then
        strA2 = strCompName & Mid(strA2, Len(strRepName) + 1)
    End If
    strA1 = PUB_GetSimpleName(strA2)  '¥h±¼«D­^¼Æ¦rªºÀÉ¦W
    
    'Added by Lydia 2020/01/20 §ï¨ì¼Ò²Õ¤W¶ÇÀÉ®×®É,¦b§ó¦W
    If Not (strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× And txtPath(2).Tag <> "") Then
        Select Case pPath
              Case "0" '¹q¤l°e¥ó¼È¦s°Ï(¦³Åª¨úÅv­­)¡A¥iÂÐ»\
                     'Added by Lydia 2018/05/03  §PÂ_ÀÉ¦W­«½Æ,¸ß°Ý¬O§_¤W¶Ç
                     'Remove by Lydia 2018/05/18 ¨ú®ø¸ß°Ý;¹q¤l°e¥ó¼È¦s°Ï¤£§ó¦W,ª½±µÂÐ»\
                     'If Dir(txtPath(0).Text & "\" & strA1) <> "" Then
                     '    If MsgBox("­ìÀÉ¦W:" & pName & vbCrLf & "¤W¶ÇÀÉ¦W:" & strA1 & vbCrLf & "¦b" & txtPath(0).Text & "¦³­«½ÆÀÉ¦W¡A¬O§_¤W¶Ç¡H", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                     '         Exit Function
                     '    End If
                     'End If
                     'end 2018/05/03
                     'end 2018/05/18
              Case "1" '±M§Q®×¥ó(¨S¦³Åv­­¡A¹w¥ýÅª¤JÀÉ®×)¡A¤£¥iÂÐ»\
                     'Modified by Lydia 2018/05/18 ¤¤»¡´À´«¥»·|ª½±µÂÐ»\­ìÀÉ¡A¨ä¥LÀÉ®×¤W¶Ç«á·|¦b®×¸¹«á­±¥[¤W¶Ç¤é´Á
                     'If InStr(UCase(m_UniPathList), UCase(strA1)) > 0 Then
                     'Modified by Lydia 2020/01/16 +TXTÀÉ
                     'If bUpdate = False Or Right(UCase(strA1), Len(".FIX.DOCX")) = ".FIX.DOCX" Or Right(UCase(strA1), Len(".FIX.DOC")) = ".FIX.DOC" _
                            Or Right(UCase(strA1), Len(".COR.DOCX")) = ".COR.DOCX" Or Right(UCase(strA1), Len(".COR.DOC")) = ".COR.DOC" _
                            Or Right(UCase(strA1), Len(".DES.DOCX")) = ".DES.DOCX" Or Right(UCase(strA1), Len(".DES.DOC")) = ".DES.DOC" Then
                     If bUpdate = False Or _
                                    InStr(".FIX.DOC;.COR.DOC;.DES.DOC;.ORI.TXT;.FIX.TXT;.COR.TXT", Right(UCase(strA1), 8)) > 0 Or _
                                    InStr(".FIX.DOCX;.COR.DOCX;.DES.DOCX", Right(UCase(strA1), 9)) > 0 Then
    
                     Else
                          'Added by Lydia 2018/05/03  §PÂ_ÀÉ¦W­«½Æ,¸ß°Ý¬O§_¤W¶Ç
                          'Remove by Lydia 2018/05/18 ¨ú®ø¸ß°Ý
                          'If MsgBox("­ìÀÉ¦W:" & pName & vbCrLf & "¤W¶ÇÀÉ¦W:" & strA1 & vbCrLf & "¦b" & txtPath(1).Text & "¦³­«½ÆÀÉ¦W¡A¬O§_¤W¶Ç¡H", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                          '     Exit Function
                          'End If
                          'end 2018/05/03
                          'end 2018/05/18
                          strMid = "-" & strSrvDate(2) 'Modified by Lydia 2018/05/07 §ï¦¨¥Á°ê¦~
                          'Added by Lydia 2018/12/27 §PÂ_³Ì²×ª©¤¤»¡
                          If InStr(strA1, "-") > 0 Then
                              strA2 = strA1
                          Else
                          'end 2018/12/27
                              strA2 = Mid(strA1, 1, 9) & strMid & Mid(Replace(strA1, strCompName & strMid, strCompName), 10)  '¦P¤@¤Ñ+®É¶¡
                          End If
                          'end 2018/12/27
                          Do While InStr(UCase(m_UniPathList), UCase(strA2)) > 0
                                 intN = intN + 1
                                 strA3 = Format(ServerTime + intN, "000000")
                                 strA2 = Mid(strA1, 1, 9) & strMid & "-" & strA3 & Mid(Replace(strA1, strCompName & strMid, strCompName), 10)  '¦P¤@¤Ñ+®É¶¡
                          Loop
                          strA1 = strA2
                     End If
              Case "2" '¥~¤å¥»(¦³Åª¨úÅv­­)¡A¤£¥iÂÐ»\
                     'If Dir(txtPath(2).Text & "\" & strA1) <> "" Then 'Remove by Lydia 2018/05/18 ¨C¤@¦¸¤W¶Ç³£¦Û°Ê+¤é´Á,­Y¦³¦P¤@¤Ñ+®É¶¡
                          'Added by Lydia 2018/05/03  §PÂ_ÀÉ¦W­«½Æ,¸ß°Ý¬O§_¤W¶Ç
                          'Remove by Lydia 2018/05/18 ¨ú®ø¸ß°Ý
                          'If MsgBox("­ìÀÉ¦W:" & pName & vbCrLf & "¤W¶ÇÀÉ¦W:" & strA1 & vbCrLf & "¦b" & txtPath(2).Text & "¦³­«½ÆÀÉ¦W¡A¬O§_¤W¶Ç¡H", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                          '     Exit Function
                          'End If
                          'end 2018/05/03
                          'end 2018/05/18
                          strMid = "-" & strSrvDate(2) 'Modified by Lydia 2018/05/07 §ï¦¨¥Á°ê¦~
                          strA2 = Mid(strA1, 1, 9) & strMid & Mid(Replace(strA1, strCompName & strMid, strCompName), 10)  '¦P¤@¤Ñ+®É¶¡
                          'Added by Lydia 2020/01/20 ±M§Q®×¥ó©MEnglish_VersÀÉ®×¡G¤W¶Ç¨ì­ì©lÀÉ°Ï
                          If strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× And txtPath(2).Tag <> "" Then
JumpToSearch:
                                strR1 = "select cpf02 from casepaperfile where cpf01='" & txtPath(2).Tag & "' and cpf02='" & strA2 & "' "
                                intQ = 1
                                Set rsQuery = ClsLawReadRstMsg(intQ, strR1)
                                If intQ = 1 Then
                                     strA3 = Format(ServerTime + intN, "000000")
                                     strA2 = Mid(strA1, 1, 9) & strMid & "-" & strA3 & Mid(Replace(strA1, strCompName & strMid, strCompName), 10) '¦P¤@¤Ñ+®É¶¡
                                     GoTo JumpToSearch
                                End If
                                Set rsQuery = Nothing
                          Else
                          'end 2020/01/20
                              strA3 = Dir(txtPath(2).Text & "\" & strA2)
                              Do While strA3 <> ""
                                     strA3 = Format(ServerTime + intN, "000000")
                                     strA2 = Mid(strA1, 1, 9) & strMid & "-" & strA3 & Mid(Replace(strA1, strCompName & strMid, strCompName), 10) '¦P¤@¤Ñ+®É¶¡
                                     strA3 = Dir(txtPath(2).Text & "\" & strA2)
                              Loop
                          End If 'Added by Lydia 2020/01/20
                          strA1 = strA2
                     'End If  'Remove by Lydia 2018/05/18
              'Added by Lydia 2018/10/22
              Case "3"  'Â½Ä¶°Ñ¦Ò¥Î¤§wordª©»¡©ú®Ñ
              'end 2018/10/22
        End Select
    End If 'Added by Lydia 2020/01/20
    
    GetFinalName = strA1
    
    Exit Function
    
ErrHand01:
    If Err.Number <> 0 Then
         MsgBox Err.Description
         Resume Next
    End If
End Function

'Added by Lydia 2018/05/18 ¸Ô²Ó»¡©ú
Private Sub Cmd1_Click()
      frm880004.iStiu = 6
      frm880004.Show
End Sub
