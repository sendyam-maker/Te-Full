VERSION 5.00
Begin VB.Form frm060317_1 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "ÃÒ®Ñ¨ç"
   ClientHeight    =   3576
   ClientLeft      =   3096
   ClientTop       =   3252
   ClientWidth     =   4464
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3576
   ScaleWidth      =   4464
   Begin VB.Frame Frame2 
      Caption         =   "³]©w©Ó¿ì³æ¤Î©w½Z"
      Height          =   570
      Left            =   60
      TabIndex        =   17
      Top             =   2490
      Width           =   4335
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   750
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   18
         Top             =   210
         Width           =   3465
      End
      Begin VB.Label Label2 
         Caption         =   "¦Lªí¾÷"
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   19
         Top             =   225
         Width           =   765
      End
   End
   Begin VB.TextBox txtPDFPath 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   3660
      Width           =   4395
   End
   Begin VB.CheckBox Check1 
      Caption         =   "¥u¦C¦L©Ó¿ì³æ"
      Height          =   225
      Left            =   2580
      TabIndex        =   15
      Top             =   1410
      Width           =   1545
   End
   Begin VB.TextBox txtLetterDate 
      Height          =   264
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1380
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "³]©w¦a§}±ø"
      Height          =   660
      Left            =   60
      TabIndex        =   12
      Top             =   1740
      Width           =   4365
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   9
         Top             =   240
         Width           =   3450
      End
      Begin VB.Label Label2 
         Caption         =   "¦Lªí¾÷"
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   13
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3255
      TabIndex        =   11
      Top             =   90
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2250
      TabIndex        =   10
      Top             =   90
      Width           =   972
   End
   Begin VB.TextBox textPA04 
      Height          =   264
      Left            =   3144
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1020
      Width           =   375
   End
   Begin VB.TextBox textPA03 
      Height          =   264
      Left            =   2904
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1020
      Width           =   255
   End
   Begin VB.TextBox textPA02 
      Height          =   264
      Left            =   2064
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1020
      Width           =   855
   End
   Begin VB.TextBox textPA01 
      Height          =   264
      Left            =   1584
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "FCP"
      Top             =   1020
      Width           =   495
   End
   Begin VB.TextBox textPA21_2 
      Height          =   264
      Left            =   3144
      MaxLength       =   7
      TabIndex        =   1
      Top             =   660
      Width           =   1035
   End
   Begin VB.TextBox textPA21_1 
      Height          =   264
      Left            =   1584
      MaxLength       =   7
      TabIndex        =   0
      Top             =   660
      Width           =   1035
   End
   Begin VB.OptionButton optSel 
      Caption         =   "¥»©Ò®×¸¹¡G"
      CausesValidation=   0   'False
      Height          =   180
      Index           =   1
      Left            =   96
      TabIndex        =   3
      Top             =   1065
      Width           =   1455
   End
   Begin VB.OptionButton optSel 
      Caption         =   "¨Ó¨ç¦¬¤å¤é¡G"
      CausesValidation=   0   'False
      Height          =   180
      Index           =   0
      Left            =   96
      TabIndex        =   2
      Top             =   705
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "·|¨Ï¥Î¨ìªºÀÉ®×: \\typing2\fcp_workflow\patent certificate\Note(Concerning Patent Term Extension).pdf "
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   90
      TabIndex        =   20
      Top             =   3090
      Width           =   4305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "©w½Z¤é´Á¡G"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   1425
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2850
      X2              =   2970
      Y1              =   825
      Y2              =   825
   End
End
Attribute VB_Name = "frm060317_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/10/27 ¤é¤å¤w§ï©ñ©w½Z¤º
'Memo By Sindy 2022/3/2 Form2.0µe­±µLª«¥ó»Ý­×§ï (Printer¦C¦L¥¼§ï)
'Memo By Morgan 2012/12/10 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo by Morgan2010/12/27 ¥Ó½Ð®×¸¹Äæ¤w­×§ï
'2010/12/6 memo by sonia ­û¤u½s¸¹Äæ¤w­×§ï
'Memo by Morgan2010/8/16 ¤é´ÁÄæ¤w­×§ï
Option Explicit

Dim m_PA01 As String
Dim m_PA02 As String
Dim m_PA03 As String
Dim m_PA04 As String
Dim m_PA09 As String
Dim m_PA22 As String 'Add by Morgan 2005/6/23
Dim m_LetterLanguage As String
Dim m_LetterKind As Integer
Dim j As Integer, i As Integer

Dim m_ET01 As String
Dim m_ET02 As String
Dim m_ET03(1 To 3) As String '³B²zª¬ªp 1:³qª¾¨ç 2:Ä¶¤å 3:´Á­­ªí
Dim m_bolEmail As Boolean, m_bolPlusPaper As Boolean    'ADD BY SONIA 2014/4/11
'Dim strSpecNO As Boolean   'add by sonia 2014/4/28
'Dim bolSpecFax As Boolean  'add by Sindy 2016/2/24
Dim m_iCopys As Integer 'Add By Sindy 2015/11/18
Dim stPS As String 'Added by Morgan 2014/12/10
Dim strPrinter2 As String 'Add By Sindy 2015/7/16

Private Declare Function FindExecutable Lib "SHELL32.DLL" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Private Const MAX_FILENAME_LEN = 260
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const WAIT_TIMEOUT = &H102&

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
'Add By Sindy 2021/3/25 ³qª¾¨ç©Ó¿ì³æ³Æµù³]©w
Dim m_strFEB26 As String, m_strFEB20 As String, m_strFEB21 As String, m_strFEB22 As String
Dim m_pSpeed As String
'2021/3/25 END


Private Sub cmdExit_Click()
   Me.Enabled = False
   Unload Me
End Sub

Private Sub cmdok_Click()
Dim strTmp As String 'Added by Lydia 2019/06/17

   If CheckDataValid() = True Then
      If optSel(0).Value = True Then
         Screen.MousePointer = vbHourglass
         If QueryLetterData() = False Then
            Screen.MousePointer = vbDefault
            MsgBox "¨S¦³²Å¦X±ø¥óªº¸ê®Æ", vbOKOnly + vbCritical, "¬d¸ß¸ê®Æ"
            Exit Sub
         Else
            Screen.MousePointer = vbDefault
            MsgBox "§@·~§¹¦¨", vbOKOnly + vbInformation, "°õ¦æ§@·~"
         End If
         Clear
      Else
         ' ¥»©Ò®×¸¹
         m_PA01 = textPA01
         m_PA02 = textPA02
         m_PA03 = textPA03
         If IsEmptyText(m_PA03) Then m_PA03 = "0"
         m_PA04 = textPA04
         If IsEmptyText(m_PA04) Then m_PA04 = "00"
         
         'Modified by Lydia 2019/06/17
         'If IsExistRecord() = False Then
         If IsExistRecord(strTmp) = False Then
            MsgBox "¥»©Ò®×¸¹¤£¦s¦b", vbOKOnly + vbCritical, "¬d¸ß¸ê®Æ"
            Exit Sub
         'Added by Lydia 2019/06/17 ­Ó®×´£¿ô³¬¨÷
         ElseIf strTmp <> "" Then
              If MsgBox("¥»®×¤w³¬¨÷/¾P¨÷¡A¬O§_Ä~Äò¦C¦L©w½Z¡H", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
              End If
         'end 2019/06/17
         End If

         frm060316_2.SetData 0, m_PA01, True
         frm060316_2.SetData 1, m_PA02, False
         frm060316_2.SetData 2, m_PA03, False
         frm060316_2.SetData 3, m_PA04, False
         frm060316_2.SetData 5, "frm060317_1", False
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 ²M°£¬d¸ß¦Lªí°O¿ýÀÉÄæ¦ì
         frm060316_2.Show
         frm060316_2.QueryData
         'Add by Morgan 2014/12/10
         stPS = GetPS(m_PA01, m_PA02, m_PA03, m_PA04)
         If stPS <> "" Then
            frm060316_2.Combo1.AddItem stPS, 0
            frm060316_2.Combo1.ListIndex = 0
         End If
         'end 2014/12/10
         Me.Hide
      End If
   End If
End Sub

'Modified by Lydia 2019/06/17 +CaseType
Private Function IsExistRecord(ByRef CaseType As String) As Boolean
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   
   CaseType = "" 'Added by Lydia 2019/06/17
   
   IsExistRecord = False
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & m_PA01 & "' AND " & _
                  "PA02 = '" & m_PA02 & "' AND " & _
                  "PA03 = '" & m_PA03 & "' AND " & _
                  "PA04 = '" & m_PA04 & "' "
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsExistRecord = True
      CaseType = "" & rsTmp.Fields("PA57") & rsTmp.Fields("PA108")
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub Form_Load()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim ii As Integer
      
   MoveFormToCenter Me
   
   If optSel(0).Value = True Then
      EnableTextBox textPA21_1, True
      EnableTextBox textPA21_2, True
      EnableTextBox textPA01, False
      EnableTextBox textPA02, False
      EnableTextBox textPA03, False
      EnableTextBox textPA04, False
   Else
      EnableTextBox textPA21_1, False
      EnableTextBox textPA21_2, False
      EnableTextBox textPA01, True
      EnableTextBox textPA02, True
      EnableTextBox textPA03, True
      EnableTextBox textPA04, True
   End If
    
   'Modify by Morgan 2011/3/15 §ï¦@¥Î¥B¤£­n±Æ°£¹w³]¦Lªí¾÷
   PUB_SetPrinter Me.Name, Combo1
   'end 2011/3/1
   'Add By Sindy 2015/7/16
   PUB_SetPrinter Me.Name, Combo2, strPrinter2
   '2015/7/16 END

   txtLetterDate = strSrvDate(2) 'Add by Morgan 2013/4/24 ©w½Z¤é´Á
   SetFileAssociation 'Add By Sindy 2015/11/25
End Sub

'Added by Morgan 2012/1/13
Private Sub SetFileAssociation(Optional sFile As String)
Dim i As Integer, s2 As String
Dim bNewFile As Boolean, ff1 As Integer
Dim strReaderPath As String

'¹w³]¥Î Reader,§ä¤£¨ì¤~ÀË¬dÃöÁp³]©w
strReaderPath = FindFirstFileAPI("C:\Program Files\Adobe\", "AcroRd32.exe")
If strReaderPath <> "" Then txtPDFPath = strReaderPath: Exit Sub

'Check if the file exists
If sFile = "" Then
   sFile = "test.pdf"
End If

If Dir(sFile) = "" Or sFile = "" Then
   If ff1 > 0 Then Close #ff1
   ff1 = FreeFile
   Open sFile For Output As #ff1
   Close #ff1
   bNewFile = True
End If

If Dir(sFile) = "" Or sFile = "" Then
   MsgBox "ÀÉ®×¤£¦s¦b!", vbCritical, "PDF ÀÉÃöÁpÀË¬d"
   Exit Sub
End If

'Create a buffer
s2 = String(MAX_FILENAME_LEN, 32)
'Retrieve the name and handle of the executable, associated with this file
i = FindExecutable(sFile, vbNullString, s2)
If i > 32 Then
   txtPDFPath = Left$(s2, InStr(s2, Chr$(0)) - 1)
Else
   MsgBox "PDF ÀÉÃöÁp¤£¦s¦b¡A½Ð½T»{¬O§_¦³¦w¸Ë¬ÛÃöÀ³¥Îµ{¦¡ !", "PDF ÀÉÃöÁpÀË¬d"
End If

If bNewFile = True Then
   Kill sFile
End If
End Sub

'§äÀÉ®×,¦^¶Ç²Ä¤@­Ó§ä¨ìªº¸ô®|
Private Function FindFirstFileAPI(path As String, SearchStr As String) As String
    Dim FileName As String ' Walking filename variable...
    Dim DirName As String ' SubDirectory Name
    Dim dirNames() As String ' Buffer for directory name entries
    Dim nDir As Integer ' Number of directories in this path
    Dim i As Integer ' For-loop counter...
    Dim hSearch As Long ' Search Handle
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    Dim bolGotIt As Boolean
    
    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DirName = StripNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
                FindFirstFileAPI = FindFirstFileAPI & path & FileName
                bolGotIt = True
                Exit Do
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Loop
        Cont = FindClose(hSearch)
    End If
    ' If there are sub-directories...
    If nDir > 0 And bolGotIt = False Then
        ' Recursively walk into them...
        For i = 0 To nDir - 1
            FindFirstFileAPI = FindFirstFileAPI & FindFirstFileAPI(path & dirNames(i) & "\", SearchStr)
            If FindFirstFileAPI <> "" Then Exit For
        Next i
    End If
End Function

Private Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Public Sub Clear()
   textPA21_1 = Empty
   textPA21_2 = Empty
   textPA01 = "FCP"
   textPA02 = Empty
   textPA03 = Empty
   textPA04 = Empty
   If optSel(0).Value = True Then
      textPA21_1.SetFocus
   Else
      textPA02.SetFocus
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   g_LetterDate = "" 'Add by Morgan 2013/4/24
   'Copy from cmdExit_Click Morgan 2004/10/26
   '¦C¦L©w½Z¾ã§å¦C¦L²M³æ
   'Modified by Lydia 2020/09/24 +µ{¦¡¦WºÙ
   'PUB_PrintLetterList strUserNum, , Combo2, strPrinter2
   PUB_PrintLetterList strUserNum, , Combo2, strPrinter2, "and LL02='ÃÒ®Ñ¨ç' "
   '§R°£©w½Z¾ã§å¦C¦L¸ê®Æ
   'Modified by Lydia +¶Ç¤J§R°£±ø¥ó
   'PUB_DeleteLetterList strUserNum
   PUB_DeleteLetterList strUserNum, "and LL02='ÃÒ®Ñ¨ç' "
   '¦C¦L¦a§}±ø
   PUB_PrintAddressList strUserNum, Me.Combo1.Text
   '§R°£¦a§}±ø¦Cªí¸ê®Æ
   PUB_DeleteAddressList strUserNum
   'ªì©l¤Æ§Ç¸¹
   pub_AddressListSN = 0
   '­Y¦Lªí¾÷ÅÜ°Ê, «h§ó·s¦C¦L³]©w
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2004/10/26 end
   'Add by Sindy 2015/7/16
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   '2015/7/16 END
   Set frm060317_1 = Nothing
End Sub

Private Sub optSel_Click(Index As Integer)
   If optSel(0).Value = True Then
      textPA21_1.SetFocus
      EnableTextBox textPA21_1, True
      EnableTextBox textPA21_2, True
      EnableTextBox textPA01, False
      EnableTextBox textPA02, False
      EnableTextBox textPA03, False
      EnableTextBox textPA04, False
   Else
      textPA02.SetFocus
      EnableTextBox textPA21_1, False
      EnableTextBox textPA21_2, False
      EnableTextBox textPA01, True
      EnableTextBox textPA02, True
      EnableTextBox textPA03, True
      EnableTextBox textPA04, True
   End If
End Sub

Private Sub textPA03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' µoÃÒ¤é(°_)
Private Sub textPA21_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPA21_1) = False Then
      If CheckIsTaiwanDate(textPA21_1, False) = False Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "µoÃÒ¤é(°_)¤é´Á®æ¦¡¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textPA21_1_GotFocus
      End If
   End If
End Sub

' µoÃÒ¤é(¨´)
Private Sub textPA21_2_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textPA21_2) = False Then
      If CheckIsTaiwanDate(textPA21_2, False) = False Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "µoÃÒ¤é(¨´)¤é´Á®æ¦¡¤£¥¿½T"
         textPA21_2.SetFocus
         InverseTextBox textPA21_2
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      Else
         If Not ChkRange(textPA21_1, textPA21_2, "µoÃÒ¤é") Then
            textPA21_1.SetFocus
            InverseTextBox textPA21_1
         End If
      End If
   End If
End Sub

Private Sub textPA01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' ¥»©Ò®×¸¹ªº¨t²Î§O
Private Sub textPA01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textPA01) = False Then
      Select Case textPA01
         Case "FCP":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥»©Ò®×¸¹¤¤ªº¨t²Î§O¤£¥¿½T"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPA01_GotFocus
      End Select
   End If
End Sub

Private Sub textPA21_1_GotFocus()
   InverseTextBox textPA21_1
End Sub

Private Sub textPA21_2_GotFocus()
   InverseTextBox textPA21_2
End Sub

Private Sub textPA01_GotFocus()
   InverseTextBox textPA01
End Sub

Private Sub textPA02_GotFocus()
   InverseTextBox textPA02
End Sub

Private Sub textPA03_GotFocus()
   InverseTextBox textPA03
End Sub

Private Sub textPA04_GotFocus()
   InverseTextBox textPA04
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim Cancel As Boolean
   
   CheckDataValid = False
   
   ' ¿ï¶µ
   If optSel(0).Value = True Then
      ' ¨Ó¨ç¦¬¤å¤é¤£¥iªÅ¥Õ
      If IsEmptyText(textPA21_1) = True Or IsEmptyText(textPA21_2) = True Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "µoÃÒ¤é¤£¥iªÅ¥Õ"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         If IsEmptyText(textPA21_1) = True Then
            textPA21_1.SetFocus
         Else
            textPA21_2.SetFocus
         End If
         GoTo EXITSUB
      End If
      'Add By Cheng 2002/03/20
      If PUB_CheckKeyInDate(Me.textPA21_1) = -1 Then
         Me.textPA21_1.SetFocus
         textPA21_1_GotFocus
         GoTo EXITSUB
      End If
      If PUB_CheckKeyInDate(Me.textPA21_2) = -1 Then
         Me.textPA21_2.SetFocus
         textPA21_2_GotFocus
         GoTo EXITSUB
      End If
      
      ' ½d³ò
      If Val(textPA21_1) > Val(textPA21_2) Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "µoÃÒ¤é½d³ò¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textPA21_1.SetFocus
         GoTo EXITSUB
      End If
   Else
      ' ¥»©Ò®×¸¹
      If IsEmptyText(textPA01) = True Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¥»©Ò®×¸¹¨t²ÎÃþ§O¤£¥iªÅ¥Õ"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textPA01.SetFocus
         GoTo EXITSUB
      End If
      ' ¥»©Ò®×¸¹
      If IsEmptyText(textPA02) = True Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¥»©Ò®×¸¹¬y¤ô¸¹¤£¥iªÅ¥Õ"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textPA02.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Added by Morgan 2013/4/24
   g_LetterDate = DBDATE(txtLetterDate)
   If txtLetterDate <> "" Then
      txtLetterDate_Validate Cancel
      If Cancel = True Then GoTo EXITSUB
   End If
   'end 2013/4/24
   
   CheckDataValid = True
EXITSUB:
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ¨ú±o­n¦LªíªººØÃþ
' Input : strPA01 ==> ¥»©Ò®×¸¹¨t²ÎÃþ§O
'         strPA02 ==> ¥»©Ò®×¸¹¬y¤ô¸¹
'         strPA03 ==> ¥»©Ò®×¸¹
'         strPA04 ==> ¥»©Ò®×¸¹
' Output : 1 = ªí¤@¯ë
'          2 = ªí»âÃÒ¦Û°Ê¥NÃº
'          3 = ªí°l¥[Áp¦X
'          4 = ªí³Ì«á¤@¦~(¤@¯ë©Î»âÃÒ¦Û°Ê¥NÃº)
'          5 = ªí³Ì«á¤@¦~(°l¥[Áp¦X)
'          6 = ¿nÅé¹q¸ô
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetLetterKind(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String) As Integer
   
   Dim nKind As Integer
   
   nKind = 0
   
   strSql = " SELECT PA11,PA26,PA70,PA71,PA75 FROM PATENT " & _
      " WHERE PA01 = '" & strPA01 & "' AND " & _
            " PA02 = '" & strPA02 & "' AND " & _
            " PA03 = '" & strPA03 & "' AND " & _
            " PA04 = '" & strPA04 & "' "
   
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 ¤£¥Î dll ¤F  objLawDll.ReadRstMsg(intI, strSQL)
   If intI = 1 Then
      With adoRecordset
         'Added by Morgan 2012/1/4 ¿nÅé¹q¸ô
         If Mid("" & .Fields("PA11"), 4, 1) = "5" Then
            nKind = 6
            GoTo EXITSUB
         End If
         
'Removed by Morgan 2014/3/7 ©w½Z¨ú®ø,¤w¨S¦³°l¥[Áp¦X®×,­l¥Í³]­p©w½Z¦P³]­p®×
'         '­Y¥Ó½Ð®×¸¹½X¼Æ¶W¹L8½X«h¬°°l¥[Áp¦X®×
'         'Modify by Morgan 2010/12/27 ¥Ó½Ð®×¸¹§ï½X¼Æ
'         'If Len("" & .Fields("PA11")) > 8 Then
'         If Len("" & .Fields("PA11")) > 9 Then
'            nKind = 3
'            GoTo EXITSUB
'         End If
         
         '¬O§_³Ì«á¤@¦~
         strExc(0) = "SELECT MAX(NP09) FROM NEXTPROGRESS " & _
               "WHERE NP02 = '" & strPA01 & "' AND " & _
                     "NP03 = '" & strPA02 & "' AND " & _
                     "NP04 = '" & strPA03 & "' AND " & _
                     "NP05 = '" & strPA04 & "' AND NP07 = '605' AND NP06 IS NULL"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.RecordCount = 1 Then
              If IsNull(RsTemp.Fields(0)) Then
                  'Removed by Morgan 2014/3/7 ©w½Z¨ú®ø,¤w¨S¦³°l¥[Áp¦X®×,­l¥Í³]­p©w½Z¦P³]­p®×
                  'If nKind = 3 Then
                  '    nKind = 5
                  'Else
                      nKind = 4
                  'End If
              End If
            End If
         End If
         If nKind = 3 Or nKind = 4 Or nKind = 5 Then
            GoTo EXITSUB
         End If

         '¦~¶O¬O§_¦Û°Ê¥NÃº
         'Modified by Morgan 2025/8/28
         '§ï¥Î¨ç¼Æ PUB_GetAutoPay §PÂ_
         'If IsNull(.Fields("PA70")) = False Then
         '   If .Fields("PA70") = "Y" Then
         '      nKind = 2
         '      GoTo EXITSUB
         '   End If
         'End If
         '
         'If IsNull(.Fields("PA75")) = False Then
         '   If IsEmptyText(.Fields("PA75")) = False Then
         '      Dim strFA01 As String
         '      Dim strFA02 As String
         '      If Len(.Fields("PA75")) > 8 Then
         '         strFA01 = Mid(.Fields("PA75"), 1, 8)
         '         strFA02 = Mid(.Fields("PA75"), 9, 1)
         '      Else
         '         strFA01 = .Fields("PA75") & String(8 - Len(.Fields("PA75")), "0")
         '         strFA02 = "0"
         '      End If
         '
         '      strExc(0) = "SELECT FA41 FROM FAGENT " & _
         '               "WHERE FA01 = '" & strFA01 & "' AND " & _
         '                     "FA02 = '" & strFA02 & "' "
         '      intI = 1
         '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
         '      If intI = 1 Then
         '         If IsNull(RsTemp.Fields("FA41")) = False Then
         '            If RsTemp.Fields("FA41") = "Y" Then
         '               nKind = 2
         '               GoTo EXITSUB
         '             End If
         '          End If
         '       End If
         '       nKind = 1
         '       GoTo EXITSUB
         '    End If
         'End If
         '
         '' ¨ú«È¤áÀÉ
         'If IsNull(.Fields("PA26")) = False Then
         '   If IsEmptyText(.Fields("PA26")) = False Then
         '      Dim strCU01 As String
         '      Dim strCU02 As String
         '      If Len(.Fields("PA26")) > 8 Then
         '          strCU01 = Mid(.Fields("PA26"), 1, 8)
         '          strCU02 = Mid(.Fields("PA26"), 9, 1)
         '      Else
         '          strCU01 = .Fields("PA26") & String(8 - Len(.Fields("PA26")), "0")
         '          strCU02 = "0"
         '      End If
         '      strExc(0) = "SELECT CU74 FROM CUSTOMER " & _
         '                   "WHERE CU01 = '" & strCU01 & "' AND " & _
         '                         "CU02 = '" & strCU02 & "' "
         '      intI = 1
         '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
         '      If intI = 1 Then
         '         If IsNull(RsTemp.Fields("CU74")) = False Then
         '            If RsTemp.Fields("CU74") = "Y" Then
         '               nKind = 2
         '               GoTo EXITSUB
         '            End If
         '         End If
         '      End If
         '      nKind = 1
         '      GoTo EXITSUB
         '   End If
         'End If
         If PUB_GetAutoPay(strPA01, strPA02, strPA03, strPA04) = "Y" Then
            nKind = 2
         Else
            nKind = 1
         End If
         'end 2025/8/28
      End With
   End If
EXITSUB:
   If nKind = 0 Then nKind = 1
   GetLetterKind = nKind
End Function

'Modify By Sindy 2015/11/3 + Optional ByVal strPS As String = ""
Public Function QueryLetterData(Optional ByVal strPS As String = "") As Boolean
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim bFind As Boolean
   Dim strEfile As String 'Add By Sindy 2015/7/16
   Dim program_name As String
   Dim process_id As Long
   Dim process_handle As Long
   Dim kk As Integer 'Add By Sindy 2016/7/12
   Dim m_CP09 As String 'Added by Lydia 2019/03/04 ¦¬¤å¸¹
   Dim bBatch As Boolean 'Added by Lydia 2019/06/17 ¬O§_§ó·s
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 ²M°£¬d¸ß¦Lªí°O¿ýÀÉÄæ¦ì

   bFind = False
    'Modify By Cheng 2003/09/09
'   strSQL = "SELECT * FROM PATENT " & _
'            "WHERE PA01 = 'FCP' AND " & _
'                  "PA21 >= '" & DBDATE(textPA21_1) & "' AND " & _
'                  "PA21 <= '" & DBDATE(textPA21_2) & "' AND " & _
'                  "PA57 IS NULL "
   
   'Modify by Morgan 2004/7/20
   '¥[¤½§i¤é,±M§QºØÃþ, ¥Ó½Ð®×¸¹, ¥[ÃÒ®Ñ¸¹
   'MODIFY BY SONIA 2014/4/28 +PA75,PA26
   'Modify By Sindy 2015/7/16 +,GetEmailFlag(CP09) eMail,pa160
   If optSel(0).Value = True Then '¾ã§å
      pub_QL05 = pub_QL05 & ";" & optSel(0).Caption & textPA21_1 & "-" & textPA21_2 'Add By Sindy 2010/12/7
        
      'Modified by Lydia 2019/03/04 +CP09
      'Modified by Lydia 2019/06/17 +CP27,CP05,PA57,PA108
      'Modified by Morgan 2025/8/6 +PA178 Amy´ú¸Õµo²{,¦]¤w§ï¬°¿é¤J®É´N³æµ§²£¥X,¨S¦³¦A¨Ï¥Î¾ã§å©Ò¥H¨Sµo²{,¦ýÁÙ¬O¸É¤W
      strSql = "Select PA01,PA02,PA03,PA04,PA08,PA14,PA22,PA75,PA26,PA27,PA28,PA29,PA30,GetEmailFlag(CP09) eMail,pa160,CP09,CP27,CP05,PA57,PA108,PA178 From Caseprogress, Patent " & _
               " Where CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) And CP01 = 'FCP' " & _
               " And CP05>=" & DBDATE(Me.textPA21_1.Text) & " And CP05<=" & DBDATE(Me.textPA21_2.Text) & _
               " And CP10='1603' And PA21 Is Not Null And PA57 IS NULL" & _
               " order By eMail, PA01, PA02, PA03, PA04 "
   'Add By Sindy 2015/7/16
   Else '³æµ§
      m_PA01 = textPA01
      m_PA02 = textPA02
      m_PA03 = textPA03
      If IsEmptyText(m_PA03) Then m_PA03 = "0"
      m_PA04 = textPA04
      If IsEmptyText(m_PA04) Then m_PA04 = "00"
      
      pub_QL05 = pub_QL05 & ";" & optSel(1).Caption & m_PA01 & "-" & m_PA02 & "-" & m_PA03 & "-" & m_PA04
      
      'Modified by Lydia 2019/03/04 +CP09
      'Modified by Lydia 2019/06/17 ¦³¯S®í®×¥ó¤w³¬¨÷©|»Ý³ø§i¡A«hµ{§Ç·|­Ó®×¤U©w½Z¡A¨ì®É¦Û°Ê±Nµo¤å¤é"111111"®³±¼¡A¦p¦¹®×¥ó¤S¥i¦Û¤j§å¤Wµo¤å
      'strSql = "Select PA01, PA02, PA03, PA04, PA08 ,PA14 ,PA22, PA75, PA26,GetEmailFlag(CP09) eMail,pa160,CP09 From Caseprogress, Patent " & _
               " Where CP01='" & m_PA01 & "' And CP02='" & m_PA02 & "' And CP03='" & m_PA03 & "' And CP04='" & m_PA04 & "'" & _
               " And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) And CP01 = 'FCP' " & _
               " And CP10='1603' And PA21 Is Not Null And PA57 IS NULL"
      'Modified by Morgan 2023/2/9 +PA178
      strSql = "Select PA01,PA02,PA03,PA04,PA08,PA14,PA22,PA75,PA26,PA27,PA28,PA29,PA30,GetEmailFlag(CP09) eMail,pa160,CP09,CP27,CP05,PA57,PA108,PA178 From Caseprogress, Patent " & _
               " Where CP01='" & m_PA01 & "' And CP02='" & m_PA02 & "' And CP03='" & m_PA03 & "' And CP04='" & m_PA04 & "'" & _
               " And CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) And CP01 = 'FCP' " & _
               " And CP10='1603' And PA21 Is Not Null and cp09 in (select max(cp09) mno from caseprogress where CP01='" & m_PA01 & "' And CP02='" & m_PA02 & "' And CP03='" & m_PA03 & "' And CP04='" & m_PA04 & "' and cp10='1603' and cp159=0) " '§ì³Ì·s¤@¹Dªº±M§QÃÒ®Ñ
   End If
   '2015/7/16 END
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      'Add by Sindy 2015/7/16
      pub_OsPrinter = PUB_GetOsDefaultPrinter
      PUB_SetOsDefaultPrinter Combo2.Text
      PUB_SetWordActivePrinter
      PUB_RestorePrinter Combo2.Text
      '2015/7/16 END
      
      InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/12/7
      bFind = True
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         bBatch = False 'Added by Lydia 2019/06/17
         
         'Add By Sindy 2021/3/25 ³qª¾¨ç©Ó¿ì³æ³Æµù³]©w
         Call PUB_GetFcpEMPBillSpec(rsTmp.Fields("PA01") & rsTmp.Fields("PA02") & rsTmp.Fields("PA03") & rsTmp.Fields("PA04"), "03", _
                      "" & rsTmp.Fields("PA75"), "" & rsTmp.Fields("PA26") & "," & rsTmp.Fields("PA27") & "," & rsTmp.Fields("PA28") & "," & rsTmp.Fields("PA29") & "," & rsTmp.Fields("PA30"), _
                      , , m_pSpeed, , , , , , , "FEB26", m_strFEB26, m_strFEB20, m_strFEB21, m_strFEB22)
         '2021/3/25 END
         
         'Added by Morgan 2023/2/9
         If rsTmp.Fields("PA178") = "1" Then
            m_strFEB20 = "1" '¤£±HÃÒ®Ñ¥¿¥»
            m_strFEB26 = "Y" '¤£¦L¦a§}±ø
         End If
         'end 2023/2/9
         
         'Modified by Lydia 2019/03/04 §ó´«Ãþ§O¥N¸¹;¶Ç¤JCÃþ¦¬¤å¸¹
         'Call PUB_PrintFCPEmpBill(rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04"), "07")
         Call PUB_PrintFCPEmpBill(rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04"), "03", "" & rsTmp.Fields("CP09"))
         
         'Add By Sindy 2015/7/16 «D¥u¦C¦L©Ó¿ì³æ
         If Check1.Value = 0 Then
            m_iCopys = 2 '¥þ³¡²Î¤@¦L2¥÷
            If m_strFEB20 = "1" Then m_iCopys = 1 '¤£±H=¥÷¼Æ1¥÷ Add By Sindy 2021/3/29
            'Modify By Sindy 2022/3/9 Mark:¨Ï¥Î PUB_GetFcpEMPBillSpec:³qª¾¨ç©Ó¿ì³æ³Æµù³]©w
'            'Modify By Sindy 2020/12/17 §ï¦@¥Î¨ç¼Æ
'            Call frm060317_1_SpecNO("" & rsTmp.Fields("PA26"), "" & rsTmp.Fields("PA75"), m_iCopys, strSpecNO, bolSpecFax)
'            '2020/12/17 END
            
   'Modified by Morgan 2014/12/9 °t¦XPrintLetter§ï¦@¥Î
   '         '©w½Z»y¤å
   '         'Modify by Morgan 2006/6/2
   '         'm_LetterLanguage = GetLetterLanguage(rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04"))
   '         m_LetterLanguage = PUB_GetLanguage(rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04"))
   '         '©w½ZºØÃþ
   '         m_LetterKind = GetLetterKind(rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04"))
   '         'Add by Morgan 2005/6/23
   '         m_PA22 = "" & rsTmp.Fields("PA22")
   '         'Modify by Morgan 2004/7/20
   '         '¥[¶Ç±M§QºØÃþPA08,¤½§i¤éPA14°Ñ¼Æ
   '         PrintLetter rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04"), rsTmp.Fields("PA08"), Val("" & rsTmp.Fields("PA14"))
            'Modify By Sindy 2015/11/3
            If strPS <> "" Then
               stPS = strPS
            Else
            '2015/11/3 END
               stPS = GetPS(rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04")) 'Add by Morgan 2014/12/10
            End If
            PrintLetter rsTmp.Fields("PA01"), rsTmp.Fields("PA02"), rsTmp.Fields("PA03"), rsTmp.Fields("PA04"), rsTmp.Fields("PA08"), Val("" & rsTmp.Fields("PA14")), "" & rsTmp.Fields("PA22"), stPS
   'end 2014/12/9
            
            PUB_SetOsDefaultPrinter Combo2.Text
            PUB_SetWordActivePrinter
            'PUB_RestorePrinter Combo2.Text
            '2015/7/16 END
            PUB_PrintLetter rsTmp.Fields("PA01") & rsTmp.Fields("PA02") & rsTmp.Fields("PA03") & rsTmp.Fields("PA04") & "&1603" '¦C¦L³qª¾¨ç Add by Sindy 2015/7/16
             
            'Added by Lydia 2019/06/17 ¦³¯S®í®×¥ó¤w³¬¨÷©|»Ý³ø§i¡A«hµ{§Ç·|­Ó®×¤U©w½Z¡A¨ì®É¦Û°Ê±Nµo¤å¤é"111111"®³±¼¡A¦p¦¹®×¥ó¤S¥i¦Û¤j§å¤Wµo¤å
            cnnConnection.BeginTrans
                On Error GoTo ErrHandle:
                bBatch = True
                If optSel(1).Value = True And "" & rsTmp.Fields("CP27") = "19221111" And Trim("" & rsTmp.Fields("PA57") & rsTmp.Fields("PA108")) <> "" Then
                    '±Nµo¤å¤é"111111"®³±¼¨Ã¥B¤W©Ó¿ì´Á­­
                    strExc(1) = CompDate(2, 10, "" & rsTmp.Fields("cp05"))
                    strSql = "Update Caseprogress set cp27=null,cp48=" & IIf(Val(strExc(1)) < strSrvDate(1), strSrvDate(1), strExc(1)) & _
                                " where cp09='" & rsTmp.Fields("cp09") & "' and cp10='1603' "
                    Pub_SeekTbLog strSql
                    cnnConnection.Execute strSql, intI
                End If
            
                'Added by Lydia 2019/06/05 µ{§Ç¤j¶µ¤u§@¾ã§åµo¤å: ©w½Z²£¥Í®É¡A¦P®É±Nµe­±¤W¤§©w½Z¤é´Á§ó·s¦Ü¸Ó®×"±M§QÃÒ®Ñ1603"¤§CP85(FCP©w½Z¤é´Á)¡A¥H«K¤U¤@¥\¯à¥i¾ã§å¤Wµo¤å¤é
                strSql = "Update Caseprogress set cp85=" & DBDATE(txtLetterDate) & _
                         " where cp09=(select max(cp09) from Caseprogress where " & ChgCaseprogress(rsTmp.Fields("PA01") & rsTmp.Fields("PA02") & rsTmp.Fields("PA03") & rsTmp.Fields("PA04")) & " and cp10='1603' and cp27||cp57 is null)"
                cnnConnection.Execute strSql, intI
            
            cnnConnection.CommitTrans 'Added by Lydia 2019/06/17
            
            '¦C¦L©µªø±M§QªºNote
            If "" & rsTmp.Fields("pa160") = "A01N" Or "" & rsTmp.Fields("pa160") = "A61K" Then
               program_name = txtPDFPath
               
               '¦]¬°²Ä 2 ­Ó¥H«á¶}±Òªº Reader ¤~·|¦L§¹«á¦Û°ÊÃö³¬,©Ò¥H©T©w¥ý¶}¤@­ÓªÅªºµ{¦¡,¥þ³¡¦L§¹«á¦AÃö³¬
               process_id = SHELL(program_name, vbHide)
               process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
               
               'Modified by Lydia 2024/07/22 §ï¥ÎÅÜ¼Æ
               'strEfile = "\\typing2\fcp_workflow\patent certificate\Note(Concerning Patent Term Extension).pdf"
               strEfile = "\\" & strTyping2Path & "\fcp_workflow\patent certificate\Note(Concerning Patent Term Extension).pdf"
               If Dir(strEfile) <> "" Then
                  For kk = 1 To 2 'Modify By Sindy 2016/7/12 ¦C¦L2¥÷
                     mdiMain.tmrConnect.Tag = 0 'Add By Sindy 2016/7/12
                     PrintOnePdf program_name, " /n /t """ & strEfile & """ """ & Combo2 & """"
                  Next kk
               End If
               TerminateProcess process_handle, 0&
               CloseHandle process_handle
            End If
            
            'Add By Cheng 2003/01/29
            'If Not strSpecNO Then   'ADD BY SONIA 2014/4/28 °£¯S©w«È¤á/¥N²z¤H¥~¤£ºÞ¬O§_E¤Æ³£­n¦L¦W±ø
            'Modify By Sindy 2015/11/24 ­Yªâ»¡¨ú®ø
            'If Not (strSpecNO = True And m_iCopys = 1) Then 'Modify By Sindy 2015/11/18 §ïif±ø¥ó,¯S®í«È¤á¦C¦L1¥÷©w½Zªº¤£¦L¦W±ø
            '2015/11/24 END
'               'Add By Sindy 2015/9/21 ¤é¤å©w½Z¤~­n¦L¦a§}±ø
'               If m_LetterLanguage = "3" Or Val(¥~±M¶}µ¡«H¨ç±Ò¥Î¤é) >= Val(strSrvDate(1)) Then
'               '2015/9/21 END
               'Modify By Sindy 2017/9/14 ¦C¦L¤@¥÷ªº¤£¦L¦a§}±ø
               'Modify By Sindy 2017/10/6 + Y48651000 ¤£¦L¦a§}±ø
               'Modify By Sindy 2021/3/29 + m_strFEB26 <> "Y" ¤£¦L¦a§}±ø
'               If Val(m_iCopys) > 1 And ("" & rsTmp.Fields("PA75") <> "Y48651000") Then
               If m_strFEB26 <> "Y" Then
               '2017/9/14 END
                  '·s¼W¦a§}±ø¦Cªí¸ê®Æ
                  pub_AddressListSN = pub_AddressListSN + 1
                  'Modify By Cheng 2003/02/07
                  '¥[¶Ç¤Jºñ¥Ö¶K¯Èªº¥÷¼Æ
                  'PUB_AddNewAddressList strUserNum, "" & rsTmp.Fields("PA01").Value, "" & rsTmp.Fields("PA02").Value, "" & rsTmp.Fields("PA03").Value, "" & rsTmp.Fields("PA04").Value, "" & pub_AddressListSN
                  PUB_AddNewAddressList strUserNum, "" & rsTmp.Fields("PA01").Value, "" & rsTmp.Fields("PA02").Value, "" & rsTmp.Fields("PA03").Value, "" & rsTmp.Fields("PA04").Value, "" & pub_AddressListSN, "0", , m_strFEB21, m_strFEB22
               End If
            'End If '2014/4/28 ADD BY SONIA
            
            'Add By Cheng 2003/09/10
            '·s¼W¾ã§å©w½Z¦C¦L²M³æ¸ê®Æ
            PUB_AddNewLetterList "ÃÒ®Ñ¨ç", Me.textPA21_1.Text & "-" & Me.textPA21_2.Text, "" & rsTmp.Fields("PA01").Value, "" & rsTmp.Fields("PA02").Value, "" & rsTmp.Fields("PA03").Value, "" & rsTmp.Fields("PA04").Value
         End If
         rsTmp.MoveNext
      Loop
      'Add By Sindy 2015/7/16
      PUB_SetOsDefaultPrinter pub_OsPrinter
      PUB_RestorePrinter strPrinter2
      '2015/7/16 END
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/7
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   QueryLetterData = bFind
   
'Added by Lydia 2019/06/17
    Exit Function

ErrHandle:
    If Err.Number <> 0 Then
        If bBatch = True Then
            cnnConnection.RollbackTrans
        End If
        Resume Next
    End If
End Function

'Add By Sindy 2015/7/16
Private Sub PrintOnePdf(ByVal program_name As String, parameters As String)

Dim process_id As Long
Dim process_handle As Long
    ' Start the program.
    On Error GoTo ShellError
    
    process_id = SHELL(program_name & parameters, vbHide)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    
    Exit Sub

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub

' ¦C¦L©w½Z«e±N¨Ò¥~Äæ¦ì¥[¤J¨ì¦C¦L©w½Z¨Ò¥~Äæ¦ìÀÉ®×¤¤
'Modified by Morgan 2014/12/9 +strPS
Private Sub InsExpField(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String, ByVal strPA08 As String, Optional strPS As String)
   Dim strTemp As String
   
   'Added by Morgan 2014/12/9
   '¦C¦L³Æµù
   If strPS <> "" Then
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "'," & _
            "'¦C¦L³Æµù','" & IIf(m_LetterLanguage = "3", "°l¦ù¡G", "P.S. ") & ChgSQL(strPS) & "')"
      cnnConnection.Execute strSql
   End If
   'end 2014/12/9
   
   'Added by Morgan 2013/8/5
   '¤@®×¨â½Ð´£¿ô
   If strPA08 = "2" Then
      strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa11,pa77,pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
         " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and " & ChgCaseMap(strPA01 & strPA02 & strPA03 & strPA04, , 0) & _
         " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and " & ChgCaseMap(strPA01 & strPA02 & strPA03 & strPA04, , 1) & ") X" & _
         ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','¤@®×¨â½Ð·s«¬®×­n¦L','¡ð')"
         cnnConnection.Execute strSql
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','µo©ú®×¥Ó½Ð¸¹','" & RsTemp("pa11") & "')"
         cnnConnection.Execute strSql
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','µo©ú®×©¼©Ò®×¸¹','" & IIf(IsNull(RsTemp("pa77")), "", "" & RsTemp("pa77")) & "')"
         cnnConnection.Execute strSql
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','µo©ú®×¥»©Ò®×¸¹','" & RsTemp("CNo") & "')"
         cnnConnection.Execute strSql
      End If
   End If
   'end 2013/8/5
   
   '¦Ü¤U¤@µ{§ÇÀÉ¤¤§ä¤U¤@µ{§Ç¥N¸¹¬OÃº¦~¶O¤Î¬O§_Äò¿ì¬°ªÅ¡A¬O«h¤@¯ë¡A­YªÅªº«h¬O³Ì«á¤@¦¸¦~¶O
   strExc(0) = "SELECT np09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(strPA01 & strPA02 & strPA03 & strPA04) & _
      " AND NP07=" & ¦~¶O & " AND NP06 IS NULL"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields(0) <> "" Then
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','¦~¶Oªk©w´Á­­'," & CNULL(DBDATE(RsTemp.Fields(0))) & ")"
           cnnConnection.Execute strSql
           
         'Added by Morgan 2019/12/18
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
            "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','¦³¦~¶O´Á­­¤~¦L','¡ð')"
         cnnConnection.Execute strSql
         'end 2019/12/18
      End If
   End If
   
   '¤U¦¸Ãº¦~¶O¤é
   strExc(1) = CompDate(2, -1, GetPA14(strPA01 & strPA02 & strPA03 & strPA04))
   If Right(strExc(1), 4) = "0229" Then
      strExc(1) = Left(strExc(1), 4) & "0228"
   End If
               
    ' ©w½Z»y¤å
    Select Case m_LetterLanguage
       Case "1" ' ¤¤¤å
            
       Case "2" ' ­^¤å
          'Add By Sindy 2016/12/27
          'add by sonia 2014/4/11 «De¤Æ.¥[¦¹¥y³Æµù
          If Not m_bolEmail Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','PS','P.S. The original Patent Certificate will be sent by airmail with confirmation of this letter.')"
            cnnConnection.Execute strSql
          End If
          '2016/12/27 END
          
          strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
             "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','¤U¦¸Ãº¦~¶O¤é','" & GetEngMMDD(strExc(1)) & "')"
          cnnConnection.Execute strSql
          
          'Add by Morgan 2005/6/23
          strTemp = ""
          Select Case Left(m_PA22, 1)
            Case "I"
               strTemp = "Please note that the patent number includes ""I"" for ""Invention"" patent."
            Case "M"
               strTemp = "Please note that the patent number includes ""M"" for ""Utility Model"" patent."
            Case "D"
               strTemp = "Please note that the patent number includes ""D"" for ""Design"" patent."
         End Select
         If strTemp <> "" Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','±M§Q¸¹»¡©ú','" & strTemp & "')"
            cnnConnection.Execute strSql
         End If
          '2005/6/23 end
          
         If m_LetterKind = "3" Or m_LetterKind = "5" Then
            '¦Ü¤U¤@µ{§ÇÀÉ¤¤§ä¤U¤@µ{§Ç¥N¸¹¬OÃº¦~¶O¤Î¬O§_Äò¿ì¬°ªÅ¡A¬O«h¤@¯ë¡A­YªÅªº«h¬O³Ì«á¤@¦¸¦~¶O
            '§ì¥À®×¤§¤U¦¸Ãº¶O¤é
            'Modify by Morgan 2010/12/27 ¥Ó½Ð®×¸¹§ï½X¼Æ
            strExc(0) = "SELECT PA01||PA02||PA03||PA04 FROM PATENT WHERE PA01='" & strPA01 & "' AND PA11 = ( " & _
                        "SELECT SUBSTR(PA11,1,9) FROM PATENT WHERE PA01='" & strPA01 & "' AND " & _
                        "PA02='" & strPA02 & "' AND PA03='" & strPA03 & "' AND PA04='" & strPA04 & "') "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) <> "" Then
                  strTemp = RsTemp.Fields(0)
                  strExc(0) = "SELECT np09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(strTemp) & _
                     " AND NP07=" & ¦~¶O & " AND NP06 IS NULL"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp.Fields(0) <> "" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                           "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','¦~¶Oªk©w´Á­­'," & CNULL(DBDATE(RsTemp.Fields(0))) & ")"
                        cnnConnection.Execute strSql
                     End If
                  End If
               End If
            End If
         End If
                 
       Case "3" ' ¤é¤å
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
             "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','¤U¦¸Ãº¦~¶O¤é','" & strExc(1) & "')"
          cnnConnection.Execute strSql
          
         'Removed by Morgan 2022/10/26 §ï¦b©w½Z¤º¨Ã¥Î¥þ°ì¨Ò¥~Äæ¦ì±±¨î
         'If strPA08 = "2" Then
            'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','·s«¬§Þ³N³ø§i´£¥Ü','Çeþò¡BüÁ¥Î·s®×ÇU“¸§QªÌÇV¡BüÁ¥Î·s®×§Þ³Nµû’þ®ÑÇy´£¥ÜþêþùÄµ§iÇyþêþò«áþúÇQþäÇsÇW¡BþðÇU“¸§QÇy¦æ¨ÏþìÇrþæÇOþßþúþàÇeþîÇz¡C')"
            'cnnConnection.Execute strSql
            'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
               "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(2) & "','" & strUserNum & "','·s«¬§Þ³N³ø§iªk±ø','²Ä104†A¡@üÁ¥Î·s®×“¸ªÌÇV¡BüÁ¥Î·s®×§Þ³Nµû’þ®ÑÇy´£¥ÜþêþùÄµ§iÇyþêþò«áþúÇQþäÇsÇW¡B" & vbCrLf & "¡@¡@¡@¡@¡@þðÇU“¸§QÇy¦æ¨ÏþìÇrþæÇOþßþúþàÇQÆê¡C')"
            'cnnConnection.Execute strSql
         'End If
         'end 2022/10/26
    End Select
    
    'Added by Morgan 2022/10/27
    '¦~¶O¦Û°Ê¥NÃº
    'Removed by Morgan 2025/8/28 §ï§ìbasLetterªº¦@¥Î¨Ò¥~Äæ¦ì
    'If m_LetterKind = 2 Then
    '     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
    '         "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','¦~¶O¦Û°Ê¥NÃº¤£¦L','¡ð')"
    '     cnnConnection.Execute strSql
    '     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
    '         "('" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','¦~¶O¦Û°Ê¥NÃº­n¦L','¡ð')"
    '     cnnConnection.Execute strSql
    'End If
    'end 2025/8/28
    'end 2022/10/27
    
   'Added by Morgan 2014/11/5
   'Modified by Morgan 2015/9/15 Y52709 ®Ö¹ï¤w­ã±M§Q¨S¦³½Ð´Ú³æ(»âÃÒ®É¤w¥ý½Ð´Ú)
   'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
      " select '" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','®Ö¹ï¤w­ã±M§Q¤¤­n¦L','¡ð'" & _
      " from caseprogress where CP01='" & strPA01 & "' AND CP02='" & strPA02 & "' AND CP03='" & strPA03 & "' AND CP04='" & strPA04 & "'" & _
      " and cp10='926' and cp27||cp57 is null and rownum=1"
   'Modified by Morgan 2016/2/4 Y45149 Nikon ©ó¤W¶ÇÃÒ®Ñ®É¤w¤@¨Ö¤W¶Ç±M§Q¤½³ø
   'Modify By Sindy 2017/9/14 Y34210000,Y34210030 NGB CORPORATION ­n¦PNikon
   'Modify By Sindy 2018/5/22 Y51982000 ... (­n¦PNikon)
   'Modified by Lydia 2018/09/27 + Y34210010,Y34210020
   'Modified by Lydia 2019/03/19 + Y27696000,Y54770000 ©MY52709000¬Û¦P ®Ö¹ï¤w­ã±M§Q¨S¦³½Ð´Ú³æ(»âÃÒ®É¤w¥ý½Ð´Ú)
   'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
      " select '" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','®Ö¹ï¤w­ã±M§Q¤¤­n¦L'||decode(pa75,'Y52709000','2','Y45149000','3','Y34210000','3','Y34210030','3','Y51982000','3','Y34210010','3','Y34210020','3'),'¡ð'" & _
      " from caseprogress,patent where CP01='" & strPA01 & "' AND CP02='" & strPA02 & "' AND CP03='" & strPA03 & "' AND CP04='" & strPA04 & "'" & _
      " and cp10='926' and cp27||cp57 is null and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and rownum=1"
   'Modified by Lydia 2019/04/03 +Y20438®Ö¹ï¤w­ã±M§Q¶O¥Î¤w¥]§t¦b»âÃÒ¶O¤¤¡A¤G¦¸®Ö¹ï¤£¥t¥~½Ð´Ú
   
'Removed by Morgan 2019/12/18 ²¾¨ì basLetter.ExceptFieldData3 ¨Ã§ï¬°"¤G¦¸®Ö¹ï¤¤", "¤G¦¸®Ö¹ï¤¤/ªþ½Ð´Ú³æ", "¤G¦¸®Ö¹ï¤¤/ªþ¤½³ø", "¤G¦¸®Ö¹ï¤¤/ªþ¤½³ø/ªþ½Ð´Ú³æ";¥t Y45149000 ©ó 2019/03/19 ­×§ï®É¦³»~§ï¨ì,¤@¨Ö­×¥¿
   'strTemp = "'Y52709000','2','Y27696000','2','Y54770000','2','Y45149000','2','Y20438000','3','Y34210000','3','Y34210030','3','Y51982000','3','Y34210010','3','Y34210020','3'"
   'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
   '   " select '" & m_ET01 & "','" & m_ET02 & "','" & m_ET03(1) & "','" & strUserNum & "','®Ö¹ï¤w­ã±M§Q¤¤­n¦L'||decode(pa75," & strTemp & "),'¡ð'" & _
   '   " from caseprogress,patent where CP01='" & strPA01 & "' AND CP02='" & strPA02 & "' AND CP03='" & strPA03 & "' AND CP04='" & strPA04 & "'" & _
   '   " and cp10='926' and cp27||cp57 is null and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and rownum=1"
   'cnnConnection.Execute strSql, intI
'end 2019/12/18
   'end 2014/11/5
End Sub

'Modify by Morgan 2004/7/20
'¥[±M§QºØÃþPA08,¤½§i¤éPA14°Ñ¼Æ
'Modified by Morgan 2014/12/9 +strPA22,strPS
Public Sub PrintLetter(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String, ByVal strPA08 As String, ByVal lngPA14 As Long, ByVal strPA22 As String, Optional strPS As String)
   
   Dim stContent As String, intNo As Integer, strFolder As String, strFileName As String 'Added by Morgan 2014/12/8
   
   m_ET01 = "07"
   m_ET02 = strPA01 & strPA02 & strPA03 & strPA04 & "&1603"
   Erase m_ET03
   'add by sonia 2014/4/11
   m_bolEmail = False
   m_bolEmail = PUB_GetEMailFlag(strPA01 & strPA02 & strPA03 & strPA04, , , m_bolPlusPaper)
   
   'Added by Morgan 2014/12/9 °t¦X frm060316_2 ©I¥s¦Û QueryLetterData ²¾¨Ó
   '©w½Z»y¤å
   m_LetterLanguage = PUB_GetLanguage(strPA01, strPA02, strPA03, strPA04)
   '©w½ZºØÃþ
   m_LetterKind = GetLetterKind(strPA01, strPA02, strPA03, strPA04)
   m_PA22 = strPA22
   'end 2014/12/9
   
   ' ©w½Z»y¤å
   Select Case m_LetterLanguage
      ' ¤¤¤å
      Case "1":
         m_ET03(1) = "01": m_ET03(3) = "11"
      ' ­^¤å
      Case "2":
         Select Case m_LetterKind
            'Added by Morgan 2012/1/4
            Case 6: '¿nÅé¹q¸ô
               m_ET03(1) = "22"
               m_ET03(2) = "23"
               
'Modified by Morgan 2019/12/18 ©w½Z¦X¨Ö
'            Case 1:
'               'Removed by Morgan 2013/7/19 ¤wµL¾A¥Î®×¥ó,ÂÂ©w½Z§R°£
'               ''Modify by Morgan 2004/7/20
'               ''±±¨î93.7.1¥H«á¥Î·s©w½Z
'               'If lngPA14 > 0 And lngPA14 < 20040701 Then
'               '   m_ET03(1) = "02": m_ET03(2) = "10"
'               'end 2013/7/19
'
'               '·s«¬
'               If strPA08 = "2" Then
'                  m_ET03(1) = "14": m_ET03(2) = "16"
'               Else
'                  m_ET03(1) = "13": m_ET03(2) = "15"
'               End If
'               ' ´Á­­ªí
'               m_ET03(3) = "12"
'
'            Case 2:
'               'Removed by Morgan 2013/7/19 ¤wµL¾A¥Î®×¥ó,ÂÂ©w½Z§R°£
'               ''Modify by Morgan 2004/7/20
'               ''±±¨î93.7.1¥H«á¥Î·s©w½Z
'               'If lngPA14 > 0 And lngPA14 < 20040701 Then
'               '   m_ET03(1) = "03": m_ET03(2) = "10"
'               'end 2013/7/19
'
'               '·s«¬
'               If strPA08 = "2" Then
'                  m_ET03(1) = "18": m_ET03(2) = "16"
'               Else
'                  m_ET03(1) = "17": m_ET03(2) = "15"
'               End If
'               ' ´Á­­ªí
'               m_ET03(3) = "12"
''Modify by Morgan 2005/1/26 ¤£ºÞ¬O§_¦³¤U¦¸Ãº¶O¤é²Î¤@¥X·s©w½Z--David
'            Case 3, 5
'                  If strPA08 = "3" Then
'                     m_ET03(1) = "19"
'                  'Removed by Morgan 2013/7/19 ¤wµL°l¥[®×,©w½Z§R°£
'                  'Else
'                  '   m_ET03(1) = "21"
'                  'end 2013/7/19
'
'                  End If
'                  ' Ä¶¤å
'                  m_ET03(2) = "20"
''               End If
''2005/1/26 end
'            Case 4: 'µL¤U¦¸Ãº¶O¤é(¤@¯ë©Î»âÃÒ¦Û°Ê¥NÃº)
'               ' ¦C¦L©w½Z
'               m_ET03(1) = "05"
'
'               'Removed by Morgan 2013/7/19 ¤wµL¾A¥Î®×¥ó,ÂÂ©w½Z§R°£
'               ''Modify by Morgan 2004/7/20
'               ''±±¨î93.7.1¥H«á¥Î·s©w½Z
'               'If lngPA14 > 0 And lngPA14 < 20040701 Then
'               '   m_ET03(2) = "10" ' Ä¶¤å
'               'end 2013/7/19
'
'               '·s«¬
'               If strPA08 = "2" Then
'                  m_ET03(2) = "16" ' Ä¶¤å
'               Else
'                  m_ET03(2) = "15" ' Ä¶¤å
'               End If
            Case Else
               m_ET03(1) = "13"
               If strPA08 = "2" Then
                  m_ET03(2) = "16" ' Ä¶¤å
               Else
                  m_ET03(2) = "15" ' Ä¶¤å
               End If
               If m_LetterKind <> 4 Then
                  m_ET03(3) = "12" ' ´Á­­ªí
               End If
'end 2019/12/18

         End Select
      'Add by Morgan 2006/7/26
      Case "3" '¤é¤å
         'Modified by Morgan 2022/10/27 ©w½Z¦X¨Ö(­ì¦Û°Ê¥NÃº¥Î±M§QºØÃþ§PÂ_¤]¬O¿ùªº)
         'If strPA08 = "2" Then
         '   m_ET03(1) = "09"
         'Else
         '   m_ET03(1) = "06"
         'End If
         m_ET03(1) = "06"
         'end 2022/10/27
         'Ä¶¤å
         m_ET03(2) = "07"
         '´Á­­ªí
         m_ET03(3) = "08"
      Case Else:
   End Select
   
   For i = 1 To 3
      If m_ET03(i) <> "" Then
         EndLetter m_ET01, m_ET02, m_ET03(i), strUserNum
      End If
   Next
   
   InsExpField strPA01, strPA02, strPA03, strPA04, strPA08, strPS
   'Modify By Sindy 2015/11/18 ¥þ³¡¤£¥X¶Ç¯ucover page
   'add by sonia 2014/4/11 «De¤Æ¥[¶Ç¯u«Ê­±
   If Not m_bolEmail Then
      StartLetter m_ET01, m_ET02, m_ET03(1)  '§ì¶Ç¯u­¶¼Æ
      NowPrint m_ET02, m_ET01, "98", False, strUserNum, , , , , 1
   End If
   '2014/4/11 end
   'Modify By Sindy 2016/2/24 ¬Y¨Ç¯S®í«È¤á­n¥X¶Ç¯ucover page
   'Modify By Sindy 2021/3/29 + Or InStr(UCase(m_pSpeed), "FAX") > 0 Or InStr(m_pSpeed, "¶Ç¯u") > 0
   If InStr(UCase(m_pSpeed), "FAX") > 0 Or InStr(m_pSpeed, "¶Ç¯u") > 0 Then
   'If bolSpecFax = True Then
      StartLetter m_ET01, m_ET02, m_ET03(1)  '§ì¶Ç¯u­¶¼Æ
      NowPrint m_ET02, m_ET01, "98", False, strUserNum, , , , , 1
   End If
   '2016/2/24 END
   
   For i = 1 To 3
      If m_ET03(i) <> "" Then
         'Modify By Sindy 2015/11/18 + m_iCopys
         NowPrint m_ET02, m_ET01, m_ET03(i), False, strUserNum, 0, , , , m_iCopys
         intNo = i 'Added by Morgan 2014/12/8
      End If
   Next
   
   'Added by Morgan 2014/12/8
   '­^¤åE¤Æ®×¥ó­n²£¥Í"ÃÒ®Ñ¨ç+Ä¶¤å+¦~¶Oªí"ªºpdfÀÉ(\\typing2\fcp_workflow\patent certificate)
   'Modify By Sindy 2015/10/19 ¤é¤å©w½Z¤]­npdfÀÉ
   'If m_LetterLanguage = 2 And m_bolEmail = True Then
   'Modify By Sindy 2015/11/18 ¯S®í«È¤á¤~­n²£¥Í¹q¤lÀÉ
   'If m_bolEmail = True Then
'Modify By Sindy 2021/3/25 Mark
'110/3/24©M«F§®½T»{¦¹¬qµ{¦¡¥i¥H§@¼o,Run©w½ZºûÅ@ªºFC¶l¥ó§Y¥i,¤£»Ý­n¨Æ«e²£¥Í¸ê®Æ©ñ¦bServer¤W
'   If strSpecNO = True Then
'   '2015/11/18 END
'      strUserLevel = "µoFC¶l¥ó" 'Add By Sindy 2015/7/28 ³o¹q¤lÀÉ¬O­nEµ¹«È¤áªº,¦]¦¹¤£­n¥[»\Confirmationªº³¹
'      For i = 1 To intNo - 1
'         If m_ET03(i) <> "" Then
'            NowPrint m_ET02, m_ET01, m_ET03(i), False, strUserNum, , stContent, True, stContent
'         End If
'      Next
'      NowPrint m_ET02, m_ET01, m_ET03(i), True, strUserNum, , stContent, , , , , True, , , , , , , , True
'      strUserLevel = "" 'Add By Sindy 2015/7/28 ¨ú®ø
'
'      If Pub_StrUserSt03 = "M51" Then
'         strFolder = PUB_Getdesktop
'      Else
'         strFolder = "\\typing2\fcp_workflow\patent certificate"
'      End If
'      strFileName = strPA01 & strPA02 & IIf(strPA03 & strPA04 <> "000", strPA03 & strPA04, "") & "Letter(Patent Certificate)"
'      frmPDF.Show
'      frmPDF.StartProcess strFolder, strFileName
'      '¤Á´«¦Lªí¾÷
'      If PUB_PdfCreatorNameInWord = "" Then PUB_PdfCreatorNameInWord = PUB_GetCreatorNameInWord
'      g_WordAp.ActivePrinter = PUB_PdfCreatorNameInWord
'      g_WordAp.ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
'      frmPDF.EndtProcess
'      Unload frmPDF
'
'      g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
'      If g_WordAp.Documents.Count = 0 Then
'         g_WordAp.Quit wdDoNotSaveChanges
'      End If
'
'      If Me.optSel(1).Value = True Then
'         MsgBox "PDFÀÉ¤w¦s©ó " & strFolder & "¡I"
'      End If
'   End If
'   'end 2014/12/8
End Sub

'Add By Cheng 2003/01/19
'¨ú±o¤½§i¤é
Private Function GetPA14(strPA0104 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetPA14 = ""
StrSQLa = "Select * From Patent Where " & ChgPatent(strPA0104)
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetPA14 = "" & rsA("PA14").Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'92.1.17 add by sonia
Private Function GetEngMMDD(ByVal strValue As String) As String
Dim strTmp As String
Dim ii As Integer
Dim arrTmp
   
GetEngMMDD = ""
'­Y¦³¶Ç¤J­È
If strValue <> "" Then
    arrTmp = Split(strValue, "; ")
    For ii = 0 To UBound(arrTmp)
        Select Case Mid(arrTmp(ii), 5, 2)
           Case "01": strTmp = "January "
           Case "02": strTmp = "February "
           Case "03": strTmp = "March "
           Case "04": strTmp = "April "
           Case "05": strTmp = "May "
           Case "06": strTmp = "June "
           Case "07": strTmp = "July "
           Case "08": strTmp = "August "
           Case "09": strTmp = "September "
           Case "10": strTmp = "October "
           Case "11": strTmp = "November "
           Case "12": strTmp = "December "
        End Select
        GetEngMMDD = GetEngMMDD & strTmp & Right(strValue, 2)
    Next ii
Else
   GetEngMMDD = ""
End If
End Function

Private Sub txtLetterDate_GotFocus()
   CloseIme
   TextInverse txtLetterDate
End Sub

Private Sub txtLetterDate_Validate(Cancel As Boolean)
   If ChkDate(txtLetterDate) = False Then
      Cancel = True
   End If
End Sub

'2014/4/11 add by sonia ­pºâ¶Ç¯u«Ê­±ªº¶Ç¯u­¶¼Æ
Public Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
Dim iPage As Integer
Dim stContent As String

On Error GoTo ErrHnd

   '¶Ç¯u­¶¼Æ(³qª¾¨ç¨Ò¥~Äæ¦ì¼g¤J«á¤~¥iÅª¥X¥¿½T¤º®e)
   'modify by sonia 2014/4/24 ¦]¦P®É±HÃÒ®Ñ¥¿¥»¬G¦h¥[¤@­¶
   'modify by sonia 2014/4/25 +06
   Select Case ET03
      Case "06"    '¤é¤å©w½Z¤ºµL¸õ­¶²Å¸¹,¦ý©T©w¤G­¶
         iPage = 6
      Case "13", "14", "17", "18"
         iPage = 5
      Case Else
         iPage = 4
   End Select

   EndLetter ET01, ET02, "98", strUserNum
   NowPrint ET02, ET01, ET03, False, strUserNum, , , True, stContent
   If InStr(stContent, Chr(12)) > 0 Then
      iPage = iPage + 1
   End If

   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
      " values('" & ET01 & "','" & ET02 & "','98','" & strUserNum & "','¶Ç¯u­¶¼Æ','" & iPage & "')"
   cnnConnection.Execute strSql, intI

   Exit Sub

ErrHnd:
   MsgBox Err.Description, vbCritical

End Sub
'2014/4/11 end

'Added by Morgan 2014/12/10
Private Function GetPS(pa01 As String, pa02 As String, pa03 As String, pa04 As String) As String
   Dim stSQL As String, iR As Integer
   
   stSQL = "select pa75,pa26||pa27||pa28||pa29||pa30 from patent where pa01='" & pa01 & "'" & _
      " and pa02='" & pa02 & "' and pa03='" & pa03 & "' and pa04='" & pa04 & "'"
   iR = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      With AdoRecordSet3
      'Modify By Sindy 2019/9/9
      'If (InStr("" & .Fields(1), "X56038000") > 0 And "" & .Fields("pa75") <> "Y52418000") Or InStr("Y48309000,Y48309010,Y48309020,Y48309030,Y48309040,Y48309050,Y51326000,Y53363000,Y53392000,Y53112B30", "" & .Fields("pa75")) > 0 Then
      If (InStr("" & .Fields(1), "X56038000") > 0 And _
          "" & .Fields("pa75") <> "Y52418000") Or _
          InStr("Y48309020,Y53363000,Y53392000,Y53112B30", "" & .Fields("pa75")) > 0 Then
         GetPS = "Our Patent Certificate in original will be sent to your designated office separately by courier."
      '2019/9/9 END
      'Added by Morgan 2015/6/5 --ªLªÚ¦p
      ElseIf "" & .Fields("pa75") = "Y52960000" Then
         GetPS = "Our original patent certificate will be sent to CeramOptec GmbH."
      'end 2015/6/5
      End If
      End With
   End If
End Function

