VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm000003 
   BorderStyle     =   1  '單線固定
   Caption         =   "信頭維護"
   ClientHeight    =   7008
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10608
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7008
   ScaleWidth      =   10608
   Begin VB.TextBox txtAtt 
      Height          =   264
      Left            =   6816
      TabIndex        =   39
      Top             =   1224
      Width           =   900
   End
   Begin VB.CommandButton cmdToPath 
      Height          =   300
      Left            =   8760
      Picture         =   "frm000003.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   38
      Top             =   864
      Width           =   350
   End
   Begin VB.TextBox txtToPath 
      Height          =   324
      Left            =   792
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   840
      Width           =   7956
   End
   Begin VB.OptionButton Option1 
      Caption         =   "其他   副檔名"
      Height          =   225
      Index           =   4
      Left            =   5472
      TabIndex        =   34
      Top             =   1248
      Width           =   2268
   End
   Begin VB.OptionButton Option1 
      Caption         =   "請作單"
      Height          =   225
      Index           =   3
      Left            =   4296
      TabIndex        =   17
      Top             =   1248
      Width           =   1068
   End
   Begin VB.OptionButton Option1 
      Caption         =   "M31 財務處"
      Height          =   225
      Index           =   2
      Left            =   2880
      TabIndex        =   16
      Top             =   1236
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "刪除(&D)"
      Enabled         =   0   'False
      Height          =   390
      Index           =   4
      Left            =   2412
      TabIndex        =   13
      Top             =   45
      Width           =   750
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5460
      Left            =   48
      TabIndex        =   6
      Top             =   1524
      Width           =   8616
      _ExtentX        =   15198
      _ExtentY        =   9631
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "信頭"
      TabPicture(0)   =   "frm000003.frx":0102
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "請作單"
      TabPicture(1)   =   "frm000003.frx":011E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WebBrowser1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "測試"
      TabPicture(2)   =   "frm000003.frx":013A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdTest(2)"
      Tab(2).Control(1)=   "cmdTest(1)"
      Tab(2).Control(2)=   "txtTest(0)"
      Tab(2).Control(3)=   "txtTest(1)"
      Tab(2).Control(4)=   "cmdTest(0)"
      Tab(2).Control(5)=   "MSHFlexGrid1"
      Tab(2).Control(6)=   "TextBox5"
      Tab(2).Control(7)=   "TextBox4"
      Tab(2).Control(8)=   "TextBox3"
      Tab(2).Control(9)=   "TextBox2"
      Tab(2).Control(10)=   "ComboBox1"
      Tab(2).Control(11)=   "CommandButton2"
      Tab(2).Control(12)=   "CommandButton1"
      Tab(2).Control(13)=   "TextBox1"
      Tab(2).Control(14)=   "lblTest(0)"
      Tab(2).Control(15)=   "lblTest(1)"
      Tab(2).ControlCount=   16
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
         Height          =   525
         Index           =   2
         Left            =   -68970
         TabIndex        =   32
         Top             =   4620
         Width           =   2085
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "下載檔案"
         Height          =   375
         Index           =   1
         Left            =   -68190
         TabIndex        =   24
         Top             =   300
         Width           =   1515
      End
      Begin VB.TextBox txtTest 
         Height          =   285
         Index           =   0
         Left            =   -74010
         TabIndex        =   21
         Top             =   420
         Width           =   1305
      End
      Begin VB.TextBox txtTest 
         Height          =   285
         Index           =   1
         Left            =   -73710
         TabIndex        =   20
         Top             =   750
         Width           =   5505
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "讀取檔案清單"
         Height          =   375
         Index           =   0
         Left            =   -68160
         TabIndex        =   19
         Top             =   690
         Width           =   1515
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   60
         Width           =   615
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   4980
         Left            =   90
         ScaleHeight     =   4932
         ScaleWidth      =   8364
         TabIndex        =   7
         Top             =   420
         Width           =   8412
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   5175
         Left            =   -74955
         TabIndex        =   8
         Top             =   450
         Width           =   8475
         ExtentX         =   14949
         ExtentY         =   9128
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2175
         Left            =   -74670
         TabIndex        =   18
         Top             =   1200
         Width           =   3030
         _ExtentX        =   5355
         _ExtentY        =   3831
         _Version        =   393216
         Cols            =   4
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "V|檔案名稱|大小|最後修改時間"
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin MSForms.TextBox TextBox5 
         Height          =   2205
         Left            =   -71520
         TabIndex        =   33
         Top             =   1200
         Width           =   4725
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "8334;3889"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox4 
         Height          =   1185
         Left            =   -71010
         TabIndex        =   31
         Top             =   4050
         Width           =   1905
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "3360;2090"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox3 
         Height          =   615
         Left            =   -72840
         TabIndex        =   30
         Top             =   4590
         Width           =   1815
         VariousPropertyBits=   -1467987941
         Size            =   "3201;1085"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextBox2 
         Height          =   615
         Left            =   -74670
         TabIndex        =   29
         Top             =   4590
         Width           =   1815
         VariousPropertyBits=   -1467987941
         Size            =   "3201;1085"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox ComboBox1 
         Height          =   345
         Left            =   -74670
         TabIndex        =   28
         Top             =   4110
         Width           =   3645
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "6429;609"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.CommandButton CommandButton2 
         Height          =   405
         Left            =   -69000
         TabIndex        =   27
         Top             =   4080
         Width           =   2085
         Caption         =   "存檔Unicode檢查"
         Size            =   "3678;714"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   405
         Left            =   -69030
         TabIndex        =   26
         Top             =   3570
         Width           =   2085
         Caption         =   "Unicode ToolTips 測試"
         Size            =   "3678;714"
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   405
         Left            =   -74670
         TabIndex        =   25
         Top             =   3570
         Width           =   5535
         VariousPropertyBits=   746604571
         Size            =   "9763;714"
         Value           =   "請貼上Unicode文字後按右邊測試，再將滑鼠移到Grid看結果"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblTest 
         AutoSize        =   -1  'True
         Caption         =   "FTP IP:"
         Height          =   180
         Index           =   0
         Left            =   -74610
         TabIndex        =   23
         Top             =   480
         Width           =   525
      End
      Begin VB.Label lblTest 
         AutoSize        =   -1  'True
         Caption         =   "FTP Folder:"
         Height          =   180
         Index           =   1
         Left            =   -74610
         TabIndex        =   22
         Top             =   810
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代碼:"
         Height          =   180
         Left            =   180
         TabIndex        =   14
         Top             =   90
         Width           =   405
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "下載不預覽"
      Height          =   225
      Left            =   7944
      TabIndex        =   12
      Top             =   1248
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "M21 人事"
      Height          =   225
      Index           =   1
      Left            =   1680
      TabIndex        =   10
      Top             =   1236
      Width           =   1155
   End
   Begin VB.OptionButton Option1 
      Caption         =   "M51 電腦中心"
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   9
      Top             =   1236
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   8760
      Picture         =   "frm000003.frx":0156
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   504
      Width           =   350
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9072
      Top             =   3936
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "另存"
      Height          =   390
      Index           =   6
      Left            =   84
      TabIndex        =   4
      Top             =   45
      Width           =   660
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下載(&R)"
      Height          =   390
      Index           =   3
      Left            =   1632
      TabIndex        =   3
      Top             =   45
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "離開(&X)"
      Height          =   390
      Index           =   2
      Left            =   3216
      TabIndex        =   2
      Top             =   45
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "上傳(&S)"
      Height          =   390
      Index           =   0
      Left            =   816
      TabIndex        =   0
      Top             =   45
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Height          =   324
      Left            =   792
      TabIndex        =   1
      Top             =   492
      Width           =   7956
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AD認證測試"
      Height          =   525
      Left            =   9240
      TabIndex        =   11
      Top             =   5670
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblToPath 
      AutoSize        =   -1  'True
      Caption         =   "下載到:"
      Height          =   180
      Left            =   144
      TabIndex        =   36
      Top             =   912
      Width           =   588
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "來源檔:"
      Height          =   180
      Left            =   144
      TabIndex        =   35
      Top             =   576
      Width           =   588
   End
   Begin VB.Image Image1 
      Height          =   1032
      Left            =   8640
      OLEDropMode     =   1  '手動
      Top             =   1560
      Width           =   1860
   End
End
Attribute VB_Name = "frm000003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Memo by Morgan 2022/1/5 改成Form2.0 (無)
'Created by Morgan 2015/6/26
Option Explicit

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    
Const cpUTF8 = 65001
Const cpBig5 = 950

Private Const Base64Char As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="

'Added by Morgan 2024/9/30
'大寫鍵切換
Private Declare Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer
Private Declare Function MapVirtualKey Lib "user32" _
    Alias "MapVirtualKeyA" _
    (ByVal uCode As Long, ByVal uMapType As Long) As Long
Private Declare Function SendInput Lib "user32" _
    (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long

Private Type KeyboardInput       '   typedef struct tagINPUT {
   dwType As Long                '     DWORD type;
   wVK As Integer                '     union {MOUSEINPUT mi;
   wScan As Integer              '            KEYBDINPUT ki;
   dwFlags As Long               '            HARDWAREINPUT hi;
   dwTime As Long                '     };
   dwExtraInfo As Long           '   }INPUT, *PINPUT;
   dwPadding As Currency         '
End Type

'SendInput constants
Private Const INPUT_KEYBOARD As Long = 1
Private Const KEYEVENTF_KEYUP As Long = 2

Private Const VK_CAPITAL = &H14
'end 2024/9/30

Function MultiByteToUTF16(utf8() As Byte, CodePage As Long) As String
    Dim bufSize As Long
    bufSize = MultiByteToWideChar(CodePage, 0&, utf8(0), UBound(utf8) + 1, 0, 0)
    MultiByteToUTF16 = Space(bufSize)
    MultiByteToWideChar CodePage, 0&, utf8(0), UBound(utf8) + 1, StrPtr(MultiByteToUTF16), bufSize
End Function
 
Function UTF16ToMultiByte(UTF16 As String, CodePage As Long) As Byte()
    Dim bufSize As Long
    Dim arr() As Byte
    Dim ii As Integer, strCode As String
    
    bufSize = WideCharToMultiByte(CodePage, 0&, StrPtr(UTF16), Len(UTF16), 0, 0, 0, 0)
    ReDim arr(bufSize - 1)
    WideCharToMultiByte CodePage, 0&, StrPtr(UTF16), Len(UTF16), arr(0), bufSize, 0, 0
    UTF16ToMultiByte = arr
'    strCode = bufSize & ":"
'    For ii = 0 To bufSize - 1
'      strCode = strCode & arr(ii) & "+"
'   Next
'  Debug.Print strCode
End Function

'    MsgBox MultiByteToUTF16(UTF16ToMultiByte("ab中,c", cpUTF8), cpUTF8)

Private Sub cmdTest_Click(Index As Integer)
   Select Case Index
   Case 0
      GetFtpList
   Case 1
      GetFtpFile
   Case 2
      Test
   End Select
End Sub

Private Sub GetFtpFile()
'   Dim hConnection As Long
'   Dim sFtpIP As String, sDirRemote As String, sFileName As String
'   Dim hFile As Long
'   Dim byteBuffer(102399) As Byte
'   Dim ReadyBuffer() As Byte
'
'   sFtpIP = txtTest(0)
'   sDirRemote = txtTest(1)
'   'sFileName = "TS001828_小時光麵館_2020-10-17_2021-02-19_207.tmsearch.docx"
'   sFileName = "TS001828_來一客ONE MORE CUP_2020-10-17_2021-02-19_210.tmsearch.docx"
'
'   hConnection = InternetConnect(hOpen, sFtpIP, FTP_Port, _
'               "pgmid", "pgmpwd", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
'   bRet = FtpSetCurrentDirectory(hConnection, sDirRemote)
'   hFile = FtpOpenFile(hConnection, sFileName, &H80000000, INTERNET_FLAG_TRANSFER_BINARY + &H80000000, 0)
'   Open App.path & "\" & sFileName For Binary As F1
'   bDoLoop = True
'   While bDoLoop
'       bDoLoop = InternetReadFileByte(hFile, VarPtr(byteBuffer(0)), sBuffer, lNumberOfBytesRead)
'       If Not CBool(lNumberOfBytesRead) Then
'           bDoLoop = False
'       Else
'           ReDim ReadyBuffer(lNumberOfBytesRead - 1) As Byte
'           If lNumberOfBytesRead <> sBuffer Then
'               For oIjk = 0 To (lNumberOfBytesRead - 1)
'                   ReadyBuffer(oIjk) = byteBuffer(oIjk)
'               Next oIjk
'           Else
'               ReadyBuffer = byteBuffer
'           End If
'           Put #F1, , ReadyBuffer
'       End If
'   Wend
'   Close F1
'
'   InternetCloseHandle hFile
'   Erase byteBuffer
'   Erase ReadyBuffer
End Sub

Private Sub GetFtpList()
'   Dim strExcelPath As String
'   Dim hLocalFile As Long
'
'   SetGrid True
'   'txtTest(1) = "//M51-1/VOLUME/CASEPAPERPDF/TS001/TS001828/API"
'   BrowseFtpFolder txtTest(1), , , txtTest(0)
'   If MSHFlexGrid1.Rows > 2 Then
'      If MSHFlexGrid1.TextMatrix(1, 1) = "" Then
'         MSHFlexGrid1.RemoveItem 1
'      End If
'   End If
'   '開檔案總管
'   strExcelPath = Replace(txtTest(1), "/", "\")
'   strExcelPath = Replace(strExcelPath, "\\", "\\RN524X\")
'   ShellExecute hLocalFile, "explore", strExcelPath, vbNullString, vbNullString, 1
End Sub

Private Sub cmdToPath_Click()
   Dim fName As String, strStartFolder As String
   
   If Dir(txtToPath & "\", vbDirectory) <> "" Then strStartFolder = txtToPath
   fName = PUB_GetFolder(Me.hWnd, strStartFolder, "請選取資料夾:")
   If fName <> "" Then
      txtToPath = fName
   End If
End Sub

Private Sub Command1_Click(Index As Integer)
   Static stFileName As String
   Static stFileName1 As String
   Static stFileName2 As String
   Dim stFileTemp As String
   Dim sReturn As String
   Dim arrTmp1 As Variant 'Added by Lydia 2020/03/24
   Dim hLocalFile As Long
   Dim bolCapsLock As Boolean 'Added by Morgan 2024/9/30
   
   'Added by Morgan 2020/6/11 +M31
   Dim strIBF01 As String, strIBF02 As String, strIBF03 As String, strIBF04 As String, strIBF05 As String
   Dim arrTmp() As String
   
   If Option1(0).Value = True Or Option1(3).Value = True Then
      strIBF01 = "M51"
   ElseIf Option1(1).Value = True Then
      strIBF01 = "M21"
   ElseIf Option1(2).Value = True Then
      strIBF01 = "M31"
      
   'Added by Morgan 2024/3/27
   Else
      strIBF01 = ""
      'Check1.Value = vbChecked
   End If
   'end 2020/6/11
   
On Error GoTo ErrHnd

   Select Case Index
      Case 0
         If Text1.Text = "Text1" Or Text1.Text = "" Then
              MsgBox "請指定檔案路徑! "
              Exit Sub
         ElseIf Dir(Text1.Text) = "" Then
              MsgBox "請指定檔案路徑! "
              Exit Sub
         End If
         
         If Option1(3).Value = True Then
            intI = InStrRev(Text1, "\")
            sReturn = InputBox("請輸入請作單號", , Mid(Text1, intI + 1, InStrRev(Text1, ".") - intI - 1))
         
         'Added by Morgan 2024/9/30
         ElseIf Option1(4).Value = True Then
            intI = InStrRev(Text1, "\")
            strExc(0) = Mid(Text1, intI + 1, InStrRev(Text1, ".") - intI - 1)
            sReturn = InputBox("請輸入代碼XXX-XXXXXX-X-XX", , strExc(0))
         Else
            'Modified by Lydia 2020/03/24  加第3,4碼
            'sReturn = InputBox("請輸入代碼", , "1")
            sReturn = InputBox("請輸入代碼(IBF02)，若有IBF03請用XX-X，若有IBF04請用XX-X-XX", , "1")
            'end 2020/03/24
         End If
         
         If sReturn <> "" Then
            If Option1(3).Value = True Then
               strIBF01 = Mid(sReturn, 1, 3)
               strIBF02 = Mid(sReturn, 4)
               'Modified by Morgan 2023/9/25 +檢查公告紀錄存在才能上傳
               strExc(0) = "select * from PGMBulletin where bu01=" & (Left(sReturn, 7) + 19110000) & " and bu02=" & Mid(sReturn, 8)
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If Save2DB(Text1, strIBF02, "0", "00", "5", "5", strIBF01) = True Then
                     MsgBox "存檔完成！"
                  End If
               ElseIf intI = 0 Then
                  MsgBox "請作單號 " & sReturn & " 並無公告紀錄！", vbExclamation
               End If
               
            'Added by Morgan 2024/9/30
            ElseIf Option1(4).Value = True Then
               If InStr(sReturn, "-") = 0 Then
                  MsgBox "代碼輸入錯誤!!", vbCritical
                  Exit Sub
               End If
               arrTmp1 = Split(sReturn, "-")
               strIBF01 = arrTmp1(0)
               strIBF02 = arrTmp1(1)
               strIBF03 = "0"
               strIBF04 = "00"
               If UBound(arrTmp1) > 1 Then
                  strIBF03 = arrTmp1(2)
                  If UBound(arrTmp1) > 2 Then
                     strIBF04 = arrTmp1(3)
                  End If
               End If
               strIBF05 = "4"
            
               If Save2DB(Text1, strIBF02, strIBF03, strIBF04, strIBF05, , strIBF01) = True Then
               
                  MsgBox "存檔完成！"
               End If
            'end 2024/9/30
            
            Else
               'Modified by Lydia 2020/03/24
               'If Save2DB(Text1, Val(sReturn), , , , , IIf(Option1(1).Value = True, "M21", "M51")) = True Then
               arrTmp1 = Split(sReturn, "-")
               strExc(3) = "0": strExc(4) = "00"
               If UBound(arrTmp1) > 0 Then
                    For intI = 1 To UBound(arrTmp1)
                        If arrTmp1(intI) <> "" Then
                            If intI = 1 Then
                               strExc(3) = arrTmp1(intI)
                            ElseIf intI = 2 Then
                               strExc(4) = arrTmp1(intI)
                            End If
                        End If
                    Next intI
               End If
               If Save2DB(Text1, "" & arrTmp1(0), strExc(3), strExc(4), , , strIBF01) = True Then
               'end 2020/03/24
                  MsgBox "存檔完成！"
               End If
            End If
         End If
      
      '望遠鏡(開啟)
      Case 1
         
         'cd1.Filter = "All files|*.*|Bitmap files (*.bmp)|*.bmp|GIF files (*.gif)|*.gif|JPEG files (*.jpg)|*.jpg|PNG files (*.png)|*.png|TIFF files (*.tif)|*.tif|WMF files (*.wmf)|*.wmf|PDF files (*.pdf)"
         cd1.Filter = "All files|*.*|Picture (*.bmp;*.gif;*.jpg;*.png;*.tif;*.wmf)|*.bmp;*.gif;*.jpg;*.png;*.tif;*.wmf|PDF (*.pdf)|*.pdf|Word (*.doc;*.docx)|*.doc;*.docx"
         cd1.FilterIndex = 0
         
         'Added by Morgan 2024/10/22
         strExc(0) = ""
         intI = InStrRev(Text1, "\")
         If intI > 0 Then
            strExc(0) = Left(Text1, intI - 1)
            If Dir(strExc(0), vbDirectory) = "" Then
               strExc(0) = ""
            End If
         End If
         If strExc(0) <> "" Then
            cd1.InitDir = strExc(0)
         Else
            cd1.InitDir = PUB_Getdesktop
         End If
         'end 2024/10/22
         
         cd1.ShowOpen
         If Trim(cd1.FileName) <> "" Then
            Screen.MousePointer = vbHourglass
            Text1 = cd1.FileName
            'Set Image1.Picture = LoadPicture(cd1.FileName)
            'Me.Width = Picture1.Width + 200
            'Me.Height = Picture1.Height + 1000
            If InStr(UCase("*.bmp;*.gif;*.JPEG;*.jpg;*.png;*.tif;*.wmf;"), UCase("*." & Mid(Text1, InStrRev(Text1, ".") + 1) & ";")) > 0 Then
               Text2 = "": Command1(4).Enabled = False 'Added by Morgan 2020/4/1
               SSTab1.Tab = 0
               Set Picture1.Picture = LoadPicture(cd1.FileName)
               If Picture1.Width + 200 > 12795 Then
                  Me.Width = Picture1.Width + 280
               End If
               Me.Height = Picture1.Height + 1600
               SSTab1.Width = Me.Width - 150
               SSTab1.Height = Me.Height - 1000
            
            ElseIf UCase(Right(Text1, 4)) = UCase(".pdf") Then
               Option1(3).Value = True
               SSTab1.Tab = 1
               SSTab1.Width = WebBrowser1.Width + 150
               SSTab1.Height = WebBrowser1.Height + 600
               Me.Width = SSTab1.Width + 150
               Me.Height = SSTab1.Height + 1600
               WebBrowser1.Navigate cd1.FileName
               
            End If
            Screen.MousePointer = vbDefault
         End If
         
      Case 2
         Unload Me
         
      Case 3
         If Option1(4).Value = True Then
            strExc(0) = "請輸入代碼!!" & vbCrLf & vbCrLf & "格式:XXX-XXXXXX-X-XX" & vbCrLf & "範例:M51-000400-0-01"
         
         ElseIf Option1(3).Value = True Then
            strExc(0) = "請輸入請作單號!!"
            
         Else
            strExc(0) = "請輸入代碼!!"
         End If
         
         bolCapsLock = CapsLock
         If bolCapsLock = False Then PressCapsLock
         Do
            sReturn = InputBox(strExc(0), , "0")
         Loop While (sReturn = "0")
         If bolCapsLock = False Then PressCapsLock
         
         If sReturn <> "" Then
            strIBF03 = "0"
            strIBF04 = "00"
            
            If InStr(sReturn, "-") > 0 Then
               arrTmp = Split(sReturn, "-")
               'Added by Morgan 2024/3/27
               If strIBF01 = "" Then
                  strIBF01 = arrTmp(0)
                  strIBF02 = arrTmp(1)
                  If UBound(arrTmp) > 1 Then
                     strIBF03 = arrTmp(2)
                     If UBound(arrTmp) > 2 Then
                        strIBF04 = arrTmp(3)
                     End If
                  End If
               Else
               'end 2024/3/27
                  strIBF02 = arrTmp(0)
                  strIBF03 = arrTmp(1)
                  If UBound(arrTmp) > 1 Then
                     strIBF04 = arrTmp(2)
                  End If
               End If
            
            'Added by Morgan 2024/10/22
            ElseIf Option1(3).Value = True Then
               strIBF01 = Left(sReturn, 3)
               strIBF02 = Mid(sReturn, 4)
               strIBF03 = "0"
               strIBF04 = "00"
               stFileTemp = txtToPath & "\" & sReturn & ".pdf"
            'end 2024/10/22
            Else
               If strIBF01 = "" Then strIBF03 = "M51"
               strIBF02 = Val(sReturn)
            End If
            
            If ReadDB2File(stFileTemp, strIBF02, strIBF03, strIBF04, strIBF05, , strIBF01) = True Then
               Text1 = stFileTemp
               If strIBF05 = "4" Or strIBF05 = "5" Then
                  stFileName1 = stFileTemp
               Else
                  stFileName = stFileTemp
               End If
                  
               If Check1.Value = vbUnchecked Then
                  If strIBF05 = "4" Then
                     If MsgBox("檔案已下載至[ " & Text1 & " ]，是否要開啟？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                        ShellExecute hLocalFile, "open", Text1, vbNullString, vbNullString, 1
                     End If
                  ElseIf strIBF05 = "5" Then
                     Me.SSTab1.Tab = 1
                     SSTab1.Width = WebBrowser1.Width + 150
                     SSTab1.Height = WebBrowser1.Height + 600
                     Me.Width = SSTab1.Width + 150
                     Me.Height = SSTab1.Height + 1600
                     WebBrowser1.Navigate Text1
                     stFileName2 = sReturn & ".pdf"
                  Else
                     Me.SSTab1.Tab = 0
                     Set Picture1.Picture = LoadPicture(Text1)
                     If Picture1.Width + 200 > 12795 Then
                        Me.SSTab1.Width = Picture1.Width + 200
                        Me.Width = Me.SSTab1.Width + 200
                     Else
                        Me.SSTab1.Width = 12795
                        Me.Width = Me.SSTab1.Width + 200
                     End If
                     Me.SSTab1.Height = Picture1.Height + 500
                     Me.Height = Me.SSTab1.Height + 1500
                     
                     Text2 = sReturn: Command1(4).Enabled = True 'Added by Morgan 2020/4/1
                  End If
                  
               End If
            End If
         End If
      
       'Added by Morgan 2020/4/1
      Case 4 '刪除
         If Val(Text2) > 0 Then
            Me.SSTab1.Tab = 0
            
            strIBF03 = "0"
            strIBF04 = "00"
            If InStr(Text2, "-") > 0 Then
               arrTmp = Split(Text2, "-")
               strIBF02 = arrTmp(0)
               strIBF03 = arrTmp(1)
               If UBound(arrTmp) > 1 Then
                  strIBF04 = arrTmp(2)
               End If
            Else
               strIBF02 = Val(Text2)
            End If
               
            If DeletePic(strIBF02, strIBF03, strIBF04, , , strIBF01) = True Then
               Set Picture1.Picture = LoadPicture
               Text2 = "": Command1(4).Enabled = False
               MsgBox "已刪除！"
            End If
         End If
      'end 2020/4/1

      Case 6
         If Me.SSTab1.Tab = 0 Then
            If stFileName = "" Then
               MsgBox "尚未載入信頭，無法另存！", vbCritical
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         Else
            If stFileName1 = "" Then
               MsgBox "尚未載入請作單，無法另存！", vbCritical
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
         
         cd1.Filter = "All files|*.*|Bitmap files (*.bmp)|*.bmp|GIF files (*.gif)|*.gif|JPEG files (*.jpg)|*.jpg|PNG files (*.png)|*.png|TIFF files (*.tif)|*.tif|WMF files (*.wmf)|*.wmf|PDF files (*.pdf)|*.pdf"
         cd1.FilterIndex = 0
         
         If Me.SSTab1.Tab = 0 Then
            cd1.FileName = Picture1.Name
         Else
            cd1.FileName = stFileName2
         End If
         cd1.ShowOpen
         If Trim(cd1.FileName) <> "" Then
            Screen.MousePointer = vbHourglass
            Text1 = cd1.FileName
            'SavePicture Picture1.Image, cd1.FileName
            If Me.SSTab1.Tab = 0 Then
               If Dir(stFileName) <> "" Then
                  FileCopy stFileName, cd1.FileName
               End If
            Else
               If Dir(stFileName1) <> "" Then
                  FileCopy stFileName1, cd1.FileName
               End If
            End If
            Screen.MousePointer = vbDefault
         End If
         
   End Select
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
   Screen.MousePointer = vbDefault
End Sub

'從資料庫讀出檔案
Private Function ReadDB2File(ByRef p_FileName As String, pIbf02 As String, Optional pIbf03 As String = "0", Optional pIbf04 As String = "00", Optional ByRef pIbf05 As String, Optional pIbf06 As String = "2", Optional pIbf01 As String = "M51") As Boolean

   Dim iFileNo As Integer
   Dim bytes() As Byte
      
On Error GoTo ErrHnd
   
   If Left(pIbf01, 1) = "M" Then pIbf02 = Format(pIbf02, "00000#")
   strExc(0) = "select * from ImgByteFile where ibf01='" & pIbf01 & "' and ibf02='" & pIbf02 & "' and ibf03='" & pIbf03 & "' and ibf04='" & pIbf04 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      pIbf05 = "" & RsTemp("ibf05")
      If p_FileName = "" Then
         p_FileName = txtToPath & "\" & RsTemp.Fields("ibf01") & "-" & RsTemp.Fields("ibf02") & "-" & RsTemp.Fields("ibf03") & "-" & RsTemp.Fields("ibf04")
         If pIbf05 = "5" Then
            p_FileName = p_FileName & ".pdf"
         ElseIf pIbf05 = "4" Then
            p_FileName = p_FileName & ".doc"
         End If
      End If
      If Dir(p_FileName) <> "" Then
         If MsgBox("檔案已存在，是否要覆蓋？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         Else
            Kill p_FileName
         End If
      End If
      'Add By Sindy 2017/8/10
'      If "" & RsTemp.Fields("IBF15") <> "" Then
         ReadDB2File = PUB_GetFtpFile(RsTemp.Fields("IBF15"), p_FileName, UCase("ImgByteFile"))
'      Else
'      '2017/8/10 END
'         With RsTemp
'            ReDim bytes(Val(.Fields("ibf13").Value))
'            bytes() = .Fields("ibf14").GetChunk(Val(.Fields("ibf13").Value))
'            iFileNo = FreeFile
'            Open App.path & "\TempFile" For Binary Access Write As #iFileNo
'            Put #iFileNo, , bytes()
'            Close #iFileNo
'            ReadDB2File = True
'         End With
'      End If
   Else
      MsgBox "無此檔案!!", vbExclamation
   End If
   Exit Function
   
ErrHnd:
   If Err.Number = 53 Then Resume Next
   MsgBox Err.Description
   
End Function
'Added by Morgan 2020/4/1
'刪除
Private Function DeletePic(pIbf02 As String, Optional pIbf03 As String = "0", Optional pIbf04 As String = "00", Optional pIbf05 As String = "1", Optional pIbf06 As String = "2", Optional pIbf01 As String = "M51") As Boolean
   Dim stSQL As String, intR As Integer
   
On Error GoTo ErrHnd1
   pIbf02 = Format(pIbf02, "00000#")
   stSQL = "update ImgByteFile set ibf05=ibf05 where ibf01='" & pIbf01 & "' and ibf02='" & pIbf02 & "' and ibf03='" & pIbf03 & "' and ibf04='" & pIbf04 & "'"
   cnnConnection.Execute stSQL, intR
   If intR <> 1 Then
      MsgBox "資要讀取失敗！", vbCritical
      Exit Function
   End If
   
   If MsgBox("是否確定要刪除？" & vbCrLf & vbCrLf & pIbf01 & "-" & pIbf02 & "-" & pIbf03 & "-" & pIbf04, vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then
      Exit Function
   End If
      
   cnnConnection.BeginTrans
On Error GoTo ErrHnd2
   stSQL = "delete ImgByteFile where ibf01='" & pIbf01 & "' and ibf02='" & pIbf02 & "' and ibf03='" & pIbf03 & "' and ibf04='" & pIbf04 & "'"
   cnnConnection.Execute stSQL, intR
   If intR = 1 Then
      PUB_DelFtpFile2 pIbf01 & "-" & pIbf02 & "-" & pIbf03 & "-" & pIbf04 & "-" & pIbf05, , UCase("ImgByteFile")
      cnnConnection.CommitTrans
      DeletePic = True
   Else
      cnnConnection.RollbackTrans
      MsgBox "刪除失敗！", vbCritical
   End If
   
   Exit Function
      
ErrHnd2:
   cnnConnection.RollbackTrans
ErrHnd1:
   MsgBox Err.Description, vbCritical
   
End Function
'將圖檔存到資料庫 2006/8/24
Private Function Save2DB(p_FileName As String, pIbf02 As String, Optional pIbf03 As String = "0", Optional pIbf04 As String = "00", Optional pIbf05 As String = "1", Optional pIbf06 As String = "2", Optional pIbf01 As String = "M51") As Boolean
   
   Dim rstRecordset As New ADODB.Recordset
   Dim iFileNo As Integer
   Dim bytes() As Byte
   Dim lngSize As Long '檔案大小
   Dim strFtpPath As String
      
On Error GoTo ErrHnd

   iFileNo = FreeFile
   Open p_FileName For Binary Access Read As #iFileNo
   lngSize = LOF(iFileNo)
   'ReDim bytes(lngSize)
   'Get #iFileNo, , bytes()
   
   If Me.SSTab1.Tab = 0 Then pIbf02 = Format(pIbf02, "00000#")
   strExc(0) = "select * from ImgByteFile where ibf01='" & pIbf01 & "' and ibf02='" & pIbf02 & "' and ibf03='" & pIbf03 & "' and ibf04='" & pIbf04 & "'"
   
   rstRecordset.CursorLocation = adUseClient
   rstRecordset.Open strExc(0), cnnConnection, adOpenStatic, adLockOptimistic
   
   'Added by Morgan 2020/3/25
   If rstRecordset.RecordCount > 0 Then
      If MsgBox("檔案已存在，是否要覆蓋？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         GoTo ExitPort
      End If
   End If
   'end 2020/3/25
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHnd1
   
   With rstRecordset
      If .RecordCount = 0 Then
         .AddNew
         .Fields("ibf01").Value = pIbf01
         .Fields("ibf02").Value = pIbf02
         .Fields("ibf03").Value = pIbf03
         .Fields("ibf04").Value = pIbf04
         .Fields("ibf05").Value = pIbf05
         .Fields("ibf06").Value = pIbf06
         .Fields("ibf07").Value = strUserNum
         .Fields("ibf08").Value = Format(Now, "yyyymmdd")
         .Fields("ibf09").Value = Format(Now, "hhmm")
      Else
         .Fields("ibf10").Value = strUserNum
         .Fields("ibf11").Value = Format(Now, "yyyymmdd")
         .Fields("ibf12").Value = Format(Now, "hhmm")
'         .Fields("ibf14").Value = Null
         'Add By Sindy 2017/8/10 檔案改放 FTP,必須在DB資料刪除前執行
         PUB_DelFtpFile2 pIbf01 & "-" & pIbf02 & "-" & pIbf03 & "-" & pIbf04 & "-" & pIbf05, , UCase("ImgByteFile")
      End If
      .Fields("ibf13").Value = lngSize
'      .Fields("ibf14").AppendChunk bytes()
      'Modify By Sindy 2017/8/10
      '檔案改放FTP
      PUB_PutFtpFile p_FileName, pIbf01 & "-" & pIbf02 & "-" & pIbf03 & "-" & pIbf04 & "-" & pIbf05, pIbf01 & "-" & pIbf02 & "-" & pIbf03 & "-" & pIbf04 & "-" & pIbf05, strFtpPath, UCase("imgbytefile")
      If strFtpPath <> "" Then
         .Fields("ibf15") = strFtpPath
      End If
      '2017/8/10 END
      .UPDATE
   End With
   
   cnnConnection.CommitTrans
   Save2DB = True
   Set rstRecordset = Nothing
   If iFileNo <> 0 Then Close #iFileNo
   Exit Function
   
ErrHnd1:
   cnnConnection.RollbackTrans
   
ErrHnd:
   MsgBox Err.Description
   
ExitPort:
   Set rstRecordset = Nothing
   If iFileNo <> 0 Then Close #iFileNo
End Function

Private Sub Command2_Click()
   If PUB_CheckAD("92012", "morgan") Then
      MsgBox "AD認證成功!"
   Else
      MsgBox "AD認證失敗!"
   End If
End Sub

Public Function PUB_CheckAD(pID As String, pPWD As String) As Boolean
   Dim DName, ADCN, ADRS
   Dim Id As String, Pwd As String
   
   DName = "DOMAIN" '網域名稱
On Error Resume Next

   Set ADCN = CreateObject("ADODB.Connection")
   Set ADRS = CreateObject("ADODB.Recordset")
   ADCN.Open = "Provider=ADSDSOObject;User ID=" & DName & "\" & Id & ";Password=" & Pwd & _
          ";Data Source=Active Directory Provider;Mode=Read;Bind Flags=0;ADS_SECURE_AUTHENTICATION"
   Set ADRS = ADCN.Execute("SELECT * FROM 'LDAP://" & DName & "'")
   If Not ADRS.EOF Then PUB_CheckAD = True
   Set ADRS = Nothing
   ADCN.Close
On Error GoTo 0

End Function

Private Sub CommandButton1_Click()
   MSHFlexGrid1.TextMatrix(1, 1) = Me.TextBox1.Text
   MSHFlexGrid1.ToolTipText = ""
End Sub

Private Sub CommandButton2_Click()
   If PUB_ChkUniText(Me, , True) = True Then
      MsgBox "檢查成功！", vbExclamation
   Else
      MsgBox "檢查失敗！", vbExclamation
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   WebBrowser1.Navigate "about:blank"
   WebBrowser1.Width = WebBrowser1.Width * 1.2
   WebBrowser1.Height = WebBrowser1.Width * (29.7 / 21) / 2
   Label1.Tag = Label1 'Added by Morgan 2020/4/1
      
   'Added by Morgan 2024/9/30
   '讀取前次設定路徑
   '下載到:
   txtToPath.Text = GetSetting("TAIE", "P", UCase(Me.Name) & "ToPath", "")
   txtToPath.Tag = txtToPath.Text
   If txtToPath <> "" Then
      If PUB_ChkDir(txtToPath) = False Then
        txtToPath = ""
      End If
   End If
   If txtToPath = "" Then
      txtToPath = PUB_Getdesktop
   End If
   'end 2024/9/30
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveSetting "TAIE", "M51", UCase(Me.Name) & "ToPath", txtToPath
   
   DestroyToolTip '清除物件
   Set frm000003 = Nothing
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   arrGridHeadWidth = Array(240, 4000, 1500, 2000)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   .FixedCols = 0
   .FormatString = "V|檔案名稱|大小|最後修改時間"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         If iCol = 2 Then
            .ColAlignment(iCol) = flexAlignRightCenter
         Else
            .ColAlignment(iCol) = flexAlignLeftCenter
         End If
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub


Private Sub BrowseFtpFolder(pFtpPath As String, Optional pErrMsg As String, Optional pRaiseErr As Boolean = True, Optional ByVal pFtpIp As String = "")
   Dim hConnection As Long
   Dim pData As WIN32_FIND_DATA
   Dim hFind As Long, LRet  As Long, stFileName As String
   Dim tZone As TIME_ZONE_INFORMATION
   Dim bias As Long
   Dim ft As SYSTEMTIME
   Dim tmpDate As Date

   hConnection = PUB_GetFtpConnect(pErrMsg, , , pFtpIp)
   If hConnection <> 0 Then
      If FtpSetCurrentDirectory(hConnection, pFtpPath) = 1 Then
         pData.cFileName = String(MAX_PATH, 0)
         hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
         If hFind <> 0 Then
            Call GetTimeZoneInformation(tZone)
            bias = tZone.bias
            Do
               stFileName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
               stFileName = MultiByteToUTF16(UTF16ToMultiByte(stFileName, 950), cpUTF8)
               If stFileName <> "." And stFileName <> ".." Then
                  FileTimeToSystemTime pData.ftLastWriteTime, ft
                  tmpDate = CDate(ft.wYear & "-" & ft.wMonth & "-" & ft.wDay & " " & ft.wHour & ":" & ft.wMinute & ":" & ft.wSecond) - TimeSerial(0, bias, 0)
                  MSHFlexGrid1.AddItem "" & vbTab & stFileName & vbTab & Format(Round(pData.nFileSizeLow / 1024), "#,###") & " KB" & vbTab & tmpDate
               End If
               LRet = InternetFindNextFile(hFind, pData)
            Loop While LRet <> 0
            InternetCloseHandle hFind
            hFind = 0
         End If
      End If
   End If
   
OutPort:
   If Err.Number <> 0 Then pErrMsg = Err.Description
   If hConnection <> 0 Then InternetCloseHandle (hConnection)
   If hFind <> 0 Then InternetCloseHandle hFind
   
   If pErrMsg <> "" And pRaiseErr = True Then
      Err.Raise 999, , pErrMsg
   End If
End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
With MSHFlexGrid1
If .MouseRow <> 0 And .MouseCol > 0 Then
   If iRow <> .MouseRow Or iCol <> .MouseCol Then
      'Debug.Print Now & "->" & .MouseRow & "," & .MouseCol & "=" & .TextMatrix(.MouseRow, .MouseCol)
      CreateToolTip GetHWndForToolTip(MSHFlexGrid1), .TextMatrix(.MouseRow, .MouseCol), , , , , , True
      'CreateToolTip GetHWndForToolTip(MSHFlexGrid1), MSHFlexGrid1.Text
      iRow = .MouseRow: iCol = .MouseCol
   End If
End If

End With
End Sub

Private Sub TextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   CreateToolTip GetHWndForToolTip(TextBox1), TextBox1.Text
End Sub

Private Sub OutlookTest()
   Dim ii As Integer, jj As Integer
   Dim olApp, myNamespace, myFolder, oMailItem
   Dim strSubject As String
      
   Set olApp = CreateObject("Outlook.Application")
   Set myNamespace = olApp.GetNamespace("MAPI")
   For ii = 1 To myNamespace.Folders.Count
      If myNamespace.Folders(ii) = "公用資料夾 - 92012@taie.com.tw" Then
         Set myFolder = myNamespace.Folders(ii)
         Exit For
      End If
   Next
   
   For ii = 1 To myFolder.Folders.Count
      If myFolder.Folders(ii) = "所有公用資料夾" Then
         Set myFolder = myFolder.Folders(ii)
         Exit For
      End If
   Next
      
   For ii = 1 To myFolder.Folders.Count
      If myFolder.Folders(ii) = "CACK(確收處理區)" Then
         Set myFolder = myFolder.Folders(ii)
         Exit For
      End If
   Next
   
   If myFolder = "CACK(確收處理區)" Then
      ii = 0
      Do While myFolder.Items.Count > ii
         ii = ii + 1
         strSubject = "主旨:" & myFolder.Items(ii).Subject
         TextBox5.Text = strSubject & vbCrLf & vbCrLf & TextBox5.Text
         
         'If UniMsgBox("是否刪除下列郵件：" & vbCrLf & vbCrLf & TextBox5.Text, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
         '   myFolder.Items(1).Delete
         '   Exit For
         'End If
         If UploadCACK(myFolder.Items(ii)) = True Then
            myFolder.Items(ii).Delete
            ii = ii - 1
         Else
            strExc(1) = App.path & "\CACK\" & strSrvDate(1) & ServerTime & ".msg"
            myFolder.Items(ii).SaveAs strExc(1)
            Set oMailItem = olApp.CreateItemFromTemplate(strExc(1))
            
            oMailItem.Recipients.add myFolder.Items(ii)
            MsgBox "確收失敗！", vbCritical
         End If
         
         If myFolder.Items.Count > ii Then
            If UniMsgBox(strSubject & vbCrLf & vbCrLf & "是否讀取下一筆郵件？", vbYesNo + vbQuestion + vbDefaultButton1) = vbNo Then
               Exit Do
            End If
         End If
      Loop
   End If
   
   Set olApp = Nothing
   Set myNamespace = Nothing
   Set myFolder = Nothing

End Sub

Private Function UploadCACK(ByRef pMailItem As Object) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stSaveName As String, stSaveName1 As String, stSaveName2 As String
   Dim stSavePath As String
   Dim oFileSys As New FileSystemObject
   Dim oFile
   Dim boInTrans As Boolean
   
On Error GoTo ErrHnd

   stSavePath = App.path & "\CACK"
   If Dir(stSavePath, vbDirectory) = "" Then
      MkDir stSavePath
   End If
   
   intQ = 1
   stSQL = "select lp01,cp01,cp02,cp03,cp04,cp10,to_char(sysdate,'YYYYMMDDHH24MISS') TT" & _
      " From letterprogress,smailbackup c,caseprogress" & _
      " where lp39>=" & (strSrvDate(1) - 10000) & _
      " and smb01(+)=lp01 and smb02(+)=lp39 and smb03(+)=lp40" & _
      " and instr('" & ChgSQL(pMailItem.Subject) & "',smb07)>0 and cp09(+)=lp01"
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      stSaveName = PUB_CaseNo2FileName(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04")) & "." & .Fields("cp10")
      stSaveName2 = stSaveName & "." & .Fields("TT") & ".CACK.msg"
      stSaveName = stSaveName & ".CACK.msg"
      pMailItem.SaveAs stSavePath & "\" & stSaveName2
      Set oFile = oFileSys.GetFile(stSavePath & "\" & stSaveName2)
      cnnConnection.BeginTrans
      boInTrans = True
      
      '上確收紀錄
      stSQL = "update letterprogress set lp46='QPGMR',lp47=to_char(sysdate,'YYYYMMDD'),lp48=to_char(sysdate,'HH24MISS') where lp01='" & .Fields("LP01") & "' and lp47=0"
      cnnConnection.Execute stSQL, intQ
      '檢查卷宗區是否已有確收信
      stSQL = "update CasePaperPDF set cpp01=cpp01 where cpp01='" & .Fields("LP01") & "' and upper(cpp02)=upper('" & stSaveName & "')"
      cnnConnection.Execute stSQL, intQ
      If intQ > 0 Then
         stSaveName = stSaveName2
      End If
      SaveAttFile_PDF .Fields("lp01"), stSavePath & "\" & stSaveName2, stSaveName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True
      cnnConnection.CommitTrans
      UploadCACK = True
      End With
   Else
      
      
   End If

ErrHnd:
   If Err.Number <> 0 Then
      If boInTrans Then cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
   Set rsQuery = Nothing
   Set oFileSys = Nothing
   Set oFile = Nothing
End Function

Private Sub Test()
   Dim arrChar() As Byte
   Dim ii As Integer
   
   PUB_PrintUnicodeText TextBox4.Text, 500, 500
   Printer.EndDoc
   Exit Sub
   
   OutlookTest
   Exit Sub
   
   arrChar = TextBox4.Text
   TextBox5.Text = ""
   For ii = LBound(arrChar) To UBound(arrChar)
      TextBox5.Text = TextBox5.Text & Right("0" & Hex(arrChar(ii)), 2)
   Next
   'TextBox5.Text = UTF8EncodeHex(TextBox4.Text)
   'TextBox5.Text = Base64Decode(TextBox4.Text)
   'TextBox4 = Base64Encode(TextBox2.Text)
   
End Sub

'字串轉UTF8編碼(回傳16進位碼)
Public Function UTF8EncodeHex(ByVal strText As String) As String

    On Error Resume Next
    
    Dim bArray() As Byte
    Dim ii As Integer
    
    If strText = "" Then Exit Function
    
    bArray() = ConvertStringToUtf8Bytes(strText)
    
    For ii = 0 To UBound(bArray)
      UTF8EncodeHex = UTF8EncodeHex & Hex(bArray(ii))
    Next
    
End Function

Public Function Base64Encode(ByVal strToEncode As String) As String

    On Error Resume Next
    
    Dim bArray() As Byte, i As Long, n1 As Long, n2 As Long, n3 As Long, C1 As Long, C2 As Long, c3 As Long, c4 As Long
    Dim ReByte() As Byte
    
    If strToEncode = "" Then Exit Function
    
    bArray() = ConvertStringToUtf8Bytes(strToEncode)
    
    For i = 0 To UBound(bArray) Step 3
    
        n1 = CLng(bArray(i))
        If i + 1 <= UBound(bArray) Then n2 = CLng(bArray(i + 1)) Else n2 = -1
        If i + 2 <= UBound(bArray) Then n3 = CLng(bArray(i + 2)) Else n3 = -1
        C1 = -1: C2 = -1: c3 = -1: c4 = -1
        C1 = n1 \ 4
        C2 = (n1 And 3) * 16
        If n2 >= 0 Then C2 = C2 + (n2 \ 16): c3 = (n2 And 15) * 4
        If n3 >= 0 Then c3 = c3 + (n3 \ 64): c4 = n3 And 63
        Base64Encode = Base64Encode & Mid$(Base64Char, C1 + 1, 1)
        Base64Encode = Base64Encode & Mid$(Base64Char, C2 + 1, 1)
        If c3 >= 0 Then Base64Encode = Base64Encode & Mid$(Base64Char, c3 + 1, 1)
        If c4 >= 0 Then Base64Encode = Base64Encode & Mid$(Base64Char, c4 + 1, 1)
    
    Next
    
    Base64Encode = Base64Encode & String$(((UBound(bArray) + 1) * 8) Mod 3, "=")
    
    
    
End Function

Public Function Base64Decode(ByVal strToDecode As String) As String

    On Error Resume Next
    Dim DecodedBytes() As Byte, Length As Long, w1 As Long, w2 As Long, w3 As Long, w4 As Long, C1 As Long, C2 As Long, c3 As Long, i As Long, j As Long
    
    If strToDecode = "" Then Exit Function
    
    strToDecode = RemoveElseCharacters(strToDecode)
    Length = Int(Len(Replace$(strToDecode, "=", "")) * 0.75)
    
    ReDim DecodedBytes(Length - 1) As Byte
    
    j = 0
    
    For i = 1 To Len(strToDecode) Step 4
        w1 = InStr(Base64Char, Mid$(strToDecode, i, 1)) - 1
        w2 = InStr(Base64Char, Mid$(strToDecode, i + 1, 1)) - 1
        If Mid$(strToDecode, i + 2, 1) <> "=" Then w3 = InStr(Base64Char, Mid$(strToDecode, i + 2, 1)) - 1 Else w3 = -1
        If Mid$(strToDecode, i + 3, 1) <> "=" Then w4 = InStr(Base64Char, Mid$(strToDecode, i + 3, 1)) - 1 Else w4 = -1
        C1 = -1: C2 = -1: c3 = -1
        C1 = w1 * 4 + (w2 \ 16)
        C2 = (w2 And 15) * 16
        If w3 >= 0 Then
            C2 = C2 + (w3 \ 4)
            c3 = (w3 And 3) * 64
        End If
        If w4 >= 0 Then
            c3 = c3 + w4
        End If
        DecodedBytes(j) = CByte(C1 And &HFF)
        If UBound(DecodedBytes) >= j + 1 Then DecodedBytes(j + 1) = CByte(C2 And &HFF)
        If c3 >= 0 Then DecodedBytes(j + 2) = CByte(c3 And &HFF)
        j = j + 3
    Next
    
    Base64Decode = ConvertUtf8BytesToString(DecodedBytes())

    
End Function

Private Function RemoveElseCharacters(ByVal strToProcess As String) As String
    On Error Resume Next
    Static oRegExp As Object
    Dim sProcess As String, i As Long
    If ObjPtr(oRegExp) = 0 Then Set oRegExp = CreateObject("VBScript.RegExp")
    If ObjPtr(oRegExp) Then
        oRegExp.Global = True
        oRegExp.Pattern = "[^A-Za-z0-9\+\/\=]"
        RemoveElseCharacters = oRegExp.Replace(strToProcess, "")
    Else
        For i = 1 To Len(strToProcess)
            If InStr(Base64Char, Mid$(strToProcess, i, 1)) Then
                sProcess = sProcess & Mid$(strToProcess, i, 1)
            End If
        Next
        RemoveElseCharacters = sProcess
    End If
End Function


'Declare Need:
'Microsoft ActiveX data Objects 2.5 Library
Public Function ConvertStringToUtf8Bytes(ByRef strText As String) As Byte()

    Dim objStream As ADODB.Stream
    Dim data() As Byte
   
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.Type = adTypeText
    objStream.Open
   
    ' write bytes into stream
    objStream.WriteText strText
    objStream.Flush
   
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = adTypeBinary
    data = objStream.Read(3)
    data = objStream.Read()
   
    ' close up and return
    objStream.Close
    ConvertStringToUtf8Bytes = data


End Function


Public Function ConvertUtf8BytesToString(ByRef data() As Byte) As String


    Dim objStream As ADODB.Stream
    Dim strTmp As String
   
    ' init stream
    Set objStream = New ADODB.Stream
    objStream.Charset = "utf-8"
    objStream.Mode = adModeReadWrite
    objStream.Type = adTypeBinary
    objStream.Open
   
    ' write bytes into stream
    objStream.Write data
    objStream.Flush
   
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = adTypeText
    strTmp = objStream.ReadText
   
    ' close up and return
    objStream.Close
    ConvertUtf8BytesToString = strTmp


End Function

Private Sub txtAtt_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Morgan 2024/9/30
Public Function CapsLock() As Boolean
   ' Determine whether CAPSLOCK key is toggled on.
   CapsLock = CBool(GetKeyState(VK_CAPITAL) And 1)
End Function

Public Sub SetCapsLockState(bEnabled As Boolean)
    'CapsLock is already in desired state. Nothing to do.
    If CapsLock = bEnabled Then Exit Sub

    PressCapsLock
End Sub

Private Sub PressCapsLock()
    GenerateKeyboardEvent VK_CAPITAL, 0
    GenerateKeyboardEvent VK_CAPITAL, KEYEVENTF_KEYUP
End Sub

Private Sub GenerateKeyboardEvent(VirtualKey As Long, Flags As Long)
    Dim kevent As KeyboardInput

    With kevent
        .dwType = INPUT_KEYBOARD
        .wScan = MapVirtualKey(VirtualKey, 0)
        .wVK = VirtualKey
        .dwTime = 0
        .dwFlags = Flags
    End With
    SendInput 1, kevent, Len(kevent)
End Sub
'end 2024/9/30
