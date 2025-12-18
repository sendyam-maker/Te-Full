VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm12040153 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利資料匯入(代繳年費用)"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8325
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   315
      TabIndex        =   11
      Top             =   3210
      Width           =   7350
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   2145
      Left            =   270
      TabIndex        =   4
      Top             =   990
      Visible         =   0   'False
      Width           =   7395
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   5040
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1830
         Width           =   1995
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   1170
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1830
         Width           =   1995
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   390
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   1200
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "已過時間："
         Height          =   180
         Index           =   1
         Left            =   4050
         TabIndex        =   14
         Top             =   1875
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "開始時間："
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   1875
         Width           =   900
      End
      Begin VB.Label lblMessage2 
         Alignment       =   2  '置中對齊
         Caption         =   "NowFile"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   10
         Top             =   990
         Width           =   6855
      End
      Begin VB.Label lblMessage1 
         Alignment       =   2  '置中對齊
         Caption         =   "Nowfolder"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   9
         Top             =   180
         Width           =   6855
      End
      Begin VB.Label lblProgress2 
         Alignment       =   2  '置中對齊
         Caption         =   "(0/0)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   8
         Top             =   1470
         Width           =   6855
      End
      Begin VB.Label lblProgress1 
         Alignment       =   2  '置中對齊
         Caption         =   "(0/0)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   7
         Top             =   750
         Width           =   6855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "匯入"
      Height          =   375
      Left            =   6615
      TabIndex        =   3
      Top             =   570
      Width           =   1050
   End
   Begin VB.TextBox txtSource 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   5910
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "選擇目錄..."
      Height          =   315
      Left            =   6615
      TabIndex        =   0
      Top             =   240
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "來源："
      Height          =   180
      Left            =   135
      TabIndex        =   2
      Top             =   270
      Width           =   540
   End
End
Attribute VB_Name = "frm12040153"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/5 改成Form2.0 (無)
'Memo By Sonia 2012/12/6 智權人員欄已修改
Option Explicit

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Function BrowseForFolder(Optional sCaption As String = "Select a folder", Optional sDefault As String) As String
    Const BIF_RETURNONLYFSDIRS = 1
    Const MAX_PATH = 260
    Dim lPos As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, tBrowse As BrowseInfo

    With tBrowse
        'Set the owner window
        .hwndOwner = GetActiveWindow        'Me.hWnd in VB
        .lpszTitle = sCaption
        .ulFlags = BIF_RETURNONLYFSDIRS     'Return only if the user selected a directory
    End With

    'Show the dialog
    lpIDList = SHBrowseForFolder(tBrowse)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        lPos = InStr(sPath, vbNullChar)
        If lPos Then
            BrowseForFolder = Left$(sPath, lPos - 1)
            If Right$(BrowseForFolder, 1) <> "\" Then
                BrowseForFolder = BrowseForFolder & "\"
            End If
        End If
    Else
        'User cancelled, return default path
        BrowseForFolder = sDefault
    End If
End Function

Private Sub cmdBrowse_Click()
   txtSource = BrowseForFolder(, txtSource.Text)
End Sub


Private Sub Command1_Click()
   Dim sPath As String
   Dim sFolder As String, sFile As String, sFilePath As String
   Dim sFolders() As String, iFolderUp As Integer
   Dim lFileCount As Long
   Dim ii As Integer
   Dim dt1 As Date, dt2 As Date
   
   If txtSource.Text = "" Then
      MsgBox "請選擇來源目錄!"
      Exit Sub
   ElseIf Dir(txtSource.Text, vbDirectory) = "" Then
      MsgBox "來源目錄不存在!"
      Exit Sub
   End If
   
   iFolderUp = 0
   sPath = txtSource.Text
   If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
   sFolder = Dir(sPath, vbDirectory)
   Do While sFolder <> ""
      If sFolder <> "." And sFolder <> ".." Then
         If GetAttr(sPath & sFolder) = vbDirectory Then
         iFolderUp = iFolderUp + 1
         ReDim Preserve sFolders(iFolderUp) As String
         sFolders(iFolderUp) = sFolder
         End If
      End If
      sFolder = Dir
   Loop
   
   If iFolderUp = 0 Then
      iFolderUp = 1
      ReDim Preserve sFolders(iFolderUp) As String
      sFolders(iFolderUp) = "."
   End If
   
   ProgressBar1.max = iFolderUp
   ProgressBar1.Min = 0
   ProgressBar1.Value = 0
   lblProgress1 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
   
   dt1 = Now
   Text1(0).Text = Format(dt1, "HH:mm:ss")
   Text1(1).Text = ""
   
   Frame1.Visible = True
   
   For ii = 1 To iFolderUp
   
      lFileCount = 0
      
      sFile = Dir(sPath & sFolders(ii) & "\*.txt")
      Do While sFile <> ""
         lFileCount = lFileCount + 1
         sFile = Dir
      Loop
      
      lblMessage1.Caption = sFolders(ii)
      ProgressBar2.max = lFileCount
      ProgressBar2.Min = 0
      ProgressBar2.Value = 0
      lblProgress2 = "( " & ProgressBar2.Value & "/" & ProgressBar2.max & " )"
   
      sFile = Dir(sPath & sFolders(ii) & "\*.txt")
      Do While sFile <> ""
         sFilePath = sPath & sFolders(ii) & "\" & sFile
         lblMessage2.Caption = sFile
         If UpdateRecord(sFilePath, strExc(1), sFolders(ii)) = False Then
            List1.AddItem sFolders(ii) & "\" & sFile & ":" & strExc(1), 0
         End If
         
         dt2 = Now
         Text1(1).Text = GetDiff(dt1, dt2)
   
         DoEvents
         ProgressBar2.Value = ProgressBar2.Value + 1
         lblProgress2 = "( " & ProgressBar2.Value & "/" & ProgressBar2.max & " )"
         sFile = Dir
      Loop
      
      ProgressBar1.Value = ProgressBar1.Value + 1
      lblProgress1 = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
   Next
   MsgBox "匯入結束"
End Sub

Private Function GetDiff(pDt1 As Date, pDt2 As Date) As String
   Dim db1 As Double, sRtn As String
   
   db1 = DateDiff("s", pDt1, pDt2)
   sRtn = db1 Mod 60
   db1 = db1 \ 60
   
   sRtn = (db1 Mod 60) & ":" & sRtn
   db1 = db1 \ 60
   
   sRtn = (db1 Mod 24) & ":" & sRtn
   db1 = db1 \ 24
   sRtn = db1 & ":" & sRtn
   
   GetDiff = sRtn
End Function


Private Function UpdateRecord(sFileName As String, stErrMsg As String, Optional pNo As String) As Boolean
   Dim iStart As Integer, iEnd As Integer, sValue(3) As String
   Dim stSQL As String
   Dim objStream As Object
   Set objStream = CreateObject("ADODB.Stream")
   Dim stContent As String
   
On Error GoTo ErrHnd
   
   With objStream
      .Type = 2
      .Mode = 3
      .Open
      .Charset = "UTF-8" ' 或其他編碼
      '.Charset = "UTF-16"
      .LoadFromFile sFileName
      stContent = .ReadText
      ' PS : 也可透過 .SaveToFile 方法把檔案存檔
      .Close
   End With
   
   iStart = InStr(stContent, "<專利編號>") + 6
   iEnd = InStr(stContent, "</專利編號>") - 1
   If iEnd >= iStart Then
      sValue(1) = Mid(stContent, iStart, iEnd - iStart + 1)
   Else
      stErrMsg = "無法讀取<專利編號>!"
      Exit Function
   End If
   
   iStart = InStr(stContent, "<申請號>") + 6
   iEnd = InStr(stContent, "</申請號>") - 1
   If iEnd > iStart Then
      sValue(2) = Mid(stContent, iStart, iEnd - iStart + 1)
   Else
      stErrMsg = "無法讀取<申請號>!"
      Exit Function
   End If
   
   iStart = InStr(stContent, "<公告/公開日>") + 8
   iEnd = InStr(stContent, "</公告/公開日>") - 1
   If iEnd > iStart Then
      sValue(3) = Mid(stContent, iStart, iEnd - iStart + 1)
   Else
      stErrMsg = "無法讀取</公告/公開日>!"
      Exit Function
   End If
   
   stSQL = "insert into Patent4Fee(pf01,pf02,pf03,pf04) values('" & sValue(1) & "','" & sValue(2) & "'," & DBDATE(sValue(3)) & ",'" & pNo & "')"
   cnnConnection.Execute stSQL, intI
   
   UpdateRecord = True
   Exit Function
   
ErrHnd:
   stErrMsg = Err.Description
   
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040153 = Nothing
End Sub

