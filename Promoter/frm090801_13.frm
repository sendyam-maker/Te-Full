VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090801_13 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子收文-附件區"
   ClientHeight    =   5592
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5592
   ScaleWidth      =   7320
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2175
      Left            =   0
      TabIndex        =   17
      Top             =   2700
      Width           =   7215
      Begin VB.ListBox lstAtt 
         Height          =   1128
         Index           =   1
         ItemData        =   "frm090801_13.frx":0000
         Left            =   150
         List            =   "frm090801_13.frx":0007
         MultiSelect     =   2  '進階多重選取
         Sorted          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   420
         Width           =   7080
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   345
         Index           =   1
         Left            =   2490
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1770
         Width           =   675
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "加入"
         Height          =   345
         Index           =   1
         Left            =   1050
         TabIndex        =   21
         Top             =   1770
         Width           =   675
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "移除"
         Height          =   345
         Index           =   1
         Left            =   1770
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1770
         Width           =   675
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "全選"
         Height          =   345
         Index           =   1
         Left            =   3210
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1770
         Width           =   675
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1080
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   18
         Text            =   "frm090801_13.frx":0013
         Top             =   30
         Width           =   5415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "內部文件："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   180
         Left            =   30
         TabIndex        =   24
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   645
      Left            =   30
      TabIndex        =   14
      Top             =   4920
      Width           =   7215
      Begin VB.TextBox TextNote 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   15
         Text            =   "frm090801_13.frx":0079
         Top             =   210
         Width           =   5955
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "備註：匯入時，檔案將搬移至系統中。"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   60
         TabIndex        =   16
         Top             =   0
         Width           =   3060
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frm090801_13.frx":00F2
      Top             =   540
      Width           =   6135
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全選"
      Height          =   345
      Index           =   0
      Left            =   3210
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2310
      Width           =   675
   End
   Begin VB.CommandButton cmdRemAtt 
      Caption         =   "移除"
      Height          =   345
      Index           =   0
      Left            =   1770
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2310
      Width           =   675
   End
   Begin VB.CommandButton cmdAddAtt 
      Caption         =   "加入"
      Height          =   345
      Index           =   0
      Left            =   1050
      TabIndex        =   7
      Top             =   2310
      Width           =   675
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   345
      Index           =   0
      Left            =   2490
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2310
      Width           =   675
   End
   Begin VB.ListBox lstAtt 
      Height          =   1128
      Index           =   0
      ItemData        =   "frm090801_13.frx":0158
      Left            =   150
      List            =   "frm090801_13.frx":015F
      MultiSelect     =   2  '進階多重選取
      Sorted          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   7080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdAtt 
      Height          =   600
      Left            =   1620
      TabIndex        =   4
      Top             =   5790
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7049
      _ExtentY        =   1058
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "本所案號|原路徑檔名|檔名"
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4950
      TabIndex        =   1
      Top             =   60
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   5940
      TabIndex        =   0
      Top             =   60
      Width           =   930
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4290
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblCRL01 
      Caption         =   "lblCRL01"
      Height          =   225
      Left            =   1320
      TabIndex        =   13
      Top             =   270
      Width           =   1935
   End
   Begin VB.Label LblTcrl01 
      AutoSize        =   -1  'True
      Caption         =   "接洽單編號："
      Height          =   180
      Left            =   150
      TabIndex        =   12
      Top             =   270
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "送件文書："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   180
      Left            =   30
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   150
      TabIndex        =   3
      Top             =   30
      Width           =   900
   End
   Begin VB.Label lblCaseNo 
      Caption         =   "lblCaseNo"
      Height          =   225
      Left            =   1200
      TabIndex        =   2
      Top             =   30
      Width           =   2055
   End
End
Attribute VB_Name = "frm090801_13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2022/10/11
Option Explicit

Public bolNewCase As Boolean
Public m_strCRL01 As String
Public strCaseNA239 As String
Public m_strSaveFiles As String
Public m_strSaveFiles2 As String

Dim m_MousePointer As Integer
Dim ii As Integer
Dim m_FilesRemoved() As String
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Public m_blnCallQuery As Boolean
Dim m_AttachPath As String '附件宣告區
Dim m_PrevForm As Form '前一畫面
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String


Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim bolCancel As Boolean, k As Integer
Dim stFileName As String
   
   '確定
   If Index = 0 Then
      If Frame2.Visible = True Then
         If lstAtt(0).ListCount = 0 And lstAtt(1).ListCount = 0 Then
            If m_PrevForm.m_strSaveFiles <> "" Then '原有檔案
               m_PrevForm.m_strSaveFiles = "" '確定取消檔名
               bolCancel = True
            End If
            If m_PrevForm.m_strSaveFiles2 <> "" Then '原有檔案
               m_PrevForm.m_strSaveFiles2 = "" '確定取消檔名
               bolCancel = True
            End If
            If bolCancel = True Then
               Unload Me
               Screen.MousePointer = m_MousePointer
               Exit Sub
            End If
            
            MsgBox "請加入附件！", vbExclamation
            Exit Sub
         End If
      Else
         If lstAtt(0).ListCount = 0 Then
            If m_PrevForm.m_strSaveFiles <> "" Then '原有檔案
               m_PrevForm.m_strSaveFiles = "" '確定取消檔名
               Unload Me
               Screen.MousePointer = m_MousePointer
               Exit Sub
            End If
            
            MsgBox "請加入附件！", vbExclamation
            Exit Sub
         End If
      End If
      
      For k = 0 To 1
         stFileName = ""
         For ii = 0 To lstAtt(k).ListCount - 1
            stFileName = stFileName & "&" & lstAtt(k).List(ii)
         Next ii
         If k = 0 Then m_PrevForm.m_strSaveFiles = Mid(stFileName, 2)
         If k = 1 Then m_PrevForm.m_strSaveFiles2 = Mid(stFileName, 2)
      Next k
      If Frame2.Visible = False Then m_PrevForm.m_strSaveFiles2 = ""
   End If
   
   If lblCRL01.Visible = True Then
      If QueryData(False) = True Then
         m_PrevForm.m_strSaveFiles = "Y"
      End If
   End If
   
   Unload Me
   Screen.MousePointer = m_MousePointer
End Sub

Private Sub KillAttach()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   m_MousePointer = Screen.MousePointer
   Screen.MousePointer = vbDefault
   lstAtt(0).Clear
   lstAtt(1).Clear
   
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath")
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   KillAttach
   
   Frame1.Caption = "": Frame1.BorderStyle = 0
   Frame2.Caption = "": Frame2.BorderStyle = 0
   
   '舊案只顯示官方文件區
   If bolNewCase = False Then
      Frame1.Top = 3000
      Frame2.Visible = False
      'Add By Sindy 2023/2/7
      'Label2.Caption = "案件回覆單"
      Label2.Visible = False
      '2023/2/7 END
      Text1.Visible = False
      Me.Height = 4275
      TextNote.Visible = True
   Else
      TextNote.Visible = False
      Label2.Visible = True 'Add By Sindy 2023/2/7
   End If
   
   LblTcrl01.Visible = False
   lblCRL01.Visible = False
   If m_strCRL01 <> "" Then
      LblTcrl01.Visible = True
      lblCRL01.Visible = True
      lblCRL01.Caption = m_strCRL01
   End If
   
   If m_blnCallQuery = True Then
      cmdAddAtt(0).Visible = False
      cmdAddAtt(1).Visible = False
      cmdRemAtt(0).Visible = False
      cmdRemAtt(1).Visible = False
      cmdOK(0).Visible = False '確定
      cmdOK(1).Caption = "回前畫面"
      
   ElseIf m_strCRL01 <> "" Then
      '直接異動DB
      cmdAddAtt(0).Caption = "新增"
      cmdAddAtt(1).Caption = "新增"
      cmdRemAtt(0).Caption = "刪除"
      cmdRemAtt(1).Caption = "刪除"
      cmdOK(0).Visible = False '確定 (隱藏確定,是因其它按鍵直接異動DB)
      cmdOK(1).Caption = "回前畫面"
   End If
End Sub

Public Function QueryData(bolAddList As Boolean) As Boolean
Dim rsA As New ADODB.Recordset
Dim sFile
   
   QueryData = False
   
   If m_strCRL01 = "" Then '尚無接洽單編號
      If m_strSaveFiles <> "" Then
         sFile = Split(m_strSaveFiles, "&")
         For ii = 0 To UBound(sFile)
            lstAtt(0).AddItem sFile(ii), 0
            SetListScroll lstAtt(0)
         Next ii
      End If
      If m_strSaveFiles2 <> "" Then
         sFile = Split(m_strSaveFiles2, "&")
         For ii = 0 To UBound(sFile)
            lstAtt(1).AddItem sFile(ii), 0
            SetListScroll lstAtt(1)
         Next ii
      End If
      QueryData = True
   Else
      Str01 = CStr(SystemNumber(lblCaseNo, 1))
      Str02 = CStr(SystemNumber(lblCaseNo, 2))
      Str03 = CStr(SystemNumber(lblCaseNo, 3))
      Str04 = CStr(SystemNumber(lblCaseNo, 4))
      
      lstAtt(0).Clear
      lstAtt(1).Clear
      
      strExc(0) = "Select * " & _
                    "From CasePaperPDF " & _
                  "Where cpp11 ='" & Trim(m_strCRL01) & "' and length(cpp01)<>9 and cpp10<>'D' and substr(upper(cpp02),-4)<>upper('.del') " & _
                  "order by cpp15 asc,cpp06 asc,cpp07 asc"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If bolAddList = True Then
            Do While Not rsA.EOF
               If "" & rsA.Fields("cpp15") <> "" Then
                  If rsA.Fields("cpp15") = "1" Or rsA.Fields("cpp15") = "3" Then '1.官方文件 3.案件回覆單
                     lstAtt(0).AddItem rsA.Fields("cpp02"), 0
                  Else
                     lstAtt(1).AddItem rsA.Fields("cpp02"), 0
                  End If
               Else
                  lstAtt(0).AddItem rsA.Fields("cpp02"), 0
               End If
               rsA.MoveNext
            Loop
         End If
         QueryData = True
      End If
      rsA.Close
   End If
   
   Set rsA = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing
   Set frm090801_13 = Nothing
End Sub

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String
   Dim bolIsSelect As Boolean
   Dim strKey As String
   Dim stFullName As String
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   
   If bolNewCase = False Then
      strKey = Replace(lblCaseNo, "-", "") '案號
   ElseIf m_strCRL01 <> "" Then
      strKey = lblCRL01.Caption '接洽單編號
   End If
   strAtt = lstAtt(Index).Text

   If strAtt = "" Then
      MsgBox "請選擇欲開啟的附件！", vbExclamation
   Else
      For ii = 0 To lstAtt(Index).ListCount - 1
         If lstAtt(Index).Selected(ii) Then
            bolIsSelect = True
            stFileName = lstAtt(Index).List(ii)
            If InStrRev(stFileName, " (") > 0 Then
               'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
               If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
               '2021/8/6 END
                  stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
               End If
            End If
            
            If InStr(stFileName, "\") = 0 Then
               stFullName = m_AttachPath & "\" & stFileName
               If PUB_GetAttachFile_CPP(strKey, stFileName, stFullName, True) = False Then
                  MsgBox "接洽單電子檔(" & stFileName & ")下載失敗！", vbCritical
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            
            SetAttr stFileName, vbReadOnly '檔案設定為唯讀屬性
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      Next ii
      If bolIsSelect = False Then
         MsgBox "請選擇欲開啟的附件！", vbExclamation
      End If
   End If

   Screen.MousePointer = vbDefault
End Sub

'全選
Private Sub cmdSelect_Click(Index As Integer)
   Dim oList As ListBox

   Set oList = lstAtt(Index)
   For ii = 0 To oList.ListCount - 1
      lstAtt(Index).Selected(ii) = True
   Next
End Sub

'新增
Private Sub cmdAddAtt_Click(Index As Integer)
   If AddFile(Index) = False Then
      Exit Sub
   Else
      If Me.cmdAddAtt(Index).Caption = "新增" Then
         Call Me.QueryData(True)
      End If
   End If
End Sub

'刪除
Private Sub cmdRemAtt_Click(Index As Integer)
   If Me.cmdRemAtt(Index).Caption = "刪除" Then
      '直接刪除DB
      Call RemoveList_DB(lstAtt(Index))
   Else
      Call RemoveList(lstAtt(Index))
   End If
End Sub

Private Function RemoveList_DB(oList As ListBox) As Boolean
Dim ii As Integer
Dim bolDel As Boolean
Dim strKey As String
   
   If bolNewCase = False Then
      strKey = Replace(lblCaseNo, "-", "") '案號
   ElseIf m_strCRL01 <> "" Then
      strKey = lblCRL01.Caption '接洽單編號
   End If
   
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            
            If MsgBox("確定要刪除" & GetFileName(oList.List(ii)) & "電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Function
            
'            If oList.ItemData(ii) > 0 Then
'               intI = UBound(m_FilesRemoved) + 1
'               ReDim Preserve m_FilesRemoved(intI) As String
'               m_FilesRemoved(intI) = GetFileName(oList.List(ii))
'            End If
            
            '直接從資料庫刪除檔案
            bolDel = DelAttFile_PDF("", strKey, GetFileName(oList.List(ii)))
            
            If bolDel = True Then
               oList.RemoveItem ii
               SetListScroll oList
               RemoveList_DB = True
               ii = ii - 1
            End If
         End If
         ii = ii + 1
      Loop
   End If
End Function

Private Function RemoveList(oList As ListBox) As Boolean
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            oList.RemoveItem ii
            SetListScroll oList
            RemoveList = True
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
   End If
End Function

Private Sub SetListScroll(oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long

   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next

   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Function AddListX(oList As ListBox, stNewItem As String) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String

   If stNewItem <> "" Then
'      For idx = 0 To oList.ListCount - 1
'         stFileName = GetFileName(oList.List(idx))
'         If UCase(GetFileName(stNewItem)) = UCase(stFileName) Then
'            MsgBox "附件 " & stFileName & " 已存在！", vbExclamation
'            AddListX = False
'            bFound = True
'            Exit For
'         End If
'      Next
'      If bFound = False Then
         oList.AddItem stNewItem, 0
         'SetListScroll oList
         AddListX = True
'      End If
   End If
End Function

'檢查檔案是否已存在
Private Function ChkListFileExists(stNewItem As String) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
   
   ChkListFileExists = False
   If stNewItem <> "" Then
      For idx = 0 To lstAtt(0).ListCount - 1
         stFileName = GetFileName(lstAtt(0).List(idx))
         If UCase(GetFileName(stNewItem)) = UCase(stFileName) Then
            MsgBox "附件 " & stFileName & " 已存在！", vbExclamation
            ChkListFileExists = True
            Exit Function
         End If
      Next
      For idx = 0 To lstAtt(1).ListCount - 1
         stFileName = GetFileName(lstAtt(1).List(idx))
         If UCase(GetFileName(stNewItem)) = UCase(stFileName) Then
            MsgBox "附件 " & stFileName & " 已存在！", vbExclamation
            ChkListFileExists = True
            Exit Function
         End If
      Next
   End If
End Function

'新增檔案
Private Function AddFile(Index As Integer) As Boolean
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim strFile As String
   
On Error GoTo ErrHnd

   AddFile = False
   
   If Index = 1 Then
      stFileName = "*.*"
   Else
      stFileName = "*.PDF"
   End If
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      If Index = 1 Then
         .Filter = "All Files (*.*)|*.*"
      Else
         .Filter = "All Files (*.PDF)|*.PDF"
      End If
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      'Modify By Sindy 2023/4/19 取消 Or cdlOFNNoDereferenceLinks
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer 'Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
            
            'Added by Lydia 2018/03/28
            If InStr(CStr(sFile(0)), "&") > 0 Then
                 MsgBox CStr(sFile(0)) & "\" & CStr(sFile(1)) & vbCrLf & vbCrLf & "【&】符號為系統保留字，不可使用於資料夾名稱！", vbExclamation
                 Exit Function
            End If
            'end 2018/03/28
            
            For ii = 1 To UBound(sFile)
               If InStr(CStr(sFile(ii)), "#") > 0 Or InStr(CStr(sFile(ii)), "&") > 0 Or InStr(CStr(sFile(ii)), " (") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#和&及 (】符號為系統保留字，不可使用於檔案命名！", vbExclamation
                  Exit Function
               End If
               
               '檢查檔名規則
               If lblCaseNo.Visible = True And bolNewCase = False Then
                   If PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), EMP_會修, "") = False Then
                      Exit Function
                   End If
               End If
               
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               
               If Index = 0 Then
                  '只可加入PDF檔
                  If UCase(Mid(sFile(ii), InStrRev(sFile(ii), ".") + 1)) <> "PDF" Then
                     MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "只可加入PDF檔！", vbExclamation
                     Exit Function
                  End If
               End If
               
               'Add By Sindy 2015/3/6 檢查檔案是否正在使用中
               If PUB_ChkFileOpening(stFileName) = True Then
                  MsgBox stFileName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
                  Exit Function
               End If
               '2015/3/6 END
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               '檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  Exit Function
'               'Modify By Sindy 2015/2/11 改控制可以放5M以下的檔案
''               ElseIf f.Size > 512000 Then
''                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "檔案大小不可超過500KB！", vbExclamation
'               'Modified by Lydia 2018/02/07 +是否限制檔案大小 And bolMax = True
'               ElseIf f.Size > 5242880 Then
'                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "檔案大小不可超過 5MB！", vbExclamation
'               '2015/2/11 END
'                  Exit Function
               'Modify By Sindy 2023/5/18 改彈提醒不鎖住
               ElseIf f.Size > 5242880 Then
                  If MsgBox("檔案過大（容量超過5MB），確認是否要上傳？", vbYesNo, "警告") = vbNo Then
                     Exit Function
                  End If
                  '2023/5/18 END
               End If
               
               If ChkListFileExists(stFileName) = True Then Exit Function '檢查檔案是否已存在
               
               If Me.cmdAddAtt(Index).Caption = "新增" Then
                  If PUB_SaveCRLFile(IIf(Index = 0, stFileName, ""), IIf(Index = 1, stFileName, ""), Str01, Str02, Str03, Str04, _
                     m_strCRL01, bolNewCase, , True) = False Then
                     GoTo ErrHnd
                  End If
               Else
                  AddListX lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS")
               End If
            Next ii
            
         Else
            'Added by Lydia 2018/03/28 資料夾名稱排除"&"
            If InStr(.FileName, "&") > 0 Then
               MsgBox .FileName & vbCrLf & vbCrLf & "【&】符號為系統保留字，不可使用於資料夾名稱！", vbExclamation
               Exit Function
            End If
            'end 2018/03/28
            
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Or InStr(strFile, "&") > 0 Or InStr(strFile, " (") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "【#和&及 (】符號為系統保留字，不可使用於檔案命名！", vbExclamation
               Exit Function
            End If
            
            If Index = 0 Then
               '只可加入PDF檔
               If UCase(Mid(strFile, InStrRev(strFile, ".") + 1)) <> "PDF" Then
                  MsgBox strFile & vbCrLf & vbCrLf & "只可加入PDF檔！", vbExclamation
                  Exit Function
               End If
            End If
            
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     Exit For
                  End If
               Next ii
            End If
            
            '檢查檔名規則
            If lblCaseNo.Visible = True And bolNewCase = False Then
                If PUB_ChkEmpFlowFNMRule(lblCaseNo, strFile, EMP_會修, "") = False Then
                   Exit Function
                End If
            End If
            
            stFileName = .FileName
            'Add By Sindy 2015/3/6 檢查檔案是否正在使用中
            If PUB_ChkFileOpening(stFileName) = True Then
               MsgBox stFileName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
               Exit Function
            End If
            '2015/3/6 END
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            '檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg strFile & MsgText(9221)
               Exit Function
'            'Modify By Sindy 2015/2/11 改控制可以放5M以下的檔案
''            ElseIf f.Size > 512000 Then
''               MsgBox strFile & vbCrLf & vbCrLf & "檔案大小不可超過500KB！", vbExclamation
'             'Modified by Lydia 2018/02/07 +是否限制檔案大小 And bolMax = True
'            ElseIf f.Size > 5242880 Then
'               MsgBox strFile & vbCrLf & vbCrLf & "檔案大小不可超過 5MB！", vbExclamation
'            '2015/2/11 END
'               Exit Function
            'Modify By Sindy 2023/5/18 改彈提醒不鎖住
            ElseIf f.Size > 5242880 Then
               If MsgBox("檔案過大（容量超過5MB），確認是否要上傳？", vbYesNo, "警告") = vbNo Then
                  Exit Function
               End If
               '2023/5/18 END
'            'Added by Lydia 2019/12/23 控制檔名不可超過100字; ex.FCP-062455在TrackingNo資料夾放”AP rl    Instructions   - Issue Fee and new divisional patent application - Taiwanese Patent Application No 107106625 - Your Ref  P187543 TW 01 DeakM DeZeeuwR and P222272 TW 01DeakM DeZeeuwR - Our Ref  FCP-058411  ACK 601 .msg”
'            'Modified by Lydia 2020/05/06 控制檔名不可超過80字 ; 因為若不是msg檔,最終會以"案號+案件性質+原檔名"上傳
'            ElseIf CheckLengthIsOK(f.Name, 80) = False Then
'               Exit Function
            End If
            
            If ChkListFileExists(stFileName) = True Then Exit Function '檢查檔案是否已存在
            
            If Me.cmdAddAtt(Index).Caption = "新增" Then
               If PUB_SaveCRLFile(IIf(Index = 0, stFileName, ""), IIf(Index = 1, stFileName, ""), Str01, Str02, Str03, Str04, _
                  m_strCRL01, bolNewCase, , True) = False Then
                  GoTo ErrHnd
               End If
            Else
               AddListX lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS")
            End If
         End If
      End If
   End With
   
   AddFile = True
   Exit Function
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function
