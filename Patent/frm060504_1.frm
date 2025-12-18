VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060504_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件命名追蹤-上傳檔案"
   ClientHeight    =   3372
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3372
   ScaleWidth      =   7320
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdSelect 
      Caption         =   "全選"
      Height          =   345
      Index           =   0
      Left            =   3168
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2916
      Width           =   675
   End
   Begin VB.CommandButton cmdRemAtt 
      Caption         =   "移除"
      Height          =   345
      Index           =   0
      Left            =   1728
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2916
      Width           =   675
   End
   Begin VB.CommandButton cmdAddAtt 
      Caption         =   "加入"
      Height          =   345
      Index           =   0
      Left            =   1008
      TabIndex        =   5
      Top             =   2916
      Width           =   675
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   345
      Index           =   0
      Left            =   2448
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2916
      Width           =   675
   End
   Begin VB.ListBox lstAtt 
      Height          =   2208
      Index           =   0
      ItemData        =   "frm060504_1.frx":0000
      Left            =   120
      List            =   "frm060504_1.frx":0007
      MultiSelect     =   2  '進階多重選取
      Sorted          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   7080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdAtt 
      Height          =   600
      Left            =   1620
      TabIndex        =   2
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
   Begin VB.Label lblTCN01 
      Caption         =   "lblTCN01"
      Height          =   228
      Left            =   1032
      TabIndex        =   9
      Top             =   156
      Width           =   1932
   End
   Begin VB.Label Lbl01 
      Caption         =   "追蹤號："
      Height          =   204
      Left            =   180
      TabIndex        =   8
      Top             =   156
      Width           =   720
   End
End
Attribute VB_Name = "frm060504_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Lydia 2023/08/18
Option Explicit

Public m_strTCN01 As String
Public m_strSaveFiles As String

Dim m_MousePointer As Integer
Dim ii As Integer
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Dim m_AttachPath As String '附件宣告區
Dim m_PrevForm As Form '前一畫面
Dim m_FLmax As Integer '限制檔案名稱長度
Dim bolMax As Boolean '限制檔案大小

Public Sub SetParent(ByRef fm As Form, Optional ByVal mLength As Integer = 0, Optional ByVal mMax As Boolean = False)
   Set m_PrevForm = fm
   m_FLmax = mLength
   bolMax = mMax
End Sub

Private Sub cmdOK_Click(Index As Integer)
Dim bolCancel As Boolean
Dim stFileName As String
   
   '確定
   If Index = 0 Then

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
      
      stFileName = ""
      For ii = 0 To lstAtt(0).ListCount - 1
         stFileName = stFileName & "&" & lstAtt(0).List(ii)
      Next ii
      m_PrevForm.m_strSaveFiles = Mid(stFileName, 2)
   End If
   
   If lblTCN01.Caption <> "(自動編號)" Then
      If lstAtt(0).ListCount = 0 Then
         MsgBox "請新增附件！", vbExclamation
         Exit Sub
      End If
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
   
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath")
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   KillAttach
 
   If m_strTCN01 = "" Then
      lblTCN01.Caption = "(自動編號)"
   Else
      lblTCN01.Caption = m_strTCN01
      '直接異動DB
      cmdAddAtt(0).Caption = "新增"
      cmdRemAtt(0).Caption = "刪除"
      cmdOK(0).Visible = False '確定
      cmdOK(1).Caption = "回前畫面"
   End If
   
End Sub

Public Function QueryData(bolAddList As Boolean) As Boolean
Dim rsA As New ADODB.Recordset
Dim sFile
   
   QueryData = False
   
   If m_strTCN01 = "" Then '尚無追蹤號(Tracking NO.)
      If m_strSaveFiles <> "" Then
         sFile = Split(m_strSaveFiles, "&")
         For ii = 0 To UBound(sFile)
            lstAtt(0).AddItem sFile(ii), 0
            SetListScroll lstAtt(0)
         Next ii
      End If
      QueryData = True
   Else
      
      lstAtt(0).Clear
      strExc(0) = "Select * " & _
                    "From CasePaperFile " & _
                  "Where cpf01 ='" & Trim(m_strTCN01) & "' and cpf10<>'D' and substr(upper(cpf02),-4)<>upper('.del') " & _
                  "order by cpf06 asc,cpf07 asc"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If bolAddList = True Then
            Do While Not rsA.EOF
               lstAtt(0).AddItem rsA.Fields("cpf02"), 0
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
   Set frm060504_1 = Nothing
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
   
   If m_strTCN01 <> "" Then
      strKey = lblTCN01.Caption '追蹤號(Tracking NO.)
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
               '排除 C:\Program Files (x86) 狀況
               If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
                  stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
               End If
            End If
            If InStr(stFileName, "\") = 0 Then '已存入FTP
               stFullName = m_AttachPath & "\" & stFileName
               If PUB_GetAttachFile_Org(strKey, stFileName, stFullName, True) = False Then
                  MsgBox "檔案(" & stFileName & ")下載失敗！", vbCritical
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

'加入/新增
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
   
   If m_strTCN01 <> "" Then
      strKey = lblTCN01.Caption '追蹤號(Tracking NO.)
   End If
   
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            
            If MsgBox("確定要刪除" & GetFileName(oList.List(ii)) & "電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Function

            '直接從資料庫刪除檔案-->只保留最新檔案
            bolDel = DelAttFile_File("", strKey, GetFileName(oList.List(ii)), , , , True)
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
      oList.AddItem stNewItem, 0
      AddListX = True
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
   End If
End Function

'新增檔案
Private Function AddFile(Index As Integer) As Boolean
   Dim stFileName As String, stReName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f
   Dim strFile As String
   Dim strTmp As String
   
On Error GoTo ErrHnd

   AddFile = False

   stFileName = "*.*"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.*)|*.*"
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      '取消 Or cdlOFNNoDereferenceLinks
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer 'Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
            
            If InStr(CStr(sFile(0)), "&") > 0 Then
                 MsgBox CStr(sFile(0)) & "\" & CStr(sFile(1)) & vbCrLf & vbCrLf & "【&】符號為系統保留字，不可使用於資料夾名稱！", vbExclamation
                 Exit Function
            End If
            
            For ii = 1 To UBound(sFile)
               If InStr(CStr(sFile(ii)), "#") > 0 Or InStr(CStr(sFile(ii)), "&") > 0 Or InStr(CStr(sFile(ii)), " (") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#和&及 (】符號為系統保留字，不可使用於檔案命名！", vbExclamation
                  Exit Function
               End If
               
               '檢查檔名規則: 單純英數字
               strTmp = PUB_GetSimpleName(CStr(sFile(ii)), , True)
               If strTmp <> CStr(sFile(ii)) Then
                  MsgBox "檔案名稱只可為英數字、半形空白及""_"",""-""，請修改下列檔案名稱！" & CStr(sFile(ii)), vbCritical
                  Exit Function
               End If
               If m_FLmax > 0 Then
                  If GetTextLength(strTmp) > m_FLmax Then
                     MsgBox "去除檔案路徑的檔案名稱總長度超過" & m_FLmax & "字元 !" & _
                          vbCrLf & "請移除檔案後縮短檔案名稱，再重新加入!" & vbCrLf & _
                          "檔案名稱：" & strTmp
                     Exit Function
                  End If
               End If
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               stReName = strTmp 'Tracking NO 在收文後才變更檔名與卷宗區一致
               
               '檢查檔案是否正在使用中
               If PUB_ChkFileOpening(stFileName) = True Then
                  MsgBox stFileName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
                  Exit Function
               End If

               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               '檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  Exit Function
               End If
               
               If ChkListFileExists(stFileName) = True Then Exit Function '檢查檔案是否已存在
               
               If Me.cmdAddAtt(Index).Caption = "新增" Then
                  If SaveAttFile_Org(m_strTCN01, stFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), "A") = False Then
                     GoTo ErrHnd
                  End If
               Else
                  AddListX lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS")
               End If
            Next ii
         Else
            If InStr(.FileName, "&") > 0 Then
               MsgBox .FileName & vbCrLf & vbCrLf & "【&】符號為系統保留字，不可使用於資料夾名稱！", vbExclamation
               Exit Function
            End If
            
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Or InStr(strFile, "&") > 0 Or InStr(strFile, " (") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "【#和&及 (】符號為系統保留字，不可使用於檔案命名！", vbExclamation
               Exit Function
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
            
            '檢查檔名規則: 單純英數字
            strTmp = PUB_GetSimpleName(strFile, , True)
            If strTmp <> strFile Then
               MsgBox "檔案名稱只可為英數字、半形空白及""_"",""-""，請修改下列檔案名稱！" & strFile, vbCritical
               Exit Function
            End If
            If m_FLmax > 0 Then
               If GetTextLength(strTmp) > m_FLmax Then
                  MsgBox "去除檔案路徑的檔案名稱總長度超過" & m_FLmax & "字元 !" & _
                       vbCrLf & "請移除檔案後縮短檔案名稱，再重新加入!" & vbCrLf & _
                       "檔案名稱：" & strTmp
                  Exit Function
               End If
            End If
            
            stFileName = .FileName
            stReName = strTmp  'Tracking NO 在收文後才變更檔名與卷宗區一致
            
            '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(stFileName) = True Then
               MsgBox stFileName & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
               Exit Function
            End If
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            '檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg strFile & MsgText(9221)
               Exit Function
            End If
            
            If ChkListFileExists(stFileName) = True Then Exit Function '檢查檔案是否已存在
            
            If Me.cmdAddAtt(Index).Caption = "新增" Then
               If SaveAttFile_Org(m_strTCN01, stFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), "A") = False Then
                  GoTo ErrHnd
               End If
            Else
               AddListX lstAtt(Index), stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS")
            End If
         End If
      End If
   End With
   
   Set fs = Nothing
   Set f = Nothing
   
   AddFile = True
   Exit Function
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

