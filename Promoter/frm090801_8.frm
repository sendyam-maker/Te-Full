VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090801_8 
   BorderStyle     =   1  '單線固定
   Caption         =   "新增附件"
   ClientHeight    =   3360
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5532
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5532
   StartUpPosition =   3  '系統預設值
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdAtt 
      Height          =   600
      Left            =   3840
      TabIndex        =   11
      Top             =   3300
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7070
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
      Left            =   3450
      TabIndex        =   7
      Top             =   50
      Width           =   930
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "取消(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4440
      TabIndex        =   6
      Top             =   50
      Width           =   930
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   2200
      Left            =   90
      TabIndex        =   0
      Top             =   816
      Width           =   5355
      Begin VB.ListBox lstAtt 
         Height          =   1668
         ItemData        =   "frm090801_8.frx":0000
         Left            =   60
         List            =   "frm090801_8.frx":0007
         MultiSelect     =   2  '進階多重選取
         Sorted          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   60
         Width           =   5220
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   345
         Left            =   1680
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1800
         Width           =   675
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "加入"
         Height          =   345
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   675
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "移除"
         Height          =   345
         Left            =   960
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1800
         Width           =   675
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "全選"
         Height          =   345
         Left            =   2400
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1800
         Width           =   675
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "label2"
      ForeColor       =   &H00FF0000&
      Height          =   370
      Left            =   120
      TabIndex        =   12
      Top             =   355
      Width           =   5200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "備註：匯入時，檔案將搬移至系統中。"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   3090
      Width           =   3060
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   96
      TabIndex        =   9
      Top             =   84
      Width           =   900
   End
   Begin VB.Label lblCaseNo 
      Caption         =   "lblCaseNo"
      Height          =   228
      Left            =   1080
      TabIndex        =   8
      Top             =   84
      Width           =   2148
   End
End
Attribute VB_Name = "frm090801_8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 (無需修改)
'Create By Sindy 2014/7/18
Option Explicit

Public m_strSaveFiles As String
Dim m_MousePointer As Integer
Dim ii As Integer
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194
Dim m_PrevForm As Form '前一畫面
Dim bolOA As Boolean 'Added by Lydia 2015/10/30 法務-開庭通知
Public bolNotPDF As Boolean 'Add By Sindy 2016/6/24 '開啟附件的檔案類型
Dim bolMax As Boolean 'Added by Lydia 2018/02/07 限制檔案大小
Dim m_FLmax As Integer 'Added by Lydia 2018/02/07 限制檔案名稱長度
Dim bolDesc As Boolean 'Added by Lydia 2018/02/27 讀取前次記錄的排序
Dim m_Title As String 'Added by Lydia 2018/03/02
'Add by Amy 2018/06/06
Dim bolMultFile As String  '多案號多檔操作
Dim arrCaseNo
Public intFCState As Integer '1-商標/2-專利 'Add by Amy 2025/05/08

'Modified by Lydia 2018/02/07 +mLength
'Modified by Lydia 2018/02/27 +bDesc
'Modified by Lydia 2018/03/02 iTitle
Public Sub SetParent(ByRef fm As Form, Optional ByVal mLength As Integer = 0, Optional ByVal bDesc As Boolean = False, Optional ByVal iTitle As String = "")
   Set m_PrevForm = fm
   
   'Added by Lydia 2018/02/07 是否限制檔案大小和檔名長度
   m_FLmax = mLength
   'Modified by Lydia 2018/12/20 +命名作業(FRM090902_2,FRM090903_1)
   If UCase(TypeName(m_PrevForm)) = "FRM060504" Or UCase(TypeName(m_PrevForm)) = "FRM060120" _
       Or UCase(TypeName(m_PrevForm)) = "FRM090902_2" Or UCase(TypeName(m_PrevForm)) = "FRM090903_1" Then
       bolMax = False
   Else
       bolMax = True
   End If
   'end 2018/02/07
   
   bolDesc = bDesc 'Added by Lydia 2018/02/27
   m_Title = iTitle 'Added by Lydia 2018/03/02
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim stFileName As String
Dim sFile As Variant
Dim stMsg As String 'Add by Amy 2018/06/06
Dim intLen As Integer 'Add By Sindy 2021/8/6
   
   '確定
   If Index = 0 Then
      If lstAtt.ListCount = 0 Then
         'Modify By Sindy 2016/6/24
         If m_PrevForm.Name = "frm071013" Or m_PrevForm.Name = "frm075012" Then
            If m_PrevForm.m_strSaveFileType = "1" Then '開庭通知
               If m_PrevForm.m_strSaveFilesOA <> "" Then '原有檔案
                  m_PrevForm.m_strSaveFilesOA = "" '確定取消檔名
                  Unload Me
                  Screen.MousePointer = m_MousePointer
                  Exit Sub
               End If
            ElseIf m_PrevForm.m_strSaveFileType = "2" Then '開庭紀要
               If m_PrevForm.m_strSaveFilesBRIEF <> "" Then '原有檔案
                  m_PrevForm.m_strSaveFilesBRIEF = "" '確定取消檔名
                  Unload Me
                  Screen.MousePointer = m_MousePointer
                  Exit Sub
               End If
            ElseIf m_PrevForm.m_strSaveFileType = "3" Then '電子筆錄
               If m_PrevForm.m_strSaveFilesNOTE <> "" Then '原有檔案
                  m_PrevForm.m_strSaveFilesNOTE = "" '確定取消檔名
                  Unload Me
                  Screen.MousePointer = m_MousePointer
                  Exit Sub
               End If
            End If
         Else
         '2016/6/24 END
            'Modify By Sindy 2015/2/11
            If m_PrevForm.m_strSaveFiles <> "" Then '原有檔案
               m_PrevForm.m_strSaveFiles = "" '確定取消檔名
               Unload Me
               Screen.MousePointer = m_MousePointer
               Exit Sub
            End If
            '2015/2/11 END
         End If
         MsgBox "請加入附件！", vbExclamation
         Exit Sub
      End If
      
      stFileName = ""
      For ii = 0 To lstAtt.ListCount - 1
         stFileName = stFileName & "&" & lstAtt.List(ii)
      Next ii
      
      'Modify By Sindy 2016/6/24
      If m_PrevForm.Name = "frm071013" Or m_PrevForm.Name = "frm075012" Then
         If m_PrevForm.m_strSaveFileType = "1" Then '開庭通知
            m_PrevForm.m_strSaveFilesOA = Mid(stFileName, 2)
         ElseIf m_PrevForm.m_strSaveFileType = "2" Then '開庭紀要
            m_PrevForm.m_strSaveFilesBRIEF = Mid(stFileName, 2)
         ElseIf m_PrevForm.m_strSaveFileType = "3" Then '電子筆錄
            m_PrevForm.m_strSaveFilesNOTE = Mid(stFileName, 2)
         End If
         
         '檢查檔案是否有重覆選取
         If m_PrevForm.m_strSaveFilesOA <> "" And m_PrevForm.m_strSaveFilesBRIEF <> "" Then
            sFile = Split(m_PrevForm.m_strSaveFilesBRIEF, "&")
            For ii = 0 To UBound(sFile)
               If InStr(m_PrevForm.m_strSaveFilesOA, Replace(sFile(ii), "\\", "\")) > 0 Then
                  
                  'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
                  intLen = Len(sFile(ii))
                  If InStrRev(sFile(ii), " (") > 0 Then
                     If UCase(Mid(sFile(ii), InStrRev(sFile(ii), " (") + 1, Len("(X86)"))) <> "(X86)" Then
                        intLen = InStrRev(sFile(ii), "(") - 1
                     End If
                  End If
                  '2021/8/6 END

                  MsgBox Trim(Mid(sFile(ii), 1, intLen)) & " 檔案重覆選取了！", vbExclamation
                  Exit Sub
               End If
            Next ii
         End If
         If m_PrevForm.Name = "frm075012" Then
            If m_PrevForm.m_strSaveFilesOA <> "" And m_PrevForm.m_strSaveFilesNOTE <> "" Then
               sFile = Split(m_PrevForm.m_strSaveFilesNOTE, "&")
               For ii = 0 To UBound(sFile)
                  If InStr(m_PrevForm.m_strSaveFilesOA, Replace(sFile(ii), "\\", "\")) > 0 Then
                     
                     'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
                     intLen = Len(sFile(ii))
                     If InStrRev(sFile(ii), " (") > 0 Then
                        If UCase(Mid(sFile(ii), InStrRev(sFile(ii), " (") + 1, Len("(X86)"))) <> "(X86)" Then
                           intLen = InStrRev(sFile(ii), "(") - 1
                        End If
                     End If
                     '2021/8/6 END
                     
                     MsgBox Trim(Mid(sFile(ii), 1, intLen)) & " 檔案重覆選取了！", vbExclamation
                     Exit Sub
                  End If
               Next ii
            End If
            If m_PrevForm.m_strSaveFilesBRIEF <> "" And m_PrevForm.m_strSaveFilesNOTE <> "" Then
               sFile = Split(m_PrevForm.m_strSaveFilesNOTE, "&")
               For ii = 0 To UBound(sFile)
                  If InStr(m_PrevForm.m_strSaveFilesBRIEF, Replace(sFile(ii), "\\", "\")) > 0 Then
                  
                     'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
                     intLen = Len(sFile(ii))
                     If InStrRev(sFile(ii), " (") > 0 Then
                        If UCase(Mid(sFile(ii), InStrRev(sFile(ii), " (") + 1, Len("(X86)"))) <> "(X86)" Then
                           intLen = InStrRev(sFile(ii), "(") - 1
                        End If
                     End If
                     '2021/8/6 END
                     
                     MsgBox Trim(Mid(sFile(ii), 1, intLen)) & " 檔案重覆選取了！", vbExclamation
                     Exit Sub
                  End If
               Next ii
            End If
         End If
      'Add by Amy 2018/06/06 多案號多檔操作(商標延展結案說明輸入進入)
      ElseIf bolMultFile = True Then
        For ii = 1 To GrdAtt.Rows - 1
            If GrdAtt.TextMatrix(ii, 1) = MsgText(601) Then
                stMsg = stMsg & GrdAtt.TextMatrix(ii, 0) & vbCrLf
            End If
        Next ii
        If stMsg <> MsgText(601) Then
            If MsgBox(stMsg & "無對應的電子檔，確定是否要重新點選回覆單？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                Call cmdAddAtt_Click
                Exit Sub
            End If
        End If
      Else
      '2016/6/24 END
          m_PrevForm.m_strSaveFiles = Mid(stFileName, 2)
      End If
   End If
   
   Unload Me
   Screen.MousePointer = m_MousePointer
End Sub

Private Sub Form_Load()
Dim sFile
   
   MoveFormToCenter Me
   m_MousePointer = Screen.MousePointer
   Screen.MousePointer = vbDefault
   Label2.Visible = False
   If intFCState > 0 Then
      Label2.Visible = True
      Label2.Caption = "檔案命名：案號.MSG (或.PDF)" & vbCrLf & _
                                       "多檔命名：案號.1.MSG (或.PDF)、案號.2.MSG (或.PDF)..."
   End If
   lstAtt.Clear
   bolMultFile = False 'Add by Amy 2018/06/06
   'Added by Lydia 2015/10/30
   bolOA = False
   If TypeName(m_PrevForm) <> "Nothing" Then
      'Modify By Sindy 2016/6/27 + Or m_PrevForm.Name = "frm075012"
      If m_PrevForm.Name = "frm071013" Or m_PrevForm.Name = "frm075012" Then
         bolOA = True
      'Add by Amy 2018/06/06 增加多案號多檔案操作
      ElseIf UCase(m_PrevForm.Name) = "FRM100123_1" Then
         bolMultFile = True
         SetGrdAtt
         Exit Sub
      End If
   End If
   'end 2015/10/30
   
   'Added by Lydia 2018/03/02 傳入抬頭
   'Memo by Lydia 2018/03/14 因為Me.BorderStyle=4-單線固定工具視窗,會造成Win7預設佈景顯示無法顯示form.caption,所以改成1-單線固定
   If m_Title <> "" Then Me.Caption = m_Title
   
   'Memo by Amy 2025/05/08 多檔名串法使用符號不一,故有遇到再取代
   '結案單被退,從一般作業->目前表單->回覆單鈕 多檔會使用[冒號]串,但再加入檔案會有路徑C:\...,故[冒號]不可由此取代
     
   If m_strSaveFiles <> "" Then
      sFile = Split(m_strSaveFiles, "&")
      'Added by Lydia 2018/02/27 改順序
      If bolDesc = True Then
            For ii = UBound(sFile) To 0 Step -1
               lstAtt.AddItem sFile(ii), 0
               SetListScroll lstAtt
            Next ii
      Else
      'end 2018/02/27
            For ii = 0 To UBound(sFile)
               lstAtt.AddItem sFile(ii), 0
               SetListScroll lstAtt
            Next ii
      End If 'end 2018/02/26
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_PrevForm = Nothing
   Set frm090801_8 = Nothing
End Sub

'開啟附件
Private Sub cmdOpenAtt_Click()
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim strAtt As String
   Dim bolIsSelect As Boolean
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   
   strAtt = lstAtt.Text
   
   If strAtt = "" Then
      MsgBox "請選擇欲開啟的附件！", vbExclamation
   Else
      For ii = 0 To lstAtt.ListCount - 1
         If lstAtt.Selected(ii) Then
            bolIsSelect = True
            stFileName = lstAtt.List(ii)
            If bolMultFile = True Then stFileName = Mid(stFileName, InStrRev(stFileName, "->") + 2)
            If InStrRev(stFileName, " (") > 0 Then
               'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
               If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
               '2021/8/6 END
                  stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
               End If
            End If
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
Private Sub cmdSelect_Click()
   Dim oList As ListBox
   
   Set oList = lstAtt
   For ii = 0 To oList.ListCount - 1
      lstAtt.Selected(ii) = True
   Next
End Sub

'新增
Private Sub cmdAddAtt_Click()
    Dim stMsg As String
    'Modify by Amy 2018/06/06 原程式搬至AddFile_OnlyCaseNo
    If bolMultFile = True Then
        If AddFile_Mult = False Then Exit Sub
        SetGrdAtt '設定list
    Else
        If AddFile_OnlyCaseNo = False Then Exit Sub
    End If
End Sub

'刪除
Private Sub cmdRemAtt_Click()
    'Add by Amy 2018/06/06 +if
    If bolMultFile = True Then
        If DelR090801_8 = True Then
            SetGrdAtt
        End If
    Else
        Call RemoveList(lstAtt)
    End If
End Sub

Private Function RemoveList(oList As ListBox) As Boolean
Dim TempList As String 'Added by Lydia 2018/02/07 記錄去除路徑的檔案名稱

   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            oList.RemoveItem ii
            SetListScroll oList
            RemoveList = True
            ii = ii - 1
         'Added by Lydia 2018/02/07 記錄去除路徑的檔案名稱
         ElseIf m_FLmax > 0 Then
               strExc(1) = oList.List(ii)
               If InStrRev(strExc(1), " (") > 0 Then
                  'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
                  If UCase(Mid(strExc(1), InStrRev(strExc(1), " (") + 1, Len("(X86)"))) <> "(X86)" Then
                  '2021/8/6 END
                     strExc(1) = Left(strExc(1), InStrRev(strExc(1), " (") - 1)
                  End If
               End If
               TempList = TempList & Mid(strExc(1), InStrRev(strExc(1), "\") + 1) & ";"
         'end 2018/02/07
         End If
         ii = ii + 1
      Loop
   End If
   
   'Added by Lydia 2018/02/07 限制檔案名稱長度
   If TempList <> "" Then TempList = Mid(TempList, 1, Len(TempList) - 1)
   If cmdOK(0).Enabled = False And GetTextLength(TempList) < m_FLmax Then
       cmdOK(0).Enabled = True
   End If
   'end 2018/02/07
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
   
'   'Add By Sindy 2018/2/8
'   If UCase(TypeName(m_PrevForm)) = UCase("frm090801") Or _
'      UCase(TypeName(m_PrevForm)) = UCase("frm210133") Then
'      If oList.ListCount = 0 Then
'         cmdAddAtt.Visible = True
'      Else
'         cmdAddAtt.Visible = False
'      End If
'   End If
'   '2018/2/8 END
   
   If ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

Private Function AddListX(oList As ListBox, stNewItem As String) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
   
   If stNewItem <> "" Then
      For idx = 0 To oList.ListCount - 1
         stFileName = GetFileName(oList.List(idx))
         If UCase(GetFileName(stNewItem)) = UCase(stFileName) Then
            MsgBox "附件 " & stFileName & " 已存在！", vbExclamation
            AddListX = False
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem stNewItem, 0
         SetListScroll oList
         AddListX = True
      End If
   End If
End Function

'Add by Amy 2018/06/06
'設定GrdAtt及lstAtt
Private Sub SetGrdAtt()
    Dim strIns As String
    Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
    
    GrdAtt.Clear
    GrdAtt.Rows = 2
    If ChkTempTB = False Then
        '暫存檔為空值時只寫入本所案號至暫存檔中
        'Me.GrdAtt.Visible = True
        arrCaseNo = Split(m_strSaveFiles, ",")
        For ii = 0 To UBound(arrCaseNo)
            strCP01 = SystemNumber(CStr(arrCaseNo(ii)), 1)
            strCP02 = SystemNumber(CStr(arrCaseNo(ii)), 2)
            strCP03 = SystemNumber(CStr(arrCaseNo(ii)), 3)
            strCP04 = SystemNumber(CStr(arrCaseNo(ii)), 4)
            strIns = "Insert Into R090801_8 (ID,R001,R002,R003,R004,R007) Values" & _
                       "('" & strUserNum & "','" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "','" & UCase(m_PrevForm.Name) & "')"
            cnnConnection.Execute strIns
    
            If GrdAtt.Rows = 2 And GrdAtt.TextMatrix(GrdAtt.Rows - 1, 0) = "" Then
                GrdAtt.TextMatrix(GrdAtt.Rows - 1, 0) = arrCaseNo(ii) '本所案號
            Else
                GrdAtt.AddItem arrCaseNo(ii) '本所案號
            End If
            GrdAtt.TextMatrix(GrdAtt.Rows - 1, 1) = "" 'User 路徑
            GrdAtt.TextMatrix(GrdAtt.Rows - 1, 2) = "" '檔名
        Next ii
    Else
        SetlstAtt
    End If
    GrdAtt.FormatString = GrdAtt.FormatString
End Sub

Private Sub SetlstAtt()
    lstAtt.Clear
    For ii = 1 To GrdAtt.Rows - 1
        If GrdAtt.TextMatrix(ii, 2) <> MsgText(601) Then
            lstAtt.AddItem GrdAtt.TextMatrix(ii, 0) & "->" & GrdAtt.TextMatrix(ii, 1)
            SetListScroll lstAtt
        End If
    Next ii
End Sub

'確認暫存檔是否有資料
Private Function ChkTempTB() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    strQ = "Select R001||'-'||R002||'-'||R003||'-'||R004,R005,R006 From R090801_8 " & _
              "Where ID='" & strUserNum & "' And R007='" & UCase(m_PrevForm.Name) & "' " & _
              "Order by R001,R002,R003,R004"
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        ChkTempTB = True
        Set GrdAtt.Recordset = RsQ
    End If
    RsQ.Close
End Function

Private Function SaveR090801_8(stNewItem As String, strCaseNo As String, ByRef strMsg As String) As Boolean
    Dim strIns As String, strFileN As String
    Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
    
On Error GoTo ErrHand
    SaveR090801_8 = False: strMsg = ""
    
    strFileN = GetFileName(stNewItem)
    '避免先寫入Grid後存暫存檔失敗,故先寫暫存檔
    strCP01 = SystemNumber(strCaseNo, 1)
    strCP02 = SystemNumber(strCaseNo, 2)
    strCP03 = SystemNumber(strCaseNo, 3)
    strCP04 = SystemNumber(strCaseNo, 4)
    If ChkStatus(strCP01, strCP02, strCP03, strCP04) = True Then
        strIns = "Insert Into R090801_8 (ID,R001,R002,R003,R004,R005,R006,R007) Values" & _
                   "('" & strUserNum & "','" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "'," & _
                   CNULL(ChgSQL(stNewItem)) & "," & CNULL(ChgSQL(strFileN)) & ",'" & UCase(m_PrevForm.Name) & "')"
    Else
        strIns = "Update R090801_8 Set R005=" & CNULL(ChgSQL(stNewItem)) & ",R006=" & CNULL(ChgSQL(strFileN)) & _
                    " Where ID='" & strUserNum & "' And R001='" & strCP01 & "' And R002='" & strCP02 & "' And R003='" & strCP03 & "' And R004='" & strCP04 & "' "
    End If
    cnnConnection.Execute strIns
    SaveR090801_8 = True
    Exit Function
    
ErrHand:
    If Err.Number <> 32755 Then
        strMsg = stNewItem & "存檔失敗-" & Err.Description
    End If
End Function

Private Function DelR090801_8() As Boolean
    Dim i As Integer
    Dim strDel As String, strFileN As String, strCaseNo As String
    Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
On Error GoTo ErrHand

    DelR090801_8 = False: i = 0
    Do While i < lstAtt.ListCount
        If lstAtt.Selected(i) = True Then
            strCaseNo = Mid(lstAtt.List(i), 1, InStr(lstAtt.List(i), "->") - 1)
            strFileN = Replace(lstAtt.List(i), strCaseNo & "->", "")
            strCaseNo = Mid(strCaseNo, 1, InStr(lstAtt.List(i), ".") - 1)
           
            strCP01 = SystemNumber(strCaseNo, 1)
            strCP02 = SystemNumber(strCaseNo, 2)
            strCP03 = SystemNumber(strCaseNo, 3)
            strCP04 = SystemNumber(strCaseNo, 4)
            If ChkStatus(strCP01, strCP02, strCP03, strCP04, True) = True Then
                strDel = "Delete From R090801_8 Where ID='" & strUserNum & "' " & _
                              "And R001='" & strCP01 & "' And R002='" & strCP02 & "' And R003='" & strCP03 & "' And R004='" & strCP04 & "' " & _
                              "And R005=" & CNULL(ChgSQL(strFileN)) & " And R007='" & UCase(m_PrevForm.Name) & "' "
            Else
                strDel = "Update R090801_8 Set R005=null,R006=null Where ID='" & strUserNum & "' " & _
                              "And R001='" & strCP01 & "' And R002='" & strCP02 & "' And R003='" & strCP03 & "' And R004='" & strCP04 & "' " & _
                              "And R007='" & UCase(m_PrevForm.Name) & "' "
            End If
            cnnConnection.Execute strDel
         End If
         i = i + 1
    Loop
            
            
    DelR090801_8 = True
    Exit Function
    
ErrHand:
    If Err.Number <> 0 Then
        MsgBox strFileN & "移除失敗-" & Err.Description
    End If
End Function

'單一本所案號夾檔(原cmdAddAtt_Click程式)
Private Function AddFile_OnlyCaseNo() As Boolean
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim strFile As String
   Dim TempList As String 'Added by Lydia 2018/02/07 記錄去除路徑的檔案名稱
   Dim intJ As Integer, arrTmp As Variant 'Added by Lydia 2020/10/28
   
On Error GoTo ErrHnd
   
   AddFile_OnlyCaseNo = False
   'Added by Lydia 2018/02/07 讀取清單
   TempList = ""
   If lstAtt.ListCount > 0 Then
      ii = 0
      Do While ii < lstAtt.ListCount
        strExc(1) = lstAtt.List(ii)
        If InStrRev(strExc(1), " (") > 0 Then
            'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
            If UCase(Mid(strExc(1), InStrRev(strExc(1), " (") + 1, Len("(X86)"))) <> "(X86)" Then
            '2021/8/6 END
               strExc(1) = Left(strExc(1), InStrRev(strExc(1), " (") - 1)
            End If
        End If
        TempList = TempList & Mid(strExc(1), InStrRev(strExc(1), "\") + 1) & ";"
         ii = ii + 1
      Loop
   End If
   'end 2018/02/07
   'Modify by Amy 2025/05/08 +FC結案單
   If intFCState > 0 Then
      stFileName = "*.PDF;*.MSG" 'Modify by Amy 2025/06/20
   'Modify By Sindy 2016/6/24
   ElseIf bolNotPDF = True Then
      stFileName = "*.*"
   Else
   '2016/6/24 END
      stFileName = "*.PDF"
   End If
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      'Modify By Sindy 2016/6/24
      If bolNotPDF = True Then
         .Filter = "All Files (*.*)|*.*"
      Else
      '2016/6/24 END
         .Filter = "All Files (*.PDF)|*.PDF"
      End If
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      'Modify by Amy 2025/08/01
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer 'Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
'            'Add By Sindy 2018/2/8
'            If UCase(TypeName(m_PrevForm)) = UCase("frm090801") Or _
'               UCase(TypeName(m_PrevForm)) = UCase("frm210133") Then
'               MsgBox "只能選取一個電子檔!!!", vbExclamation
'               Exit Sub
'            End If
'            '2018/2/8 END
            
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            'Modified by Lydia 2015/10/30 +開庭通知
            'SaveSetting "TAIE", "FCP", EMP_回覆單 & "Dir", sFile(0)
             If bolOA Then
                 SaveSetting "TAIE", "L", "OA" & "Dir", sFile(0)
             Else
                 SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
             End If
             'Added by Lydia 2018/03/28
             If InStr(CStr(sFile(0)), "&") > 0 Then
                  MsgBox CStr(sFile(0)) & "\" & CStr(sFile(1)) & vbCrLf & vbCrLf & "【&】符號為系統保留字，不可使用於資料夾名稱！", vbExclamation
                  Exit Function
             End If
             'end 2018/03/28
            For ii = 1 To UBound(sFile)
               If InStr(CStr(sFile(ii)), "#") > 0 Or InStr(CStr(sFile(ii)), "&") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名！", vbExclamation
                  Exit Function
               End If
               
               'Add By Sindy 2015/1/27
               '檢查檔名規則
               'Modified by Lydia 2015/10/30
               'If PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), EMP_會修, "") = False Then
               'Modified by Lydia 2015/11/30 CFP常辦國家年費(延展費)預估報價-引用
               If lblCaseNo.Visible = True Then
                   If bolOA Then
                       If PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), "", "") = False Then Exit Function
                   ElseIf PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), EMP_會修, "") = False Then
                      Exit Function
                   End If
               End If 'end 2015/11/30
               '2015/1/27 END
               
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               
               'Modify by Amy 2025/05/08 +FC結案單
               If intFCState > 0 Then
                  If UCase(Mid(sFile(ii), InStrRev(sFile(ii), ".") + 1)) <> "PDF" _
                    And UCase(Mid(sFile(ii), InStrRev(sFile(ii), ".") + 1)) <> "MSG" Then
                     MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "只可加入PDF檔 或 MSG檔！", vbExclamation
                     Exit Function
                  End If
               'Modify By Sindy 2016/6/24
               ElseIf bolNotPDF = False Then
                  '只可加入PDF檔
                  If UCase(Mid(sFile(ii), InStrRev(sFile(ii), ".") + 1)) <> "PDF" Then
                     MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "只可加入PDF檔！", vbExclamation
                     Exit Function
                  End If
               End If
               '2016/6/24 END
               
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
               'Modify By Sindy 2015/2/11 改控制可以放5M以下的檔案
'               ElseIf f.Size > 512000 Then
'                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "檔案大小不可超過500KB！", vbExclamation
               'Modified by Lydia 2018/02/07 +是否限制檔案大小 And bolMax = True
               ElseIf f.Size > 5242880 And bolMax = True Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "檔案大小不可超過 5MB！", vbExclamation
               '2015/2/11 END
                  Exit Function
               End If
               TempList = TempList & Mid(stFileName, InStrRev(stFileName, "\") + 1) & ";" 'Added by Lydia 2018/02/07
               AddListX lstAtt, stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS")
            Next ii
            
         Else
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Or InStr(strFile, "&") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名！", vbExclamation
               Exit Function
            End If
            'Added by Lydia 2018/03/28 資料夾名稱排除"&"
            If InStr(.FileName, "&") > 0 Then
               MsgBox .FileName & vbCrLf & vbCrLf & "【&】符號為系統保留字，不可使用於資料夾名稱！", vbExclamation
               Exit Function
            End If
            'end 2018/03/28
            
            'Modify by Amy 2025/05/08 +FC結案單
            If intFCState > 0 Then
               If UCase(Mid(strFile, InStrRev(strFile, ".") + 1)) <> "PDF" _
                  And UCase(Mid(strFile, InStrRev(strFile, ".") + 1)) <> "MSG" Then
                  MsgBox strFile & vbCrLf & vbCrLf & "只可加入PDF檔 或 MSG檔！", vbExclamation
                  Exit Function
               End If
            'Modify By Sindy 2016/6/24
            ElseIf bolNotPDF = False Then
               '只可加入PDF檔
               If UCase(Mid(strFile, InStrRev(strFile, ".") + 1)) <> "PDF" Then
                  MsgBox strFile & vbCrLf & vbCrLf & "只可加入PDF檔！", vbExclamation
                  Exit Function
               End If
            End If
            '2016/6/24 END
            
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     'Modified by Lydia 2015/10/30 +開庭通知
                     'SaveSetting "TAIE", "FCP", EMP_回覆單 & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     If bolOA Then
                        SaveSetting "TAIE", "L", "OA" & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     Else
                        SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     End If
                     Exit For
                  End If
               Next ii
            End If
            
            'Add By Sindy 2015/1/27
            '檢查檔名規則
            'Modified by Lydia 2015/10/30
            'If PUB_ChkEmpFlowFNMRule(lblCaseNo, strFile, EMP_會修, "") = False Then
            'Modified by Lydia 2015/11/30 CFP常辦國家年費(延展費)預估報價-引用
            If lblCaseNo.Visible = True Then
                If bolOA Then
                   If PUB_ChkEmpFlowFNMRule(lblCaseNo, strFile, "", "") = False Then Exit Function
                ElseIf PUB_ChkEmpFlowFNMRule(lblCaseNo, strFile, EMP_會修, "") = False Then
                   Exit Function
                End If
            End If 'end 2015/11/30
            '2015/1/27 END
            
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
            'Modify By Sindy 2015/2/11 改控制可以放5M以下的檔案
'            ElseIf f.Size > 512000 Then
'               MsgBox strFile & vbCrLf & vbCrLf & "檔案大小不可超過500KB！", vbExclamation
             'Modified by Lydia 2018/02/07 +是否限制檔案大小 And bolMax = True
            ElseIf f.Size > 5242880 And bolMax = True Then
               MsgBox strFile & vbCrLf & vbCrLf & "檔案大小不可超過 5MB！", vbExclamation
            '2015/2/11 END
               Exit Function
            'Added by Lydia 2019/12/23 控制檔名不可超過100字; ex.FCP-062455在TrackingNo資料夾放”AP rl    Instructions   - Issue Fee and new divisional patent application - Taiwanese Patent Application No 107106625 - Your Ref  P187543 TW 01 DeakM DeZeeuwR and P222272 TW 01DeakM DeZeeuwR - Our Ref  FCP-058411  ACK 601 .msg”
            'Modified by Lydia 2020/05/06 控制檔名不可超過80字 ; 因為若不是msg檔,最終會以"案號+案件性質+原檔名"上傳
            ElseIf CheckLengthIsOK(f.Name, 80) = False Then
               Exit Function
            End If
            TempList = TempList & Mid(stFileName, InStrRev(stFileName, "\") + 1) & ";" 'Added by Lydia 2018/02/07
            AddListX lstAtt, stFileName & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS")
         End If
      End If
   End With
   
   'Added by Lydia 2018/02/07 限制檔案名稱長度
   If TempList <> "" Then TempList = Mid(TempList, 1, Len(TempList) - 1)
   If m_FLmax = 0 Or (m_FLmax > 0 And GetTextLength(TempList) < m_FLmax) Then
       cmdOK(0).Enabled = True
   Else
       'Modified by Lydia 2020/10/28 總長度分成兩種:客戶提供文件frm060120以全部檔案的名稱計算
       'MsgBox "去除檔案路徑的檔案名稱總長度超過" & m_FLmax & "字元 !" & _
                  vbCrLf & "請移除檔案後縮短檔案名稱，再重新加入!"
       'cmdOK(0).Enabled = False
       strExc(0) = ""
       If TypeName(m_PrevForm) <> "frm060120" Then '其他以各檔案計算
           arrTmp = Split(TempList, ";")
           For intJ = 0 To UBound(arrTmp)
              If Trim("" & arrTmp(intJ)) <> "" Then
                    If GetTextLength("" & arrTmp(intJ)) >= m_FLmax Then
                         strExc(0) = vbCrLf & "" & arrTmp(intJ)
                    End If
              End If
           Next intJ
       Else
           strExc(0) = vbCrLf & Replace(TempList, ";", vbCrLf)
       End If
       If strExc(0) <> "" Then
            MsgBox "去除檔案路徑的檔案名稱總長度超過" & m_FLmax & "字元 !" & _
                       vbCrLf & "請移除檔案後縮短檔案名稱，再重新加入!" & vbCrLf & _
                       "檔案名稱：" & strExc(0)
            cmdOK(0).Enabled = False
       End If
       'end 2020/10/28
   End If
   'end 2018/02/07
   AddFile_OnlyCaseNo = True
   Exit Function
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

'Add by Amy 2018/06/06
'多案號多檔案操作
Private Function AddFile_Mult() As Boolean
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim strFile As String, strFSize As String, strDateLastModified As String, strCaseNo As String
   Dim strMsg As String, strNoExists As String, strAllMsg As String, strNoCase As String
   
On Error GoTo ErrHnd
   
   AddFile_Mult = True
   stFileName = "*.PDF"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.PDF)|*.PDF"
      If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
         .InitDir = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
      Else
         .InitDir = PUB_Getdesktop
      End If
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         '選取多檔
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            SaveSetting "TAIE", "FCP", EMP_回覆單 & "Dir", sFile(0)
            For ii = 1 To UBound(sFile)
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               If ChkAddFile(sFile(ii), stFileName, strFSize, strDateLastModified, strCaseNo, strMsg, strNoExists) = False Then
                  If strMsg <> MsgText(601) Then strAllMsg = strAllMsg & strMsg & vbCrLf
                  If strNoExists <> MsgText(601) Then strNoCase = strNoCase & strNoExists & vbCrLf
               End If
               If strMsg & strNoExists = MsgText(601) Then
                  If SaveR090801_8(stFileName & " (" & Round(strFSize / 1024, 2) & " KB)" & " #" & Format(strDateLastModified, "YYYYMMDDHHMMSS"), strCaseNo, strMsg) = False Then
                      If strMsg <> MsgText(601) Then strAllMsg = strAllMsg & strMsg & vbCrLf
                  End If
               End If
            Next ii
         '只選一個檔
         Else
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
               For ii = Len(.FileName) To 1 Step -1
                  If Mid(Trim(.FileName), ii, 1) = "\" Then
                     SaveSetting "TAIE", "FCP", EMP_回覆單 & "Dir", Mid(Trim(.FileName), 1, ii - 1)
                     Exit For
                  End If
               Next ii
            End If
            stFileName = .FileName
            If ChkAddFile(strFile, stFileName, strFSize, strDateLastModified, strCaseNo, strMsg, strNoExists) = False Then
                If strMsg <> MsgText(601) Then strAllMsg = strAllMsg & strMsg & vbCrLf
                If strNoExists <> MsgText(601) Then strNoCase = strNoCase & strNoExists & vbCrLf
            End If
            If strMsg & strNoExists = MsgText(601) Then
                If SaveR090801_8(stFileName & " (" & Round(strFSize / 1024, 2) & " KB)" & " #" & Format(strDateLastModified, "YYYYMMDDHHMMSS"), strCaseNo, strMsg) = False Then
                    If strMsg <> MsgText(601) Then strAllMsg = strAllMsg & strMsg & vbCrLf
                End If
            End If
         End If
      End If
   End With
   If strAllMsg & strNoCase <> MsgText(601) Then
        MsgBox strAllMsg & IIf(strNoCase = "", "", vbCrLf & strNoCase)
   End If
   Exit Function
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
   AddFile_Mult = False '無法預期錯誤或已選取「取消」設False
End Function

'檢查使用者點選的檔案
Private Function ChkAddFile(ByVal strChkFile As String, ByVal strFullFile As String, ByRef strFSize As String, ByRef strDateLastModified As String, _
                            ByRef strCaseNo As String, ByRef strMsg As String, ByRef strNoExists As String) As Boolean
    Dim j As Integer
    Dim bolChkFileOK As Boolean, bolFileExists As Boolean
    Dim fs, f

    ChkAddFile = False: strMsg = "": strNoExists = ""
    strCaseNo = PUB_AnalysisFileNmGetCaseNO(strChkFile)
    If InStr(strChkFile, "#") > 0 Or InStr(strChkFile, "&") > 0 Then
        strMsg = strChkFile & vbCrLf & vbCrLf & "【#和&】符號為系統保留字，不可使用於檔案命名！"
        Exit Function
    End If
   
   For j = 1 To GrdAtt.Rows - 1
        If GrdAtt.TextMatrix(j, 0) = strCaseNo Then
            ChkAddFile = True
            '檢查檔名規則
            If PUB_ChkEmpFlowFNMRule(CStr(GrdAtt.TextMatrix(j, 0)), strChkFile, EMP_會修, "", , , , False) = True Then
                bolChkFileOK = True
                If UCase(GetFileName(strChkFile)) = UCase(GrdAtt.TextMatrix(j, 2)) Then
                    bolFileExists = True
                    Exit For
                End If
            End If
        End If
   Next j
   If ChkAddFile = False Then
      strNoExists = strChkFile & " 檔案有誤，無此案件！"
      Exit Function
   End If
   If bolChkFileOK = False Then
      strMsg = strChkFile & " 檔案命名不符規定，請修改檔名！"
      ChkAddFile = False
      Exit Function
   End If
   If bolFileExists = True Then
      strMsg = strChkFile & " 附件已存在！"
      ChkAddFile = False
      Exit Function
   End If
  
   '只可加入PDF檔
   If UCase(Mid(strChkFile, InStrRev(strChkFile, ".") + 1)) <> "PDF" Then
      strMsg = strChkFile & " 只可加入PDF檔！"
      ChkAddFile = False
      Exit Function
   End If
   '檢查檔案是否正在使用中
   If PUB_ChkFileOpening(strFullFile) = True Then
      strMsg = strFullFile & " 檔案正在使用中（請關閉），方可繼續操作。"
      ChkAddFile = False
      Exit Function
   End If
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile(strFullFile)
   '檔案大小為 0 KB 有誤
   If f.Size = 0 Then
      strMsg = strChkFile & MsgText(9221)
      ChkAddFile = False
      Exit Function
   '控制可以放5M以下的檔案
   ElseIf f.Size > 5242880 Then
      strMsg = strChkFile & " 檔案大小不可超過 5MB！"
      ChkAddFile = False
      Exit Function
   'Added by Lydia 2019/12/23 控制檔名不可超過100字; ex.FCP-062455在TrackingNo資料夾放”AP rl    Instructions   - Issue Fee and new divisional patent application - Taiwanese Patent Application No 107106625 - Your Ref  P187543 TW 01 DeakM DeZeeuwR and P222272 TW 01DeakM DeZeeuwR - Our Ref  FCP-058411  ACK 601 .msg”
   'Modified by Lydia 2020/05/06 控制檔名不可超過80字 ; 因為若不是msg檔,最終會以"案號+案件性質+原檔名"上傳
   ElseIf CheckLengthIsOK(f.Name, 80) = False Then
      ChkAddFile = False
      Exit Function
   End If
   strFSize = f.Size
   strDateLastModified = f.DateLastModified

   ChkAddFile = True
End Function

'判斷需新增或修改暫存檔內容
Private Function ChkStatus(ByVal stCP01 As String, ByVal stCP02 As String, ByVal stCP03 As String, ByVal stCP04 As String, Optional ByVal bolDel As Boolean = False) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    ChkStatus = False
    strQ = "Select * From R090801_8 Where ID='" & strUserNum & "' And R007='" & UCase(m_PrevForm.Name) & "' " & _
              "And R001='" & stCP01 & "' And R002='" & stCP02 & "' And R003='" & stCP03 & "' And R004='" & stCP04 & "' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If RsQ.RecordCount > 1 Then
            '增加/刪除時同案號多筆用新增/刪除語法
            ChkStatus = True
        ElseIf bolDel = False And RsQ.RecordCount = 1 And Not IsNull(RsQ.Fields("R005")) Then
            '增加時只有一筆且加過檔案用新增語法
            ChkStatus = True
        End If
    End If
    RsQ.Close
End Function
'end 2018/06/06

