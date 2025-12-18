VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm100101_2_1 
   Caption         =   "客戶附件"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7770
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7755
      Begin VB.ComboBox cboAtt 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frm100101_2_1.frx":0000
         Left            =   1140
         List            =   "frm100101_2_1.frx":0002
         Style           =   2  '單純下拉式
         TabIndex        =   1
         Top             =   30
         Width           =   6615
      End
      Begin VB.Label lblAttCnt 
         Alignment       =   2  '置中對齊
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  '單線固定
         Caption         =   " PDF:(0)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   0
         TabIndex        =   2
         Top             =   30
         Width           =   1140
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6120
      Left            =   -30
      TabIndex        =   3
      Top             =   390
      Width           =   7785
      ExtentX         =   13732
      ExtentY         =   10795
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frm100101_2_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Sonia 2022/1/24 Form2.不用改;Form2.0已檢查 (無需修改的物件)
'Created by Morgan 2019/6/20
Public stFileName As String
Public stFileDescs As String
Public stSavePath As String

Private Sub cboAtt_Click()
   Dim hLocalFile As Long
   Dim arrFileName() As String
   Dim strFile As String, strFileType As String 'Add By Sindy 2020/10/13
   
   'Add By Sindy 2020/10/13
   If InStrRev(cboAtt.List(cboAtt.ListIndex), Chr(9) & "(") > 0 Then
      strFile = Left(cboAtt.List(cboAtt.ListIndex), InStrRev(cboAtt.List(cboAtt.ListIndex), Chr(9) & "(") - 1)
   End If
   If InStrRev(strFile, ".") > 0 Then
      strFileType = UCase(Mid(strFile, InStrRev(strFile, ".") + 1))
   End If
   If strFileType <> "PDF" Then
      SetAttr stSavePath & "\" & strFile, vbReadOnly '檔案設定成唯讀屬性
      ShellExecute hLocalFile, "open", stSavePath & "\" & strFile, vbNullString, vbNullString, 1
   Else
   '2020/10/13 END
      arrFileName = Split(cboAtt.List(cboAtt.ListIndex), Chr(9))
      WebBrowser1.Navigate stSavePath & "\" & arrFileName(0): DoEvents
   End If
End Sub

Private Sub Form_Activate()
   Static bolActivated As Boolean
   
   If bolActivated = False Then
      bolActivated = True
      WebBrowser1.Navigate stSavePath & "\" & stFileName
      SetAttList stFileDescs
   End If
End Sub


Private Sub Form_Load()
   'Modified by Morgan 2021/8/18 載入前次結束時的大小及位置
   'MoveFormToCenter Me
   PUB_SetPdfForm Me
   'end 2021/8/18
   WebBrowser1.Navigate "about:blank"
End Sub

Private Sub Form_Resize()
   Dim lngWidth As Long, lngHeight As Long
   
On Error GoTo ErrHnd
   
   If Me.WindowState = 0 Or Me.WindowState = 2 Then
      lngWidth = Me.Width - WebBrowser1.Left - 150
      lngHeight = Me.Height - Frame4.Height - 450
      If lngWidth > 0 And lngHeight > 0 Then
         WebBrowser1.Width = lngWidth
         WebBrowser1.Height = lngHeight
         Frame4.Left = WebBrowser1.Left
         Frame4.Width = WebBrowser1.Width
         'cboAtt.Width = Frame4.Width - cboAtt.Left
         lngWidth = Frame4.Width - lblAttCnt.Width
         If lngWidth > 0 Then
            cboAtt.Width = lngWidth
         End If
      End If
   End If
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description, vbExclamation, "Form_Resize"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SavePdfForm Me 'Added by Morgan 2021/8/18 紀錄視窗最後的大小及位置
   PUB_KillAttach stSavePath
   Set frm100101_2_1 = Nothing
End Sub

Private Sub SetAttList(Optional pItems As String)
   Dim arrItem() As String
   Dim ii As Integer, iAttCnt As Integer

   'Modify By Sindy 2020/10/13
   If Me.Caption = "多案卷宗區附件" Then
      cboAtt.Clear: lblAttCnt = "File:(0)"
   Else
   '2020/10/13 END
      cboAtt.Clear: lblAttCnt = " PDF:(0)"
   End If
   If pItems <> "" Then
      arrItem = Split(pItems, ";")
      For ii = LBound(arrItem) To UBound(arrItem)
         If arrItem(ii) <> "" Then
            cboAtt.AddItem arrItem(ii)
            iAttCnt = iAttCnt + 1
         End If
      Next
      'Modify By Sindy 2020/10/13
      If Me.Caption = "多案卷宗區附件" Then
         lblAttCnt = "File:(" & iAttCnt & ")"
      Else
      '2020/10/13 END
         lblAttCnt = " PDF:(" & iAttCnt & ")"
      End If
   End If
End Sub

Private Sub lblAttCnt_Click()
   SendMessage cboAtt.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
End Sub
