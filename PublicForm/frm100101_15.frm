VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_15 
   BorderStyle     =   1  '單線固定
   Caption         =   "往來記錄瀏覽區"
   ClientHeight    =   5940
   ClientLeft      =   1810
   ClientTop       =   2590
   ClientWidth     =   9470
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9470
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm100101_15.frx":0000
      Height          =   3495
      Left            =   30
      TabIndex        =   26
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8696
      _ExtentY        =   6156
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
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
   Begin VB.CommandButton cmdMove 
      Caption         =   "移檔"
      Height          =   345
      Left            =   2595
      TabIndex        =   7
      Top             =   1290
      Width           =   525
   End
   Begin VB.CommandButton cmdReName 
      Caption         =   "更名"
      Height          =   345
      Left            =   1995
      TabIndex        =   6
      Top             =   1290
      Width           =   525
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "刪除"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   1395
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   1290
      Width           =   525
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "新增"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   795
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   1290
      Width           =   525
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "新增行事曆"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   2370
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   1680
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "新增往來記錄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   1170
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   1680
      Width           =   1140
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm100101_15.frx":0015
      Left            =   990
      List            =   "frm100101_15.frx":0017
      TabIndex        =   1
      Top             =   660
      Width           =   4005
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢(&Q)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   30
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   1290
      Width           =   690
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "點我展開"
      Height          =   345
      Left            =   4950
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   0
      Width           =   4515
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5595
      Left            =   4950
      TabIndex        =   20
      Top             =   330
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   9869
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   405
      Left            =   30
      TabIndex        =   21
      Top             =   5520
      Width           =   4905
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "多檔預覽"
         Height          =   315
         Index           =   1
         Left            =   30
         TabIndex        =   13
         Top             =   60
         Width           =   870
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   315
         Index           =   0
         Left            =   930
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   60
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "PS:點選聯絡人編號才會顯示聯絡人名稱"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1740
         TabIndex        =   24
         Top             =   60
         Width           =   3105
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "關閉預覽"
      Height          =   195
      Left            =   3960
      TabIndex        =   2
      Top             =   450
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆(&N)"
      CausesValidation=   0   'False
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   3195
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   1290
      Width           =   870
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "基本資料(&C)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   30
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   1680
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   4140
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   1290
      Width           =   750
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4650
      Top             =   1470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstUsers 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   220
      Index           =   0
      ItemData        =   "frm100101_15.frx":0019
      Left            =   3750
      List            =   "frm100101_15.frx":0020
      MultiSelect     =   1  '簡易多重選取
      Sorted          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1770
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSForms.TextBox txtCR06 
      Height          =   285
      Left            =   990
      TabIndex        =   0
      Top             =   960
      Width           =   4005
      VariousPropertyBits=   671105051
      Size            =   "7064;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "主　旨："
      Height          =   180
      Left            =   45
      TabIndex        =   23
      Top             =   990
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "往來類別："
      Height          =   180
      Left            =   45
      TabIndex        =   22
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人："
      Height          =   180
      Left            =   45
      TabIndex        =   19
      Top             =   360
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "往來對象："
      Height          =   180
      Left            =   45
      TabIndex        =   18
      Top             =   60
      Width           =   900
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Left            =   990
      TabIndex        =   17
      Top             =   30
      Width           =   3930
      VariousPropertyBits=   27
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   285
      Left            =   990
      TabIndex        =   16
      Top             =   360
      Width           =   2970
      VariousPropertyBits=   27
      Size            =   "5239;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm100101_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/06 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
'Create by Morgan 2007/12/14
Option Explicit

Public CRdateF As String, CRdateT As String, CRtype As String '2008/11/19 add by sonia
Public cmdState As Integer
Public m_quyDataKind As Integer '0.國外 1.國內
'Dim m_bLanguage As String      '2008/12/9 ADD BY SONIA

'附件宣告區
Dim m_AttachPath As String
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
Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

'Const GrdMaxW = 9500
'Const GrdMinW = 4930
Dim GrdMaxW As Long
Dim GrdMinW As Long
Dim m_bolDblClick As Boolean 'Add By Sindy 2019/3/6
Dim m_mouseRow As Long, m_mouseRowOld As Long 'Add By Sindy 2019/3/6
Private Const cTableName As String = "CONTACTFILE" 'Add by Sindy 2019/3/7 指定FTP資料夾名稱
Dim m_PCU51 As String 'Add By Sindy 2019/7/25 國外潛在客戶的國內外權限
Dim m_CU13 As String 'Add By Sindy 2020/5/21
Dim varTmp As Variant 'Add By Sindy 2020/9/2
Dim bolGrpSpec As Boolean 'Added by Lydia 2020/11/30 是否為特殊群組(國內智權部往來記錄查詢人員)
Dim m_strCRexcept As String 'Added by Lydia 2025/08/08
Public m_pub_QL05 As String 'Add By Sindy 2025/8/27 只記錄於此Form


Private Sub Check1_Click()
   If Check1.Value = 1 Then
      WebBrowser1.Navigate "about:blank"
      Command4.Visible = False
      WebBrowser1.Visible = False
      GRD1.Width = Me.Width - 20 'GrdMaxW
      'Call SetGrd(False)
   Else
      Command4.Visible = True
      WebBrowser1.Visible = True
      GRD1.Width = GrdMinW
      'Call SetGrd(False)
   End If
End Sub

'Add By Sindy 2019/11/6 移檔
Private Sub cmdMove_Click()
Dim strSaveFiles As String
Dim strRecvNo As String
Dim ii As Integer
Dim strTmp As String 'Added by Lydia 2020/11/30

   For ii = 1 To GRD1.Rows - 1
      GRD1.row = ii
      GRD1.col = 1
      If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v") And _
         Trim(GRD1.TextMatrix(ii, 8)) <> "" Then
         If Trim(GRD1.TextMatrix(ii, 8)) = "" Then
            MsgBox "請勾選有電子檔的資料!!!"
            Exit Sub
         End If
         
         strSaveFiles = strSaveFiles & "&" & Trim(GRD1.TextMatrix(ii, 13))
         If InStr(strRecvNo, Trim(GRD1.TextMatrix(ii, 13))) = 0 Then
            strRecvNo = strRecvNo & ",'" & Trim(GRD1.TextMatrix(ii, 13)) & "'"
         End If
         'Modified by Morgan 2023/10/17 未清除檔名後其他資訊
         'strSaveFiles = strSaveFiles & "  " & GetFileName(Trim(grd1.TextMatrix(ii, 8)))
         strSaveFiles = strSaveFiles & "  " & PUB_MGridGetValue(ii, "CF02", GRD1)
         'end 2023/10/17
         'Modify By Sindy 2024/10/1 都要抓建檔人
         'If m_quyDataKind = 1 Then
            strTmp = Trim(GRD1.TextMatrix(ii, 16)) 'Added by Lydia 2020/11/30 建檔人
         'End If
         '2024/10/1 END
      End If
   Next ii
   If strSaveFiles = "" Then
      MsgBox "請勾選一筆欲移動的電子檔!!!"
      Exit Sub
   End If
   strSaveFiles = Mid(strSaveFiles, 2)
   strRecvNo = Mid(strRecvNo, 2)
   'Added by Lydia 2020/11/30 增加權限判斷
   If m_quyDataKind = 0 Then
      'Modify By Sindy 2024/10/1 傳入建檔人
      If PUB_CheckModifyLimit_frm140402(m_PCU51, strTmp) = False Then Exit Sub
   Else
      '特殊群組只有查詢功能 = False
      If PUB_CheckModifyLimit_frm100101_19(m_CU13, Me.Tag, strTmp, False) = False Then Exit Sub
   End If
   'end 2020/11/30
   
   Call frm100101_15_1.SetParent(Me)
   frm100101_15_1.m_strSaveFiles = strSaveFiles
   frm100101_15_1.lblKey = Me.Label3
   frm100101_15_1.lblKey.Tag = Left(Me.Tag, 8)
   If frm100101_15_1.QueryData(0) = True Then
      frm100101_15_1.Show vbModal
   End If
End Sub

Public Sub cmdok_Click(Index As Integer)
   cmdState = Index
   Me.PubShowNextData
End Sub

'Add By Sindy 2019/11/6 更名
Private Sub cmdReName_Click()
Dim intChkCnt As Integer
Dim strNewFile As String, strOldCF02 As String
Dim strCF01 As String, strCF02 As String
Dim ii As Integer
Dim strTmp As String 'Added by Lydia 2020/11/30

On Error GoTo ErrHnd
   
   intChkCnt = 0
   For ii = 1 To GRD1.Rows - 1
      GRD1.row = ii
      GRD1.col = 1
      If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v") And _
         Trim(GRD1.TextMatrix(ii, 8)) <> "" Then
         strCF01 = Trim(GRD1.TextMatrix(ii, 13))
         strCF02 = Trim(GRD1.TextMatrix(ii, 14))
         'Modify By Sindy 2024/10/1 都要抓建檔人
         'If m_quyDataKind = 1 Then
            strTmp = Trim(GRD1.TextMatrix(ii, 16)) 'Added by Lydia 2020/11/30 建檔人
         'End If
         '2024/10/1 END
         If strCF01 <> "" And strCF02 <> "" Then
            intChkCnt = intChkCnt + 1
         End If
      End If
   Next ii
   strOldCF02 = strCF02
   If intChkCnt = 0 Or strOldCF02 = "" Then
      MsgBox "請勾選一筆欲更名的電子檔!!!"
      Exit Sub
   ElseIf intChkCnt > 1 Then
      MsgBox "只可勾選一筆資料做更名!!!"
      Exit Sub
   End If
   'Added by Lydia 2020/11/30 增加權限判斷
   If m_quyDataKind = 0 Then
      'Modify By Sindy 2024/10/1 傳入建檔人
      If PUB_CheckModifyLimit_frm140402(m_PCU51, strTmp) = False Then Exit Sub
   Else
      '特殊群組只有查詢功能 = False
      If PUB_CheckModifyLimit_frm100101_19(m_CU13, Me.Tag, strTmp, False) = False Then Exit Sub
   End If
   'end 2020/11/30
   
ShowInput:
   strNewFile = InputBox("確定是否「更名」？", "更名！", strOldCF02)
   If UCase(strNewFile) = UCase(strCF02) Then
      MsgBox "請輸入欲更改的電子檔名!!!"
      GoTo ShowInput
   End If
   
   If Trim(strNewFile) = "" Then
      Exit Sub
   Else
      If Right(UCase(strNewFile), 4) = UCase(".del") Then
         MsgBox "新的電子檔名最後面不能是(.del) !!!"
         strOldCF02 = strNewFile
         GoTo ShowInput
      End If
      
      For ii = 1 To GRD1.Rows - 1
         GRD1.row = ii
         GRD1.col = 1
         If GRD1.TextMatrix(ii, 13) <> "" Then GRD1.Tag = GRD1.TextMatrix(ii, 13)
         If GRD1.Tag = strCF01 Then
            If UCase(strNewFile) = UCase(GetFileName(Trim(GRD1.TextMatrix(ii, 8)))) Then
               MsgBox "檔名重覆，請輸入欲更改的電子檔名!!!"
               strOldCF02 = strNewFile
               GoTo ShowInput
            End If
         End If
      Next ii
      
      strSql = "update CONTACTFILE set cf02='" & strNewFile & "' where cf01='" & strCF01 & "' and cf02='" & strCF02 & "'"
      Pub_SaveLog strUserNum, strSql
      cnnConnection.Execute strSql
      
      Call cmdok_Click(3)
   End If
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Command4_Click()
   If Command4.Caption = "點我展開" Then
      RePosForm True
   Else
      RePosForm False
   End If
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/27 還原此Form的查詢條件記錄
   If Me.WindowState = 0 Then Me.WindowState = 2 '最大化
   'Modified by Lydia 2022/06/10 Debug: 跑執行檔，用83004的帳號查詢X77693010的往來記錄，會出現「當有強制回應表單顯示時，無法再顯示非強制回應表單」
   'If cmdState <> 2 Then 'Sindy 2022/4/28 按下查詢”基本資料”時，不要一直重新啟動查詢功能，會影響到查不出第3筆以上的資料
   If cmdState <> 2 And cmdState > 0 Then
      Call cmdok_Click(3)
   End If
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   WebBrowser1.Navigate "about:blank"
   
   GrdMaxW = Me.Width - 20
   GrdMinW = Me.GRD1.Width - 20
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   '檢查在程式執行目錄及C:\WINDOWS\system32裡,是否有合併的執行檔pdftk.exe
   'App.path = "C:\Program Files\tepatpro"
   If Dir(App.path & "\pdftk.exe") = "" And _
      Dir("C:\WINDOWS\system32\pdftk.exe") = "" Then
      cmdOpenAtt(1).Visible = False
   End If
   If Pub_StrUserSt03 = "M51" And InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") Then
      Me.Height = 6600
   Else
      Me.Height = 6120
   End If
   Me.WindowState = 2 '最大化
   Check1.Value = 1 '關閉預覽
   
   cmdState = -1
'   '2008/12/9 ADD BY SONIA
'   m_bLanguage = IsUserHasRightOfLanguage
'   '有值才可查潛在客戶往來記錄 Y不限語文 J限日文 E限非日文
   
   'Add By Sindy 2019/3/27
   'Modify By Sindy 2020/5/21 Mark
'   If CheckUse("frm140404", strExec, False) = True Or _
'      CheckUse("frm210129", strExec, False) = True Then
'      cmdOK(4).Enabled = True '新增往來記錄
'      cmdOK(6).Enabled = True '新增
'      cmdOK(7).Enabled = True '刪除
'      cmdReName.Enabled = True '更名
'      cmdMove.Enabled = True '移檔
'   Else
'      cmdOK(4).Enabled = False
'      cmdOK(6).Enabled = False
'      cmdOK(7).Enabled = False
'      cmdReName.Enabled = False
'      cmdMove.Enabled = False
'   End If

   If CheckUse("frm06010610", strExec, False) = True Then
      cmdok(5).Visible = True
   Else
      cmdok(5).Visible = False
   End If
   '2019/3/27 END
   
   Call SetCombo1 'Add By Sindy 2019/2/23
   
   'Added by Lydia 2020/11/30 是否為特殊群組
   bolGrpSpec = False
   If InStr(Pub_GetSpecMan("國內智權部往來記錄查詢人員", False) & ",", strUserNum) > 0 Then
       bolGrpSpec = True
   End If
   'end 2020/11/30
   
   m_strCRexcept = Pub_GetCRExceptNo(Me.Name) 'Added by Lydia 2025/08/08
   
   m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/27 記錄此Form的查詢條件
End Sub

'Add By Sindy 2019/2/23 往來類別
Private Sub SetCombo1()
   Combo1.Clear
   strSql = "select ac02,ac03 from allcode where ac01='11'" & _
            " order by ac02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         Combo1.AddItem RsTemp.Fields("ac02") & " " & RsTemp.Fields("ac03")
         RsTemp.MoveNext
      Loop
   End If
End Sub

Private Sub Form_Resize()
   If Me.WindowState = 0 Or Me.WindowState = 2 Then
      If Command4.Caption = "點我展開" Then
         RePosForm False
      Else
         RePosForm True
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
   Set frm100101_15 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next

   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.*"
   End If
End Sub

Private Sub RePosForm(pFull As Boolean)
   Static lngLeft As Long
   
   If Forms(0).WindowState <> 1 Then
      If lngLeft = 0 Then lngLeft = Command4.Left
      If pFull = True Then
         WebBrowser1.Left = 0
         WebBrowser1.Width = Me.Width - 90
         WebBrowser1.Height = Me.Height - Command4.Height - 350
         Command4.Caption = "點我還原"
      Else
         WebBrowser1.Left = lngLeft
         Command4.Caption = "點我展開"
      End If
      WebBrowser1.Width = Me.Width - 90 - WebBrowser1.Left
      WebBrowser1.Height = Me.Height - Command4.Height - 350
      Command4.Left = WebBrowser1.Left
      Command4.Width = WebBrowser1.Width
      
      GRD1.Height = Me.Height - GRD1.Top - Frame1.Height - 300
      If Check1.Value = 1 Then
         GRD1.Width = Me.Width - 150
      Else
         GRD1.Width = GrdMinW
      End If
      Frame1.Top = GRD1.Top + GRD1.Height - 50
   End If
End Sub

Public Sub PubShowNextData()
   Dim i As Integer, j As Integer
   Dim strTmp As String
   
   Select Case cmdState
      Case 0 '結束
         fnCloseAllFrm100
      Case 1 '下一筆
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 2
         Me.Enabled = False
         With GRD1
            For i = 1 To .Rows - 1
               .col = 1 '0 Modify By Sindy 2022/1/19
               .row = i
               'Modify By Sindy 2022/1/19
               'If Trim(.Text) = "V" Then
               'Modify By Sindy 2019/7/15
               'If Trim(.Text) <> "" Then
               If .CellBackColor = &HFFC0C0 Or Trim(GRD1.TextMatrix(i, 0)) <> "" Then
               '2022/1/19 END
               '.col = 2
               'If .CellBackColor = &HFFC0C0 Then
                  '.col = 0
               '2019/7/15 END
                  .col = 0
                  .Text = ""
                  For j = 0 To .Cols - 1
                     'If j <> 1 Then
                        .col = j
                        .CellBackColor = QBColor(15)
                     'End If
                  Next j
                  .col = 13
                  strTmp = .Text
                  If Not IsNull(strTmp) Then
                     If fnSaveParentForm(Me) = False Then
                        Me.Enabled = True
                        Exit Sub
                     End If
                     Screen.MousePointer = vbHourglass
                     'Modify By Sindy 2020/5/17
                     'If m_quyDataKind = 1 Then '國內
                     If GRD1.TextMatrix(i, 18) = "C" Then '國內
                        frm100101_19.Show
                        frm100101_19.Tag = Pub_RplStr(strTmp)
                        frm100101_19.m_pub_QL05 = ";編號：" & Me.Tag 'Add By Sindy 2025/8/27
                        frm100101_19.StrMenu
                     Else
                     '2020/5/17 END
                        '國外
                        frm100101_16.Show
                        frm100101_16.Tag = Pub_RplStr(strTmp)
                        frm100101_16.m_pub_QL05 = ";編號：" & Me.Tag 'Add By Sindy 2025/8/27
                        frm100101_16.StrMenu
                     End If
                     Screen.MousePointer = vbDefault
                     Me.Enabled = True
                     Exit Sub
                  End If
               End If
            Next i
         End With
         Me.Enabled = True
         
      Case 3 '查詢
         'Modify By Sindy 2020/5/17
         If m_quyDataKind = 1 Then
            Call StrMenu2(True) '國內往來記錄
         Else
         '2020/5/17 END
            Call StrMenu(True) '國外往來記錄
         End If
         
      'Add By Sindy 2019/3/27
      Case 4 '新增往來記錄
         'Modify By Sindy 2020/5/17
         If m_quyDataKind = 1 Then '國內往來記錄
'            If CheckUse("frm210129", strExec) = True Then
               frm210129.Show
               frm210129.OnAction vbKeyF2
               frm210129.txtCOR(3) = Me.Tag
'            End If
         Else
         '2020/5/17 END
            '國外往來記錄
            If CheckUse("frm140404", strExec) = True Then
               frm140404.Show
               frm140404.OnAction vbKeyF2
               frm140404.txtCR(3) = Me.Tag
            End If
         End If
            
      Case 5 '新增行事曆
         If CheckUse("frm06010610", strExec) = True Then
            frm06010610.Show
            frm06010610.OnAction vbKeyF2
         End If
      '2019/3/27 END
      
      'Add By Sindy 2019/5/31
      Case 6 '新增電子檔
         GRD1.row = m_mouseRow
         GRD1.col = 1
         
         'Modify By Sindy 2019/7/25
         'If PUB_CheckModifyLimit_frm140402(GRD1.TextMatrix(m_mouseRow, 16), "M") = False Then Exit Sub
         If m_quyDataKind = 0 Then 'Add By Sindy 2020/5/19 + if
            'Modify By Sindy 2024/10/1 傳入建檔人
            If PUB_CheckModifyLimit_frm140402(m_PCU51, GRD1.TextMatrix(m_mouseRow, 16)) = False Then Exit Sub
         Else
            'Modified by Lydia 2020/11/03 +判斷特殊群組= False 'Memo by Lydia 2020/11/30 因為特殊群組只有查詢功能
            'If PUB_CheckModifyLimit_frm100101_19(m_CU13, Me.Tag, GRD1.TextMatrix(m_mouseRow, 16)) = False Then Exit Sub
            If PUB_CheckModifyLimit_frm100101_19(m_CU13, Me.Tag, GRD1.TextMatrix(m_mouseRow, 16), False) = False Then
                 Exit Sub
            End If
            'end 2020/11/30
         End If
         '2019/7/25 END
         
         If m_mouseRow > 0 And GRD1.CellBackColor = &HFFC0C0 Then
            Call AddAttFile
         Else
            MsgBox "請勾選一筆欲新增電子檔的往來記錄!!!"
            Exit Sub
         End If
         
      'Add By Sindy 2019/6/5
      Case 7 '刪除電子檔
         GRD1.row = m_mouseRow
         GRD1.col = 1
         If m_mouseRow > 0 And GRD1.CellBackColor = &HFFC0C0 Then
            Call DelAttFile
         Else
            MsgBox "請至少勾選一筆欲刪除電子檔的往來記錄!!!"
            Exit Sub
         End If
   End Select
   
End Sub

'Add By Sindy 2019/6/14
'刪除電子檔
Private Sub DelAttFile()
Dim bolDel As Boolean
Dim intChkCnt As Integer
Dim bolConn As Boolean
Dim rsTmp As New ADODB.Recordset
Dim ii As Integer

On Error GoTo ErrHnd
   
   bolDel = False
   intChkCnt = 0
   For ii = 1 To GRD1.Rows - 1
      If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v") And Trim(GRD1.TextMatrix(ii, 8)) <> "" Then
         intChkCnt = intChkCnt + 1
      End If
   Next ii
   If intChkCnt <= 0 Then
      MsgBox "請至少勾選一筆欲刪除電子檔的資料!!!"
      Exit Sub
   End If

   Screen.MousePointer = vbHourglass

   For ii = 1 To GRD1.Rows - 1
      If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v") And Trim(GRD1.TextMatrix(ii, 8)) <> "" Then
         'Modify By Sindy 2019/7/25
         'If PUB_CheckModifyLimit_frm140402(GRD1.TextMatrix(ii, 16), "M") = False Then Exit Sub
         If m_quyDataKind = 0 Then 'Add By Sindy 2020/5/19 + if
            'Modified by Lydia 2020/11/30
            'If PUB_CheckModifyLimit_frm140402(m_PCU51) = False Then Exit Sub
            'Modify By Sindy 2024/10/1 傳入建檔人
            If PUB_CheckModifyLimit_frm140402(m_PCU51, GRD1.TextMatrix(ii, 16)) = False Then GoTo EXITSUB
         Else
            'Modified by Lydia 2020/11/03 +判斷特殊群組= False 'Memo by Lydia 2020/11/30 因為特殊群組只有查詢功能
            'If PUB_CheckModifyLimit_frm100101_19(m_CU13, Me.Tag, GRD1.TextMatrix(ii, 16)) = False Then Exit Sub
            'Modified by Lydia 2020/11/30
            'If PUB_CheckModifyLimit_frm100101_19(m_CU13, Me.Tag, GRD1.TextMatrix(ii, 16), False) = False Then Exit Sub
            If PUB_CheckModifyLimit_frm100101_19(m_CU13, Me.Tag, GRD1.TextMatrix(ii, 16), False) = False Then
                GoTo EXITSUB
            End If
            'end 2020/11/30
         End If
         '2019/7/25 END
         
         If MsgBox("確定要永久刪除 ( " & GetFileName(Trim(GRD1.TextMatrix(ii, 8))) & " ) 電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         '直接從資料庫刪除檔案
         If PUB_DelFtpFile2(GRD1.TextMatrix(ii, 13), GRD1.TextMatrix(ii, 15), cTableName) = False Then
            GoTo ErrHnd
         End If
         strSql = "delete from CONTACTFILE where cf01='" & GRD1.TextMatrix(ii, 13) & "' and upper(cf06)='" & UCase(GRD1.TextMatrix(ii, 15)) & "'"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
         bolDel = True
      End If
   Next ii
   
   If bolDel = True Then
      Call cmdok_Click(3)
   End If
   
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   Exit Sub

ErrHnd:
   Set rsTmp = Nothing
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
'Added by Lydia 2020/11/30
EXITSUB:
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
End Sub

'新增電子檔
Private Sub AddAttFile()
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer, jj As Integer
   Dim fs, f
   Dim bolAdd As Boolean
   Dim strFile As String
   
On Error GoTo ErrHnd
   
   bolAdd = False
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
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            '記錄路徑
            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
            For ii = 1 To UBound(sFile)
               If InStr(CStr(sFile(ii)), "#") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                  GoTo EXITSUB
               End If
               
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               
'               '檢查檔名規則
'               If Right(Trim(UCase(stFileName)), 4) = UCase(".PDF") Then
'                  '檔名中不可有中文字
'                  For jj = 1 To Len(stFileName)
'                     If Asc(Mid(stFileName, jj, 1)) <= 0 Then
'                        MsgBox stFileName & vbCrLf & vbCrLf & "檔案命名不符規定，不可有中文字!!!"
'                        GoTo EXITSUB
'                     End If
'                  Next jj
'               End If
               
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               '檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  GoTo EXITSUB
               ElseIf f.Size > 5242880 Then
                  If MsgBox("檔案過大（容量超過5MB），確認是否要上傳？", vbYesNo, "警告") = vbNo Then
                     GoTo EXITSUB
                  End If
               End If
               If IsRecordExist(GRD1.TextMatrix(m_mouseRow, 13), sFile(ii)) = False Then
                  '存檔
                  If PUB_UpdateCFData(GRD1.TextMatrix(m_mouseRow, 13), stFileName, f.Size) = False Then
                     GoTo EXITSUB
                  Else
                     bolAdd = True
                     Call PUB_DelPCOrgFile(stFileName) '一併將PC上的實體檔案刪除
                  End If
               End If
            Next ii
         Else
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
               GoTo EXITSUB
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
            stFileName = .FileName
            
'            '檢查檔名規則
'            If Right(Trim(UCase(stFileName)), 4) = UCase(".PDF") Then
'               '檔名中不可有中文字
'               For jj = 1 To Len(stFileName)
'                  If Asc(Mid(stFileName, jj, 1)) <= 0 Then
'                     MsgBox stFileName & vbCrLf & vbCrLf & "檔案命名不符規定，不可有中文字!!!"
'                     GoTo EXITSUB
'                  End If
'               Next jj
'            End If
            
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            '檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg strFile & MsgText(9221)
               GoTo EXITSUB
            ElseIf f.Size > 5242880 Then
               If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                  GoTo EXITSUB
               End If
            End If
            If IsRecordExist(GRD1.TextMatrix(m_mouseRow, 13), strFile) = False Then
               '存檔
               If PUB_UpdateCFData(GRD1.TextMatrix(m_mouseRow, 13), stFileName, f.Size) = False Then
                  GoTo EXITSUB
               Else
                  bolAdd = True
                  Call PUB_DelPCOrgFile(stFileName) '一併將PC上的實體檔案刪除
               End If
            End If
         End If
EXITSUB:
         If bolAdd = True Then
            Call cmdok_Click(3)
         End If
      End If
   End With
   
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHnd:
   If bolAdd = True Then
      Call cmdok_Click(3)
   End If
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal stKey1 As String, ByVal stFileName As String) As Boolean
Dim adoRst As ADODB.Recordset
   
   IsRecordExist = False
   
   strSql = "SELECT cf01 FROM CONTACTFILE WHERE cf01='" & stKey1 & "' and upper(cf02)=upper('" & ChgSQL(stFileName) & "')"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      IsRecordExist = True
      MsgBox "附件 " & stFileName & " 已存在！"
   End If
   
   Set adoRst = Nothing
End Function

'國外往來記錄
Public Sub StrMenu(Optional bolMySelfQuery As Boolean = False)
Dim strKey As String, strKey1 As String
Dim strCOR As String, strCR As String
'2008/11/17 ADD BY SONIA
'Dim StrCR05 As Variant
Dim StrCR04 As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer, k As Integer
Dim strCR01 As String 'Add By Sindy 2019/3/6
Dim ii As Integer 'Add By Sindy 2019/3/13
Dim dblRow As Double 'Add By Sindy 2025/8/28
   
   'Add By Sindy 2011/01/03 檢查國內外權限
   'Add By Sindy 2019/2/23
   If bolMySelfQuery = False Then
   '2019/2/23 END
      If CheckSR12(Me.Tag) = False Then
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Sub
      End If
   End If
   pub_QL05 = pub_QL05 & IIf(PUB_CheckQL05("編號：" & Me.Tag & "(往來記錄)") = "", "", ";編號：" & Me.Tag & "(往來記錄)") 'Add By Sindy 2025/8/13
   
   Me.Caption = "往來記錄瀏覽區" 'Modify By Sindy 2025/8/28 秀玲說拿掉(國外)
   Me.Tag = Replace(Me.Tag, "平台", "") 'Add By Sindy 2021/3/25
   strKey = Left(Me.Tag, 8)
   strKey1 = ""   '2008/11/17 ADD BY SONIA
   If Mid(Me.Tag, 10, 1) = "-" Then
      strKey1 = Mid(Me.Tag, 11)
   End If
   
   '往來對象,聯絡人資料
   '2008/11/17 MODIFY BY SONIA
   'strExc(0) = "select N1,nvl(pcc05,nvl(pcc03,pcc04)) N2" & _
      " from (select NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)) N1" & _
      " from customer where cu01='" & strKey & "' and cu02='0'" & _
      " union all select NVL(fa04,nvl(fa05||' '||fa63||' '||fa64||' '||fa65,fa06)) N1" & _
      " from fagent where fa01='" & strKey & "' and fa02='0'" & _
      " union all select NVL(PCU08,NVL(RTRIM(PCU03||' '||PCU04||' '||PCU05||' '||PCU06),PCU07)) N1" & _
      " from potcustomer where pcu01='" & strKey & "' and pcu02='0'" & _
      ") A,potcustcont where pcc01(+)='" & strKey & "' and pcc02(+)='" & strKey1 & "'"
   '2008/12/9 modify by sonia 加國籍才能判斷語文權限
   'Modify By Sindy 2019/7/25 + ,'' PCU51
   strExc(0) = "select N1,nvl(pcc05,nvl(pcc03,pcc04)) N2, N3, NO1, NO2, PCU51" & _
      " from (select NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) N1,CU01||CU02 NO1,CU01 NO2,CU10 N3,'' PCU51" & _
      " from customer where cu01='" & strKey & "' and cu02='0'" & _
      " union all select NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) N1,FA01||FA02 NO1,FA01 NO2,FA10 N3,'' PCU51" & _
      " from fagent where fa01='" & strKey & "' and fa02='0'" & _
      " union all select NVL(PCU08,DECODE(PCU03,NULL,PCU07,PCU03||' '||PCU04||' '||PCU05||' '||PCU06)) N1,PCU01||PCU02 NO1,PCU01 NO2,PCU09 N3,PCU51" & _
      " from potcustomer where pcu01='" & strKey & "' and pcu02='0'" & _
      " union all select cw12 N1,cw01 NO1,'' NO2,'' N3,'' PCU51" & _
      " from custweb where cw01='" & strKey & "'" & _
      ") A,potcustcont where A.NO2=pcc01(+)"
   
   If strKey1 <> "" Then
      strExc(0) = strExc(0) & " and pcc02(+)='" & strKey1 & "'"
   Else
      strExc(0) = strExc(0) & " and pcc02(+)='ZZ'"
   End If
   '2008/11/17 END

   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   m_PCU51 = "" 'Add By Sindy 2019/7/25
   If intI = 1 Then
      Label3 = strKey & " " & RsTemp.Fields(0)
      Label4 = strKey1 & " " & RsTemp.Fields(1)
      m_PCU51 = "" & RsTemp.Fields("PCU51") 'Add By Sindy 2019/7/25 國外潛在客戶的國內外權限
'      '2008/12/9 ADD BY SONIA 加語文權限
'      If m_bLanguage = "" And Mid(RsTemp.Fields(3), 1, 1) = "R" Then
'         MsgBox "您沒有查詢潛在客戶的往來記錄權限 !!!", vbInformation
'         Screen.MousePointer = vbDefault
'         Me.Enabled = True
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         Exit Sub
'      ElseIf m_bLanguage = "J" And Mid(RsTemp.Fields(3), 1, 1) = "R" And Mid(RsTemp.Fields(2), 1, 3) <> "011" Then
'         MsgBox "您沒有查詢英文組潛在客戶的往來記錄權限 !!!", vbInformation
'         Screen.MousePointer = vbDefault
'         Me.Enabled = True
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         Exit Sub
'      ElseIf m_bLanguage = "E" And Mid(RsTemp.Fields(3), 1, 1) = "R" And Mid(RsTemp.Fields(2), 1, 3) = "011" Then
'         MsgBox "您沒有查詢日文組潛在客戶的往來記錄權限 !!!", vbInformation
'         Screen.MousePointer = vbDefault
'         Me.Enabled = True
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         Exit Sub
'      End If
'      '2008/12/9 END
      'Add By Sindy 2009/04/30
      'If GetCustData(Mid(RsTemp.Fields(3), 1, 8)) = False Then
      If PUB_GetCustData(RsTemp.Fields(3)) = False Then
         'MsgBox "您沒有維護此筆潛在客戶的往來記錄權限！"
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Sub
      End If
      '2009/04/30 End
   End If
   
   If strKey1 <> "" Then
      strCR = strCR & " and instr(cr04,'" & strKey1 & "')>0"
   End If
   
   '2008/11/17 ADD BY SONIA
   '往來日期
   If Len(CRdateF) <> 0 Then
       strCR = strCR & " AND CR02>=" & Val(ChangeTStringToWString(CRdateF))
       strCOR = strCOR & " AND CoR02>=" & Val(ChangeTStringToWString(CRdateF))
   End If
   If Len(CRdateT) <> 0 Then
       strCR = strCR & " AND CR02<=" & Val(ChangeTStringToWString(CRdateT))
       strCOR = strCOR & " AND CoR02<=" & Val(ChangeTStringToWString(CRdateT))
   End If
   'Add By Sindy 2025/8/27
   If Len(CRdateF) <> 0 Or Len(CRdateT) <> 0 Then
      pub_QL05 = pub_QL05 & PUB_CheckQL05(";往來日期：" & CRdateF & "-" & CRdateT)
   End If
   '2025/8/27 END
   
   'Add By Sindy 2019/2/23
   If bolMySelfQuery = True Then
      '主旨
      If Len(txtCR06) <> 0 Then
         strCR = strCR & " AND (instr(upper(CR06),upper('" & txtCR06 & "'))>0 or instr(upper(cf13),upper('" & txtCR06 & "'))>0 or instr(upper(cf02),upper('" & txtCR06 & "'))>0)"
         strCOR = strCOR & " AND (instr(upper(CoR04),upper('" & txtCR06 & "'))>0 or instr(upper(cf13),upper('" & txtCR06 & "'))>0 or instr(upper(cf02),upper('" & txtCR06 & "'))>0)"
         pub_QL05 = pub_QL05 & PUB_CheckQL05(";主旨：" & txtCR06) 'Add By Sindy 2025/8/27
      End If
      '往來類別
      If Len(Combo1.Text) <> 0 Then
         varTmp = Split(Combo1.Text, " ")
         strCR = strCR & " AND CR05='" & Trim(varTmp(0)) & "'"
         pub_QL05 = pub_QL05 & PUB_CheckQL05(";往來類別：" & Combo1.Text) 'Add By Sindy 2025/8/27
      End If
   '2019/2/23 END
   Else
      '往來類別
      If CRtype <> "" Then
         'StrCR05 = Split(CRtype, ",")
         varTmp = Split(CRtype, " ")
         strCR = strCR & " AND CR05='" & Trim(varTmp(0)) & "'"
         pub_QL05 = pub_QL05 & PUB_CheckQL05(";往來類別：" & CRtype) 'Add By Sindy 2025/8/27
      End If
      '2008/11/17 END
   End If
   
'Modify by Toni 2008/11/06 因為PCC01 and PCU01都是8位
'   strExc(0) = "select ' ' AS V,CR01 AS 往來記錄編號,CR11 被回覆記錄編號,CR18 回覆記錄編號" & _
'      ",SQLDATEW(CR02) 往來日期,SQLDATEW(CR10) 回覆期限,CR05 往來類別,CR06 主旨,CR07 地點" & _
'      ",CR08 內容,CR04 聯絡人" & _
'      " from contactrecord where cr03='" & strKey & "'" & strCR & _
'      " order by cr01"
   strExc(0) = "select ' ' AS V,CR01 AS 往來記錄編號,CR11 被回覆記錄編號,CR18 回覆記錄編號" & _
      "," & SQLDate("CR02") & " 往來日期," & SQLDate("CR10") & " 回覆期限,AC03 往來類別,CR06 主旨" & _
      ",decode(cf02,null,'',cf02||' ('||Round(cf07 / 1024, 2)||' KB)'||'('||cf03||';'||cf04||';'||cf05||')') AS 檔案名稱" & _
      ",nvl(e1.efc03,'') AS 副檔名說明,CR07 地點" & _
      ",CR08 內容,CR04 聯絡人,CR01,CF02,CF06,CR12,CF04,'F' Dtype,decode(cr05,'A01.1P',1,'A01.2P',2,99) as sort,cf05,cr02 cor02,cr01 cor01" & _
      " from contactrecord,contactfile,efilecaption e1,allcode" & _
      " where SUBSTR(cr03,1,8)='" & strKey & "'" & strCR & _
      " and cr01=cf01(+) and ac01(+)='11' and cr05=ac02(+)" & _
      " and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cf02),'.'||e1.efc02(+)||'.')>0"
   'end 2008/11/06
   '加入Y的國內往來記錄
   'Modify by Amy 2020/09/08 Mark if 因 X79079往來記錄寫於國外(KA4000128)
   'If Left(strKey, 1) = "Y" Then
   strExc(0) = strExc(0) & " union all " & _
      "select ' ' AS V,CoR01 AS 往來記錄編號,'' 被回覆記錄編號,'' 回覆記錄編號" & _
      "," & SQLDate("CoR02") & " 往來日期,'' 回覆期限,'' 往來類別,CoR04 主旨" & _
      ",decode(cf02,null,'',cf02||' ('||Round(cf07 / 1024, 2)||' KB)'||'('||cf03||';'||cf04||';'||cf05||')') AS 檔案名稱" & _
      ",nvl(e1.efc03,'') AS 副檔名說明,'' 地點" & _
      ",CoR05 內容,'' 聯絡人,CoR01,CF02,CF06,CoR06 CR12,CF04,'C' Dtype,99 as sort,cf05,cor02,cor01" & _
      " from contactrecord1,contactfile,efilecaption e1" & _
      " where SUBSTR(cor03,1,8)='" & strKey & "'" & strCOR & _
      " and cor01=cf01(+)" & _
      " and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cf02),'.'||e1.efc02(+)||'.')>0"
   'End If
   'Modify By Sindy 2022/12/5 + Widen:以往來日期排序最新日期者在最上頭
   'Modify By Sindy 2023/10/16 + ,cor01 desc
   strExc(0) = strExc(0) & " order by sort asc,cor02 desc,cor01 desc,cf04 desc,cf05 desc"
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   GRD1.Clear
   SetGrd
   If intI = 1 Then
      dblRow = adoRecordset.RecordCount 'Add By Sindy 2025/8/28
      'Added by Lydia 2025/08/08 國外往來記錄的維護及查詢限制
      If m_strCRexcept <> "" Then
          strExc(1) = "select * from (" & Mid(UCase(strExc(0)), 1, InStr(UCase(strExc(0)), "ORDER") - 1) & ") where 往來記錄編號 not in (" & GetAddStr(m_strCRexcept) & ")" & _
                      " order by sort asc,cor02 desc,cor01 desc,cf04 desc,cf05 desc"
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
          If intI = 1 Then
             If RsTemp.RecordCount <> adoRecordset.RecordCount Then
                 pub_QL05 = pub_QL05 & "(含限閱" & adoRecordset.RecordCount - RsTemp.RecordCount & "筆)" 'Add By Sindy 2025/8/28
                 MsgBox "限閱往來記錄" & adoRecordset.RecordCount - RsTemp.RecordCount & "筆！", vbInformation
                 Set adoRecordset = RsTemp
             End If
          Else
              Screen.MousePointer = vbDefault
              Me.Enabled = True
              pub_QL05 = pub_QL05 & "(含限閱" & adoRecordset.RecordCount - RsTemp.RecordCount & "筆)" 'Add By Sindy 2025/8/28
              MsgBox "限閱往來記錄" & adoRecordset.RecordCount - RsTemp.RecordCount & "筆！", vbInformation
              Exit Sub
          End If
      End If
      'end 2025/08/08
      If pub_QL04 <> "" Then InsertQueryLog (dblRow) 'Add By Sindy 2025/8/13
      
      '2008/11/17 MODIFY BY SONIA 逐筆抓聯絡人名稱,逐筆檢查往來類別
      Set GRD1.Recordset = adoRecordset
      Me.Enabled = False
'      adoRecordset.MoveFirst
'      Do While Not adoRecordset.EOF
'         '檢查往來類別
'         k = 0
'         For i = 0 To UBound(StrCR05)
'            If InStr(UCase(adoRecordset.Fields(6)), UCase(StrCR05(i))) > 0 Then
'               k = k + 1
'            End If
'         Next i
'         If k < UBound(StrCR05) + 1 Then
'            GoTo NextRecord
'         End If
         
'         GRD1.Rows = GRD1.Rows + 1
'         GRD1.row = GRD1.Rows - 2
'
'         If Not IsNull(adoRecordset.Fields(1)) Then
'            GRD1.TextMatrix(GRD1.row, 1) = adoRecordset.Fields(1)
'         End If
'         If Not IsNull(adoRecordset.Fields(2)) Then
'            GRD1.TextMatrix(GRD1.row, 2) = adoRecordset.Fields(2)
'         End If
'         If Not IsNull(adoRecordset.Fields(3)) Then
'            GRD1.TextMatrix(GRD1.row, 3) = adoRecordset.Fields(3)
'         End If
'         If Not IsNull(adoRecordset.Fields(4)) Then
'            GRD1.TextMatrix(GRD1.row, 4) = adoRecordset.Fields(4)
'         End If
'         If Not IsNull(adoRecordset.Fields(5)) Then
'            GRD1.TextMatrix(GRD1.row, 5) = adoRecordset.Fields(5)
'         End If
'         If Not IsNull(adoRecordset.Fields(6)) Then
'            GRD1.TextMatrix(GRD1.row, 6) = adoRecordset.Fields(6)
'         End If
'         If Not IsNull(adoRecordset.Fields(7)) Then
'            GRD1.TextMatrix(GRD1.row, 7) = adoRecordset.Fields(7)
'         End If
'         If Not IsNull(adoRecordset.Fields(8)) Then
'            GRD1.TextMatrix(GRD1.row, 8) = adoRecordset.Fields(8)
'         End If
'         If Not IsNull(adoRecordset.Fields(9)) Then
'            GRD1.TextMatrix(GRD1.row, 9) = adoRecordset.Fields(9)
'         End If
      For ii = 1 To GRD1.Rows - 1
         If Not IsNull(GRD1.TextMatrix(ii, 12)) Then
            StrCR04 = GRD1.TextMatrix(ii, 12)
         Else
            StrCR04 = ""
         End If
         GRD1.TextMatrix(ii, 12) = ""
         If StrCR04 <> "" Then
            Set rsTmp = New ADODB.Recordset
            strSql = "SELECT nvl(pcc05,nvl(pcc03,pcc04)) NM FROM PotCustCont " & _
                     "WHERE PCC01 = '" & strKey & "'" & " AND PCC02 IN (" & StrCR04 & ") ORDER BY PCC02"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               For i = 1 To rsTmp.RecordCount
                  GRD1.TextMatrix(ii, 12) = GRD1.TextMatrix(ii, 12) & rsTmp.Fields("NM") & ";"
                  rsTmp.MoveNext
               Next i
            End If
         End If
      Next ii
'NextRecord:
'         'Added by Lydia 2018/12/22 統一靠左
'         For i = 0 To GRD1.Cols - 1
'            GRD1.col = i
'            GRD1.CellAlignment = flexAlignLeftCenter
'         Next i
'         'end 2018/12/22
'         adoRecordset.MoveNext
'      Loop
      'GRD1.Rows = GRD1.Rows - 1
      
      If GRD1.row = 0 Then
         ShowNoData
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         'Add By Sindy 2019/3/7
         If bolMySelfQuery = False Then
         '2019/3/7 END
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
         End If
         Exit Sub
      End If
      '2008/11/17 END
      
      'Add By Sindy 2019/3/6
      strCR01 = ""
      For i = 1 To GRD1.Rows - 1
         If GRD1.RowHeight(i) > 0 Then
            If strCR01 = Trim(GRD1.TextMatrix(i, 13)) Then
               GRD1.TextMatrix(i, 1) = ""
               GRD1.TextMatrix(i, 4) = ""
               GRD1.TextMatrix(i, 6) = ""
               GRD1.TextMatrix(i, 7) = "" '主旨
               GRD1.TextMatrix(i, 10) = ""
               GRD1.TextMatrix(i, 11) = "" '內容
               GRD1.TextMatrix(i, 12) = ""
            End If
         End If
         strCR01 = Trim(GRD1.TextMatrix(i, 13))
      Next i
      Me.Enabled = True
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
      ShowNoData
      Screen.MousePointer = vbDefault
      Me.Enabled = True
'      'Add By Sindy 2019/3/7
'      If bolMySelfQuery = False Then
'      '2019/3/7 END
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      End If
      Exit Sub
   End If
   GRD1.col = 0
   GRD1.row = 0
End Sub

'國內往來記錄
Public Sub StrMenu2(Optional bolMySelfQuery As Boolean = False)
Dim strKey As String, strKey1 As String
Dim strCOR As String, strCR As String
'2008/11/17 ADD BY SONIA
'Dim StrCR05 As Variant
Dim StrCR04 As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer, k As Integer
Dim strCR01 As String 'Add By Sindy 2019/3/6
Dim ii As Integer 'Add By Sindy 2019/3/13
   
   'Add By Sindy 2011/01/03 檢查國內外權限
   'Add By Sindy 2019/2/23
   If bolMySelfQuery = False Then
   '2019/2/23 END
      If CheckSR12(Me.Tag) = False Then
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Sub
      End If
   End If
   pub_QL05 = pub_QL05 & IIf(PUB_CheckQL05("編號：" & Me.Tag & "(往來記錄)") = "", "", ";編號：" & Me.Tag & "(往來記錄)") 'Add By Sindy 2025/8/13
   
   Me.Caption = "往來記錄瀏覽區" 'Modify By Sindy 2025/8/28 秀玲說拿掉(國內)
   strKey = Left(Me.Tag, 8)
   strKey1 = ""
   If Mid(Me.Tag, 10, 1) = "-" Then
      strKey1 = Mid(Me.Tag, 11)
   End If
      
   '往來對象,聯絡人資料
   strExc(0) = "select N1, N2, N3, NO1, NO2, PCU51" & _
      " from (select NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) N1,CU13 N2,CU01||CU02 NO1,CU01 NO2,CU10 N3,'' PCU51" & _
      " from customer where cu01='" & strKey & "' and cu02='0'" & _
      " union all select NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) N1,FA120 N2,FA01||FA02 NO1,FA01 NO2,FA10 N3,'' PCU51" & _
      " from fagent where fa01='" & strKey & "' and fa02='0'" & _
      " union all select NVL(Poc03,DECODE(Poc23,NULL,Poc27,Poc23||' '||Poc24||' '||Poc25||' '||Poc26)) N1,Poc13 N2,Poc01||Poc02 NO1,Poc01 NO2,Poc04 N3,'' PCU51" & _
      " from potcustomer1 where poc01='" & strKey & "' and poc02='0'" & _
      ") A,potcustcont where A.NO2=pcc01(+)"
   If strKey1 <> "" Then
      strExc(0) = strExc(0) & " and pcc02(+)='" & strKey1 & "'"
   Else
      strExc(0) = strExc(0) & " and pcc02(+)='ZZ'"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   m_PCU51 = "" 'Add By Sindy 2019/7/25
   If intI = 1 Then
      Label3 = strKey & " " & RsTemp.Fields(0)
      'Label4 = strKey1 & " " & RsTemp.Fields(1)
      Label4.Caption = "" '聯絡人
      m_PCU51 = "" & RsTemp.Fields("PCU51") 'Add By Sindy 2019/7/25 國外潛在客戶的國內外權限
      
      'Modify by Amy 2017/07/17 X54363010 因為智權為MCTF開頭 run PUB_Id2Num 會錯
      If InStr(Left("" & RsTemp.Fields("N2"), 4), "MCTF") > 0 Then
        m_CU13 = Replace(Pub_GetSpecMan(RsTemp.Fields("N2"), False), ";", ",")
      Else
        m_CU13 = "" & RsTemp.Fields("N2")
      End If
      'SetlstUsers 0, "" & RsTemp.Fields("N2")
      SetlstUsers 0, m_CU13
      'end 2017/07/17
      'Modified by Lydia 2020/11/03 建檔人改成必傳(COR06或ADD新增用)
      'If PUB_GetCustData_frm100101_19(Me.Tag) = False Then
      'Modified by Lydia 2020/11/30 +判斷特殊群組
      'If PUB_GetCustData_frm100101_19(Me.Tag, "ADD", False) = False Then
      If PUB_GetCustData_frm100101_19(Me.Tag, "ADD", bolGrpSpec) = False Then
      'end 2020/11/30
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Sub
      End If
   End If
   
   If strKey1 <> "" Then
      strCR = strCR & " and instr(cr04,'" & strKey1 & "')>0"
   End If
   
   '往來日期
   If Len(CRdateF) <> 0 Then
       strCOR = strCOR & " AND COR02>=" & Val(ChangeTStringToWString(CRdateF))
       strCR = strCR & " AND CR02>=" & Val(ChangeTStringToWString(CRdateF))
   End If
   If Len(CRdateT) <> 0 Then
       strCOR = strCOR & " AND COR02<=" & Val(ChangeTStringToWString(CRdateT))
       strCR = strCR & " AND CR02<=" & Val(ChangeTStringToWString(CRdateT))
   End If
   'Add By Sindy 2025/8/27
   If Len(CRdateF) <> 0 Or Len(CRdateT) <> 0 Then
      pub_QL05 = pub_QL05 & PUB_CheckQL05(";往來日期：" & CRdateF & "-" & CRdateT)
   End If
   '2025/8/27 END
   
   'Add By Sindy 2019/2/23
   If bolMySelfQuery = True Then
      '主旨
      If Len(txtCR06) <> 0 Then
         strCOR = strCOR & " AND (instr(upper(COR04),upper('" & txtCR06 & "'))>0 or instr(upper(cf13),upper('" & txtCR06 & "'))>0 or instr(upper(cf02),upper('" & txtCR06 & "'))>0)"
         strCR = strCR & " AND (instr(upper(CR06),upper('" & txtCR06 & "'))>0 or instr(upper(cf13),upper('" & txtCR06 & "'))>0 or instr(upper(cf02),upper('" & txtCR06 & "'))>0)"
         pub_QL05 = pub_QL05 & PUB_CheckQL05(";主旨：" & txtCR06) 'Add By Sindy 2025/8/27
      End If
      '往來類別
      If Len(Combo1.Text) <> 0 Then
         varTmp = Split(Combo1.Text, " ")
         strCR = strCR & " AND CR05='" & Trim(varTmp(0)) & "'"
         pub_QL05 = pub_QL05 & PUB_CheckQL05(";往來類別：" & Combo1.Text) 'Add By Sindy 2025/8/27
      End If
   '2019/2/23 END
   Else
      '往來類別
      If CRtype <> "" Then
         varTmp = Split(CRtype, " ")
         strCR = strCR & " AND CR05='" & Trim(varTmp(0)) & "'"
         pub_QL05 = pub_QL05 & PUB_CheckQL05(";往來類別：" & CRtype) 'Add By Sindy 2025/8/27
      End If
      '2008/11/17 END
   End If
   
   strExc(0) = "select ' ' AS V,CoR01 AS 往來記錄編號,'' 被回覆記錄編號,'' 回覆記錄編號" & _
      "," & SQLDate("CoR02") & " 往來日期,'' 回覆期限,'' 往來類別,CoR04 主旨" & _
      ",decode(cf02,null,'',cf02||' ('||Round(cf07 / 1024, 2)||' KB)'||'('||cf03||';'||cf04||';'||cf05||')') AS 檔案名稱" & _
      ",nvl(e1.efc03,'') AS 副檔名說明,'' 地點" & _
      ",CoR05 內容,'' 聯絡人,CoR01,CF02,CF06,CoR06 CR12,CF04,'C' Dtype,99 as sort,cf05,cor02 cr02,cor01 cr01" & _
      " from contactrecord1,contactfile,efilecaption e1" & _
      " where SUBSTR(cor03,1,8)='" & strKey & "'" & strCOR & _
      " and cor01=cf01(+)" & _
      " and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cf02),'.'||e1.efc02(+)||'.')>0"
   '加入Y的國外往來記錄
   'Modify by Amy 2020/09/08 Mark if 因 X79079往來記錄寫於國外(KA4000128)
   'If Left(strKey, 1) = "Y" Then
   strExc(0) = strExc(0) & " union all " & _
      "select ' ' AS V,CR01 AS 往來記錄編號,CR11 被回覆記錄編號,CR18 回覆記錄編號" & _
      "," & SQLDate("CR02") & " 往來日期," & SQLDate("CR10") & " 回覆期限,AC03 往來類別,CR06 主旨" & _
      ",decode(cf02,null,'',cf02||' ('||Round(cf07 / 1024, 2)||' KB)'||'('||cf03||';'||cf04||';'||cf05||')') AS 檔案名稱" & _
      ",nvl(e1.efc03,'') AS 副檔名說明,CR07 地點" & _
      ",CR08 內容,CR04 聯絡人,CR01,CF02,CF06,CR12,CF04,'F' Dtype,decode(cr05,'A01.1P',1,'A01.2P',2,99) as sort,cf05,cr02,cr01" & _
      " from contactrecord,contactfile,efilecaption e1,allcode" & _
      " where SUBSTR(cr03,1,8)='" & strKey & "'" & strCR & _
      " and cr01=cf01(+) and ac01(+)='11' and cr05=ac02(+)" & _
      " and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cf02),'.'||e1.efc02(+)||'.')>0"
   'End If
   'Modify By Sindy 2022/12/5 + Widen:以往來日期排序最新日期者在最上頭
   'Modify By Sindy 2023/10/16 + ,cr01 desc
   strExc(0) = strExc(0) & " order by sort asc,cr02 desc,cr01 desc,cf04 desc,cf05 desc"
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
   GRD1.Clear
   SetGrd
   If intI = 1 Then
      If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/13
      '2008/11/17 MODIFY BY SONIA 逐筆抓聯絡人名稱,逐筆檢查往來類別
      Set GRD1.Recordset = adoRecordset
      Me.Enabled = False
'      For ii = 1 To GRD1.Rows - 1
'         If Not IsNull(GRD1.TextMatrix(ii, 12)) Then
'            StrCR04 = GRD1.TextMatrix(ii, 12)
'         Else
'            StrCR04 = ""
'         End If
'         GRD1.TextMatrix(ii, 12) = ""
'         If StrCR04 <> "" Then
'            Set rsTmp = New ADODB.Recordset
'            strSql = "SELECT nvl(pcc05,nvl(pcc03,pcc04)) NM FROM PotCustCont " & _
'                     "WHERE PCC01 = '" & strKey & "'" & " AND PCC02 IN (" & StrCR04 & ") ORDER BY PCC02"
'            rsTmp.CursorLocation = adUseClient
'            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsTmp.RecordCount > 0 Then
'               For i = 1 To rsTmp.RecordCount
'                  GRD1.TextMatrix(ii, 12) = GRD1.TextMatrix(ii, 12) & rsTmp.Fields("NM") & ";"
'                  rsTmp.MoveNext
'               Next i
'            End If
'         End If
'      Next ii
      
      If GRD1.row = 0 Then
         ShowNoData
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         'Add By Sindy 2019/3/7
         If bolMySelfQuery = False Then
         '2019/3/7 END
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
         End If
         Exit Sub
      End If
      '2008/11/17 END
      
      'Add By Sindy 2019/3/6
      strCR01 = ""
      For i = 1 To GRD1.Rows - 1
         If GRD1.RowHeight(i) > 0 Then
            If strCR01 = Trim(GRD1.TextMatrix(i, 13)) Then
               GRD1.TextMatrix(i, 1) = ""
               GRD1.TextMatrix(i, 4) = ""
               GRD1.TextMatrix(i, 6) = ""
               GRD1.TextMatrix(i, 7) = "" '主旨
               GRD1.TextMatrix(i, 11) = "" '內容
            End If
         End If
         strCR01 = Trim(GRD1.TextMatrix(i, 13))
      Next i
      Me.Enabled = True
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
      ShowNoData
      Screen.MousePointer = vbDefault
      Me.Enabled = True
'      'Add By Sindy 2019/3/7
'      If bolMySelfQuery = False Then
'      '2019/3/7 END
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      End If
      Exit Sub
   End If
   
   GRD1.col = 0
   GRD1.row = 0
End Sub
Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
   If p_stNums <> "" Then
      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         arrID = Split(p_stNums, ",")
         With RsTemp
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("st01") = arrID(intI) Then
                  lstUsers(p_idx).AddItem "" & .Fields(1), 0
                  lstUsers(p_idx).ITEMDATA(0) = PUB_Id2Num(.Fields(0)) '員工編號
                  .MoveLast
               End If
               .MoveNext
            Loop
         Next
         End With
      End If
   End If
End Sub

Private Sub SetGrd(Optional bolSetRow As Boolean = True)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2019/2/23 + "檔案名稱", "副檔名說明"
   'Modify By Sindy 2020/11/13 + "Sort"
   arrGridHeadText = Array("V", "記錄編號", "被回覆記錄編號(多)", "回覆記錄編號(多)", "往來日期", "回覆期限" _
      , "往來類別", "主旨", "檔案名稱", "副檔名說明", "地點" _
      , "內容", "聯絡人", "CR01", "CF02", "CF06" _
      , "CR12", "CF04", "Dtype", "Sort", "CF05", "CR02")
      
   '2008/11/10 ADD BY SONIA 回覆功能先鎖住以後再用
   'Modify By Sindy 2019/2/23
   'Modify By Sindy 2020/5/17
   If m_quyDataKind = 1 Then '1.國內
      arrGridHeadWidth = Array(200, 200, 0, 0, 800, 0 _
         , 0, 3000, 2500, 1000, 0 _
         , 3000, 0, 0, 0, 0 _
         , 0, 0, 0, 0, 0, 0)
   Else
   '2020/5/17 END
      '0.國外
      arrGridHeadWidth = Array(200, 200, 0, 0, 800, 0 _
         , 900, 1000, 2500, 800, 1000 _
         , 2000, 4000, 0, 0, 0 _
         , 0, 0, 0, 0, 0, 0)
   End If
'   arrGridHeadWidth = Array(200, 200, 0, 0, 800, 0 _
'         , 900, 1000, 1000, 800, 1000 _
'         , 2000, 4000, 1000, 0, 0 _
'         , 0, 0, 1000, 0, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   If bolSetRow = True Then
      GRD1.Rows = 2
   End If
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      If bolSetRow = True Then
         GRD1.Text = arrGridHeadText(iRow)
      End If
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      If bolSetRow = True Then
         'GRD1.CellAlignment = flexAlignCenterCenter
         GRD1.CellAlignment = flexAlignLeftCenter
      End If
   Next
   GRD1.Visible = True
End Sub

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
Dim hLocalFile As Long
Dim stFileName As String, strSaveFileName As String
Dim bolIsSelect As Boolean
Dim strCmd As String
Dim process_id As Long
Dim process_handle As Long
Dim strMergeFN As String, strMergeName As String
Dim strFileName As String, strFileType As String
Dim ii As Integer, jj As Integer
   
   strMergeFN = "" '組欲合併的檔案
'   If Index = 1 Then
'      If Check1.Value = 1 Then
'         Check1.Value = 0
'      End If
'      'WebBrowser1.Navigate "about:blank"
'   End If
   '切換至來源目錄
   If m_AttachPath <> "." Then ChDir m_AttachPath

   KillAttach
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   If m_bolDblClick = True Then
      strFileName = GRD1.TextMatrix(m_mouseRow, 14) 'GetFileName()
      strFileType = Right(strFileName, Len(strFileName) - InStrRev(strFileName, ".") + 1)
      If strFileName <> "" Then
         If UCase(strFileType) = UCase(".PDF") Then
            GRD1.TextMatrix(m_mouseRow, 0) = "V"
         Else
            GRD1.TextMatrix(m_mouseRow, 0) = "v"
         End If
      End If
   End If
   Check1.Tag = 0 'False Add By Sindy 2017/7/24
   For ii = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v" Then
         If m_bolDblClick = False Or (m_bolDblClick = True And ii = m_mouseRow) Then
            If Trim(GRD1.TextMatrix(ii, 14)) <> "" Then
               bolIsSelect = True
               If GRD1.TextMatrix(ii, 0) = "V" Then

                  'Add By Sindy 2017/7/24 選取PDF檔時,才開啟預覽視窗
                  If Index = 1 Then
                     'If Check1.Value = 1 Then
                     If Check1.Value = 1 And Check1.Tag = 0 Then
                        Check1.Value = 0
                        Check1.Tag = 1 'True Add By Sindy 2017/7/24
                     End If
                     'WebBrowser1.Navigate "about:blank"
                  End If

                  '切換至來源目錄
                  If m_AttachPath <> "." Then ChDir m_AttachPath
                  '清除反白
                  If UCase(cmdOpenAtt(1).Tag) <> UCase("call") Then
                     GRD1.col = 0
                     GRD1.row = ii
                     GRD1.TextMatrix(ii, 0) = ""
                     For jj = 1 To GRD1.Cols - 1
                        GRD1.col = jj
                        GRD1.CellBackColor = QBColor(15)
                     Next jj
                  End If
                  '讀取檔案名稱
                  If Index = 0 Then
                     stFileName = Trim(GRD1.TextMatrix(ii, 14))
                  Else
                     If InStr(Trim(GRD1.TextMatrix(ii, 15)), "/") > 0 Then
                        stFileName = Mid(GRD1.TextMatrix(ii, 15), InStrRev(Trim(GRD1.TextMatrix(ii, 15)), "/") + 1)
                     Else
                        stFileName = Trim(GRD1.TextMatrix(ii, 15))
                     End If
                  End If
                  'Modify By Sindy 2021/10/14 + And InStrRev(stFileName, " (") > InStrRev(stFileName, ".")
                  If InStrRev(stFileName, " (") > 0 And InStrRev(stFileName, " (") > InStrRev(stFileName, ".") Then
                     stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
                  End If
                  strSaveFileName = stFileName
                  If m_bolDblClick = True Then
                     strSaveFileName = Left(stFileName, InStrRev(stFileName, ".") - 1) & ServerTime & ".pdf"
                  End If
                  strMergeFN = strMergeFN & IIf(strMergeFN <> "", " ", "") & ".\" & strSaveFileName
                  If InStr(stFileName, "\") = 0 Then
                     If GetAttachFile(Trim(GRD1.TextMatrix(ii, 15)), stFileName, m_AttachPath & "\" & strSaveFileName) = False Then
                        'MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                     End If
                  End If
                  '開啟檔案
                  If Index = 0 Then
                     ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
                  End If
               End If
            End If
         End If
      End If
   Next ii
   
   If bolIsSelect = False Then
      MsgBox "無檔案可開啟！"
   Else
      If Index = 1 And strMergeFN <> "" Then
         'Modified by Morgan 2018/6/26 有時會發生錯誤,改單檔不合併
         'If m_bolDblClick = True Then
         If m_bolDblClick = True Or InStr(strMergeFN, " ") = 0 Then
            WebBrowser1.Navigate stFileName
         Else
            '合併
            strMergeName = "merge" & ServerTime & ".pdf"
            'strCmd = pub_PdftkEXE & " " & strMergeFN & " cat output " & m_AttachPath & "\" & strMergeName
            strCmd = pub_PdftkEXE & " " & strMergeFN & " cat output .\" & strMergeName
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
               If intI > 10 Then
                  TerminateProcess process_handle, 0&
                  CloseHandle process_handle
                  MsgBox "合併PDF失敗！"
                  GoTo ErrHnd
               Else
                  CloseHandle process_handle
               End If
            Else
               MsgBox "合併PDF失敗！"
               GoTo ErrHnd
            End If
            WebBrowser1.Navigate m_AttachPath & "\" & strMergeName
         End If
      End If
      'Add By Sindy 2015/5/27 開啟非PDF的電子檔
      For ii = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(ii, 0) = "v" Then
            '清除反白
            If UCase(cmdOpenAtt(1).Tag) <> UCase("call") Then
               GRD1.col = 0
               GRD1.row = ii
               GRD1.TextMatrix(ii, 0) = ""
               For jj = 1 To GRD1.Cols - 1
                  GRD1.col = jj
                  GRD1.CellBackColor = QBColor(15)
               Next jj
            End If
            If m_bolDblClick = False Or (m_bolDblClick = True And ii = m_mouseRow) Then
               If Trim(GRD1.TextMatrix(ii, 14)) <> "" Then
                  '讀取檔案名稱
                  'Modify By Sindy 2021/10/14 例如:
                  '  grd1.TextMatrix(ii, 14)= 20210818 1443 Reply (w POA).REPLY.msg
                  '  grd1.TextMatrix(ii, 15)= KB00/KB0000845/KB0000845_20210819.103115.msg.001
'                  If Index = 0 Then
                     stFileName = Trim(GRD1.TextMatrix(ii, 14))
'                  Else
'                     If InStr(Trim(grd1.TextMatrix(ii, 15)), "/") > 0 Then
'                        stFileName = Mid(grd1.TextMatrix(ii, 15), InStrRev(Trim(grd1.TextMatrix(ii, 15)), "/") + 1)
'                     Else
'                        stFileName = Trim(grd1.TextMatrix(ii, 15))
'                     End If
'                  End If
                  'Modify By Sindy 2021/10/14 + And InStrRev(stFileName, " (") > InStrRev(stFileName, ".")
                  If InStrRev(stFileName, " (") > 0 And InStrRev(stFileName, " (") > InStrRev(stFileName, ".") Then
                     stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
                  End If
                  Call PUB_ChkFileTypeOpenExE(stFileName) 'Add By Sindy 2017/9/13
                  If InStr(stFileName, "\") = 0 Then
                     If GetAttachFile(Trim(GRD1.TextMatrix(ii, 15)), stFileName, m_AttachPath & "\" & stFileName) = False Then
                        'MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                     End If
                  End If
                  '開啟檔案
                  ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
               End If
            End If
         End If
      Next ii
   End If
   
ErrHnd:
   m_bolDblClick = False
   Screen.MousePointer = vbDefault
   ChDir App.path '目錄切回
End Sub

Private Function GetAttachFile(ByVal strCF06 As String, ByRef pFileName As String, Optional pSavePath As String) As Boolean
   Dim stAttPath As String
   
On Error GoTo ErrHnd

   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
   Else
      '改傳完整的檔案路徑:路徑+檔名
      If InStr(pSavePath, m_AttachPath) > 0 Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
      End If
      stAttPath = pSavePath
   End If
   
   GetAttachFile = PUB_GetFtpFile(strCF06, stAttPath, cTableName)

   pFileName = stAttPath 'Add By Sindy 2020/5/19
   
   Exit Function
   
ErrHnd:
   If Err.Number = 70 Then
      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub GRD1_DblClick()
   If GRD1.row > 0 Then
      m_bolDblClick = True
      cmdOpenAtt_Click 1
   End If
End Sub
Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
Dim strFileName As String
Dim strFileType As String
Dim jj As Integer

GRD1.row = GRD1.MouseRow
GRD1.col = GRD1.MouseCol
nRow = GRD1.row
nCol = GRD1.col

GRD1.Visible = False
If nRow > 0 And Trim(GRD1.TextMatrix(nRow, 13)) <> "" Then
   m_mouseRow = nRow 'GRD1.MouseRow '記錄目前Row
   
   '先將上筆有反白的資料列復恢
   If m_mouseRowOld > 0 And m_mouseRowOld <= (GRD1.Rows - 1) Then
      GRD1.row = m_mouseRowOld
      GRD1.col = 1
      For jj = 1 To GRD1.Cols - 1
         GRD1.col = jj
         GRD1.CellBackColor = QBColor(15)
      Next jj
      If GRD1.MouseCol <> 0 Then
         GRD1.TextMatrix(GRD1.row, 0) = ""
      End If
      'Call recovercolor(CInt(m_mouseRowOld))
   End If
   
   '將點選的資料列反白
   GRD1.row = m_mouseRow
   GRD1.col = 1
   
   'Add By Sindy 2019/12/6 +選到編號欄=複製
   If nCol = 1 Then
      GRD1.Visible = True
      GRD1.CellForeColor = &H0 '黑色
      If GRD1.TextMatrix(m_mouseRow, nCol) <> "" Then
          '複製編號至剪貼簿
          Clipboard.SetText GRD1.TextMatrix(m_mouseRow, nCol)
          GRD1.CellBackColor = QBColor(7)
          MsgBox "編號已複製", , MsgText(21)
      End If
      GRD1.Visible = False
   End If
   '2019/12/6 END
   
   strFileName = GRD1.TextMatrix(GRD1.row, 14) 'GetFileName()
   strFileType = Right(strFileName, Len(strFileName) - InStrRev(strFileName, ".") + 1)
   '資料列反白
   For jj = 1 To GRD1.Cols - 1
      GRD1.col = jj
      GRD1.CellBackColor = &HFFC0C0
   Next jj
   m_mouseRowOld = GRD1.row
   If GRD1.TextMatrix(GRD1.row, 0) = "V" Or GRD1.TextMatrix(GRD1.MouseRow, 0) = "v" Then
      GRD1.TextMatrix(GRD1.row, 0) = ""
   Else
      If GRD1.TextMatrix(GRD1.row, 14) <> "" Then
         If UCase(strFileType) = UCase(".PDF") Then
            GRD1.TextMatrix(GRD1.row, 0) = "V"
         Else
            GRD1.TextMatrix(GRD1.row, 0) = "v"
         End If
      End If
   End If
End If
GRD1.Visible = True
End Sub
'Add By Sindy 2019/3/6
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
GRD1.ToolTipText = ""
If GRD1.MouseRow <> 0 And GRD1.MouseCol > 0 Then
   If GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
      GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
   End If
End If
End Sub
'Private Sub Grd1_Click()
'   Dim i As Integer
'   With GRD1
'      .Visible = False
'      .row = .MouseRow
'      .col = 0
'      If .row <> 0 Then
'         If .Text = "V" Then
'            .Text = ""
'            For i = 0 To .Cols - 1
'               If i <> 1 Then
'                  .col = i
'                  .CellBackColor = QBColor(15)
'               End If
'            Next i
'         Else
'            .Text = "V"
'            For i = 0 To .Cols - 1
'               If i <> 1 Then
'                  .col = i
'                  .CellBackColor = &HFFC0C0
'               End If
'            Next i
'         End If
'      End If
'      .Visible = True
'   End With
'End Sub

''Modify By Sindy 2009/04/30
'Private Function GetCustData(p_stCust As String) As Boolean
'Dim strName As String
'
'   GetCustData = False
'
'   Select Case Left(p_stCust, 1)
'      Case "X"
'         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3,CU81 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "'"
'      Case "Y"
'         strExc(0) = "select fa31,fa04,rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) fa05,fa06,FA10 N3,FA46 from fagent where fa01='" & Left(p_stCust, 8) & "' and fa02='" & Right(p_stCust, 1) & "'"
'      Case "R"
'         strExc(0) = "select pcu36,pcu08,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06) pcu03,pcu07,PCU09 N3,PCU41 from potcustomer where pcu01='" & Left(p_stCust, 8) & "' and pcu02='" & Right(p_stCust, 1) & "'"
'      Case Else
'         MsgBox "往來對象必須為 X、Y 或 R 開頭", vbCritical + vbOKOnly, "檢核資料"
'         Exit Function
'   End Select
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''   lbl1 = ""
'   If intI = 1 Then
'      For intI = 1 To 3
'         If Not IsNull(RsTemp(intI)) Then
'            strName = RsTemp(intI)
'            Exit For
'         End If
'      Next
'
'      '依LoginUser和輸入人員之部門第一碼判斷部門權限, 相同者才可輸入查詢
'      '但M51不受限制
'      strExc(0) = "SELECT A.ST03,B.ST03 FROM STAFF A,STAFF B " & _
'                         "WHERE A.ST01 = '" & strUserNum & "' " & _
'                              "AND B.ST01 = '" & Trim(RsTemp(5)) & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If Trim(RsTemp(0)) <> "M51" And _
'            Left(Trim(RsTemp(0)), 1) <> Left(Trim(RsTemp(1)), 1) Then
'            MsgBox "您沒有維護此筆潛在客戶的往來記錄權限！"
'            Exit Function
'         End If
'      End If
'   Else
'      MsgBox "往來對象輸入錯誤！"
'      Exit Function
'   End If
''   lbl1 = strName
'
'   GetCustData = True
'End Function
