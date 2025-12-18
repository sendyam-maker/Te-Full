VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_L 
   BorderStyle     =   1  '單線固定
   Caption         =   "卷宗區查詢"
   ClientHeight    =   6468
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9060
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6468
   ScaleWidth      =   9060
   Tag             =   "加班資料"
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5415
      Left            =   4440
      TabIndex        =   15
      Top             =   330
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   9551
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
   Begin VB.CommandButton cmdUpload 
      BackColor       =   &H0000FFFF&
      Caption         =   "轉出"
      Height          =   315
      Left            =   540
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   360
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdMonitor 
      Caption         =   "商標監控"
      Height          =   315
      Left            =   30
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FF00&
      Caption         =   "點我展開"
      Height          =   345
      Left            =   4440
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   0
      Width           =   4515
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3675
      Left            =   0
      TabIndex        =   8
      Top             =   1650
      Width           =   4395
      _ExtentX        =   7747
      _ExtentY        =   6477
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V  |  總收文號  |  收文日  |  案件性質  |  檔案名稱| 最後修改日期"
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
      _Band(0).Cols   =   6
   End
   Begin VB.CheckBox Check1 
      Caption         =   "關閉預覽"
      Height          =   195
      Left            =   3270
      TabIndex        =   7
      Top             =   720
      Width           =   1065
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H008080FF&
      Caption         =   "？"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4140
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   1140
      Width           =   285
   End
   Begin VB.CommandButton cmdReName 
      Caption         =   "更名"
      Height          =   315
      Left            =   2550
      TabIndex        =   4
      Top             =   30
      Width           =   525
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "取消全選"
      Height          =   315
      Left            =   570
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdAddAtt 
      Caption         =   "新增"
      Height          =   315
      Left            =   1500
      TabIndex        =   2
      Top             =   30
      Width           =   525
   End
   Begin VB.CommandButton cmdRemAtt 
      Caption         =   "刪除"
      Height          =   315
      Left            =   2010
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   30
      Width           =   525
   End
   Begin VB.CommandButton cmdSaveAtt 
      Caption         =   "下載"
      Height          =   315
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   525
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   3630
      TabIndex        =   6
      Top             =   30
      Width           =   765
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "移檔"
      Height          =   315
      Left            =   3090
      TabIndex        =   5
      Top             =   30
      Width           =   525
   End
   Begin VB.CheckBox ChkDelC 
      Caption         =   "含客戶提供文件(紅色列)"
      Height          =   195
      Left            =   1650
      TabIndex        =   21
      Top             =   1200
      Width           =   2235
   End
   Begin VB.CheckBox ChkDelMsg 
      Caption         =   "不顯示郵件"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   90
      TabIndex        =   22
      Top             =   1440
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2610
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000FF00&
      Caption         =   "補看確認(&O)"
      Height          =   315
      Index           =   1
      Left            =   3075
      Style           =   1  '圖片外觀
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdCopy 
      BackColor       =   &H0000C0C0&
      Caption         =   "複製到..."
      Height          =   315
      Left            =   2025
      Style           =   1  '圖片外觀
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   360
      Width           =   1035
   End
   Begin VB.CommandButton cmdEMail 
      Caption         =   "EMail(&E)"
      Height          =   315
      Left            =   990
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   1065
      Left            =   30
      TabIndex        =   16
      Top             =   5340
      Width           =   4755
      Begin VB.Frame Frame2 
         Caption         =   "電腦中心使用"
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   30
         TabIndex        =   29
         Top             =   390
         Width           =   4395
         Begin VB.CommandButton cmdFlag 
            Caption         =   "上合併註記"
            Height          =   315
            Left            =   3180
            Style           =   1  '圖片外觀
            TabIndex        =   32
            Top             =   240
            Width           =   1125
         End
         Begin VB.CommandButton cmdSwap 
            BackColor       =   &H00C0C0FF&
            Caption         =   "抽換"
            Height          =   315
            Left            =   2220
            Style           =   1  '圖片外觀
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox Check2 
            Caption         =   "單檔也用合併方式開啟"
            Height          =   225
            Left            =   60
            TabIndex        =   34
            Top             =   240
            Width           =   2115
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Left            =   1830
            Style           =   2  '單純下拉式
            TabIndex        =   33
            Top             =   450
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.CheckBox ChkDelF 
            Caption         =   "含已刪除檔 (紅色列)"
            Height          =   195
            Left            =   1710
            TabIndex        =   31
            Top             =   30
            Width           =   1905
         End
      End
      Begin VB.CommandButton cmdEmpFlow 
         Caption         =   "承辦歷程"
         Height          =   315
         Left            =   1500
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   60
         Width           =   885
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   315
         Index           =   0
         Left            =   930
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   60
         Width           =   555
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "多檔預覽"
         Height          =   315
         Index           =   1
         Left            =   30
         TabIndex        =   17
         Top             =   60
         Width           =   870
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(註:雙擊單檔預覽)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   1230
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   19
      Left            =   150
      TabIndex        =   12
      Top             =   750
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   18
      Left            =   150
      TabIndex        =   11
      Top             =   990
      Width           =   960
   End
   Begin VB.Label lblCaseNo 
      Height          =   180
      Left            =   1140
      TabIndex        =   10
      Top             =   750
      Width           =   1830
   End
   Begin MSForms.Label lblCaseName 
      Height          =   285
      Left            =   1125
      TabIndex        =   9
      Top             =   990
      Width           =   3255
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblFtpPath 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "轉出至路徑: "
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3360
      TabIndex        =   28
      Top             =   1470
      Visible         =   0   'False
      Width           =   990
   End
End
Attribute VB_Name = "frm100101_L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/4/23 Form2.0已修改
'Create by Sindy 2013/6/11
Option Explicit

' 變數宣告區
Public m_CP09 As String
Public m_CP10Nm As String
Public m_Appl As String
Public m_Nation As String
Public m_strKey As String '本所案號 或 多筆總收文號
Public m_CPP11 As String 'Add By Sindy 2023/2/18 電子表單單號

Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_CP10 As String
Dim m_CP13 As String 'Add By Sindy 2014/12/11
Dim m_CP14 As String 'Add By Sindy 2014/12/11

Dim ii As Integer, jj As Integer
Dim m_PrevForm As Form '前一畫面
'Added by Lydia 2020/02/20
Dim m_PrevFormOld As Form  '重複呼叫的再前一畫面
Dim bolActive As Boolean '是否已觸發 Form Active 事件
'end 2020/02/20

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

''列印宣告區
'Private Declare Function ShellExecute Lib _
'"shell32.dll" Alias "ShellExecuteA" ( _
'ByVal hwnd As Long, ByVal lpOperation As String, _
'ByVal lpFile As String, ByVal lpParameters As String, _
'ByVal lpDirectory As String, _
'ByVal nShowCmd As Long) As Long
'Private Const SW_HIDE = 0

Dim m_bolDblClick As Boolean
Dim m_mouseRow As Long, m_mouseRowOld As Long
Const GrdMaxW = 11640 '8865
Const GrdMinW = 4395
Dim bolReadOnlyRev As Boolean
Dim m_identity As String '身份
Dim m_Flag1 As Integer, m_Flag2 As Integer 'Added by Lydia 2018/04/12 記錄判斷變色的欄位
Dim m_QueryEfile As String 'Add By Sindy 2020/9/7 鎖可查詢的系統別及副檔名
Dim m_LimitType As String, m_RecvNo As String 'Add By Sindy 2020/12/31
Dim m_strFolder As String 'Add By Sindy 2022/11/1
Dim strYYMMSql As String 'Added by Lydia 2023/12/22 限制年月
Private nfrm090128_New As Form 'Added by Lydia 2024/11/11 查名單(網中)：查名單明細作業

Private Sub SetGrd(Optional bolSetRow As Boolean = True)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2014/5/23 +發文室,寄件方式
   'Modify By Sindy 2014/7/24 +CPP03
   'Modify By Sindy 2015/5/26 +cpp12,cpp05,cpp06
   'Modify By Sindy 2017/4/6 +cpp04
   'Modify By Sindy 2018/11/20 +cpp07
   'Modify by Sindy 2019/10/30 發文日改專業發文日
   'Modify by Sindy 2023/3/3 +cpp11
   'Modify by Sindy 2024/3/29 +R01001
   'Modified by Morgan 2025/7/21 修改日期時間->上傳時間(要顯示)
   '                        0    1           2             3           4           5             6       7       8               9                   10              11      12        13        14          15                16       17      18      19      20      21       22       23       24       25       26      27       28       29       30       31       32       33       34       35
   arrGridHeadText = Array("V", "總收文號", "專業發文日", "案件性質", "檔案名稱", "副檔名說明", "CP09", "CP10", "檔案修改時間", "檔案重覆另外命名", "修改日期時間", "CP82", "收文日", "發文室", "寄件方式", "CPP10_Flag註記", "CPP03", "CP05", "CP66", "CP67", "sort", "cpp12", "cpp05", "cpp06", "cpp08", "cpp09", "cp43", "cpp04", "cpp02", "cpp07", "cpp15", "cpp11", "cpp16", "cpp17", "cpp18", "R01001")
   'Modify By Sindy 2017/9/20 方便檢視某些欄位值
   'Modify By Sindy 2019/11/15
   'If Pub_StrUserSt03 = "M51" And strUserNum <> "74001" Then
   If Pub_StrUserSt03 = "M51" Then
   '2019/11/15 END
      arrGridHeadWidth = Array(330, 200, 800, 1000, 2500, 1000, 0, 0, 1200, 0, 1200, 0, 800, 800, 500, 500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   '2017/9/20 END
   ElseIf Check1.Value = 0 Then
      arrGridHeadWidth = Array(330, 200, 800, 1000, 1850, 1000, 0, 0, 1200, 0, 1200, 0, 800, 800, 500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   Else
      arrGridHeadWidth = Array(330, 1000, 800, 1000, 2500, 1000, 0, 0, 1200, 0, 1200, 0, 800, 800, 500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   End If
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
         GRD1.CellAlignment = flexAlignCenterCenter
      End If
   Next
   'Added by Lydia 2018/04/12 記錄判斷變色的欄位
   If m_Flag1 = 0 Then
       m_Flag1 = PUB_MGridGetId("CPP10_Flag註記", GRD1)
   End If
   If m_Flag2 = 0 Then
       m_Flag2 = PUB_MGridGetId("cp43", GRD1)
   End If
   'end 2018/04/12
   
   GRD1.Visible = True
End Sub

'檢查權限
Private Function CheckLimits() As Boolean
Dim strMsgTxt As String 'Add By Sindy 2022/7/18
Dim bolSpecCase As Boolean 'Add By Sindy 2023/4/27
   
   CheckLimits = False '無權限
   m_QueryEfile = "" '鎖可查詢的系統別及副檔名
   
   'Modify By Sindy 2023/4/27 教威無法查ACS-000166投標:應該是先做CheckSR09檢查權限，有權限的人再檢查ACS特殊案的權限，非特殊案就不用再檢查了。
   bolSpecCase = PUB_ChkCPPAndCPFLimits_Spec(m_CP01, m_CP02, m_CP03, m_CP04, m_LimitType, m_RecvNo, strMsgTxt)
   'Modify By Sindy 2023/5/2
   If bolSpecCase = False Then
   '2023/5/2 END
      If CheckSR09(strUserNum, m_CP01, "Y", , m_CP01, m_CP02, m_CP03, m_CP04) = False Then
         Exit Function
      End If
   End If
   '2023/4/27 END
   
   'Add by Sindy 2020/12/31 原始檔及卷宗區特殊權限
   'If PUB_ChkCPPAndCPFLimits_Spec(m_CP01, m_CP02, m_CP03, m_CP04, m_LimitType, m_RecvNo, strMsgTxt) = True Then
   If bolSpecCase = True Then
      If m_LimitType = "" Then
         'Modify By Sindy 2022/11/9 開放專利商標程序人員「僅具有閱覽接洽單之權限」。
         'Modify By Sindy 2022/12/31 開放財務處人員「僅具有閱覽接洽單之權限」。
         'Modify By Sindy 2023/2/8 開放就個案之智權同仁及其區主管與杜協理閱覽ACS案件卷宗區『案件接洽單』之權限
         If Pub_StrUserSt03 = "P12" Or _
            Pub_StrUserSt03 = "P22" Or _
            (m_CP01 = "ACS" And _
             (Pub_StrUserSt03 = "M31" Or m_identity = "S" Or Left(Pub_StrUserSt03, 1) = "S" Or _
              InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0) _
            ) Then
            m_QueryEfile = " AND (e1.efc01='ALL' OR e1.efc01='" & m_CP01 & "') AND (instr(e1.efc03,'接洽單')>0 OR instr(e1.efc04,'接洽單')>0) AND instr(upper(cpp02),'.'||e1.EFC02||'.')>0"
         Else
         '2022/11/9 END
            'Modify By Sindy 2021/8/19
            'MsgBox "您沒有查詢案件明細的權限", vbOKOnly, "檢核資料"
            MsgBox "此案與ACS" & strMsgTxt & "有關，您無權限查詢，您可請顧服組協助！", vbOKOnly, "檢核資料"
            '2021/8/19 END
            Exit Function
         End If
      End If
   Else
   '2020/12/31 END
      'Add By Sindy 2020/9/7 調整程式只開放W2.顧服組人員及W0.顧服組主管
      '可以查P案卷宗區的說明書，其他檔案及原始檔區都不開。
      'A5024.王�[娟 除外
      If m_CP01 = "P" Then
         If (PUB_GetST05(strUserNum) = "W2" Or PUB_GetST05(strUserNum) = "W0") _
            And strUserNum <> "A5024" Then
            m_QueryEfile = " AND (e1.efc01='ALL' OR e1.efc01='P') AND (instr(e1.efc03,'說明書')>0 OR instr(e1.efc04,'說明書')>0) AND instr(upper(cpp02),'.'||e1.EFC02||'.')>0"
         End If
      End If
      '2020/9/7 END
   End If
   CheckLimits = True '有權限
End Function

'Modify by Amy 2025/05/19 +CPP11
Public Function QueryData(Optional ByVal stCPP11 As String = "") As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim varSplit As Variant
Dim pYYMM As String 'Added by Lydia 2023/12/22 限制收文年月

   '清空及預設欄位值
   GRD1.Clear
   SetGrd
   lblCaseNo.Caption = Empty
   lblCaseName.Caption = Empty

   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   If m_strKey = "" Then Exit Function 'Add By Sindy 2015/12/17
   
   If InStr(m_strKey, "-") = 0 Then '總收文號
      bolReadOnlyRev = True
      pub_QL05 = ";總收文號：" & m_strKey & "(卷宗區)" 'Add By Sindy 2025/8/7
      varSplit = Split(m_strKey, ",")
      strSql = "Select cp01,cp02,cp03,cp04,CP13,CP14" & _
               " From CaseProgress" & _
               " Where CP09='" & varSplit(0) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         m_CP01 = rsTmp.Fields("cp01")
         m_CP02 = rsTmp.Fields("cp02")
         m_CP03 = rsTmp.Fields("cp03")
         m_CP04 = rsTmp.Fields("cp04")
         'm_CP13 = "" & rsTmp.Fields("CP13") 'Add By Sindy 2014/12/11
         m_CP14 = "" & rsTmp.Fields("CP14") 'Add By Sindy 2014/12/11
         
'         '檢查權限
'         If CheckLimits() = False Then
'            'tmpBol = fnCancelNowFormAndShowParentForm(Me)
'            Screen.MousePointer = vbDefault
'            Set rsTmp = Nothing
'            Call cmdExit_Click
'            Exit Function
'         End If
         
         If InStr(m_strKey, ",") > 0 Then
            m_strKey = " cp09 in('" & Replace(Trim(m_strKey), ",", "','") & "') "
         Else
            m_strKey = " cp09 in('" & m_strKey & "') "
         End If
      Else
         Screen.MousePointer = vbDefault
         ShowNoData
         rsTmp.Close
         Set rsTmp = Nothing
         Call cmdExit_Click
         Exit Function
      End If
      rsTmp.Close
      pub_QL05 = ";本所案號：" & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & pub_QL05 'Add By Sindy 2025/8/7
      
   Else '本所案號
      bolReadOnlyRev = False
      pub_QL05 = ";本所案號：" & m_strKey & "(卷宗區)" 'Add By Sindy 2025/8/7
      m_CP01 = SystemNumber(m_strKey, 1)
      m_CP02 = SystemNumber(m_strKey, 2)
      m_CP03 = SystemNumber(m_strKey, 3)
      m_CP04 = SystemNumber(m_strKey, 4)
            
'Added by Lydia 2023/12/22 當以本所案號查詢時，若條件為TT-999999則點進度時再增加輸入收文年月對話框(預設當月，取消表示全部)以減少等待時間。
    pYYMM = ""
    If InStr("TT999999,LA999999", m_CP01 & m_CP02) > 0 Then
JumpToReInput:
        pYYMM = InputBox("請輸入收文年月以減少等待時間，不輸入年月或按取消表示查詢全部資料。", "輸入收文年月", Left(strSrvDate(2), 5))
        If pYYMM <> "" Then
           If Val(Left(pYYMM, 5)) > Val(Left(strSrvDate(2), 5)) Then
               MsgBox "收文年月不可大於系統年月！", vbInformation
               GoTo JumpToReInput
           End If
           strYYMMSql = strYYMMSql & " AND CP05>=" & Val(pYYMM & "01") + 19110000 & " AND CP05<=" & Val(pYYMM & "31") + 19110000 & " "
           Me.Caption = Me.Caption & "-收文年月：" & pYYMM
        End If
    End If
'end 2023/12/22

'      '檢查權限
'      If CheckLimits = False Then
'         'tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         Screen.MousePointer = vbDefault
'         Set rsTmp = Nothing
'         Call cmdExit_Click
'         Exit Function
'      End If
      
      'Modify By Sindy 2017/2/9 +,cp14
      'Modified by Lydia 2023/12/22 + strYYMMSql
      strSql = "Select distinct cp09,cp14" & _
               " From CaseProgress,Acc090" & _
               " Where CP01='" & m_CP01 & "' and CP02='" & m_CP02 & "' and CP03='" & m_CP03 & "' and CP04='" & m_CP04 & "' AND CP12=A0901(+)" & strYYMMSql
      'Add By Sindy 2020/5/12
      '以操作人員之st15再抓ACC090之A0911，再以進度CP12再抓ACC090之A0911，若與操作人員帶出之A0911相同才可顯示進度。卷宗區也是。
      'st05 in ('00',’01’)人員不受上述限制
      'Modify By Sindy 2022/5/23 再加入系統特殊設定「全所智權部主管」的人員也不限制。
      If InStr("00,01", Pub_strUserST05) = 0 And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) = 0 And m_CP01 = "TT" And m_CP02 = "999999" Then
         strSql = strSql + " AND A0911=(select a90.A0911 from acc090 a90 where a90.a0901='" & Pub_StrUserSt15 & "') "
      End If
      '2020/5/12 END
      strSql = strSql & " order by cp09 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         m_strKey = ""
         Do While Not rsTmp.EOF
            'm_strKey = m_strKey & "'" & Trim(rsTmp.Fields("cp09")) & "'," 'Modify By Sindy 2021/10/15 Mark
            'Modify By Sindy 2017/2/9
            If InStr(m_CP14, Trim(rsTmp.Fields("cp14"))) = 0 Then
               m_CP14 = m_CP14 & "," & Trim(rsTmp.Fields("cp14"))
            End If
            '2017/2/9 END
            rsTmp.MoveNext
         Loop
         If m_CP14 <> "" Then m_CP14 = Mid(m_CP14, 2) 'Add By Sindy 2017/2/9
         'Modify By Sindy 2021/10/15
         'm_strKey = Left(m_strKey, Len(m_strKey) - 1)
         'Modified by Lydia 2023/12/22 + strYYMMSql
         m_strKey = " cp09 in(select cp09 from caseprogress,Acc090 where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "' AND CP12=A0901(+)" & strYYMMSql
         'Add By Sindy 2020/5/12
         '以操作人員之st15再抓ACC090之A0911，再以進度CP12再抓ACC090之A0911，若與操作人員帶出之A0911相同才可顯示進度。卷宗區也是。
         'st05 in ('00',’01’)人員不受上述限制
         'Modify By Sindy 2022/5/23 再加入系統特殊設定「全所智權部主管」的人員也不限制。
         If InStr("00,01", Pub_strUserST05) = 0 And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) = 0 And m_CP01 = "TT" And m_CP02 = "999999" Then
            m_strKey = m_strKey + " AND A0911=(select a90.A0911 from acc090 a90 where a90.a0901='" & Pub_StrUserSt15 & "') "
         End If
         '2020/5/12 END
         m_strKey = m_strKey + ") "
         '2021/10/15 END
      Else
         Screen.MousePointer = vbDefault
         ShowNoData
         rsTmp.Close
         Set rsTmp = Nothing
         Call cmdExit_Click
         Exit Function
      End If
      rsTmp.Close
   End If
   
   'Added by Lydia 2018/02/01 FCP含已發文之客戶提供文件
   ChkDelC.Visible = False
   strSql = "select count(*) from custsupportdoc where csd01='" & m_CP01 & "' and csd02='" & m_CP02 & "' and csd03='" & m_CP03 & "' and csd04='" & m_CP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
       If Val("" & rsTmp(0)) > 0 Then
           ChkDelC.Visible = True
       End If
   End If
   rsTmp.Close
   'end 2018/02/01
   
   '案件資料
   strSql = "Select PA01||'-'||PA02||'-'||PA03||'-'||PA04 as 本所案號,PA05||PA06||PA07 as 案件名稱,PA09 as 申請國家,pa26 as 申請人1" & _
            " From Patent" & _
            " Where PA01='" & m_CP01 & "' And PA02='" & m_CP02 & "' And PA03='" & m_CP03 & "' And PA04='" & m_CP04 & "'"
   strSql = strSql & " union Select TM01||'-'||TM02||'-'||TM03||'-'||TM04 as 本所案號,TM05 as 案件名稱,TM10 as 申請國家,tm23 as 申請人1" & _
            " From Trademark" & _
            " Where TM01='" & m_CP01 & "' And TM02='" & m_CP02 & "' And TM03='" & m_CP03 & "' And TM04='" & m_CP04 & "'"
   strSql = strSql & " union Select SP01||'-'||SP02||'-'||SP03||'-'||SP04 as 本所案號,SP05||SP06||SP07 as 案件名稱,SP09 as 申請國家,sp08 as 申請人1" & _
            " From Servicepractice" & _
            " Where SP01='" & m_CP01 & "' And SP02='" & m_CP02 & "' And SP03='" & m_CP03 & "' And SP04='" & m_CP04 & "'"
   strSql = strSql & " union Select HC01||'-'||HC02||'-'||HC03||'-'||HC04 as 本所案號,HC06 as 案件名稱,'000' as 申請國家,hc05 as 申請人1" & _
            " From Hirecase" & _
            " Where HC01='" & m_CP01 & "' And HC02='" & m_CP02 & "' And HC03='" & m_CP03 & "' And HC04='" & m_CP04 & "'"
   strSql = strSql & " union Select LC01||'-'||LC02||'-'||LC03||'-'||LC04 as 本所案號,LC05||LC06||LC07 as 案件名稱,LC15 as 申請國家,lc11 as 申請人1" & _
            " From Lawcase" & _
            " Where LC01='" & m_CP01 & "' And LC02='" & m_CP02 & "' And LC03='" & m_CP03 & "' And LC04='" & m_CP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("本所案號")) Then lblCaseNo.Caption = rsTmp.Fields("本所案號")
      If Not IsNull(rsTmp.Fields("案件名稱")) Then lblCaseName.Caption = rsTmp.Fields("案件名稱")
      m_Nation = ""
      If Not IsNull(rsTmp.Fields("申請國家")) Then m_Nation = rsTmp.Fields("申請國家")
      m_Appl = ""
      If Not IsNull(rsTmp.Fields("申請人1")) Then m_Appl = rsTmp.Fields("申請人1")
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Call cmdExit_Click
      Exit Function
   End If
   rsTmp.Close
   
   m_CP13 = ShowCurrCP13(m_CP01, m_CP02, m_CP03, m_CP04, m_Nation) 'Modify By Sindy 2014/5/26
   Call ChkModifyLimits
   'Modify By Sindy 2023/2/8 改檢查權限的位置,從上方Move至此處
   '檢查權限
   If CheckLimits() = False Then
      'tmpBol = fnCancelNowFormAndShowParentForm(Me)
      'Add By Sindy 2024/5/10 因杜協理的案件職代是北所的智權部主管輪職值,會有跨所權限不足的問題
      '                       發生時,開放此結案單和NP01文號的回覆單可以查看
      If UCase(m_PrevForm.Name) = UCase("frm210148_1") Then
         m_strKey = " cp09 in('" & m_PrevForm.m_F0303 & "') "
      Else
      '2024/5/10 END
         Screen.MousePointer = vbDefault
         Set rsTmp = Nothing
         Call cmdExit_Click
         Exit Function
      End If
   End If
   '2023/2/8 END
   
   'Modify By Sindy 2015/3/9 Mark:不分系統別都可以查看
'   If m_CP01 <> "P" And m_CP01 <> "CFP" And m_CP01 <> "PS" And m_CP01 <> "CPS" Then
'      Screen.MousePointer = vbDefault
'      MsgBox "無卷宗區可查詢！"
'      Set rsTmp = Nothing
'      Call cmdExit_Click
'      Exit Function
'   End If
   
   'Modify by Amy 2025/05/19 +stCPP11,FC結案單需抓.MSG
   Call ReadAttachFile(stCPP11)
   
   'Added by Morgan 2018/10/29
   If strSrvDate(1) >= e化客戶啟用日 Then
      'Modified by Morgan 2021/12/2
      '外商程序承辦由此處EMail通知函給客戶
      If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "F11" Or Pub_StrUserSt03 = "F12" Then
         'Modified by Morgan 2019/5/2 控制E化客戶案件才可EMail(指定信箱)
         cmdEmail.Visible = PUB_ChkECustCase(m_CP01, m_CP02, m_CP03, m_CP04)
      End If
   End If
   'end 2018/10/29
   
   'Add By Sindy 2022/11/1 專利處程序分案會操作轉出
   If (UCase(m_PrevForm.Name) = UCase("frm040101_1") Or UCase(m_PrevForm.Name) = UCase("frm050101_2")) And _
      strSrvDate(1) >= 接洽單電子收文啟用日 Then
      cmdUpload.Left = cmdMonitor.Left
      cmdUpload.Visible = True
      If UCase(m_PrevForm.Name) = UCase("frm040101_1") Then
         If m_Nation = "000" Then
            m_strFolder = str_P_OrderPath & "\POA"
         Else
            m_strFolder = str_P_OrderPath & "\POA_CN"
         End If
      Else
         m_strFolder = str_CFP_OrderPath & "\POA"
      End If
      If Dir(m_strFolder, vbDirectory) = "" Then MkDir m_strFolder
      LblFtpPath.Caption = "轉出路徑:" & m_strFolder
      LblFtpPath.Visible = True
   End If
   
   QueryData = True
   Screen.MousePointer = vbDefault
   Me.Enabled = True

EXITSUB:
   Set rsTmp = Nothing
End Function

'檢查可新增刪除的權限
Private Sub ChkModifyLimits()
   'm_identity.身份：A.業務助理 C.電腦中心 S.智權人員 E.承辦人/工程師 F.程序人員 W.檔案室
   'Modify By Sindy 2020/3/16 改成共用函數
   m_identity = PUB_ChkCPPAndCPFLimits(m_CP01, m_CP02, m_CP03, m_CP04, m_CP13)
   
   cmdAddAtt.Enabled = False '新增
   cmdRemAtt.Enabled = False '刪除
   cmdReName.Enabled = False '更名 'Add By Sindy 2014/11/27
   cmdMove.Enabled = False '搬檔 'Add By Sindy 2015/5/27
   cmdCopy.Enabled = False '複製 'Add By Sindy 2021/10/21
   If m_identity <> "" And InStr("A,C,S,E,F", m_identity) > 0 Then
      cmdAddAtt.Enabled = True '新增
      cmdRemAtt.Enabled = True '刪除
      
      If m_identity = "S" And m_CP13 = strUserNum Then cmdReName.Enabled = True 'Add By Sindy 2023/1/13 開啟更名功能->供智權人員使用(限自已的案件)
      
      If InStr("C,E,F", m_identity) > 0 Then
         cmdReName.Enabled = True '更名 'Add By Sindy 2014/11/27
         cmdMove.Enabled = True '搬檔 'Add By Sindy 2015/5/27
         cmdCopy.Enabled = True '複製 'Add By Sindy 2021/10/21
      End If
   'Add By Sindy 2015/7/27
   ElseIf m_identity <> "" And InStr("W", m_identity) > 0 Then '檔案室
      cmdAddAtt.Enabled = True '新增
      cmdReName.Enabled = True '更名
   '2015/7/27 END
   End If
End Sub

'查詢附件檔
'Modify By Sindy 2015/5/26
'Private Function ReadAttachFile() As Boolean
'Modify by Amy 2025/05/19 +stCPP11
Public Function ReadAttachFile(Optional ByVal stCPP11 As String) As Boolean
'2015/5/26 END
Dim rsTmp As New ADODB.Recordset
Dim strConSql As String 'Add By Sindy 2014/5/16
Dim strCP10NoShow As String 'Add By Sindy 2014/12/11
Dim strChkLostCP09 As String 'Add By Sindy 2014/12/11
Dim strCon2 As String 'Added by Lydia 2018/09/06 每一個語法都加上
Dim strSqlwhere As String 'Add By Sindy 2021/10/19
Dim strSqlCPP11 As String 'Add by Amy 2025/05/19

   ReadAttachFile = True
   
   'Add By Sindy 2024/3/27
   strSql = "delete from R100101_L where ID='" & strUserNum & "'"
   cnnConnection.Execute strSql
   '2024/3/27 END
   
   'Add By Sindy 2014/5/16 電腦中心及專業部才可以看到承辦單
   'Modify By Sindy 2016/7/12 + 開放檔案室可以看全部附件(Or m_identity = "W")
   strConSql = ""
   'Modify By Sindy 2018/7/18 改鎖智權人員不可以看其附件(承辦單,.altr代理人來函)
'   If Pub_StrUserSt03 = "M51" Or m_identity = "W" Or _
'      (Pub_StrUserSt03 >= "P10" And Pub_StrUserSt03 < "P19") Then
'      '全部附件都可看
'   Else
   'Modified by Morgan 2019/7/22 +W1開頭(客戶服務組)--秀玲
   'If Left(Pub_StrUserSt03, 1) = "S" Then
   'Modify By Sindy 2020/9/7 顧服組另外控制
   'If Left(Pub_StrUserSt03, 1) = "S" Or Left(Pub_StrUserSt03, 2) = "W1" Then
   If Left(Pub_StrUserSt03, 1) = "S" Then
   '2020/9/7 END
   '2018/7/18 END
      'strConSql = strConSql & " and (instr(upper(cpp02),upper('" & EMP_承辦單 & "'))=0 or cpp02 is null)"
      'Modify By Sindy 2014/11/26 加 電子送件時,不要看到.DWG
      'Modify By Sindy 2020/10/30 + EMP_多案承辦單,不要看到
      'Modify By Sindy 2024/3/19 卷宗區開放智權部人員可以查看 "非PDF" 的承辦單。 + & ".pdf'))=0
      strConSql = strConSql & " and (instr(upper(cpp02),upper('." & EMP_承辦單 & ".pdf'))=0 or cpp02 is null)" & _
                              " and (instr(upper(cpp02),upper('." & EMP_多案承辦單 & ".pdf'))=0 or cpp02 is null)" & _
                              " and (((cp118 is not null and instr(upper(cpp02),upper('.DWG'))=0) or cpp02 is null) or cp118 is null)"
      '2024/3/19 END
      
      'Add By Sindy 2015/1/15 加 CFP的.Altr代理人來函不可看到
      'If m_CP01 = "CFP" Then
         strConSql = strConSql & " and (instr(upper(cpp02),upper('.altr'))=0 or cpp02 is null)"
      'End If
      'Add By Sindy 2018/11/1 桂英:限制 ALTR 和 INVOICE 智權人員不可隨意開啟
      strConSql = strConSql & " and (instr(upper(cpp02),upper('.INVOICE'))=0 or cpp02 is null)"
      'Add By Sindy 2020/2/15 玫音:報價：quote (控管同"altr"，不開放給智權同仁)
      strConSql = strConSql & " and (instr(upper(cpp02),upper('.QUOTE'))=0 or cpp02 is null)"
      'Add By Sindy 2021/4/22 雅娟:陸代郵件：PAT (控管同"altr"，不開放給智權同仁)
      strConSql = strConSql & " and (instr(upper(cpp02),upper('.PAT.'))=0 or cpp02 is null)"
      
      'Added by Morgan 2019/7/22 電子化客戶函未判發不可看--郭
      strConSql = strConSql & " and not (instr(upper(cpp02),upper('.cus.'))>0 and lp01 is not null and lp05=0)"
   End If
   '2014/5/16 END
   'Added by Lydia 2015/11/16 查名單電子化
   'Modifie by Lydia 2016/04/25 +TS案
   If m_CP01 = "T" Or m_CP01 = "TS" Then
      '不可直接看結果附件,要經由查覆明細畫面
      strConSql = strConSql & " and instr(upper(cpp02),upper('." & UCase(TMQ_查名作業 & ".pdf") & "'))=0"
   End If
   'end 2015/11/16
   
   'Add By Sindy 2016/6/22 檢查是否有查看開庭紀要(brief)及電子筆錄(note)的權限
   ' or cpp02 is null) ==> 因文號有可能都沒有放電子檔,但還是要顯示出來
   If PUB_GetLimitToBRIEF(m_CP01, m_CP02, m_CP03, m_CP04, m_strKey) = False Then
      'Modify By Sindy 2017/1/4 ex.LA-003005
'      strConSql = strConSql & " and ((instr(upper(cpp02),upper('.brief.'))=0" & _
'                                   " and instr(upper(cpp02),upper('.note.'))=0" & _
'                                   " and substr(upper(cpp02),1,1)<>'L' and substr(upper(cpp02),1,3)<>'FCL' and substr(upper(cpp02),1,3)<>'CFL') " & _
'                                   " or cpp02 is null)"
      strConSql = strConSql & " and ((instr(upper(cpp02),upper('.brief.'))=0 and instr(upper(cpp02),upper('.note.'))=0) or cpp02 is null)"
   End If
   '2016/6/22 END
   
   'Added by Lydia 2018/09/06
   strCon2 = ""
   If ChkDelMsg.Value = 1 Then '排除郵件檔
       '因為FCP案有匯入外來和寄出郵件,造成卷宗區檔案眾多;在FCP不印說明書之後,造成核稿人一開始無法很容易找到翻譯用之最終提申本,所以勾選此項排除郵件檔,方便尋找
       strCon2 = strCon2 & " and Upper(cpp02) not like '%.MSG' "
   End If
   'end 2018/09/06
   
   Screen.MousePointer = vbHourglass
   KillAttach
   GRD1.Clear
   SetGrd
   GRD1.FixedCols = 0
'   select cpp01,cp01,cp02,cp03,cp04 from casepaperpdf,caseprogress where cpp10='D'
'   and cpp01=cp09(+)
   'Modify By Sindy 2014/7/24 +CPP03
   'Modify By Sindy 2014/11/25 +副檔名說明
'   strSql = "Select ' ' as V,cp09 as 總收文號,sqldatet(cp27) as 專業發文日,Decode('" & m_Nation & "','000',CPM03,CPM04) as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,cp09,CP10,sqldatet(cpp08)||' '||sqltime(cpp09) as 檔案修改時間,' ',cpp08||cpp09 as 修改日期時間,CP82,sqldatet(cp05) as 收文日,sqldatet(cp127) as 發文室,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,CPP10,CPP03" & _
'            " From CasepaperPDF,caseprogress,Casepropertymap,LetterProgress" & _
'            " Where cp09 in(" & m_strKey & ")" & _
'            " And cp01='" & m_CP01 & "' And cp02='" & m_CP02 & "' And cp03='" & m_CP03 & "' And cp04='" & m_CP04 & "'" & _
'            " And cp09=cpp01(+) and (cpp10 is null or cpp10='X' or cpp10='Y'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, " or cpp10='D'", "") & ")" & _
'            " And cp01=cpm01(+) And cp10=cpm02(+)" & strConSql & _
'            " and cp09=LP01(+)" & _
'            " order by SQLDatet2(CP05) desc,CP66 desc,CP67 desc,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3') desc,CP09 desc,cpp11 asc,cpp08 desc,cpp09 desc"
   
   'Modify By Sindy 2014/12/11 過濾某些案件性質不可看到附件
   '查詢權限僅電腦中心、智權人員(若智權人員離職才開放其區主管)、承辦人、游經理、王副總
   '211.準備程序 212.言詞辯論 213.現場堪察 226.配合開庭 408.面詢
   strCP10NoShow = ""
   'Modify By Sindy 2015/1/8 73.內專程序主管 或 75.內專程序 也可以看
   'Modify By Sindy 2016/7/5 83.CFP程序主管 或 85.CFP程序 也可以看
   'Modify By Sindy 2016/7/12 + 開放檔案室可以看全部附件(Or m_identity = "W")
   'Modify By Sindy 2017/2/9 Or strUserNum = m_CP14 ==> Or InStr(m_CP14, strUserNum) > 0
   'Modified by Lydia 2017/03/28 chkmailid取得可能不只一人 strUserNum = ChkMailId(m_CP13)=> InStr(ChkMailId(m_CP13), strUserNum) > 0
   'Modify By Sindy 2018/7/18 外專這些案件性質不需要鎖附件
   'Modify By Sindy 2020/2/25 游經理:請開啟張偉城( 89026 )可查看配合訴訟案件內容
   'Modify By Sindy 2023/3/28 游經理:開放法律所人員可查看案件性質408(面詢)、213(現場勘查)、211(準備程序)、212(言詞辯論)、226(配合開庭)等案件性質的附件。
   '                          以使與律師配合的訴訟案件可順利運作。 ex:P-120983
   'Modified by Morgan 2025/2/4 +P10部門
   If Not (Pub_StrUserSt03 = "P10" Or Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 2) = "F2" Or strUserNum = m_CP13 Or InStr(ChkMailId(m_CP13), strUserNum) > 0 Or m_identity = "W" _
           Or InStr(m_CP14, strUserNum) > 0 Or PUB_GetST05(strUserNum) = "72" Or PUB_GetST05(strUserNum) = "71" _
           Or PUB_GetST05(strUserNum) = "73" Or PUB_GetST05(strUserNum) = "75" _
           Or PUB_GetST05(strUserNum) = "83" Or PUB_GetST05(strUserNum) = "85" _
           Or InStr("'89026'", strUserNum) > 0 Or Left(Pub_StrUserSt03, 1) = "L") Then
      'Add By Sindy 2017/9/27 游經理:CFP案的案件性質408(面詢)、213(現場勘查)、211(準備程序)、212(言詞辯論)、226(配合開庭)等予以開放不再控管。
      If m_CP01 = "P" Or m_CP01 = "PS" Then
      '2017/9/27 END
         'Modify By Sindy 2024/8/6 游經理:面詢，目前的設定為" 限制閱覽"，請予以解除。
         'strCP10NoShow = "'211','212','213','226','408'"
         strCP10NoShow = "'211','212','213','226'"
      End If
   End If
   'Modify By Sindy 2017/6/14 cp01=efc01(+) ==> instr(cp01||',ALL',efc01(+))>0
   'Modified by Morgan 2017/8/15 +副檔名說明先判斷 CPP15
   'Modify By Sindy 2018/5/21 修改sort
   'Modified by Lydia 2018/09/06 +strCon2
   'Modify By Sindy 2021/10/15 cp09 in(" & m_strKey & ") ==> m_strKey
   'Modify By Sindy 2022/11/1 +,cpp15
   'Modify by Sindy 2023/3/3 +,cpp11
   'Modified by Lydia 2023/12/22 + strYYMMSql
   'Modified by Morgan 2025/7/7 檔案修改時間->上傳時間
   'Modified by Morgan 2025/7/21 上傳時間(8)->檔案修改時間(FTP下載已增加更正檔案修改時間), 修改日期時間(10)->上傳時間
   strSql = "Select distinct '" & strUserNum & "',cp09,cp27 as 專業發文日,Decode('" & m_Nation & "','000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質" & _
            ",decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',nvl(e9.efc03,nvl(e2.efc03,e1.efc03)),nvl(e9.efc04,nvl(e2.efc04,e1.efc04))),decode(length(cp10),4,decode(sign(instr(upper(cpp02),'.'||cp10||'.PDF')),1,'官方來函',''),'')) as 副檔名說明,CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,cp163,sqldatet(cpp17)||' '||sqltime(cpp18) as 上傳時間,CP82,cp05 as 收文日,cp127 as 發文室,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,CPP10 as CPP10_Flag註記,CPP03,CP05,CP66,CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(nvl(e9.efc05,nvl(e2.efc05,e1.efc05)),decode(length(cp10),4,decode(sign(instr(upper(cpp02),'.'||cp10||'.PDF')),1,15,999)))) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,cp43,cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18"
   strSqlwhere = _
            " From (select * From CasepaperPDF,caseprogress,Casepropertymap,LetterProgress" & _
            " Where " & m_strKey & strYYMMSql & _
            " And cp01='" & m_CP01 & "' And cp02='" & m_CP02 & "' And cp03='" & m_CP03 & "' And cp04='" & m_CP04 & "'" & _
            " And cp09=cpp01(+) and (cpp10 is null or cpp10='X' or cpp10='Y'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, " or cpp10='D'", "") & ")" & _
            IIf(strCP10NoShow <> "", " And cp10 not in(" & strCP10NoShow & ")", "") & _
            " And cp01=cpm01(+) And cp10=cpm02(+)" & strConSql & strCon2 & _
            " and cp09=LP01(+)),efilecaption e1,efilecaption e2,efilecaption e9" & _
            " Where instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0 and e2.efc01(+)='999' and e2.efc02(+)=cpp15 and instr(','||cp01,','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & _
            m_QueryEfile
   strSql = strSql & strSqlwhere
   '檢查是否有被條件過濾掉,而應查出來的文號
   strExc(0) = "select cp09 from(" & strSql & ") group by cp09 order by cp09 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   strChkLostCP09 = m_strKey
   If intI = 1 Then
      'Modify By Sindy 2021/10/19 屬查詢整筆案號
      If InStr(UCase(strChkLostCP09), UCase("from")) > 0 Then
         'Modify By Sindy 2022/5/23 + and" & m_strKey
         'Modified by Lydia 2023/12/22 + strYYMMSql
         strChkLostCP09 = "cp09 not in(select cp09 " & strSqlwhere & ") and" & m_strKey & strYYMMSql
      Else
      '2021/10/19 END
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            strChkLostCP09 = Replace(strChkLostCP09, RsTemp.Fields("cp09"), "")
            RsTemp.MoveNext
         Loop
      End If
   End If
   If InStr(UCase(strChkLostCP09), UCase("from")) = 0 Then
      strChkLostCP09 = Replace(strChkLostCP09, "'',", "")
      strChkLostCP09 = Replace(strChkLostCP09, ",''", "")
      strChkLostCP09 = Replace(strChkLostCP09, "''", "")
   End If
   'Modify By Sindy 2021/10/15 + And UCase(Trim(strChkLostCP09)) <> UCase("cp09 in()")
   If strChkLostCP09 <> "" And UCase(Trim(strChkLostCP09)) <> UCase("cp09 in()") Then
      'Modified by Lydia 2017/10/11 debug ('' as CPP10 as CPP10_Flag註記) => ('' as CPP10_Flag註記)
      'Modify By Sindy 2018/5/21 修改sort
      'Modify By Sindy 2021/10/19 cp09 in(" & strChkLostCP09 & ") => strChkLostCP09 ( strChkLostCP09: ex: cp09 in('AB0043312','AB0043311','AB0043310') )
      'Modify By Sindy 2023/3/3 +,'' as cpp11
      'Modified by Morgan 2025/7/7 檔案修改時間->上傳時間
      'Modified by Morgan 2025/7/21 上傳時間(8)->檔案修改時間(FTP下載已增加更正檔案修改時間), 修改日期時間(10)->上傳時間
      strSql = strSql & " union " & _
            "Select distinct '" & strUserNum & "',cp09,cp27 as 專業發文日,Decode('" & m_Nation & "','000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質," & _
            "'' as 檔案名稱,'' as 副檔名說明,CP10,'' as 檔案修改時間,cp163,'' as 上傳時間,CP82,cp05 as 收文日,cp127 as 發文室,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,'' as CPP10_Flag註記,'' as CPP03,CP05,CP66,CP67,0 as sort,'','',0,0,0,cp43,'' as cpp04,'' as cpp02,0,'','' as cpp11,'' as cpp16,0 as cpp17,0 as cpp18" & _
            " From caseprogress,Casepropertymap,LetterProgress" & _
            " Where " & strChkLostCP09 & _
            " And cp01='" & m_CP01 & "' And cp02='" & m_CP02 & "' And cp03='" & m_CP03 & "' And cp04='" & m_CP04 & "'" & _
            " And cp01=cpm01(+) And cp10=cpm02(+)" & _
            " and cp09=LP01(+)"
   End If
   'Add By Sindy 2015/1/28 增加讀取暫存的電子檔,如.回覆單
   'Modify By Sindy 2015/2/12 29991231 CP05==>是為了排在第一筆顯示
   If bolReadOnlyRev = False Then
      'Modify By Sindy 2017/6/14 '" & m_CP01 & "'=efc01(+) ==> instr('" & m_CP01 & ",ALL',efc01(+))>0
      'Modify By Sindy 2018/5/21 修改sort
      'Modified by Lydia 2018/09/06 +strCon2
      'Modify By Sindy 2022/11/1 +,cpp15
      'Modify By Sindy 2023/3/3 +,cpp11
      'Modify by Amy 2025/05/19 FC結案單,會有.Msg檔,上線後都以結案單號抓資料-與Sindy 討論
      If strSrvDate(1) >= FCP結案單電子化啟用日 And stCPP11 <> "" Then
         'Modif by Amy 2025/05/22 由共同查詢進入不會傳cpp11
         If stCPP11 <> "" Then
            strSqlCPP11 = " and cpp11='" & stCPP11 & "'"
         End If
      'Modify By Sindy 2025/9/12 mark; 目前回覆單在未收文時,從案號進來的卷宗區應該都不需要顯示了
      '                          1.因接洽單在簽核的過程中有自己的看附件畫面
      '                          2.結案單是從待處理區->檢視回覆單查看的(上列程式段)
'      Else
'         strSqlCPP11 = " and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0"
      '2025/9/12 END
      'End If
         'Modified by Morgan 2025/7/7 檔案修改時間->上傳時間
         'Modified by Morgan 2025/7/21 上傳時間(8)->檔案修改時間(FTP下載已增加更正檔案修改時間), 修改日期時間(10)->上傳時間
         strSql = strSql & " union " & _
            "Select distinct '" & strUserNum & "',cpp01 as 總收文號,0 as 專業發文日,' ' as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',nvl(e9.efc03,e1.efc03),nvl(e9.efc04,e1.efc04)),'') as 副檔名說明,' ' CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,'',sqldatet(cpp17)||' '||sqltime(cpp18) as 上傳時間,0 CP82,0 as 收文日,0 as 發文室,' ' as 寄件方式,CPP10 as CPP10_Flag註記,CPP03,29991231 CP05,0 CP66,0 CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(e9.efc05,nvl(e1.efc05,999))) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,'',cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18" & _
            " From CasepaperPDF,efilecaption e1,efilecaption e9" & _
            " Where cpp01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, "", " and (cpp10 is null or cpp10<>'D')") & strCon2 & _
            " and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0 and instr('," & m_CP01 & "',','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & _
            m_QueryEfile & strSqlCPP11
         'end 2025/05/19
      End If 'Modify By Sindy 2025/9/12 +
      'Added by Morgan 2016/7/22
      '未收文或已刪除的來函附件
      If Pub_StrUserSt03 = "M51" Then
         'Modify By Sindy 2017/6/14 '" & m_CP01 & "'=efc01(+) ==> instr('" & m_CP01 & ",ALL',efc01(+))>0
         'Modify By Sindy 2018/5/21 修改sort
         'Modified by Lydia 2018/09/06 +strCon2
         'Modify By Sindy 2022/11/1 +,cpp15
         'Modify By Sindy 2023/3/3 +,cpp11
         'Modified by Morgan 2025/7/7 檔案修改時間->上傳時間
         'Modified by Morgan 2025/7/21 上傳時間(8)->檔案修改時間(FTP下載已增加更正檔案修改時間), 修改日期時間(10)->上傳時間
         'Modify By Sindy 2025/9/12 取消 and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))=0 回覆單在此也要抓出來
         strSql = strSql & " union " & _
            "Select distinct '" & strUserNum & "',cpp01 as 總收文號,0 as 專業發文日,' ' as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',nvl(e9.efc03,e1.efc03),nvl(e9.efc04,e1.efc04)),'') as 副檔名說明,' ' CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,'',sqldatet(cpp17)||' '||sqltime(cpp18) as 上傳時間,0 CP82,0 as 收文日,0 as 發文室,' ' as 寄件方式,CPP10 as CPP10_Flag註記,CPP03,29991231 CP05,0 CP66,0 CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(e9.efc05,nvl(e1.efc05,999))) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,'',cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18" & _
            " From CasepaperPDF,efilecaption e1,efilecaption e9" & _
            " Where cpp01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, "", " and (cpp10 is null or cpp10<>'D')") & strCon2 & _
            " and cpp10 in ('C','U') and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0 and instr('," & m_CP01 & "',','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & _
            m_QueryEfile
      End If
      'end 2016/7/22
   End If
   '2015/1/28 END
   
   'Added by Lydia 2018/02/01 FCP含已發文之客戶提供文件
   If ChkDelC.Value = 1 Then
        '令CPP10='D'
        'Modified by Lydia 2018/04/12 取消CPP10='D'
         'strSql = strSql & " union " & _
            "Select distinct ' ' as V,cpp01 as 總收文號,sqldatet(csd11) as 專業發文日,'客戶提供文件' as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',efc03,efc04),'') as 副檔名說明,cpp01 cp09,' ' CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,' ',cpp08||cpp09 as 修改日期時間,0 CP82,sqldatet(0) as 收文日,' ' as 發文室,' ' as 寄件方式,'D' as CPP10_Flag註記,CPP03,29991231 CP05,0 CP66,0 CP67,nvl(efc05,999) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,'',cpp04,cpp02" & _
            " From CasepaperPDF,EFileCaption, CustSupportDoc" & _
            " Where csd01='" & m_CP01 & "' and csd02='" & m_CP02 & "' and csd03='" & m_CP03 & "' and csd04='" & m_CP04 & "'" & _
            " and cpp01=csd05 and nvl(csd11,0) > 0  and instr('," & m_CP01 & ",ALL',','||efc01(+))>0 and instr(upper(cpp02),'.'||efc02(+)||'.')>0"
         'Modify By Sindy 2018/5/21 修改sort
         'Modified by Lydia 2018/09/06 +strCon2
         'Modified by Lydia 2020/06/23 比照一般, 已刪除的資料預設不顯示
         'strSql = strSql & " union " & _
            "Select distinct ' ' as V,cpp01 as 總收文號,sqldatet(csd11) as 專業發文日,'客戶提供文件' as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',nvl(e9.efc03,e1.efc03),nvl(e9.efc04,e1.efc04)),'') as 副檔名說明,cpp01 cp09,' ' CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,' ',cpp08||cpp09 as 修改日期時間,0 CP82,sqldatet(0) as 收文日,' ' as 發文室,' ' as 寄件方式,CPP10 as CPP10_Flag註記,CPP03,29991231 CP05,0 CP66,0 CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(e9.efc05,nvl(e1.efc05,999))) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,'D',cpp04,cpp02,cpp07,cpp15" & _
            " From CasepaperPDF,EFileCaption e1,CustSupportDoc,efilecaption e9" & _
            " Where csd01='" & m_CP01 & "' and csd02='" & m_CP02 & "' and csd03='" & m_CP03 & "' and csd04='" & m_CP04 & "'" & _
            " and cpp01=csd05 and nvl(csd11,0) > 0  and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0 and instr('," & m_CP01 & "',','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & strCon2
         'Modify By Sindy 2022/11/1 +,cpp15
         'Modify By Sindy 2023/3/3 +,cpp11
         'Modified by Morgan 2025/7/7 檔案修改時間->上傳時間
         'Modified by Morgan 2025/7/21 上傳時間(8)->檔案修改時間(FTP下載已增加更正檔案修改時間), 修改日期時間(10)->上傳時間
         strSql = strSql & " union " & _
            "Select distinct '" & strUserNum & "',cpp01 as 總收文號,csd11 as 專業發文日,'客戶提供文件' as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',nvl(e9.efc03,e1.efc03),nvl(e9.efc04,e1.efc04)),'') as 副檔名說明,' ' CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,'',sqldatet(cpp17)||' '||sqltime(cpp18) as 上傳時間,0 CP82,0 as 收文日,0 as 發文室,' ' as 寄件方式,CPP10 as CPP10_Flag註記,CPP03,29991231 CP05,0 CP66,0 CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(e9.efc05,nvl(e1.efc05,999))) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,'D',cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18" & _
            " From CasepaperPDF,EFileCaption e1,CustSupportDoc,efilecaption e9" & _
            " Where csd01='" & m_CP01 & "' and csd02='" & m_CP02 & "' and csd03='" & m_CP03 & "' and csd04='" & m_CP04 & "'" & _
            " and cpp01=csd05 and (cpp10 is null or cpp10='X' or cpp10='Y'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, " or cpp10='D'", "") & ")" & _
            " and nvl(csd11,0) > 0  and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0 and instr('," & m_CP01 & "',','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & strCon2 & _
            m_QueryEfile
   End If
   'end 2018/02/01
   'Add By Sindy 2024/3/27 改存入暫存檔,以便後續使用
   strSql = "insert into R100101_L(ID,CP09,CP27,CP10Nm,CPP02Nm,EFCNm,CP10,FileMDT,CP163,ModifyDT,CP82,CP05,CP127,SendType" & _
            ",CPP10,CPP03,CP05Sort,CP66,CP67,SORT,CPP12,CPP05,CPP06,CPP08,CPP09,CP43,CPP04,CPP02" & _
            ",CPP07,CPP15,CPP11,CPP16,CPP17,CPP18) " & strSql
   cnnConnection.Execute strSql, intI
   '增加顯示複合歷程歸卷的附件
   'Modified by Morgan 2025/7/7 檔案修改時間->上傳時間
   'Modified by Morgan 2025/7/21 上傳時間(8)->檔案修改時間(FTP下載已增加更正檔案修改時間), 修改日期時間(10)->上傳時間
   strSql = "Select distinct '" & strUserNum & "',L.cp09 as 總收文號,L.cp27 as 專業發文日,cp10nm as 案件性質" & _
            ",decode(cpp02,null,'',replace(cpp02,M.cp01||M.cp02||'.','" & m_CP01 & m_CP02 & ".')||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱" & _
            ",nvl(Decode('000','000',nvl(e9.efc03,e1.efc03),nvl(e9.efc04,e1.efc04)),'') as 副檔名說明" & _
            ",L.CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')')||' : '||M.cp01||M.cp02 as 檔案修改時間" & _
            ",'',sqldatet(cpp17)||' '||sqltime(cpp18) as 上傳時間,L.CP82,L.cp05,L.cp127,sendtype,CPP10,CPP03,L.CP05,L.CP66" & _
            ",L.CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(e9.efc05,nvl(e1.efc05,999))) as sort" & _
            ",cpp12,cpp05,cpp06,cpp08,cpp09,L.cp43,cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18,M.cp09" & _
            " From CasepaperPDF,efilecaption e1,efilecaption e9,caseprogress M,LetterProgress," & _
            "(select cp09,cp27,cp10nm,cpp02nm,cp10,cp82,cp05,cp127,sendtype,cp05sort,cp66,cp67,cp43,cp163" & _
            " from R100101_L Where id='" & strUserNum & "' and cp09<>cp163 and cp163 is not null) L" & _
            " Where L.cp163=cpp01 and cpp01=M.cp09(+) and cpp01=LP01(+)" & strConSql & strCon2 & _
            " and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0" & _
            " and instr('," & m_CP01 & "',','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & _
            " and cpp12='S' and cpp03>0 and instr(cpp02nm,replace(cpp02,M.cp01||M.cp02||'.','" & m_CP01 & m_CP02 & ".'))=0" & _
            " and substr(upper(cpp02),-4)<>'.DEL'" & _
            m_QueryEfile
   strSql = "insert into R100101_L(ID,CP09,CP27,CP10Nm,CPP02Nm,EFCNm,CP10,FileMDT,CP163,ModifyDT,CP82,CP05,CP127,SendType" & _
            ",CPP10,CPP03,CP05Sort,CP66,CP67,SORT,CPP12,CPP05,CPP06,CPP08,CPP09,CP43,CPP04,CPP02" & _
            ",CPP07,CPP15,CPP11,CPP16,CPP17,CPP18,R01001) " & strSql
   cnnConnection.Execute strSql, intI
   'Add By Sindy 2024/8/30 增加顯示複合歷程寄件備份的郵件
   'Modified by Morgan 2025/7/7 檔案修改時間->上傳時間
   'Modified by Morgan 2025/7/21 上傳時間(8)->檔案修改時間(FTP下載已增加更正檔案修改時間), 修改日期時間(10)->上傳時間
   strSql = "Select distinct '" & strUserNum & "',L.cp09 as 總收文號,L.cp27 as 專業發文日,cp10nm as 案件性質" & _
            ",decode(cpp02,null,'',replace(cpp02,M.cp01||M.cp02||'.','" & m_CP01 & m_CP02 & ".')||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱" & _
            ",nvl(Decode('000','000',nvl(e9.efc03,e1.efc03),nvl(e9.efc04,e1.efc04)),'') as 副檔名說明" & _
            ",L.CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')')||' : '||M.cp01||M.cp02 as 檔案修改時間" & _
            ",'',sqldatet(cpp17)||' '||sqltime(cpp18) as 上傳時間,L.CP82,L.cp05,L.cp127,sendtype,CPP10,CPP03,L.CP05,L.CP66" & _
            ",L.CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(e9.efc05,nvl(e1.efc05,999))) as sort" & _
            ",cpp12,cpp05,cpp06,cpp08,cpp09,L.cp43,cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18,M.cp09" & _
            " From CasepaperPDF,efilecaption e1,efilecaption e9,caseprogress M,LetterProgress," & _
            "(select cp09,cp27,cp10nm,cpp02nm,cp10,cp82,cp05,cp127,sendtype,cp05sort,cp66,cp67,cp43,cp163" & _
            " from R100101_L Where id='" & strUserNum & "' and cp09<>cp163 and cp163 is not null) L,appform" & _
            " Where L.cp163=cpp01 and cpp01=M.cp09(+) and cpp01=LP01(+)" & strConSql & strCon2 & _
            " and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0" & _
            " and instr('," & m_CP01 & "',','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & _
            " and instr(cpp02nm,replace(cpp02,M.cp01||M.cp02||'.','" & m_CP01 & m_CP02 & ".'))=0" & _
            " and substr(upper(cpp02),-4)<>'.DEL' and af01=cpp01 and af11=cpp08 and af12=cpp09" & _
            m_QueryEfile
   strSql = "insert into R100101_L(ID,CP09,CP27,CP10Nm,CPP02Nm,EFCNm,CP10,FileMDT,CP163,ModifyDT,CP82,CP05,CP127,SendType" & _
            ",CPP10,CPP03,CP05Sort,CP66,CP67,SORT,CPP12,CPP05,CPP06,CPP08,CPP09,CP43,CPP04,CPP02" & _
            ",CPP07,CPP15,CPP11,CPP16,CPP17,CPP18,R01001) " & strSql
   cnnConnection.Execute strSql, intI
   '2024/8/30 END
      
   'Add By Sindy 2024/9/27 依系統別調整案件性質的顯示內容
   '專利
   strSql = "update R100101_L R1" & _
            " set CP10Nm=(select decode(cp01||cp10,'P1604',decode(sign(instr(cp64,'專利權消滅')),1,substr(cp64,instr(cp64,'專利權消滅'),15),decode(sign(instr(cp64,'消滅')),1,substr(cp64,instr(cp64,'消滅'),12),CP10Nm)),'CFP1604',decode(sign(instr(cp64,'消滅')),1,substr(cp64,instr(cp64,'消滅'),12),CP10Nm),CP10Nm) from caseprogress where cp09=R1.cp09)" & _
            " where ID='" & strUserNum & "'" & _
            " and cp09 in(select CP1.cp09 from caseprogress CP1 where cp09=CP1.cp09 and CP1.cp01 in('CFP','FCP','P'))"
   cnnConnection.Execute strSql, intI
   '法務
   strSql = "update R100101_L R1" & _
            " set CP10Nm=(select substr(decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||CP64||'('||CP10Nm||')',1,500) from caseprogress,SystemKind where cp09=R1.cp09 AND CP01=SK01(+))" & _
            " where ID='" & strUserNum & "'" & _
            " and cp09 in(select CP1.cp09 from caseprogress CP1 where cp09=CP1.cp09 and CP1.cp01 in('CFL','FCL','L','LIN'))"
   cnnConnection.Execute strSql, intI
   '顧問
   strSql = "update R100101_L R1" & _
            " set CP10Nm=(select substr(decode(sign(instr('3,4',sk02)),1,decode(cp46,19221111,'回執退件日:'||sqldatet(cp47)||';',19220101,'回執未回郵局送達日:'||sqldatet(cp47)||';',''),'')||decode(cp10,'0',sqldatet(cp53)||'--'||sqldatet(cp54),CP64||'('||CP10Nm||')'),1,50) from caseprogress,SystemKind where cp09=R1.cp09 AND CP01=SK01(+))" & _
            " where ID='" & strUserNum & "'" & _
            " and cp09 in(select CP1.cp09 from caseprogress CP1 where cp09=CP1.cp09 and CP1.cp01 in('LA'))"
   cnnConnection.Execute strSql, intI
   '2024/9/27 END
   
   '抓暫存檔的資料
   'Modified by Morgan 2025/7/7 檔案修改時間->上傳時間
   'Modified by Morgan 2025/7/21 上傳時間(8)->檔案修改時間(FTP下載已增加更正檔案修改時間), 修改日期時間(10)->上傳時間
   strSql = "Select distinct ' ' as V,cp09 as 總收文號,sqldatet(cp27) as 專業發文日,CP10Nm as 案件性質" & _
            ",CPP02Nm as 檔案名稱,EFCNm as 副檔名說明,cp09,CP10,FileMDT as 檔案修改時間,' ',ModifyDT as 上傳時間,CP82,sqldatet(cp05) as 收文日,sqldatet(cp127) as 發文室,SendType as 寄件方式,CPP10 as CPP10_Flag註記" & _
            ",CPP03,CP05Sort,CP66,CP67,Sort,cpp12,cpp05,cpp06,cpp08,cpp09,cp43,cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18,R01001" & _
            " from R100101_L where ID='" & strUserNum & "'"
   '2024/3/27 END
   '資料排序
   'Modify By Sindy 2017/7/11 + ,cpp02 desc
   strSql = strSql & " order by CP05Sort DESC, CP66 DESC, CP67 DESC, CP09 DESC,sort asc,cpp02 desc"
            '" order by cp09 asc,cpp08 desc,cpp09 desc"
   'Modified by Lydia 2023/09/11 測試客戶提供文件(frm060120)發生錯誤，改寫法
   'rsTmp.CursorLocation = adUseClient
   'rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'If rsTmp.RecordCount > 0 Then
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
   'end 2023/09/11
      Set GRD1.Recordset = rsTmp
      Call QueryDelData
      'Add By Sindy 2024/4/3
      If m_mouseRow > 0 And GRD1.Rows - 1 >= m_mouseRow Then GRD1.TopRow = m_mouseRow
      '2024/4/3 END
   Else
      If QueryDelData = False Then
         rsTmp.Close
         Set rsTmp = Nothing
         ReadAttachFile = False
         Exit Function
      End If
   End If
   If pub_QL04 <> "" Then InsertQueryLog (GRD1.Rows - 1) 'Add By Sindy 2025/8/7
   rsTmp.Close
   
EXITSUB:
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
End Function

''Modify By Sindy 2024/3/27 舊函數先暫時保留
''查詢附件檔
''Modify By Sindy 2015/5/26
'Public Function ReadAttachFile_Old() As Boolean
''2015/5/26 END
'Dim rsTmp As New ADODB.Recordset
'Dim strConSql As String 'Add By Sindy 2014/5/16
'Dim strCP10NoShow As String 'Add By Sindy 2014/12/11
'Dim strChkLostCP09 As String 'Add By Sindy 2014/12/11
'Dim strCon2 As String 'Added by Lydia 2018/09/06 每一個語法都加上
'Dim strSqlwhere As String 'Add By Sindy 2021/10/19
'
'   ReadAttachFile_Old = True
'
'   'Add By Sindy 2014/5/16 電腦中心及專業部才可以看到承辦單
'   'Modify By Sindy 2016/7/12 + 開放檔案室可以看全部附件(Or m_identity = "W")
'   strConSql = ""
'   'Modify By Sindy 2018/7/18 改鎖智權人員不可以看其附件(承辦單,.altr代理人來函)
''   If Pub_StrUserSt03 = "M51" Or m_identity = "W" Or _
''      (Pub_StrUserSt03 >= "P10" And Pub_StrUserSt03 < "P19") Then
''      '全部附件都可看
''   Else
'   'Modified by Morgan 2019/7/22 +W1開頭(客戶服務組)--秀玲
'   'If Left(Pub_StrUserSt03, 1) = "S" Then
'   'Modify By Sindy 2020/9/7 顧服組另外控制
'   'If Left(Pub_StrUserSt03, 1) = "S" Or Left(Pub_StrUserSt03, 2) = "W1" Then
'   If Left(Pub_StrUserSt03, 1) = "S" Then
'   '2020/9/7 END
'   '2018/7/18 END
'      'strConSql = strConSql & " and (instr(upper(cpp02),upper('" & EMP_承辦單 & "'))=0 or cpp02 is null)"
'      'Modify By Sindy 2014/11/26 加 電子送件時,不要看到.DWG
'      'Modify By Sindy 2020/10/30 + EMP_多案承辦單,不要看到
'      'Modify By Sindy 2024/3/19 卷宗區開放智權部人員可以查看 "非PDF" 的承辦單。 + & ".pdf'))=0
'      strConSql = strConSql & " and (instr(upper(cpp02),upper('." & EMP_承辦單 & ".pdf'))=0 or cpp02 is null)" & _
'                              " and (instr(upper(cpp02),upper('." & EMP_多案承辦單 & ".pdf'))=0 or cpp02 is null)" & _
'                              " and (((cp118 is not null and instr(upper(cpp02),upper('.DWG'))=0) or cpp02 is null) or cp118 is null)"
'      '2024/3/19 END
'
'      'Add By Sindy 2015/1/15 加 CFP的.Altr代理人來函不可看到
'      'If m_CP01 = "CFP" Then
'         strConSql = strConSql & " and (instr(upper(cpp02),upper('.altr'))=0 or cpp02 is null)"
'      'End If
'      'Add By Sindy 2018/11/1 桂英:限制 ALTR 和 INVOICE 智權人員不可隨意開啟
'      strConSql = strConSql & " and (instr(upper(cpp02),upper('.INVOICE'))=0 or cpp02 is null)"
'      'Add By Sindy 2020/2/15 玫音:報價：quote (控管同"altr"，不開放給智權同仁)
'      strConSql = strConSql & " and (instr(upper(cpp02),upper('.QUOTE'))=0 or cpp02 is null)"
'      'Add By Sindy 2021/4/22 雅娟:陸代郵件：PAT (控管同"altr"，不開放給智權同仁)
'      strConSql = strConSql & " and (instr(upper(cpp02),upper('.PAT.'))=0 or cpp02 is null)"
'
'      'Added by Morgan 2019/7/22 電子化客戶函未判發不可看--郭
'      strConSql = strConSql & " and not (instr(upper(cpp02),upper('.cus.'))>0 and lp01 is not null and lp05=0)"
'   End If
'   '2014/5/16 END
'   'Added by Lydia 2015/11/16 查名單電子化
'   'Modifie by Lydia 2016/04/25 +TS案
'   If m_CP01 = "T" Or m_CP01 = "TS" Then
'      '不可直接看結果附件,要經由查覆明細畫面
'      strConSql = strConSql & " and instr(upper(cpp02),upper('." & UCase(TMQ_查名作業 & ".pdf") & "'))=0"
'   End If
'   'end 2015/11/16
'
'   'Add By Sindy 2016/6/22 檢查是否有查看開庭紀要(brief)及電子筆錄(note)的權限
'   ' or cpp02 is null) ==> 因文號有可能都沒有放電子檔,但還是要顯示出來
'   If PUB_GetLimitToBRIEF(m_CP01, m_CP02, m_CP03, m_CP04, m_strKey) = False Then
'      'Modify By Sindy 2017/1/4 ex.LA-003005
''      strConSql = strConSql & " and ((instr(upper(cpp02),upper('.brief.'))=0" & _
''                                   " and instr(upper(cpp02),upper('.note.'))=0" & _
''                                   " and substr(upper(cpp02),1,1)<>'L' and substr(upper(cpp02),1,3)<>'FCL' and substr(upper(cpp02),1,3)<>'CFL') " & _
''                                   " or cpp02 is null)"
'      strConSql = strConSql & " and ((instr(upper(cpp02),upper('.brief.'))=0 and instr(upper(cpp02),upper('.note.'))=0) or cpp02 is null)"
'   End If
'   '2016/6/22 END
'
'   'Added by Lydia 2018/09/06
'   strCon2 = ""
'   If ChkDelMsg.Value = 1 Then '排除郵件檔
'       '因為FCP案有匯入外來和寄出郵件,造成卷宗區檔案眾多;在FCP不印說明書之後,造成核稿人一開始無法很容易找到翻譯用之最終提申本,所以勾選此項排除郵件檔,方便尋找
'       strCon2 = strCon2 & " and Upper(cpp02) not like '%.MSG' "
'   End If
'   'end 2018/09/06
'
'   Screen.MousePointer = vbHourglass
'   KillAttach
'   GRD1.Clear
'   SetGrd
'   GRD1.FixedCols = 0
''   select cpp01,cp01,cp02,cp03,cp04 from casepaperpdf,caseprogress where cpp10='D'
''   and cpp01=cp09(+)
'   'Modify By Sindy 2014/7/24 +CPP03
'   'Modify By Sindy 2014/11/25 +副檔名說明
''   strSql = "Select ' ' as V,cp09 as 總收文號,sqldatet(cp27) as 專業發文日,Decode('" & m_Nation & "','000',CPM03,CPM04) as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,cp09,CP10,sqldatet(cpp08)||' '||sqltime(cpp09) as 檔案修改時間,' ',cpp08||cpp09 as 修改日期時間,CP82,sqldatet(cp05) as 收文日,sqldatet(cp127) as 發文室,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,CPP10,CPP03" & _
''            " From CasepaperPDF,caseprogress,Casepropertymap,LetterProgress" & _
''            " Where cp09 in(" & m_strKey & ")" & _
''            " And cp01='" & m_CP01 & "' And cp02='" & m_CP02 & "' And cp03='" & m_CP03 & "' And cp04='" & m_CP04 & "'" & _
''            " And cp09=cpp01(+) and (cpp10 is null or cpp10='X' or cpp10='Y'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, " or cpp10='D'", "") & ")" & _
''            " And cp01=cpm01(+) And cp10=cpm02(+)" & strConSql & _
''            " and cp09=LP01(+)" & _
''            " order by SQLDatet2(CP05) desc,CP66 desc,CP67 desc,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3') desc,CP09 desc,cpp11 asc,cpp08 desc,cpp09 desc"
'
'   'Modify By Sindy 2014/12/11 過濾某些案件性質不可看到附件
'   '查詢權限僅電腦中心、智權人員(若智權人員離職才開放其區主管)、承辦人、游經理、王副總
'   '211.準備程序 212.言詞辯論 213.現場堪察 226.配合開庭 408.面詢
'   strCP10NoShow = ""
'   'Modify By Sindy 2015/1/8 73.內專程序主管 或 75.內專程序 也可以看
'   'Modify By Sindy 2016/7/5 83.CFP程序主管 或 85.CFP程序 也可以看
'   'Modify By Sindy 2016/7/12 + 開放檔案室可以看全部附件(Or m_identity = "W")
'   'Modify By Sindy 2017/2/9 Or strUserNum = m_CP14 ==> Or InStr(m_CP14, strUserNum) > 0
'   'Modified by Lydia 2017/03/28 chkmailid取得可能不只一人 strUserNum = ChkMailId(m_CP13)=> InStr(ChkMailId(m_CP13), strUserNum) > 0
'   'Modify By Sindy 2018/7/18 外專這些案件性質不需要鎖附件
'   'Modify By Sindy 2020/2/25 游經理:請開啟張偉城( 89026 )可查看配合訴訟案件內容
'   'Modify By Sindy 2023/3/28 游經理:開放法律所人員可查看案件性質408(面詢)、213(現場勘查)、211(準備程序)、212(言詞辯論)、226(配合開庭)等案件性質的附件。
'   '                          以使與律師配合的訴訟案件可順利運作。 ex:P-120983
'   If Not (Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 2) = "F2" Or strUserNum = m_CP13 Or InStr(ChkMailId(m_CP13), strUserNum) > 0 Or m_identity = "W" _
'           Or InStr(m_CP14, strUserNum) > 0 Or PUB_GetST05(strUserNum) = "72" Or PUB_GetST05(strUserNum) = "71" _
'           Or PUB_GetST05(strUserNum) = "73" Or PUB_GetST05(strUserNum) = "75" _
'           Or PUB_GetST05(strUserNum) = "83" Or PUB_GetST05(strUserNum) = "85" _
'           Or InStr("'89026'", strUserNum) > 0 Or Left(Pub_StrUserSt03, 1) = "L") Then
'      'Add By Sindy 2017/9/27 游經理:CFP案的案件性質408(面詢)、213(現場勘查)、211(準備程序)、212(言詞辯論)、226(配合開庭)等予以開放不再控管。
'      If m_CP01 = "P" Or m_CP01 = "PS" Then
'      '2017/9/27 END
'         strCP10NoShow = "'211','212','213','226','408'"
'      End If
'   End If
'   'Modify By Sindy 2017/6/14 cp01=efc01(+) ==> instr(cp01||',ALL',efc01(+))>0
'   'Modified by Morgan 2017/8/15 +副檔名說明先判斷 CPP15
'   'Modify By Sindy 2018/5/21 修改sort
'   'Modified by Lydia 2018/09/06 +strCon2
'   'Modify By Sindy 2021/10/15 cp09 in(" & m_strKey & ") ==> m_strKey
'   'Modify By Sindy 2022/11/1 +,cpp15
'   'Modify by Sindy 2023/3/3 +,cpp11
'   'Modified by Lydia 2023/12/22 + strYYMMSql
'   strSql = "Select distinct ' ' as V,cp09 as 總收文號,sqldatet(cp27) as 專業發文日,Decode('" & m_Nation & "','000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質" & _
'            ",decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',nvl(e9.efc03,nvl(e2.efc03,e1.efc03)),nvl(e9.efc04,nvl(e2.efc04,e1.efc04))),decode(length(cp10),4,decode(sign(instr(upper(cpp02),'.'||cp10||'.PDF')),1,'官方來函',''),'')) as 副檔名說明,cp09,CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,' ',cpp08||cpp09 as 修改日期時間,CP82,sqldatet(cp05) as 收文日,sqldatet(cp127) as 發文室,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,CPP10 as CPP10_Flag註記,CPP03,CP05,CP66,CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(nvl(e9.efc05,nvl(e2.efc05,e1.efc05)),decode(length(cp10),4,decode(sign(instr(upper(cpp02),'.'||cp10||'.PDF')),1,15,999)))) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,cp43,cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18"
'   strSqlwhere = _
'            " From (select * From CasepaperPDF,caseprogress,Casepropertymap,LetterProgress" & _
'            " Where " & m_strKey & strYYMMSql & _
'            " And cp01='" & m_CP01 & "' And cp02='" & m_CP02 & "' And cp03='" & m_CP03 & "' And cp04='" & m_CP04 & "'" & _
'            " And cp09=cpp01(+) and (cpp10 is null or cpp10='X' or cpp10='Y'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, " or cpp10='D'", "") & ")" & _
'            IIf(strCP10NoShow <> "", " And cp10 not in(" & strCP10NoShow & ")", "") & _
'            " And cp01=cpm01(+) And cp10=cpm02(+)" & strConSql & strCon2 & _
'            " and cp09=LP01(+)),efilecaption e1,efilecaption e2,efilecaption e9" & _
'            " Where instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0 and e2.efc01(+)='999' and e2.efc02(+)=cpp15 and instr(','||cp01,','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & _
'            m_QueryEfile
'   strSql = strSql & strSqlwhere
'   '檢查是否有被條件過濾掉,而應查出來的文號
'   strExc(0) = "select cp09 from(" & strSql & ") group by cp09 order by cp09 asc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   strChkLostCP09 = m_strKey
'   If intI = 1 Then
'      'Modify By Sindy 2021/10/19 屬查詢整筆案號
'      If InStr(UCase(strChkLostCP09), UCase("from")) > 0 Then
'         'Modify By Sindy 2022/5/23 + and" & m_strKey
'         'Modified by Lydia 2023/12/22 + strYYMMSql
'         strChkLostCP09 = "cp09 not in(select cp09 " & strSqlwhere & ") and" & m_strKey & strYYMMSql
'      Else
'      '2021/10/19 END
'         RsTemp.MoveFirst
'         Do While Not RsTemp.EOF
'            strChkLostCP09 = Replace(strChkLostCP09, RsTemp.Fields("cp09"), "")
'            RsTemp.MoveNext
'         Loop
'      End If
'   End If
'   If InStr(UCase(strChkLostCP09), UCase("from")) = 0 Then
'      strChkLostCP09 = Replace(strChkLostCP09, "'',", "")
'      strChkLostCP09 = Replace(strChkLostCP09, ",''", "")
'      strChkLostCP09 = Replace(strChkLostCP09, "''", "")
'   End If
'   'Modify By Sindy 2021/10/15 + And UCase(Trim(strChkLostCP09)) <> UCase("cp09 in()")
'   If strChkLostCP09 <> "" And UCase(Trim(strChkLostCP09)) <> UCase("cp09 in()") Then
'      'Modified by Lydia 2017/10/11 debug ('' as CPP10 as CPP10_Flag註記) => ('' as CPP10_Flag註記)
'      'Modify By Sindy 2018/5/21 修改sort
'      'Modify By Sindy 2021/10/19 cp09 in(" & strChkLostCP09 & ") => strChkLostCP09 ( strChkLostCP09: ex: cp09 in('AB0043312','AB0043311','AB0043310') )
'      'Modify By Sindy 2023/3/3 +,'' as cpp11
'      strSql = strSql & " union " & _
'            "Select distinct ' ' as V,cp09 as 總收文號,sqldatet(cp27) as 專業發文日,Decode('" & m_Nation & "','000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質," & _
'            "'' as 檔案名稱,'' as 副檔名說明,cp09,CP10,'' as 檔案修改時間,' ','' as 修改日期時間,CP82,sqldatet(cp05) as 收文日,sqldatet(cp127) as 發文室,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,'' as CPP10_Flag註記,'' as CPP03,CP05,CP66,CP67,0 as sort,'','',0,0,0,cp43,'' as cpp04,'' as cpp02,0,'','' as cpp11,'' as cpp16,0 as cpp17,0 as cpp18" & _
'            " From caseprogress,Casepropertymap,LetterProgress" & _
'            " Where " & strChkLostCP09 & _
'            " And cp01='" & m_CP01 & "' And cp02='" & m_CP02 & "' And cp03='" & m_CP03 & "' And cp04='" & m_CP04 & "'" & _
'            " And cp01=cpm01(+) And cp10=cpm02(+)" & _
'            " and cp09=LP01(+)"
'   End If
'   'Add By Sindy 2015/1/28 增加讀取暫存的電子檔,如.回覆單
'   'Modify By Sindy 2015/2/12 29991231 CP05==>是為了排在第一筆顯示
'   If bolReadOnlyRev = False Then
'      'Modify By Sindy 2017/6/14 '" & m_CP01 & "'=efc01(+) ==> instr('" & m_CP01 & ",ALL',efc01(+))>0
'      'Modify By Sindy 2018/5/21 修改sort
'      'Modified by Lydia 2018/09/06 +strCon2
'      'Modify By Sindy 2022/11/1 +,cpp15
'      'Modify By Sindy 2023/3/3 +,cpp11
'      strSql = strSql & " union " & _
'         "Select distinct ' ' as V,cpp01 as 總收文號,sqldatet(0) as 專業發文日,' ' as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',nvl(e9.efc03,e1.efc03),nvl(e9.efc04,e1.efc04)),'') as 副檔名說明,cpp01 cp09,' ' CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,' ',cpp08||cpp09 as 修改日期時間,0 CP82,sqldatet(0) as 收文日,' ' as 發文室,' ' as 寄件方式,CPP10 as CPP10_Flag註記,CPP03,29991231 CP05,0 CP66,0 CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(e9.efc05,nvl(e1.efc05,999))) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,'',cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18" & _
'         " From CasepaperPDF,efilecaption e1,efilecaption e9" & _
'         " Where cpp01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, "", " and (cpp10 is null or cpp10<>'D')") & strCon2 & _
'         " and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0 and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0 and instr('," & m_CP01 & "',','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & _
'         m_QueryEfile
'      'Added by Morgan 2016/7/22
'      '未收文或已刪除的來函附件
'      If Pub_StrUserSt03 = "M51" Then
'         'Modify By Sindy 2017/6/14 '" & m_CP01 & "'=efc01(+) ==> instr('" & m_CP01 & ",ALL',efc01(+))>0
'         'Modify By Sindy 2018/5/21 修改sort
'         'Modified by Lydia 2018/09/06 +strCon2
'         'Modify By Sindy 2022/11/1 +,cpp15
'         'Modify By Sindy 2023/3/3 +,cpp11
'         strSql = strSql & " union " & _
'            "Select distinct ' ' as V,cpp01 as 總收文號,sqldatet(0) as 專業發文日,' ' as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',nvl(e9.efc03,e1.efc03),nvl(e9.efc04,e1.efc04)),'') as 副檔名說明,cpp01 cp09,' ' CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,' ',cpp08||cpp09 as 修改日期時間,0 CP82,sqldatet(0) as 收文日,' ' as 發文室,' ' as 寄件方式,CPP10 as CPP10_Flag註記,CPP03,29991231 CP05,0 CP66,0 CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(e9.efc05,nvl(e1.efc05,999))) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,'',cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18" & _
'            " From CasepaperPDF,efilecaption e1,efilecaption e9" & _
'            " Where cpp01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, "", " and (cpp10 is null or cpp10<>'D')") & strCon2 & _
'            " and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))=0 and cpp10 in ('C','U') and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0 and instr('," & m_CP01 & "',','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & _
'            m_QueryEfile
'      End If
'      'end 2016/7/22
'   End If
'   '2015/1/28 END
'
'   'Added by Lydia 2018/02/01 FCP含已發文之客戶提供文件
'   If ChkDelC.Value = 1 Then
'        '令CPP10='D'
'        'Modified by Lydia 2018/04/12 取消CPP10='D'
'         'strSql = strSql & " union " & _
'            "Select distinct ' ' as V,cpp01 as 總收文號,sqldatet(csd11) as 專業發文日,'客戶提供文件' as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',efc03,efc04),'') as 副檔名說明,cpp01 cp09,' ' CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,' ',cpp08||cpp09 as 修改日期時間,0 CP82,sqldatet(0) as 收文日,' ' as 發文室,' ' as 寄件方式,'D' as CPP10_Flag註記,CPP03,29991231 CP05,0 CP66,0 CP67,nvl(efc05,999) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,'',cpp04,cpp02" & _
'            " From CasepaperPDF,EFileCaption, CustSupportDoc" & _
'            " Where csd01='" & m_CP01 & "' and csd02='" & m_CP02 & "' and csd03='" & m_CP03 & "' and csd04='" & m_CP04 & "'" & _
'            " and cpp01=csd05 and nvl(csd11,0) > 0  and instr('," & m_CP01 & ",ALL',','||efc01(+))>0 and instr(upper(cpp02),'.'||efc02(+)||'.')>0"
'         'Modify By Sindy 2018/5/21 修改sort
'         'Modified by Lydia 2018/09/06 +strCon2
'         'Modified by Lydia 2020/06/23 比照一般, 已刪除的資料預設不顯示
'         'strSql = strSql & " union " & _
'            "Select distinct ' ' as V,cpp01 as 總收文號,sqldatet(csd11) as 專業發文日,'客戶提供文件' as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',nvl(e9.efc03,e1.efc03),nvl(e9.efc04,e1.efc04)),'') as 副檔名說明,cpp01 cp09,' ' CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,' ',cpp08||cpp09 as 修改日期時間,0 CP82,sqldatet(0) as 收文日,' ' as 發文室,' ' as 寄件方式,CPP10 as CPP10_Flag註記,CPP03,29991231 CP05,0 CP66,0 CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(e9.efc05,nvl(e1.efc05,999))) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,'D',cpp04,cpp02,cpp07,cpp15" & _
'            " From CasepaperPDF,EFileCaption e1,CustSupportDoc,efilecaption e9" & _
'            " Where csd01='" & m_CP01 & "' and csd02='" & m_CP02 & "' and csd03='" & m_CP03 & "' and csd04='" & m_CP04 & "'" & _
'            " and cpp01=csd05 and nvl(csd11,0) > 0  and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0 and instr('," & m_CP01 & "',','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & strCon2
'         'Modify By Sindy 2022/11/1 +,cpp15
'         'Modify By Sindy 2023/3/3 +,cpp11
'         strSql = strSql & " union " & _
'            "Select distinct ' ' as V,cpp01 as 總收文號,sqldatet(csd11) as 專業發文日,'客戶提供文件' as 案件性質,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱,nvl(Decode('" & m_Nation & "','000',nvl(e9.efc03,e1.efc03),nvl(e9.efc04,e1.efc04)),'') as 副檔名說明,cpp01 cp09,' ' CP10,sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 檔案修改時間,' ',cpp08||cpp09 as 修改日期時間,0 CP82,sqldatet(0) as 收文日,' ' as 發文室,' ' as 寄件方式,CPP10 as CPP10_Flag註記,CPP03,29991231 CP05,0 CP66,0 CP67,decode(substr(upper(cpp02),-4),'.MSG',1,nvl(e9.efc05,nvl(e1.efc05,999))) as sort,cpp12,cpp05,cpp06,cpp08,cpp09,'D',cpp04,cpp02,cpp07,cpp15,cpp11,cpp16,cpp17,cpp18" & _
'            " From CasepaperPDF,EFileCaption e1,CustSupportDoc,efilecaption e9" & _
'            " Where csd01='" & m_CP01 & "' and csd02='" & m_CP02 & "' and csd03='" & m_CP03 & "' and csd04='" & m_CP04 & "'" & _
'            " and cpp01=csd05 and (cpp10 is null or cpp10='X' or cpp10='Y'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, " or cpp10='D'", "") & ")" & _
'            " and nvl(csd11,0) > 0  and instr(',ALL',','||e1.efc01(+))>0 and instr(upper(cpp02),'.'||e1.efc02(+)||'.')>0 and instr('," & m_CP01 & "',','||e9.efc01(+))>0 and instr(upper(cpp02),'.'||e9.efc02(+)||'.')>0" & strCon2 & _
'            m_QueryEfile
'   End If
'   'end 2018/02/01
'
'   '資料排序
'   'Modify By Sindy 2017/7/11 + ,cpp02 desc
'   strSql = strSql & " order by CP05 DESC, CP66 DESC, CP67 DESC, CP09 DESC,sort asc,cpp02 desc"
'            '" order by cp09 asc,cpp08 desc,cpp09 desc"
'   'Modified by Lydia 2023/09/11 測試客戶提供文件(frm060120)發生錯誤，改寫法
'   'rsTmp.CursorLocation = adUseClient
'   'rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   'If rsTmp.RecordCount > 0 Then
'   intI = 1
'   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'   'end 2023/09/11
'      Set GRD1.Recordset = rsTmp
'      Call QueryDelData
'   Else
'      If QueryDelData = False Then
'         rsTmp.Close
'         Set rsTmp = Nothing
'         ReadAttachFile_Old = False
'         Exit Function
'      End If
'   End If
'   rsTmp.Close
'
'EXITSUB:
'   Screen.MousePointer = vbDefault
'   Set rsTmp = Nothing
'End Function

'欄位變色
Private Sub recovercolor(intRow As Integer)
Dim j As Integer
   
   GRD1.row = intRow
   'Modified by Lydia 2018/04/12 記錄判斷變色的欄位
   'If Trim(GRD1.TextMatrix(intRow, 15)) = "D" Then
   If Trim(GRD1.TextMatrix(intRow, m_Flag1)) = "D" Or Trim(GRD1.TextMatrix(intRow, m_Flag2)) = "D" Then
      For j = 1 To GRD1.Cols - 1
         GRD1.col = j
         GRD1.CellBackColor = &H8080FF '已刪除顯示紅色
      Next j
   End If
   'Add By Sindy 2022/11/1 專利處程序分案,官方文件區顯示黃色
   If UCase(m_PrevForm.Name) = UCase("frm040101_1") Or _
      UCase(m_PrevForm.Name) = UCase("frm050101_2") Then
      If Trim(GRD1.TextMatrix(intRow, 30)) = "1" Then 'CPP15=1
         For j = 1 To 4 'GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = &HFFFF& '黃色
         Next j
      End If
   End If
   '2022/11/1 END
End Sub

Private Function QueryDelData() As Boolean
Dim strFileName As String, strFileName_jj As String, intLen As Integer 'Add By Sindy 2013/9/24
Dim strCP09 As String
   
   QueryDelData = False
   GRD1.Visible = False
   
   strCP09 = ""
   For ii = 1 To GRD1.Rows - 1
      If GRD1.RowHeight(ii) > 0 Then 'Add By Sindy 2015/1/13 +if
         If strCP09 = Trim(GRD1.TextMatrix(ii, 6)) Then
            GRD1.TextMatrix(ii, 1) = ""
            GRD1.TextMatrix(ii, 2) = ""
            GRD1.TextMatrix(ii, 3) = "" '案件性質
            GRD1.TextMatrix(ii, 12) = "" '收文日
            GRD1.TextMatrix(ii, 13) = "" '發文室
            GRD1.TextMatrix(ii, 14) = "" '寄件方式
         End If
         strCP09 = Trim(GRD1.TextMatrix(ii, 6))
         'Modify By Sindy 2015/11/16 檢查是否有同文號同檔名的狀況,若有,只顯示一筆
         For jj = ii + 1 To GRD1.Rows - 1
            If strCP09 <> GRD1.TextMatrix(jj, 6) Then
               Exit For
            End If
            If GRD1.TextMatrix(ii, 4) = GRD1.TextMatrix(jj, 4) Then
               GRD1.RowHeight(jj) = 0
            End If
            'Add By Sindy 2024/3/29 檔名中有 .cdata. 有重覆檔名者, 且有電子檔歸卷文號者不顯示
            If GetFileName(GRD1.TextMatrix(ii, 4)) = GetFileName(GRD1.TextMatrix(jj, 4)) And InStr(UCase(GRD1.TextMatrix(jj, 4)), UCase(".cdata.")) > 0 And GRD1.TextMatrix(jj, 35) <> "" Then
               GRD1.RowHeight(jj) = 0
            End If
            '2024/3/29 END
         Next jj
         '2015/11/16 END
         '檢查檔案重覆另外命名
         If Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
            strFileName = Trim(GRD1.TextMatrix(ii, 4))
            If InStrRev(strFileName, " (") > 0 Then
               'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
               If UCase(Mid(strFileName, InStrRev(strFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
               '2021/8/6 END
                  strFileName = Left(strFileName, InStrRev(strFileName, " (") - 1)
               End If
'            Else
'               strFileName = Trim(GRD1.TextMatrix(ii, 4))
            End If
            GRD1.TextMatrix(ii, 9) = strFileName
            For jj = 1 To ii - 1
               strFileName_jj = Trim(GRD1.TextMatrix(jj, 4))
               If InStrRev(strFileName_jj, " (") > 0 Then
                  'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
                  If UCase(Mid(strFileName_jj, InStrRev(strFileName_jj, " (") + 1, Len("(X86)"))) <> "(X86)" Then
                  '2021/8/6 END
                     strFileName_jj = Left(strFileName_jj, InStrRev(strFileName_jj, " (") - 1)
                  End If
'               Else
'                  strFileName_jj = Trim(GRD1.TextMatrix(jj, 4))
               End If
               
               If strFileName_jj = strFileName Then
                  If InStrRev(UCase(strFileName), "DWG.") > 0 Then
                     intLen = InStrRev(UCase(strFileName), "DWG.")
                  Else
                     intLen = InStrRev(strFileName, ".")
                  End If
                  'Modified by Morgan 2025/7/21 原欄位(10)[修改日期時間]改為上傳時間，改抓欄位(24)[cpp08]&欄位(25)[cpp09]
                  'GRD1.TextMatrix(ii, 9) = Replace(Left(strFileName, intLen - 1) & "_" & Trim(GRD1.TextMatrix(ii, 10)) & Right(strFileName, Len(strFileName) - (intLen - 1)), "._", ".")
                  GRD1.TextMatrix(ii, 9) = Replace(Left(strFileName, intLen - 1) & "_" & Trim(GRD1.TextMatrix(ii, 24) & GRD1.TextMatrix(ii, 25)) & Right(strFileName, Len(strFileName) - (intLen - 1)), "._", ".")
                  'end 2025/7/21
                  Exit For
               End If
            Next jj
         End If
         
         Call recovercolor(ii) 'Add By Sindy 2014/6/26
      End If
   Next ii
   
   '若有資料游標停在第一筆
'      If rsTmp.RecordCount > 0 Then
'         For ii = 4 To GRD1.Cols - 1
'            GRD1.col = ii
'            GRD1.CellBackColor = &HFFC0C0
'         Next ii
'      End If
   
   GRD1.col = 0
   GRD1.row = 1
   GRD1.FixedCols = 1
   GRD1.Visible = True
End Function

'Add By Sindy 2014/5/13
Private Sub Check1_Click()
   If Check1.Value = 1 Then
      WebBrowser1.Navigate "about:blank"
      Command4.Visible = False
      WebBrowser1.Visible = False
      GRD1.Width = GrdMaxW
      Call SetGrd(False)
   Else
      Command4.Visible = True
      WebBrowser1.Visible = True
      GRD1.Width = GrdMinW
      Call SetGrd(False)
   End If
End Sub

Private Sub ChkDelF_Click()
   Call ReadAttachFile
End Sub

'Add By Sindy 2021/10/21 複製到...
Private Sub cmdCopy_Click()
Dim strSaveFiles As String
Dim strRecvNo As String
   
   m_CP09 = ""
   For ii = 1 To GRD1.Rows - 1
      GRD1.row = ii
      GRD1.col = 0
      If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v") And _
         Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
         
         '只能勾選是總收文號的資料列做新增
         If InStr("A,B,C,D", Left(Trim(GRD1.TextMatrix(ii, 6)), 1)) = 0 Then
            MsgBox "該筆資料並非總收文號，不可複製！"
            Exit Sub
         ElseIf Len(Trim(Trim(GRD1.TextMatrix(ii, 6)))) <> 9 Then '總收文號+D
            MsgBox "該筆資料並非總收文號，不可複製！"
            Exit Sub
         End If
         
         If m_CP09 <> "" And m_CP09 <> Trim(GRD1.TextMatrix(ii, 6)) Then
            MsgBox "可點選多筆附件做複製，但須同一筆總收文號！"
            Exit Sub
         End If
         
         m_CP09 = Trim(GRD1.TextMatrix(ii, 6))
         m_CP10 = Trim(GRD1.TextMatrix(ii, 7))
         m_CP10Nm = Trim(GRD1.TextMatrix(ii, 3))
         strSaveFiles = strSaveFiles & "&" & Trim(GRD1.TextMatrix(ii, 6))
         If InStr(strRecvNo, Trim(GRD1.TextMatrix(ii, 6))) = 0 Then
            strRecvNo = strRecvNo & ",'" & Trim(GRD1.TextMatrix(ii, 6)) & "'"
         End If
         strSaveFiles = strSaveFiles & "  " & GetFileName(Trim(GRD1.TextMatrix(ii, 4)))
      End If
   Next ii
   If strSaveFiles = "" Then
      MsgBox "請至少勾選一筆欲複製的電子檔！"
      Exit Sub
   End If
   strSaveFiles = Mid(strSaveFiles, 2)
   strRecvNo = Mid(strRecvNo, 2)
   
   Call frm100101_L_4.SetParent(Me)
   frm100101_L_4.m_strSaveFiles = strSaveFiles
   frm100101_L_4.strRecvNo = strRecvNo
   frm100101_L_4.m_CP10 = m_CP10
   frm100101_L_4.Show vbModal
End Sub

'Added by Morgan 2018/10/29
Private Sub cmdEmail_Click()
   Dim ii As Integer, strCPP02 As String, bolContinue As Boolean
   
   With GRD1
   For ii = 1 To GRD1.Rows - 1
      If UCase(.TextMatrix(ii, 0)) = "V" Then
         strCPP02 = Trim(GRD1.TextMatrix(ii, 4))
         strExc(0) = GetFileName(strCPP02)
         If Right(LCase(strExc(0)), 8) = ".cus.pdf" Then
            m_CP09 = Trim(GRD1.TextMatrix(ii, 6))
            bolContinue = True
         Else
            MsgBox "請點選【通知函】來寄發 EMail", vbExclamation
         End If
         Exit For
      End If
   Next ii
   End With
   If bolContinue Then
      'Modified by Morgan 2021/12/2 +是否Mail紀錄檢查
      If PUB_ChkEmailBackUp(m_CP09) = True Then
         If PUB_SendECustLetter(m_CP09) = True Then
            ReadAttachFile
         End If
      End If
   End If
End Sub

'Add By Sindy 2015/9/18
Private Sub cmdEmpFlow_Click()
Dim rsTmp As New ADODB.Recordset
   
   For ii = 1 To GRD1.Rows - 1
      'If GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v" Then
      GRD1.row = ii
      GRD1.col = 1
      If GRD1.CellBackColor = &HFFC0C0 Then
         strSql = "Select eep01" & _
                  " From EmpElectronProcess" & _
                  " Where eep01='" & Trim(GRD1.TextMatrix(ii, 6)) & "'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount = 0 Then
            ShowNoData
            rsTmp.Close
            Set rsTmp = Nothing
            Exit Sub
         End If
         rsTmp.Close
         '清除反白
         For jj = 1 To GRD1.Cols - 1
            GRD1.col = jj
            GRD1.CellBackColor = QBColor(15)
         Next jj
         Call recovercolor(ii)
         GRD1.TextMatrix(ii, 0) = ""
         '查詢承辦歷程
         frm100101_F_2.Hide
         frm100101_F_2.m_EEP01 = Trim(GRD1.TextMatrix(ii, 6)) '總收文號
         frm100101_F_2.SetParent Me
         If frm100101_F_2.QueryData = True Then
            frm100101_F_2.Show
         End If
         Me.Hide
      End If
   Next ii
   
   Set rsTmp = Nothing
End Sub

'結束
Private Sub cmdExit_Click()
'   m_PrevForm.Hide
'   If UCase(m_PrevForm.Name) = UCase("frm090202_4") Then
'      m_PrevForm.QueryData
'   End If
'   m_PrevForm.Show
   'Modify By Sindy 2019/1/15
   If Not m_PrevForm Is Nothing Then  'Added by Lydia 2021/02/22
      'Modify By Sindy 2025/6/5 防止操作系統Menu的”共用查詢”變灰無法使用
'      If UCase(m_PrevForm.Name) = UCase("frm100102_2") Or _
'         UCase(m_PrevForm.Name) = UCase("frm100114_2") Then
      If Left(UCase(m_PrevForm.Name), 6) = "FRM100" Then
      '2025/6/5 END
         tmpBol = fnCancelNowFormAndShowParentForm(Me) '下一筆
      'Added by Lydia 2021/02/22
      Else
          Unload Me
      End If
      'end 2021/02/22
   Else
   '2019/1/15 END
      Unload Me '結束
   End If
End Sub

Public Sub SetParent(ByRef fm As Form)
   'Added by Lydia 2020/02/020 記錄-重複呼叫的再前一畫面
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
        Set m_PrevFormOld = m_PrevForm
   End If
   'end 2020/02/20
   
   Set m_PrevForm = fm
End Sub

'Add By Sindy 2018/1/11
Private Sub cmdFlag_Click()
Dim rsTmp As New ADODB.Recordset
Dim bolUpdCPP10 As Boolean
   
   bolUpdCPP10 = False
   For ii = 1 To GRD1.Rows - 1
      GRD1.row = ii
      GRD1.col = 1
      'Modify By Sindy 2021/8/17
      'If grd1.CellBackColor = &HFFC0C0 Then
      If Trim(GRD1.TextMatrix(ii, 0)) = "V" Or Trim(GRD1.TextMatrix(ii, 0)) = "v" Then
      '2021/8/17 END
         If Trim(GRD1.TextMatrix(ii, 15)) <> "Y" Then
            'Modify By Sindy 2021/7/21
            If Trim(GRD1.TextMatrix(ii, 15)) = "D" Then
               If MsgBox("檔案[ " & Trim(GRD1.TextMatrix(ii, 28)) & " ]，已註記刪除，確定要救回來嗎？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Sub
               End If
               If Right(UCase(Trim(GRD1.TextMatrix(ii, 28))), 4) = ".DEL" Then
                  MsgBox "檔案[ " & Trim(GRD1.TextMatrix(ii, 28)) & " ]，檔名最後還是掛.DEL，" & vbCrLf & vbCrLf & "請先更名後再恢復註記[X]。", vbExclamation
                  Exit Sub
               End If
            End If
            '2021/7/21 END
            strSql = "Select cpp01,cpp10" & _
                     " From CasePaperPDF" & _
                     " Where cpp01='" & Trim(GRD1.TextMatrix(ii, 6)) & "'" & _
                     " and cpp02='" & Trim(GRD1.TextMatrix(ii, 28)) & "'" & _
                     " and cpp10='" & IIf(Trim(GRD1.TextMatrix(ii, 15)) = "X", "X", "D") & "'"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount = 1 Then
               'Modify By Sindy 2019/2/11 + and substr(upper(cpp02),-4)<>'.DEL'
               strSql = "update CasePaperPDF set cpp10='" & IIf(Trim(GRD1.TextMatrix(ii, 15)) = "X", "Y", "X") & "'" & _
                        " Where cpp01='" & Trim(GRD1.TextMatrix(ii, 6)) & "'" & _
                        " and cpp02='" & Trim(GRD1.TextMatrix(ii, 28)) & "'" & _
                        " and substr(upper(cpp02),-4)<>'.DEL'" & _
                        " and cpp10='" & IIf(Trim(GRD1.TextMatrix(ii, 15)) = "X", "X", "D") & "'"
               cnnConnection.Execute strSql
               bolUpdCPP10 = True
            End If
            rsTmp.Close
         End If
         
         '清除反白
         For jj = 1 To GRD1.Cols - 1
            GRD1.col = jj
            GRD1.CellBackColor = QBColor(15)
         Next jj
         Call recovercolor(ii)
         GRD1.TextMatrix(ii, 0) = ""
      End If
   Next ii
   
   '有異動資料重新查詢
   If bolUpdCPP10 = True Then
      Call ReadAttachFile
   End If
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2014/11/27
Private Sub cmdHelp_Click()
   Me.Hide
   frm100101_L_1.m_CP01 = m_CP01
   frm100101_L_1.Show
End Sub

''產生整份卷宗批次
'Private Sub cmdMerge_Click()
'   Call PDFMergeFile_Batch(False)
'End Sub

'整份卷宗
'Private Sub cmdOpenAll_Click()
'   Dim hLocalFile As Long
'   Dim stFileName As String
'   Dim strCPP01 As String, strCPP02 As String, strCPP03 As String
'
'   '檢查是否有整份卷宗檔
'   strExc(0) = "select cpp01,cpp02,cpp03 from casepaperpdf where cpp01='000000000' and upper(cpp02)='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & ".PDF'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If RsTemp.RecordCount > 0 Then
'         strCPP01 = RsTemp.Fields("cpp01")
'         strCPP02 = RsTemp.Fields("cpp02")
'         strCPP03 = "" & RsTemp.Fields("cpp03")
'         '檢查是否合併完整
'         strExc(0) = "select count(*) from caseprogress,casepaperpdf where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "' and cp09=cpp01(+) and cpp10='X'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If RsTemp.Fields(0) > 0 Then
'               MsgBox "卷宗尚有資料待合併！"
'            End If
'         End If
'      Else
'         MsgBox "無整份卷宗！"
'         Exit Sub
'      End If
'   Else
'      MsgBox "無整份卷宗！"
'      Exit Sub
'   End If
'
'   KillAttach
'   Screen.MousePointer = vbHourglass
'   '讀取檔案名稱
'   stFileName = strCPP02 & " (" & Round(Val(strCPP03) / 1024, 2) & " KB)"
'   If InStrRev(stFileName, " (") > 0 Then
'      stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
'   End If
'   If InStr(stFileName, "\") = 0 Then
'      If GetAttachFile(strCPP01, stFileName) = False Then
'         MsgBox "無法儲存檔案[ " & stFileName & " ]！"
'      End If
'   End If
'   '開啟檔案
'   ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
'   Screen.MousePointer = vbDefault
'End Sub

Public Function PubShowNextData(Optional intRow As Integer = 0) As Boolean
Dim i As Integer
Dim strNo As String
Dim stCP43 As String, stNP22 As String, strMsg As String 'Add by Amy 2022/06/17
Dim intFCState As String, strST15 As String, strSysKind As String, strNation As String, strF0316 As String 'Add by Amy 2025/05/08
Dim strCCM18 As String 'Add by Amy 2025/06/19
   
   PubShowNextData = False
'   For i = 1 To GRD1.Rows - 1
'      GRD1.row = i
'      GRD1.col = 1
'      If GRD1.CellBackColor = &HFFC0C0 And InStr(UCase(GRD1.TextMatrix(i, 4)), UCase("." & EMP_接洽單 & ".menu")) > 0 And _
'         Val(GRD1.TextMatrix(i, 16)) = 0 Then
      If InStr(UCase(GRD1.TextMatrix(intRow, 4)), UCase("." & EMP_接洽單 & ".menu")) > 0 And _
         Val(GRD1.TextMatrix(intRow, 16)) = 0 Then
         GRD1.row = intRow
         PubShowNextData = True
         strExc(0) = "select * from caseprogress where cp09='" & Trim(GRD1.TextMatrix(intRow, 6)) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strNo = RsTemp.Fields("cp140")
'            '清除反白
'            For jj = 1 To GRD1.Cols - 1
'               GRD1.col = jj
'               GRD1.CellBackColor = QBColor(15)
'            Next jj
'            Call recovercolor(i)
            '查詢接洽記錄單
            'Modify By Sindy 2022/12/23 改用共用函數
            Call PUB_Queryfrm090801(strNo, "" & RsTemp.Fields("cp05"), Me)
'            'Modify By Sindy 2022/9/16
'            If DBDATE(RsTemp.Fields("cp05")) >= 接洽單電子收文啟用日 Then
'               'If UCase(m_PrevForm.Name) = UCase("frm210156") Then
'                  '畫面存在
'                  If PUB_CheckFormExist("frm090801_Q") = True Then
'                     MsgBox "接洽單已開啟中！", vbInformation
'                     Exit Function
'                  End If
'               'End If
'               frm090801_Q.SetParent Me
'               frm090801_Q.m_blnCallPrint = True
'               frm090801_Q.Text5 = strNo
'               Call frm090801_Q.cmdOK_Click(4)
'               'frm090801_Q.ZOrder
'               frm090801_Q.Show vbModal
'            Else
'            '2022/9/16 END
'               frm090801.SetParent Me
'               frm090801.m_blnCallPrint = True 'Add By Sindy 2022/10/19
'               frm090801.Text5 = strNo
'               frm090801.m_blnCallPrint_CRL119 = True '是否列印特殊收據頁
'               Call frm090801.cmdOK_Click(4)
'               frm090801.cmdOK(2).Visible = False
'               frm090801.cmdOK(0).Visible = False
'               frm090801.txtPCnt.Visible = False
'               Me.Hide
'            End If
            '2022/12/23 END
         Else
            MsgBox Trim(GRD1.TextMatrix(intRow, 1)) & "無接洽記錄單！", vbExclamation
            Exit Function
         End If
      'Add By Sindy 2015/1/16
      ElseIf InStr(UCase(GRD1.TextMatrix(intRow, 4)), UCase("." & EMP_結案單 & ".menu")) > 0 And _
         Val(GRD1.TextMatrix(intRow, 16)) = 0 Then
         GRD1.row = intRow
         PubShowNextData = True
         strExc(0) = "select * from caseprogress where cp09='" & Trim(GRD1.TextMatrix(intRow, 6)) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strNo = RsTemp.Fields("cp140")
'            '清除反白
'            For jj = 1 To GRD1.Cols - 1
'               GRD1.col = jj
'               GRD1.CellBackColor = QBColor(15)
'            Next jj
'            Call recovercolor(i)
            
            'Add by Amy 2022/06/17 +T延展結案,close.Menu
            If Left(strNo, 2) = "T-" Then
                'Modify by Amy 2022/06/22 +Trim(GRD1.TextMatrix(intRow, 6)-當筆的總收文號
                stCP43 = GetTI02(Mid(strNo, 3), Trim(GRD1.TextMatrix(intRow, 6)), stNP22)
                If stCP43 = MsgText(601) Then strMsg = strMsg & "總收文號為空"
                If stNP22 = MsgText(601) Then
                    If strMsg <> MsgText(601) Then strMsg = strMsg & vbCrLf
                    strMsg = strMsg & "下一程序號為空"
                End If
                If strMsg <> MsgText(601) Then
                    MsgBox "T延展結案" & strMsg & vbCrLf & "請洽電腦中心！"
                    Exit Function
                End If
            End If
            '查詢結案單
            'Add by Amy 2025/04/10 +FC結案單
            intFCState = 0 '非FC結案單
            strF0316 = Pub_GetField("Flow003", "F0301='" & strNo & "'", "F0316")
            strST15 = PUB_GetStaffST15(strF0316, 1)
            strSysKind = Mid(lblCaseNo, 1, Val(InStr(lblCaseNo, "-")) - 1)
            strNation = GetPrjNation1(lblCaseNo)
            If strSrvDate(1) >= FCP結案單電子化啟用日 Then
               'Modify by Amy 2025/06/19 發現舊資料會頁籤判斷會有問題FCP-065275 由P案轉入,1080730閉卷-年費時,由林士堯結案
               '       ex:FCP-065275 由P案轉入,1080730閉卷-年費時,由林士堯結案 / 外商承辦使用國內結案單操作結案 ex:T-242111(結案單號11203939)
               strCCM18 = Pub_GetField("CloseCaseMain", "CCM01='" & strNo & "'", "CCM18")
               If strCCM18 = "F" Then
                  If strSysKind = "FCP" Or strSysKind = "FG" Or strSysKind = "P" Or strSysKind = "CFP" Then
                     intFCState = 2
                  Else
                     intFCState = 1
                  End If
               End If
               'end 2025/06/19
            End If
            frm210147_1.intFCState = intFCState
            frm210147_1.m_NP07 = GRD1.TextMatrix(i, PUB_MGridGetId("CP10", GRD1))
            'end 2025/04/10
            
            Call frm210147_1.SetParent(Me)
            frm210147_1.Hide
            frm210147_1.cmdModify.Visible = False
            frm210147_1.cmdDel.Visible = False
            frm210147_1.cmdFile.Visible = False '檢視回覆單按鈕隱藏
            frm210147_1.m_bolCallCloseMenu = True 'Add By Sindy 2020/12/25 外部呼叫查詢:卷宗區
            frm210147_1.txtF0301 = strNo
            'Add by Amy 2022/06/17 +T延展結案,close.Menu
            If Left(strNo, 2) = "T-" Then
                 frm210147_1.m_stNP01 = stCP43
                 frm210147_1.m_stNP22 = stNP22
            End If
            frm210147_1.Show
            frm210147_1.QueryData
            Me.Hide
         Else
            MsgBox Trim(GRD1.TextMatrix(intRow, 1)) & "無結案單！", vbExclamation
            Exit Function
         End If
      'Add By Sindy 2015/9/11
      ElseIf InStr(UCase(GRD1.TextMatrix(intRow, 4)), UCase("." & EMP_承辦單 & ".menu")) > 0 And _
             Val(GRD1.TextMatrix(intRow, 16)) = 0 Then
         GRD1.row = intRow
         PubShowNextData = True
'         '清除反白
'         For jj = 1 To GRD1.Cols - 1
'            GRD1.col = jj
'            GRD1.CellBackColor = QBColor(15)
'         Next jj
'         Call recovercolor(i)
         
         '查詢承辦歷程
         Me.Hide
         frm100101_F_2.Hide
         frm100101_F_2.m_EEP01 = Trim(GRD1.TextMatrix(intRow, 6)) '總收文號
         frm100101_F_2.SetParent Me
         If frm100101_F_2.QueryData = True Then
            frm100101_F_2.Show
         End If
         
      'Add By Sindy 2020/9/23
      ElseIf InStr(UCase(GRD1.TextMatrix(intRow, 4)), UCase("." & EMP_多案承辦單 & ".menu")) > 0 And _
             Val(GRD1.TextMatrix(intRow, 16)) = 0 Then
         GRD1.row = intRow
         PubShowNextData = True
'         '清除反白
'         For jj = 1 To GRD1.Cols - 1
'            GRD1.col = jj
'            GRD1.CellBackColor = QBColor(15)
'         Next jj
'         Call recovercolor(i)
         
         strExc(0) = "select cp163 from caseprogress where cp09='" & Trim(GRD1.TextMatrix(intRow, 6)) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '查詢承辦歷程
            Me.Hide
            frm100101_F_2.Hide
            frm100101_F_2.m_EEP01 = "" & RsTemp.Fields("cp163")
            frm100101_F_2.m_CP09q = Trim(GRD1.TextMatrix(intRow, 6))
            frm100101_F_2.SetParent Me
            If frm100101_F_2.QueryData = True Then
               'Add By Sindy 2020/12/9
               frm100101_F_2.Caption = "承辦歷程資料查詢 （本所案號：" & lblCaseNo & " " & GetCaseTypeName(m_CP01, Trim(GRD1.TextMatrix(intRow, 7)), IIf(m_Nation = "000", 0, 1)) & "）"
               '2020/12/9 END
               frm100101_F_2.Show
            End If
         Else
            MsgBox "找不到多案歷程收文號！", vbExclamation
         End If
         
      ElseIf InStr(UCase(GRD1.TextMatrix(intRow, 4)), UCase("." & EMP_Email & ".menu")) > 0 And _
         Val(GRD1.TextMatrix(intRow, 16)) = 0 Then
         GRD1.row = intRow
         PubShowNextData = True
'         '清除反白
'         For jj = 1 To GRD1.Cols - 1
'            GRD1.col = jj
'            GRD1.CellBackColor = QBColor(15)
'         Next jj
'         Call recovercolor(i)
         
         '查詢寄件備份
         frm880019.Hide
         frm880019.Caption = "寄件備份"
         frm880019.cmdExit.Caption = "結束"
         frm880019.cmdSend.Visible = False
         frm880019.cmdAttach.Visible = False
         frm880019.cmdReceiver(0).Visible = False
         frm880019.cmdReceiver(1).Visible = False
         frm880019.cmdReceiver(2).Visible = False 'Add By Sindy 2018/5/14
         frm880019.txtReceiver.Locked = True
         frm880019.txtCopy.Locked = True
         frm880019.txtBCC.Locked = True 'Add By Sindy 2020/10/16
         frm880019.txtSubject.Locked = True
         frm880019.txtAttachment.Locked = True
         frm880019.txtContent.Locked = True
         frm880019.FramePrint.Visible = True 'Add By Sindy 2015/10/16
         frm880019.m_CP09 = Trim(GRD1.TextMatrix(intRow, 6)) '總收文號
         frm880019.m_SMB02 = Trim(GRD1.TextMatrix(intRow, 24)) '寄件日期
         frm880019.m_SMB03 = Trim(GRD1.TextMatrix(intRow, 25)) '寄件時間
         frm880019.SetParent Me
         If frm880019.QueryData = True Then
            frm880019.Show vbModal
         End If
         Unload frm880019
      '2015/9/11 END
      'Added By Lydia 2015/11/16 查名單電子化
      ElseIf InStr(UCase(GRD1.TextMatrix(intRow, 4)), UCase("." & TMQ_查名作業 & ".menu")) > 0 And _
         Val(GRD1.TextMatrix(intRow, 16)) = 0 Then
            strExc(0) = GRD1.TextMatrix(intRow, 4)
            '抓出委查單號
            strNo = Mid(strExc(0), 1, InStr(UCase(strExc(0)), UCase("." & TMQ_查名作業 & ".menu")) - 1)
            strNo = Left(Mid(strNo, InStr(strNo, "H")), 9)
            'Added by Lydia 2024/11/11 查名單(網中)：查名單明細作業
            If Left(strNo, 2) = "H1" And strSrvDate(1) >= 查名單網中系統平行測試 Then
               Set nfrm090128_New = Forms(0).GetForm("frm090128_New")
               If Not nfrm090128_New Is Nothing Then
                  nfrm090128_New.SetParent Me, True, strNo, 0, "Q", 0, "" & Trim(GRD1.TextMatrix(intRow, 6))
                  nfrm090128_New.Show
                  If nfrm090128_New.QueryData = True Then
                     Me.Hide
                  Else
                     Unload nfrm090128_New
                  End If
               Else
                  MsgBox "無法載入查名單(網中)明細，請連絡電腦中心！", vbCritical, "查名作業"
                  Exit Function
               End If
            Else
            'end 2024/11/11
               If Mid(strNo, 1, 1) = "H" Then
                  PubShowNextData = True
                  frm090128.m_NoList = strNo
                  frm090128.ShowCP09 = "" & Trim(GRD1.TextMatrix(intRow, 6)) 'Added by Lydia 2016/04/12 顯示目前案件進度的收文號
                  frm090128.R_type = "Q"
                  frm090128.iStiu = 0
                  frm090128.SetParent Me
                  frm090128.m_NoIdx = 0
                  frm090128.mbolCall = True
                  frm090128.Show
                  If frm090128.QueryData = True Then
                     Me.Hide
                  Else
                     Unload frm090128
                  End If
               End If
            End If 'Added by Lydia 2024/11/11
      'Added by Lydia 2017/12/11 FCP案件命名電子化
      ElseIf InStr(UCase(GRD1.TextMatrix(intRow, 4)), UCase("." & FCP命名記錄)) > 0 And _
         Val(GRD1.TextMatrix(intRow, 16)) = 0 Then
            strExc(0) = Trim(GRD1.TextMatrix(intRow, 6))
            frm090902_2.SetParent Me, m_CP01 & m_CP02 & m_CP03 & m_CP04, strExc(0), strUserNum, "Q"
            frm090902_2.Show
            'Modified by Lydia 2018/04/26 +檢查名稱
            'If frm090902_2.ReadData = True Then
            If frm090902_2.ReadData(True) = True Then
               Me.Hide
            Else
               Unload frm090902_2
            End If
      'Added by Lydia 2018/03/06 FCP客戶提供文件
      ElseIf InStr(UCase(GRD1.TextMatrix(intRow, 4)), UCase("." & FCP提供文件)) > 0 And _
         Val(GRD1.TextMatrix(intRow, 16)) = 0 Then
            'Added by Lydia 2021/02/22 判斷客戶提供文件處理中；FCP-64083程序在做內部收文時，另外叫出卷宗區的客戶提供文件.Menu，之後就無法觸發客戶提供文件處理部份
            If PUB_CheckFormExist("frm060121_1") Then
                MsgBox "請先關閉〔客戶提供文件處理〕畫面！"
                Exit Function
            End If
            'end 2021/02/22
            strExc(0) = Trim(GRD1.TextMatrix(intRow, 6))
            Call frm060121_1.SetParent(Me, m_CP01 & m_CP02 & m_CP03 & m_CP04, strExc(0))
            frm060121_1.Show
            If frm060121_1.ReadData = True Then
               Me.Hide
            Else
               Unload frm060121_1
            End If
      End If
End Function

Private Sub cmdMonitor_Click()
   Dim strAPIPath As String, strFtpIp As String
   Dim hLocalFile As Long
   
   strFtpIp = Pub_GetSpecMan("FTP_VOL_IP")
   strAPIPath = PUB_GetFtpTableDir() & "/" & PUB_GetFtpDir2(m_CP01, m_CP02, m_CP03, m_CP04) & "/API"
   strAPIPath = Replace(strAPIPath, "/", "\")
   strAPIPath = Replace(strAPIPath, "\\", "\\" & strFtpIp & "\")
   ShellExecute hLocalFile, "explore", strAPIPath, vbNullString, vbNullString, 1
End Sub

'Add By Sindy 2015/5/25
Private Sub cmdMove_Click()
Dim strSaveFiles As String
Dim strRecvNo As String
   
   For ii = 1 To GRD1.Rows - 1
      GRD1.row = ii
      GRD1.col = 1
      'If GRD1.CellBackColor = &HFFC0C0 Then
      'Modify By Sindy 2019/2/19 + 寄件備份可以移檔
      If (GRD1.TextMatrix(ii, 0) = "V" Or _
          GRD1.TextMatrix(ii, 0) = "v" Or _
          (GRD1.CellBackColor = &HFFC0C0 And _
           UCase(Right(GetFileName(Trim(GRD1.TextMatrix(ii, 4))), Len(".Email.menu"))) = UCase(".Email.menu") And _
           GRD1.TextMatrix(ii, 23) <> "" And _
           GRD1.TextMatrix(ii, 29) <> "" And _
           GRD1.TextMatrix(ii, 6) <> "") _
         ) And _
         Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
         
'         If Trim(GRD1.TextMatrix(ii, 4)) = "" Then
'            MsgBox "請勾選有電子檔的資料！"
'            Exit Sub
'         End If

         'Add By Sindy 2024/4/8
         If Trim(GRD1.TextMatrix(ii, 35)) <> "" Then
            MsgBox "此電子檔非此文號所有，無法移檔！", vbExclamation
            Exit Sub
         End If
         '2024/4/8 END
         
         'Add By Sindy 2019/2/19 歷程的寄件備份似乎沒有移檔的需要,暫時不開放
         If UCase(Right(GetFileName(Trim(GRD1.TextMatrix(ii, 4))), Len(".Email.menu"))) = UCase(".Email.menu") Then
            '檢查寄件備份
            strExc(0) = "select smb01,smb11 from smailbackup where smb01='" & GRD1.TextMatrix(ii, 6) & "' and smb02=" & GRD1.TextMatrix(ii, 23) & " and smb03=" & GRD1.TextMatrix(ii, 29)
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If RsTemp.RecordCount > 0 Then
               If RsTemp.RecordCount = 1 Then
                  If Val("" & RsTemp.Fields("smb11")) > 0 Then
                     MsgBox GetFileName(Trim(GRD1.TextMatrix(ii, 4))) & vbCrLf & vbCrLf & "歷程的寄件備份似乎沒有移檔的需要,暫時不開放，請查明！", vbExclamation
                     Exit Sub
                  End If
               Else
                  MsgBox GetFileName(Trim(GRD1.TextMatrix(ii, 4))) & vbCrLf & vbCrLf & "寄件備份不只一筆資料，請查明！", vbExclamation
                  Exit Sub
               End If
            Else
               MsgBox GetFileName(Trim(GRD1.TextMatrix(ii, 4))) & vbCrLf & vbCrLf & "寄件備份查無資料，請查明！", vbExclamation
               Exit Sub
            End If
         End If
         
         strSaveFiles = strSaveFiles & "&" & Trim(GRD1.TextMatrix(ii, 6))
         If InStr(strRecvNo, Trim(GRD1.TextMatrix(ii, 6))) = 0 Then
            strRecvNo = strRecvNo & ",'" & Trim(GRD1.TextMatrix(ii, 6)) & "'"
         End If
         strSaveFiles = strSaveFiles & "  " & GetFileName(Trim(GRD1.TextMatrix(ii, 4)))
      End If
   Next ii
   If strSaveFiles = "" Then
      MsgBox "請勾選一筆欲移動的電子檔！"
      Exit Sub
   End If
   strSaveFiles = Mid(strSaveFiles, 2)
   strRecvNo = Mid(strRecvNo, 2)
   
   Call frm100101_L_3.SetParent(Me)
   frm100101_L_3.m_strSaveFiles = strSaveFiles
   frm100101_L_3.strRecvNo = strRecvNo
   frm100101_L_3.m_Nation = m_Nation
   If frm100101_L_3.QueryData(0) = True Then
      frm100101_L_3.Show vbModal
   End If
End Sub

'Added by Morgan 2018/12/17
Private Sub cmdok_Click(Index As Integer)
   '發後補看
   'Modify By Sindy 2019/3/25 + 結案單
   If UCase(m_PrevForm.Name) = UCase("frm040117") Or _
      UCase(m_PrevForm.Name) = UCase("frm210148_1") Then
      m_PrevForm.cmdAction = Index
      Unload Me
   End If
End Sub

'Add By Sindy 2014/11/27 更名
Private Sub cmdReName_Click()
Dim strReName As String
Dim intChkCnt As Integer
Dim strNewFile As String
Dim strCPP02 As String, strChkCPP02 As String
Dim i As Integer
   
On Error GoTo ErrHnd
   
   intChkCnt = 0
   For ii = 1 To GRD1.Rows - 1
      GRD1.row = ii
      GRD1.col = 4
      'If GRD1.CellBackColor = &HFFC0C0 Then
      If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v" Or (UCase(Right(Trim(GRD1.TextMatrix(ii, 4)), 12)) = ".MENU (0 KB)" And GRD1.CellBackColor = &HFFC0C0)) And _
         Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
         
         'Add By Sindy 2024/4/8
         If Trim(GRD1.TextMatrix(ii, 35)) <> "" Then
            MsgBox "此電子檔非此文號所有，不可更名！", vbExclamation
            Exit Sub
         End If
         '2024/4/8 END
         
         m_CP09 = Trim(GRD1.TextMatrix(ii, 6))
         m_CP10 = Trim(GRD1.TextMatrix(ii, 7))
         strCPP02 = Trim(GRD1.TextMatrix(ii, 4))
         If m_CP09 <> "" And strCPP02 <> "" Then
            intChkCnt = intChkCnt + 1
         End If
      End If
   Next ii
   strCPP02 = GetFileName(strCPP02)
   strChkCPP02 = strCPP02
   If intChkCnt = 0 Or strCPP02 = "" Then
      MsgBox "請勾選一筆欲更名的電子檔！"
      Exit Sub
   ElseIf intChkCnt > 1 Then
      MsgBox "只可勾選一筆資料做更名！"
      Exit Sub
   End If
   
   'Added by Morgan 2020/7/28
   If Right(UCase(strCPP02), 14) = ".ENCRYPTED.ZIP" Then
      MsgBox "加密壓縮檔不可更名！", vbExclamation
      Exit Sub
   End If
   'end 2020/7/28
   
ShowInput:
   strNewFile = InputBox("確定是否「更名」？" & vbCrLf & vbCrLf & _
   "檔案名稱不可以包含下列任意字元:" & vbCrLf & _
   "\ / : * ? "" < > |", "更名！", strCPP02)
   If UCase(strNewFile) = UCase(strChkCPP02) Then
      MsgBox "請輸入欲更改的電子檔名！"
      strCPP02 = strNewFile
      GoTo ShowInput
   End If
   
   If Trim(strNewFile) = "" Then
      Exit Sub
   Else
      'Add By Sindy 2025/5/8
      If LCase(Right(strNewFile, Len(strNewFile) - InStrRev(strNewFile, "."))) = "" Then
         MsgBox "檔名輸入錯誤，不可沒有副檔名類型！"
         strCPP02 = strCPP02
         GoTo ShowInput
      End If
      '2025/5/8 END
      
      'Add By Sindy 2022/6/13
      If UCase(Right(strCPP02, 5)) = ".MENU" And UCase(Right(strNewFile, 5)) <> ".MENU" Then
         MsgBox "檔名輸入錯誤，副檔名必須為.MENU！"
         strCPP02 = strCPP02
         GoTo ShowInput
      End If
      '2022/6/13 END
      
      If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, strNewFile, strReName, True, 1) = False Then
         strCPP02 = strNewFile
         GoTo ShowInput
      End If
      
      If UCase(strReName) = UCase(strChkCPP02) Then
         MsgBox "請輸入欲更改的電子檔名！"
         strCPP02 = strReName
         GoTo ShowInput
      End If
      'Add By Sindy 2018/6/12
      If Right(UCase(strReName), 4) = UCase(".del") Then
         MsgBox "新的電子檔名最後面不能是(.del)！"
         strCPP02 = strReName
         GoTo ShowInput
      End If
      '2018/6/12 END
      
      'Add By Sindy 2021/10/4
      '考慮卷宗區仍以英文為主，宜續行管制，故請更名功能加入限制不可輸入中文，
      '而僅開放電腦中心人員無限制，以供特殊狀況可輸入中文。
      If Pub_StrUserSt03 <> "M51" Then
         For i = 1 To Len(strReName)
            If Asc(Mid(strReName, i, 1)) <= 0 Then
               MsgBox "檔案命名不符規定，不可有中文字！"
               strCPP02 = strReName
               GoTo ShowInput
            End If
         Next i
      End If
      '2021/10/4 END
      
      'Add By Sindy 2020/12/22
      If InStr(strReName, "\") > 0 Or _
         InStr(strReName, "/") > 0 Or _
         InStr(strReName, ":") > 0 Or _
         InStr(strReName, "*") > 0 Or _
         InStr(strReName, "?") > 0 Or _
         InStr(strReName, """") > 0 Or _
         InStr(strReName, "<") > 0 Or _
         InStr(strReName, ">") > 0 Or _
         InStr(strReName, "|") > 0 Then
         MsgBox "檔案名稱不可以包含下列任意字元:" & vbCrLf & _
                "\ / : * ? "" < > |"
         strCPP02 = strReName
         GoTo ShowInput
      End If
      '2018/12/22 END
      
      'Add By Sindy 2016/7/27
      For ii = 1 To GRD1.Rows - 1
         GRD1.row = ii
         GRD1.col = 1
         If GRD1.TextMatrix(ii, 1) <> "" Then GRD1.Tag = GRD1.TextMatrix(ii, 1)
         If GRD1.Tag = m_CP09 Then
            If UCase(strReName) = UCase(GetFileName(Trim(GRD1.TextMatrix(ii, 4)))) Then
               MsgBox "檔名重覆，請輸入欲更改的電子檔名！"
               strCPP02 = strReName
               GoTo ShowInput
            End If
         End If
      Next ii
      '2016/7/27 END
      
      'Add By Sindy 2025/1/2 電子檔名為.del時,若拿掉.del視同為恢復非刪除狀況
      strSql = ""
      If Right(UCase(strChkCPP02), 4) = UCase(".del") Then
         strExc(0) = "select * from casepaperpdf where cpp01='" & m_CP09 & "' and cpp02='" & strChkCPP02 & "' and cpp10='D'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = ",cpp10='X'"
         End If
      End If
      strSql = "update casepaperpdf set cpp02='" & strReName & "'" & IIf(strSql = "", "", strSql) & " where cpp01='" & m_CP09 & "' and cpp02='" & strChkCPP02 & "'"
      '2025/1/2 END
      Pub_SaveLog strUserNum, strSql 'Add By Sindy 2016/7/28
      cnnConnection.Execute strSql
      'Added by Lydia 2019/03/06 FCP之公告公報1228增加判斷是否有公告本
      If m_CP01 = "FCP" And m_CP10 = "1228" And InStr(UCase(strChkCPP02 & "||" & strNewFile), ".GAZ.PDF") > 0 Then
           Call UpdateCP121(m_CP09, m_CP10, "GAZ")
      End If
      'end 2018/03/06
      
      Call ReadAttachFile
   End If
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2021/8/18
'抽換電子檔
Private Sub cmdSwap_Click()
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer, jj As Integer
   Dim fs, f, s
   Dim stReName As String
   Dim bolSwap As Boolean
   Dim intChkCnt As Integer
   Dim strFile As String
   Dim strCaseNoName As String
   Dim strCP82 As String
   Dim strCP09 As String, strCPP02 As String
   Dim bolComp As Boolean, varTmp As Variant, strChkCPP02 As String
   
On Error GoTo ErrHnd
   
   Combo2.Clear
   intChkCnt = 0
   For ii = 1 To GRD1.Rows - 1
      GRD1.row = ii
      GRD1.col = 1
      If GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v" Then
      'If grd1.CellBackColor = &HFFC0C0 Then
         strCP09 = Trim(GRD1.TextMatrix(ii, 6))
         m_CP10 = Trim(GRD1.TextMatrix(ii, 7))
         m_CP10Nm = Trim(GRD1.TextMatrix(ii, 3))
         strCP82 = Trim(GRD1.TextMatrix(ii, 11))
         strCPP02 = Trim(GRD1.TextMatrix(ii, 28))
         
         Combo2.AddItem strCP09 & " " & strCPP02
   
'         '只能勾選是總收文號的資料列做抽換
'         If InStr("A,B,C,D", Left(Trim(strCP09), 1)) = 0 Then
'            MsgBox strCPP02 & "該筆資料並非已收文，不可以做抽換！"
'            Exit Sub
'         ElseIf Len(Trim(strCP09)) <> 9 Then
'            MsgBox strCPP02 & "該筆資料並非已收文，不可以做抽換！"
'            Exit Sub
'         End If
         
         intChkCnt = intChkCnt + 1
      End If
   Next ii
   If intChkCnt = 0 Then
      MsgBox "請至少選取一筆欲抽換電子檔的資料列！"
      Exit Sub
   End If
   
   'Add By Sindy 110/8/19
   For ii = 0 To Combo2.ListCount - 1
      varTmp = Split(Combo2.List(ii), " ")
      strChkCPP02 = varTmp(1)
      For jj = ii + 1 To Combo2.ListCount - 1
         If InStr(UCase(Combo2.List(jj)), UCase(strChkCPP02)) > 0 Then
            MsgBox "複選欲抽換電子檔的檔名不能重覆！"
            Exit Sub
         End If
      Next jj
   Next ii
   '110/8/19 END
   
   bolSwap = False
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
            
            If intChkCnt <> UBound(sFile) Then
               MsgBox "選取的電子檔數量，與要抽換電子檔的數量不符！"
               Exit Sub
            End If
            
            '記錄路徑
            SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", sFile(0)
            
            For ii = 1 To UBound(sFile)
               bolComp = False
               For jj = 0 To Combo2.ListCount - 1
                  If InStr(UCase(Combo2.List(jj)), UCase(CStr(sFile(ii)))) > 0 Then
                     bolComp = True
                     Exit For
                  End If
               Next jj
               If bolComp = False Then
                  MsgBox "選取的電子檔名( " & CStr(sFile(ii)) & " )，與要抽換電子檔的檔名不符！"
                  Exit Sub
               End If
            Next ii
            
            For ii = 1 To UBound(sFile)
               For jj = 0 To Combo2.ListCount - 1
                  If InStr(UCase(Combo2.List(jj)), UCase(CStr(sFile(ii)))) > 0 Then
                     varTmp = Split(Combo2.List(jj), " ")
                     strCP09 = varTmp(0)
                     strCPP02 = varTmp(1)
                     Combo2.RemoveItem jj
                     Exit For
                  End If
               Next jj
               
               '直接從資料庫刪除檔案
               If DeleteFile(strCP09, strCPP02) = True Then
                  If InStr(sFile(ii), "\") > 0 Then
                     stFileName = sFile(ii)
                  Else
                     stFileName = sFile(0) & "\" & sFile(ii)
                  End If
                  Set fs = CreateObject("Scripting.FileSystemObject")
                  Set f = fs.GetFile(stFileName)
                  '檔案大小為 0 KB 有誤
                  If f.Size = 0 Then
                     ShowMsg sFile(ii) & MsgText(9221)
                     GoTo EXITSUB
                  End If
                  'If AddListX(stFileName & " (" & Round(f.Size / 1024, 2) & " KB)") = True Then
                     '存檔
                     If SaveAttFile_PDF(strCP09, stFileName, CStr(sFile(ii)), Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = True Then
                        bolSwap = True
                     Else
                        GoTo EXITSUB
                     End If
                     Call PUB_DelPCOrgFile(stFileName) '一併將PC上的實體檔案刪除
                     Pub_SaveLog strUserNum, "抽換卷宗區附件：" & sFile(ii), m_CP01, m_CP02, m_CP03, m_CP04, strCP09
                  'End If
               End If
            Next ii
            
         Else
            If intChkCnt <> 1 Then
               MsgBox "選取的電子檔數量，與要抽換電子檔的數量不符！"
               Exit Sub
            End If
            
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            
            If UCase(strFile) <> UCase(strCPP02) Then
               MsgBox "選取的電子檔名，與要抽換電子檔的檔名不符！"
               Exit Sub
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

            '直接從資料庫刪除檔案
            If DeleteFile(strCP09, strCPP02) = True Then
               stFileName = .FileName
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               '檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg strFile & MsgText(9221)
                  GoTo EXITSUB
               End If
               'If AddListX(stFileName & " (" & Round(f.Size / 1024, 2) & " KB)") = True Then
                  '存檔
                  If SaveAttFile_PDF(strCP09, stFileName, strFile, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = True Then
                     bolSwap = True
                  Else
                     GoTo EXITSUB
                  End If
                  Call PUB_DelPCOrgFile(stFileName) '一併將PC上的實體檔案刪除
                  Pub_SaveLog strUserNum, "抽換卷宗區附件：" & strFile, m_CP01, m_CP02, m_CP03, m_CP04, strCP09
               'End If
            End If
         End If
EXITSUB:
         If bolSwap = True Then
            Call ReadAttachFile
         End If
      End If
   End With
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

Private Sub Form_Activate()
   If Screen.ActiveForm.Name <> Me.Name Then Exit Sub 'Added by Morgan 2023/6/2
   If Me.WindowState = 0 Then Me.WindowState = 2 '最大化
   
   'Added by Lydia 2020/02/20
   If bolActive = True Then '只啟動一次
      Exit Sub
   Else
        If UCase(TypeName(m_PrevFormOld)) <> "NOTHING" Then '關閉-重複呼叫的再前一畫面; 避免主MENU的共同查詢無法Enabled
             If Left(UCase(m_PrevFormOld.Name), 6) = "FRM100" Then
                 fnCloseAllFrm100 '會關閉所有frm100含本表單，所以前一畫面要再重新按一次按鈕進入
             Else
                 Unload m_PrevFormOld
             End If
             bolActive = True
        End If
   End If
   'end 2020/02/10
   
   'Added by Morgan 2021/3/26
   If InStr(m_CP01, "T") > 0 Then
      cmdMonitor.Visible = True
   End If
End Sub

'Add By Sindy 2022/11/1 轉出
Private Sub cmdUpload_Click()
Dim pSavePath As String
Dim strCP09 As String, strCPP02 As String
Dim bolHadData As Boolean

   '測試抓桌面的相同資料夾以免誤刪真實檔案
   If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then
      m_strFolder = PUB_Getdesktop & "\" & Mid(m_strFolder, InStrRev(m_strFolder, "\") + 1)
      If Dir(m_strFolder, vbDirectory) = "" Then MkDir m_strFolder
   End If
   '下載路徑
   pSavePath = App.path & "\" & strUserNum
   If Dir(pSavePath, vbDirectory) = "" Then MkDir pSavePath
   
   bolHadData = False
   'P,PS,CFP,CPS 1.官方文件(電子收文),要搬移電子檔到File Server
   For ii = 1 To GRD1.Rows - 1
      'Modify By Sindy 2023/4/7 不控管只有1.官方文件才能轉出,因人員有可能放到內部文件區 And GRD1.TextMatrix(ii, 30) = "1"
      If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v") _
         And Trim(GRD1.TextMatrix(ii, 4)) <> "" And Val(GRD1.TextMatrix(ii, 16)) > 0 Then
         bolHadData = True
         
         strCP09 = Trim(GRD1.TextMatrix(ii, 6))
         strCPP02 = Trim(GRD1.TextMatrix(ii, 28))
         strExc(10) = strCPP02
         '下載電子檔
         If PUB_GetAttachFile_CPP(strCP09, strExc(10), pSavePath) = False Then
            MsgBox strCP09 & " : " & strCPP02 & " 電子檔下載失敗！", vbCritical
            Exit Sub
         End If
         
         If pSavePath & "\" & strCPP02 <> "" Then
            '上傳File Server
            FileCopy pSavePath & "\" & strCPP02, m_strFolder & "\" & Replace(strCPP02, "." & Trim(Trim(GRD1.TextMatrix(ii, 7))) & ".", ".")
            strSql = "update casepaperpdf set cpp02=cpp02||'.del',cpp10='D' where cpp01='" & strCP09 & "' and cpp02='" & strCPP02 & "'"
            cnnConnection.Execute strSql, intI
         End If
      End If
   Next ii
   If bolHadData = False Then
      MsgBox "請點選官方文件方可轉出！", vbCritical
      Exit Sub
   Else
      Me.m_strKey = Me.m_CP09
      Call QueryData
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   WebBrowser1.Navigate "about:blank"
   
   ReDim m_FilesRemoved(0)
   'm_AttachPath = App.Path & "\SeminarAttach"
   'm_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath")
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
   
   Call ChangSelect
   
   'Add By Sindy 2014/6/25
   Check2.Visible = False
   If Pub_StrUserSt03 = "M51" Then
      Frame2.Visible = True
      ChkDelF.Visible = True
      cmdFlag.Visible = True: cmdSwap.Visible = True
      Check2.Visible = True
      Check2.Value = 1
   Else
      Frame2.Visible = False
      ChkDelF.Visible = False
      cmdFlag.Visible = False: cmdSwap.Visible = False
      'Added by Lydia 2018/02/01 調整Grid高度
      'Modified by Lydia 2018/09/06
      'Grd1.Height = 4305
      'Grd1.Top = 1080
      'end 2018/02/01
      'Removed by Morgan 2018/11/1
      'GRD1.Height = 4065
      'GRD1.Top = 1320
      'end 2018/11/1
      'end 2018/09/06
   End If
   '2014/6/25 END
   
   If Pub_StrUserSt03 = "M51" And InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") Then
      Me.Height = 6600
      'cmdMerge.Visible = True
      'cmdOpenAll.Visible = True 'Add By Sindy 2013/10/23 取消
      'cmdPrintAtt.Visible = True
   Else
      Me.Height = 6120
      'cmdMerge.Visible = False
      'cmdOpenAll.Visible = False 'Add By Sindy 2013/10/23 取消
      'cmdPrintAtt.Visible = False
   End If
   
   Me.WindowState = 2 '最大化
   Check1.Value = 1 '關閉預覽
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   DestroyToolTip '清除物件
   Set frm100101_L = Nothing
   
'   'Add by Sindy 2022/12/17 若接洽單已開需關閉
'   If PUB_CheckFormExist("frm090801_Q") = True Then
'      Unload frm090801_Q
'   End If
'   '2022/12/17 END
   
   'Added by Lydia 2020/02/20
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then  'Added by Lydia 2020/02/20
      m_PrevForm.Show
      'Modify By Sindy 2016/6/28 + UCase(m_PrevForm.Name) = UCase("frm072002")
      'Modified by Morgan 2018/12/17 +frm040117
      'Modify By Sindy 2019/3/25 + frm210148_1.結案單
      If UCase(m_PrevForm.Name) = UCase("frm100123") Or _
         UCase(m_PrevForm.Name) = UCase("frm040118") Or _
         UCase(m_PrevForm.Name) = UCase("frm072002") Or _
         UCase(m_PrevForm.Name) = UCase("frm040117") Or _
         UCase(m_PrevForm.Name) = UCase("frm210148_1") Then
         m_PrevForm.PubShowNextData
      'Add By Sindy 2021/2/22
      ElseIf UCase(TypeName(m_PrevForm)) = UCase("frm090201_10") Then
         strSql = "Update CaseProgress Set CP143=" & strSrvDate(1) & " Where CP09='" & Me.Tag & "' "
         cnnConnection.Execute strSql
         Call m_PrevForm.cmdok_Click(1)
      '2021/2/22 END
      End If
      
      Set m_PrevForm = Nothing
   End If 'Added by Lydia 2020/02/20
   
   Set nfrm090128_New = Nothing 'Added by Lydia 2024/11/11
   
End Sub

Private Sub KillAttach()
On Error Resume Next
   If Dir(m_AttachPath & "\.") <> "" Then
      Kill m_AttachPath & "\*.pdf"
   End If
End Sub

'Add By Sindy 2015/1/28
Public Sub FrmCallOpenFile(intRow As Integer, bolReadLastRow As Boolean)
Dim strFileName As String
Dim strFileType As String

GRD1.Visible = False
If intRow > 0 Then
   m_mouseRow = intRow
   GRD1.row = intRow
   GRD1.col = 1
'   If GRD1.CellBackColor = &HFFC0C0 Then
'      '清除反白
'      GRD1.TextMatrix(intRow, 0) = ""
'      For jj = 1 To GRD1.Cols - 1
'         GRD1.col = jj
'         GRD1.CellBackColor = QBColor(15)
'      Next jj
'      Call recovercolor(intRow)
'   Else
      '資料列反白
      If Val(GRD1.TextMatrix(intRow, 16)) > 0 Then
         strFileName = GetFileName(GRD1.TextMatrix(intRow, 4))
         strFileType = Right(strFileName, Len(strFileName) - InStrRev(strFileName, ".") + 1)
         If UCase(strFileType) = UCase(".PDF") Then
            GRD1.TextMatrix(intRow, 0) = "V"
         Else
            GRD1.TextMatrix(intRow, 0) = "v"
         End If
      End If
      For jj = 1 To GRD1.Cols - 1
         GRD1.col = jj
         GRD1.CellBackColor = &HFFC0C0
      Next jj
'   End If
   'If PubShowNextData(GRD1.row) = False Then
   If bolReadLastRow = True Then
      'm_bolDblClick = True
      cmdOpenAtt(1).Tag = "call"
      cmdOpenAtt_Click 1
   End If
End If
GRD1.Visible = True
End Sub

Private Sub GRD1_DblClick()
   If GRD1.row > 0 Then
      Screen.MousePointer = vbHourglass 'Add By Sindy 2019/3/8
      If PubShowNextData(GRD1.row) = False Then
         m_bolDblClick = True
         cmdOpenAtt_Click 1
      End If
      Screen.MousePointer = vbDefault 'Add By Sindy 2019/3/8
   End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
Dim strFileName As String
Dim strFileType As String

'Modify By Sindy 2016/1/30
'getGrdColRow GRD1, x, y, nCol, nRow
'If nCol < 0 Then nCol = 0
'GRD1.col = nCol
'GRD1.row = nRow
GRD1.row = GRD1.MouseRow
GRD1.col = GRD1.MouseCol
nRow = GRD1.row
nCol = GRD1.col
'2016/1/30 END

'Modify By Sindy 2014/7/25
'GRD1.Visible = False
If nRow > 0 And Trim(GRD1.TextMatrix(nRow, 6)) <> "" Then
   m_mouseRow = nRow 'GRD1.MouseRow '記錄目前Row
   
   '先將上筆有反白的資料列復恢
   'Modify By Sindy 2016/1/30
   'If m_mouseRowOld > 0 Then
   If m_mouseRowOld > 0 And m_mouseRowOld <= (GRD1.Rows - 1) Then
   '2016/1/30 END
      GRD1.row = m_mouseRowOld
      GRD1.col = 1
      For jj = 1 To GRD1.Cols - 1
         GRD1.col = jj
         GRD1.CellBackColor = QBColor(15)
      Next jj
      'Add By Sindy 2021/3/4 按 Shift 一樣可以多筆選取
      If Shift = 0 Then
      '2021/3/4 END
         If GRD1.MouseCol <> 0 Then
            GRD1.TextMatrix(GRD1.row, 0) = ""
         End If
      End If
      Call recovercolor(CInt(m_mouseRowOld))
   End If
   
   '將點選的資料列反白
   GRD1.row = m_mouseRow
   GRD1.col = 1
   strFileName = GetFileName(GRD1.TextMatrix(GRD1.row, 4))
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
      If Val(GRD1.TextMatrix(GRD1.row, 16)) > 0 Then
         If UCase(strFileType) = UCase(".PDF") Then
            GRD1.TextMatrix(GRD1.row, 0) = "V"
         Else
            GRD1.TextMatrix(GRD1.row, 0) = "v"
         End If
      End If
   End If
   
'   If GRD1.CellBackColor = &HFFC0C0 Then
'      If GRD1.TextMatrix(GRD1.MouseRow, 0) = "" And GRD1.MouseCol = 0 Then
'         If Val(GRD1.TextMatrix(GRD1.MouseRow, 16)) > 0 Then
'            If UCase(strFileType) = UCase(".PDF") Then
'               GRD1.TextMatrix(GRD1.MouseRow, 0) = "V"
'            Else
'               GRD1.TextMatrix(GRD1.MouseRow, 0) = "v"
'            End If
'         End If
'      Else
'         '清除反白
'         GRD1.TextMatrix(GRD1.MouseRow, 0) = ""
'         'GRD1.row = GRD1.MouseRow
'         For jj = 1 To GRD1.Cols - 1
'            GRD1.col = jj
'            GRD1.CellBackColor = QBColor(15)
'         Next jj
'         Call recovercolor(GRD1.MouseRow) 'Add By Sindy 2014/6/26
'      End If
'   Else
'      '資料列反白
'      If GRD1.MouseCol = 0 Then
'         If Val(GRD1.TextMatrix(GRD1.MouseRow, 16)) > 0 Then
'            If UCase(strFileType) = UCase(".PDF") Then
'               GRD1.TextMatrix(GRD1.MouseRow, 0) = "V"
'            Else
'               GRD1.TextMatrix(GRD1.MouseRow, 0) = "v"
'            End If
'         End If
'      End If
'      'GRD1.row = GRD1.MouseRow
'      For jj = 1 To GRD1.Cols - 1
'         GRD1.col = jj
'         GRD1.CellBackColor = &HFFC0C0
'      Next jj
'      m_mouseRowOld = m_mouseRow
'   End If
End If
'GRD1.Visible = True
End Sub

'Modify By Sindy 2021/4/23
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   'grd1.ToolTipText = ""
   If GRD1.MouseRow <> 0 Then
      If iRow <> GRD1.MouseRow Or iCol <> GRD1.MouseCol Then
         'Modify By Sindy 2017/4/6 因FC郵件之故,檔案名稱欄則顯示主旨內容
         'Add By Sindy 2021/3/4
         If GRD1.MouseCol = 0 Then
            'grd1.ToolTipText = "按 Shift + 資料列 一樣可以多筆選取"
            CreateToolTip GetHWndForToolTip(GRD1), "按 Shift + 資料列 一樣可以多筆選取"
         '2021/3/4 END
         ElseIf GRD1.MouseCol = 4 And GRD1.TextMatrix(GRD1.MouseRow, 27) <> "" Then
            CreateToolTip GetHWndForToolTip(GRD1), GRD1.TextMatrix(GRD1.MouseRow, 27)
         ElseIf GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
         '2017/4/6 END
            'grd1.ToolTipText = grd1.TextMatrix(grd1.MouseRow, grd1.MouseCol)
            CreateToolTip GetHWndForToolTip(GRD1), GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
         End If
         iRow = GRD1.MouseRow
         iCol = GRD1.MouseCol
      End If
   End If
End Sub

'Private Sub grd1_SelChange()
'GRD1.Visible = False
'If GRD1.MouseRow <> 0 And Trim(GRD1.TextMatrix(GRD1.MouseRow, 6)) <> "" Then
'   If GRD1.TextMatrix(GRD1.MouseRow, 0) = "V" Then
'      '清除反白
'      GRD1.TextMatrix(GRD1.MouseRow, 0) = ""
'      GRD1.row = GRD1.MouseRow
'      For jj = 1 To GRD1.Cols - 1
'         GRD1.col = jj
'         GRD1.CellBackColor = QBColor(15)
'      Next jj
'   Else
'      '資料列反白
'      GRD1.TextMatrix(GRD1.MouseRow, 0) = "V"
'      GRD1.row = GRD1.MouseRow
'      For jj = 1 To GRD1.Cols - 1
'         GRD1.col = jj
'         GRD1.CellBackColor = &HFFC0C0
'      Next jj
'   End If
'End If
'GRD1.Visible = True
'End Sub

Private Function DeleteFile(strCP09 As String, strFileName As String) As Boolean
   
On Error GoTo ErrHand
   
   DeleteFile = True
   Screen.MousePointer = vbHourglass
   If DelAttFile_PDF(lblCaseNo.Caption, strCP09, strFileName) = False Then GoTo ErrHand
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   DeleteFile = False
   Screen.MousePointer = vbDefault
End Function

Private Function BrowseForFolder(Optional sCaption As String = "請選擇欲儲存的位置", Optional sDefault As String) As String
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

''列印
'Private Sub cmdPrintAtt_Click()
'   Dim stFileName As String
'   Dim bolIsSelect As Boolean
'
'   bolIsSelect = False
'   Screen.MousePointer = vbHourglass
'
'   For ii = 1 To GRD1.Rows - 1
'      If GRD1.TextMatrix(ii, 0) = "V" Then
'         '清除反白
'         GRD1.col = 0
'         GRD1.row = ii
'         GRD1.TextMatrix(ii, 0) = ""
'         For jj = 1 To GRD1.Cols - 1
'            GRD1.col = jj
'            GRD1.CellBackColor = QBColor(15)
'         Next jj
'         If Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
'            '讀取檔案名稱
'            bolIsSelect = True
'            stFileName = Trim(GRD1.TextMatrix(ii, 4))
'            If InStrRev(stFileName, " (") > 0 Then
'               stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
'            End If
'            If InStr(stFileName, "\") = 0 Then
'               m_CP09 = Trim(GRD1.TextMatrix(ii, 6))
'               If GetAttachFile(m_CP09, stFileName, m_AttachPath & "\" & Trim(GRD1.TextMatrix(ii, 9))) = False Then
'                  MsgBox "無法儲存檔案[ " & stFileName & " ]！"
'               End If
'            End If
'            '列印
'            ShellExecute Me.hwnd, "print", stFileName, vbNullString, vbNullString, 1
'         End If
'      End If
'   Next ii
'   If bolIsSelect = False Then
'      MsgBox "無檔案可列印！"
'   End If
'
'   Screen.MousePointer = vbDefault
'End Sub

Private Function GetAttachFile(ByVal strCP09 As String, ByRef pFileName As String, Optional pSavePath As String) As Boolean
   Dim stAttPath As String
   
'Removed by Morgan 2015/3/24
'   Dim lngSize As Long
'   Dim iFileNo As Integer
'   Dim bytes() As Byte

On Error GoTo ErrHnd

   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
'      '檔案已存在時不必重新下載
'      If Dir(stAttPath) <> "" Then
'         'Kill stAttPath
'         pFileName = stAttPath
'         GetAttachFile = True
'         Exit Function
'      End If
   Else
      'Add By Sindy 2013/12/27 改傳完整的檔案路徑:路徑+檔名
      If InStr(pSavePath, m_AttachPath) > 0 Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
      End If
      '2013/12/27 END
      stAttPath = pSavePath
   End If
   
'Modified by Morgan 2015/3/24 讀取檔案改呼叫共用函數(要改為FTP方式)
'   strExc(0) = "select * from casepaperpdf where cpp01='" & strCP09 & "' and cpp02='" & ChgSQL(pFileName) & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If Dir(stAttPath) <> "" Then Kill stAttPath
'      With RsTemp
'         lngSize = Val(.Fields("cpp03").Value)
'         ReDim bytes(lngSize)
'         If lngSize > 0 Then
'            bytes() = .Fields("cpp04").GetChunk(lngSize)
'         End If
'      End With
'      iFileNo = FreeFile
'      Open stAttPath For Binary Access Write As #iFileNo
'      If lngSize > 0 Then Put #iFileNo, , bytes()
'      Close #iFileNo
'
'      pFileName = stAttPath
'      GetAttachFile = True
'   End If
   GetAttachFile = PUB_GetAttachFile_CPP(strCP09, pFileName, stAttPath, True)
'end 2015/3/24

   Exit Function
   
ErrHnd:
   If Err.Number = 70 Then
      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
   
'Removed by Morgan 2015/3/24
'   If iFileNo > 0 Then Close #iFileNo
End Function

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
Dim m_UpdCPP0102 As String, bolUpdCPP0102 As Boolean 'Add By Sindy 2017/10/31
Dim strZipSrc As String, strZipFile As String, strSrcFile As String 'Added by Morgan 2020/7/28
Dim varTemp As Variant 'Add By Sindy 2021/2/26
   
   'Add By Sindy 2014/12/5
   For ii = 1 To GRD1.Rows - 1
      GRD1.row = ii
      GRD1.col = 1
      If GRD1.CellBackColor = &HFFC0C0 And InStr(UCase(GRD1.TextMatrix(ii, 4)), UCase("." & EMP_接洽單 & ".menu")) > 0 And _
         Val(GRD1.TextMatrix(ii, 16)) = 0 Then 'CPP03=0
         If PubShowNextData(ii) = True Then Exit For
      End If
   Next ii
   '2014/12/5 END
   
   strMergeFN = "" '組欲合併的檔案
'   If Index = 1 Then
'      If Check1.Value = 1 Then
'         Check1.Value = 0
'      End If
'      'WebBrowser1.Navigate "about:blank"
'   End If
   'Add By Sindy 2024/3/29 預設先關閉 WebBrowser1 物件
   '因重覆操作同一個電子檔時,會出現電子檔已開啟訊息
   Check1.Value = 1
   Call Check1_Click
   '2024/3/29 END
   '切換至來源目錄
   If m_AttachPath <> "." Then ChDir m_AttachPath
   
   KillAttach
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   If m_bolDblClick = True Then
      strFileName = GetFileName(GRD1.TextMatrix(m_mouseRow, 4))
      'Added by Morgan 2020/7/28
      If Right(UCase(strFileName), 14) = ".ENCRYPTED.ZIP" Then
         strFileName = Left(strFileName, Len(strFileName) - 14)
      End If
      'end 2020/7/28
      strFileType = Right(strFileName, Len(strFileName) - InStrRev(strFileName, ".") + 1)
      If strFileName <> "" Then
         If UCase(strFileType) = UCase(".PDF") Then
            GRD1.TextMatrix(m_mouseRow, 0) = "V"
         Else
            GRD1.TextMatrix(m_mouseRow, 0) = "v"
         End If
      End If
   End If
   Screen.MousePointer = vbHourglass
   Check1.Tag = 0 'False Add By Sindy 2017/7/24
   m_UpdCPP0102 = "": bolUpdCPP0102 = False 'Add By Sindy 2017/10/31
   Me.cmdAddAtt.Tag = "" 'Add By Sindy 2021/2/26
   For ii = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v" Then
         If m_bolDblClick = False Or (m_bolDblClick = True And ii = m_mouseRow) Then
            If Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
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
                     'Modify By Sindy 2023/1/4 不清除反白
'                     GRD1.col = 0
'                     GRD1.row = ii
'                     GRD1.TextMatrix(ii, 0) = ""
'                     For jj = 1 To GRD1.Cols - 1
'                        GRD1.col = jj
'                        GRD1.CellBackColor = QBColor(15)
'                     Next jj
                  End If
                  '讀取檔案名稱
                  'Modify By Sindy 2024/3/29
'                  stFileName = Trim(GRD1.TextMatrix(ii, 4))
'                  If InStrRev(stFileName, " (") > 0 Then
'                     'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
'                     If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
'                     '2021/8/6 END
'                        stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
'                     End If
'                  End If
'                  strSaveFileName = stFileName
'                  If m_bolDblClick = True Then
'                     strSaveFileName = Left(stFileName, InStrRev(stFileName, ".") - 1) & ServerTime & ".pdf"
'                  End If
                  strSaveFileName = Trim(GRD1.TextMatrix(ii, 9))
                  '2024/3/29 END
                  strMergeFN = strMergeFN & IIf(strMergeFN <> "", " ", "") & ".\" & strSaveFileName
                  stFileName = Trim(GRD1.TextMatrix(ii, 28)) 'Add By Sindy 2024/3/29 取得CPP02檔名
                  If InStr(stFileName, "\") = 0 Then
                     'Add By Sindy 2024/3/29
                     If Trim(GRD1.TextMatrix(ii, 35)) <> "" Then '電子檔歸卷文號
                        m_CP09 = Trim(GRD1.TextMatrix(ii, 35))
                     Else
                     '2024/3/29 END
                        m_CP09 = Trim(GRD1.TextMatrix(ii, 6))
                     End If
                     'Added by Morgan 2020/7/28 HP加密壓縮檔要先解密
                     strZipSrc = ""
                     If Right(UCase(stFileName), 14) = ".ENCRYPTED.ZIP" Then
                        strZipSrc = stFileName
                        stFileName = Left(stFileName, Len(stFileName) - 14)
                        strSrcFile = App.path & "\" & stFileName
                        strZipFile = App.path & "\$ENCRYPTED.ZIP"
                        
                        If GetAttachFile(m_CP09, strZipSrc, strZipFile) = False Then
                           GoTo ErrHnd
                        End If
                        
                        If PUB_UnZipFile2(strZipFile, App.path, cHPZipPwd) = False Then
                           GoTo ErrHnd
                        End If
                        
                        'Modified by Morgan 2025/3/28
                        'stFileName = m_AttachPath & "\" & strSaveFileName
                        stFileName = m_AttachPath & "\" & stFileName
                        'end 2025/3/28
                        
                        'Added by Morgan 2025/3/28
                        
                        '壓縮檔內的檔名可能會少0(因曾經有改過規則)
                        If Left(m_CP02, 1) = "0" Then
                           If Dir(strSrcFile) = "" Then
                              strSrcFile = Replace(strSrcFile, m_CP02, Mid(m_CP02, 2))
                           End If
                        End If
                        'end 2025/3/28
                  
                        FileCopy strSrcFile, stFileName
                        Kill strZipFile
                        Kill strSrcFile
                     Else
                     'end 2020/7/28
                     
                        If GetAttachFile(m_CP09, stFileName, m_AttachPath & "\" & strSaveFileName) = False Then
                           GoTo ErrHnd 'Add By Sindy 2020/6/30
                           'MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                        Else
                           'Add By Sindy 2017/10/31
                           '記錄總收文號
                           If Index = 1 And Trim(GRD1.TextMatrix(ii, 15)) = "X" Then
                              cmdAddAtt.Tag = cmdAddAtt.Tag & "," & ii 'Add By Sindy 2021/2/26
                              m_UpdCPP0102 = m_UpdCPP0102 & ",'" & m_CP09 & GRD1.TextMatrix(ii, 28) & "'"
                           End If
                           '2017/10/31 END
                        End If
                        
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
   If m_UpdCPP0102 <> "" Then m_UpdCPP0102 = Mid(m_UpdCPP0102, 2) 'Add By Sindy 2017/10/31
   If cmdAddAtt.Tag <> "" Then cmdAddAtt.Tag = Mid(cmdAddAtt.Tag, 2) 'Add By Sindy 2021/2/26
   
   If bolIsSelect = False Then
      MsgBox "無檔案可開啟！"
   Else
      If Index = 1 And strMergeFN <> "" Then
         'Modified by Morgan 2018/6/26 有時會發生錯誤,改單檔不合併
         'If m_bolDblClick = True Then
         'Modify By Sindy 2020/6/30 單檔也用合併方式開啟;為測試合併是否成功
         If m_bolDblClick = True Or InStr(strMergeFN, " ") = 0 _
            And Not (Check2.Visible = True And Check2.Value = 1) Then
            WebBrowser1.Navigate stFileName
         Else
            '合併
            'Modify By Sindy 2019/11/6 淑華要求用本所案號命名
            'strMergeName = "merge" & ServerTime & ".pdf"
            strMergeName = m_CP01 & m_CP02 & IIf(m_CP03 & m_CP04 = "000", "", m_CP03 & m_CP04) & ".pdf"
            '2019/11/6 END
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
            'Add By Sindy 2017/10/31
            '上合併成功註記
            If m_UpdCPP0102 <> "" And Dir(m_AttachPath & "\" & strMergeName) <> "" Then
               'Modify By Sindy 2019/2/11 + and substr(upper(cpp02),-4)<>'.DEL'
               strSql = "update casepaperpdf set cpp10='Y' where cpp01||cpp02 in(" & m_UpdCPP0102 & ") and cpp10='X' and substr(upper(cpp02),-4)<>'.DEL'"
               cnnConnection.Execute strSql
               bolUpdCPP0102 = True
            End If
            '2017/10/31 END
         End If
      End If
      'Add By Sindy 2015/5/27 開啟非PDF的電子檔
      For ii = 1 To GRD1.Rows - 1
         If GRD1.TextMatrix(ii, 0) = "v" Then
            '清除反白
            If UCase(cmdOpenAtt(1).Tag) <> UCase("call") Then
               'Modify By Sindy 2023/1/4 不清除反白
'               GRD1.col = 0
'               GRD1.row = ii
'               GRD1.TextMatrix(ii, 0) = ""
'               For jj = 1 To GRD1.Cols - 1
'                  GRD1.col = jj
'                  GRD1.CellBackColor = QBColor(15)
'               Next jj
            End If
            If m_bolDblClick = False Or (m_bolDblClick = True And ii = m_mouseRow) Then
               If Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
                  '讀取檔案名稱
                  'Modify By Sindy 2024/3/29
'                  stFileName = Trim(GRD1.TextMatrix(ii, 4))
'                  If InStrRev(stFileName, " (") > 0 Then
'                     'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
'                     If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
'                     '2021/8/6 END
'                        stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
'                     End If
'                  End If
                  strSaveFileName = Trim(GRD1.TextMatrix(ii, 9))
                  stFileName = Trim(GRD1.TextMatrix(ii, 28)) '取得CPP02檔名
                  '2024/3/29 END
                  If InStr(stFileName, "\") = 0 Then
                     'Add By Sindy 2024/3/29
                     If Trim(GRD1.TextMatrix(ii, 35)) <> "" Then '電子檔歸卷文號
                        m_CP09 = Trim(GRD1.TextMatrix(ii, 35))
                     Else
                     '2024/3/29 END
                        m_CP09 = Trim(GRD1.TextMatrix(ii, 6))
                     End If
                     
                     'Added by Morgan 2020/7/28 HP加密壓縮檔要先解密
                     strZipSrc = ""
                     If Right(UCase(stFileName), 14) = ".ENCRYPTED.ZIP" Then
                        strZipSrc = stFileName
                        stFileName = Left(stFileName, Len(stFileName) - 14)
                        strSrcFile = App.path & "\" & stFileName
                        strZipFile = App.path & "\$ENCRYPTED.ZIP"
                        
                        If GetAttachFile(m_CP09, strZipSrc, strZipFile) = False Then
                           GoTo ErrHnd
                        End If
                        
                        If PUB_UnZipFile2(strZipFile, m_AttachPath, cHPZipPwd) = False Then
                           GoTo ErrHnd
                        End If
                        stFileName = m_AttachPath & "\" & stFileName
                        Kill strZipFile
                     Else
                     'end 2020/7/28
                     
                        'If GetAttachFile(m_CP09, stFileName, m_AttachPath & "\" & stFileName) = False Then
                        If GetAttachFile(m_CP09, stFileName, m_AttachPath & "\" & strSaveFileName) = False Then
                           GoTo ErrHnd 'Add By Sindy 2020/6/30
                           'MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                        End If
                     End If 'Added by Morgan 2020/7/28
                  End If
                  
                  Call PUB_ChkFileTypeOpenExE(stFileName) 'Add By Sindy 2017/9/13
                  '開啟檔案
                  ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
               End If
            End If
         End If
      Next ii
   End If
   'm_mouseRowOld = 0 'Add By Sindy 2015/6/10
   'Add By Sindy 2017/10/31
   '有異動資料重新查詢
   If bolUpdCPP0102 = True Then
      'Modify By Sindy 2021/2/26
      'Call ReadAttachFile
      varTemp = Split(cmdAddAtt.Tag, ",")
      For ii = 0 To UBound(varTemp)
         If Val(varTemp(ii)) > 0 Then
            GRD1.col = 15
            GRD1.row = varTemp(ii)
            GRD1.TextMatrix(varTemp(ii), 15) = "Y"
         End If
      Next ii
      cmdAddAtt.Tag = ""
      '2021/2/26 END
   End If
   '2017/10/31 END
   
ErrHnd:
   m_bolDblClick = False
   Screen.MousePointer = vbDefault
   ChDir App.path '目錄切回
End Sub

Private Sub ChangSelect()
   If cmdSelect.Caption = "全選" Then
      cmdSelect.Caption = "取消全選"
   ElseIf cmdSelect.Caption = "取消全選" Then
      cmdSelect.Caption = "全選"
   End If
End Sub

Private Sub cmdSelect_Click()
Dim i As Integer, k As Integer
Dim strFileName As String
Dim strFileType As String
   
   '先全部清除
   GRD1.Visible = False
   For k = 1 To GRD1.Rows - 1
      GRD1.col = 1
      GRD1.row = k
      GRD1.TextMatrix(k, 0) = ""
      For i = 1 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   Next k
   GRD1.Visible = True
   If cmdSelect.Caption = "全選" Then
      GRD1.Visible = False
      For k = 1 To GRD1.Rows - 1
         GRD1.col = 0
         GRD1.row = k
         'Add By Sindy 2023/11/20
         If GRD1.RowHeight(k) > 0 Then
         '2023/11/20 END
            If Trim(GRD1.Text) = "" And Val(GRD1.TextMatrix(k, 16)) > 0 Then
               strFileName = GetFileName(GRD1.TextMatrix(k, 4))
               strFileType = Right(strFileName, Len(strFileName) - InStrRev(strFileName, ".") + 1)
               If UCase(strFileType) = UCase(".PDF") Then
                  GRD1.Text = "V"
               Else
                  GRD1.Text = "v"
               End If
            End If
         End If
      Next k
      GRD1.Visible = True
   End If
   m_mouseRowOld = 0 'Add By Sindy 2015/6/10
   Call ChangSelect
End Sub

'下載
Private Sub cmdSaveAtt_Click()
   Dim stFileName As String, stFolderPath As String, stFullName As String
   Dim bMultiFile As Boolean
   Dim ii As Integer
   'Added by Lydia 2019/12/03
   Dim strCP10 As String
   Dim strName2 As String
   Dim strZipSrc As String, strZipFile As String, strSrcFile As String 'Added by Morgan 2020/7/28
   
   'Add By Sindy 2020/1/14
   '讀取前次設定路徑
   stFolderPath = GetSetting("TAIE", "P", UCase(Me.Name) & "Dir", "")
   If stFolderPath <> "" Then
      If PUB_ChkDir(stFolderPath) = False Then
         stFolderPath = PUB_Getdesktop
      End If
   Else
      stFolderPath = PUB_Getdesktop
   End If
   stFolderPath = PUB_GetFolder(Me.hWnd, stFolderPath, "請選取資料夾:")
   If Trim(stFolderPath) <> "" Then 'they did not hit cancel
      SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir", stFolderPath
   Else
      Exit Sub
   End If
   If Right(Trim(stFolderPath), 1) <> "\" Then
      stFolderPath = Trim(stFolderPath) & "\"
   End If
   '2020/1/14 END
   
   stFileName = ""
   bMultiFile = False
   For ii = 1 To GRD1.Rows - 1
      If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v") And Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
         If stFileName <> "" Then
            bMultiFile = True
            Exit For
         Else
            'Modify By Sindy 2024/4/8
            'stFileName = Trim(grd1.TextMatrix(ii, 4))
            stFileName = Trim(GRD1.TextMatrix(ii, 28))
            '2024/4/8 END
            'Add By Sindy 2024/4/8
            If Trim(GRD1.TextMatrix(ii, 35)) <> "" Then '電子檔歸卷文號
               m_CP09 = Trim(GRD1.TextMatrix(ii, 35))
            Else
            '2024/4/8 END
               m_CP09 = Trim(GRD1.TextMatrix(ii, 6))
            End If
            strCP10 = Trim(GRD1.TextMatrix(ii, 7)) 'Added by Lydia 2019/12/03
         End If
      End If
   Next ii
   
   Screen.MousePointer = vbHourglass
   If stFileName = "" Then
      MsgBox "無檔案可存檔！"
   Else
      '多選
      If bMultiFile Then
         'stFolderPath = BrowseForFolder() 'Modify By Sindy 2020/1/14 Mark
         If stFolderPath <> "" Then
            For ii = 1 To GRD1.Rows - 1
               If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v") And Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
                  'Modify By Sindy 2024/4/8
                  'stFileName = Trim(grd1.TextMatrix(ii, 4))
                  stFileName = Trim(GRD1.TextMatrix(ii, 28))
                  '2024/4/8 END
                  If InStrRev(stFileName, " (") > 0 Then
                     'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
                     If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
                     '2021/8/6 END
                        stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
                     End If
                  End If
                  'stFullName = stFolderPath & stFileName
                  stFullName = stFolderPath & Trim(GRD1.TextMatrix(ii, 9))
                  'Add By Sindy 2024/4/8
                  If Trim(GRD1.TextMatrix(ii, 35)) <> "" Then '電子檔歸卷文號
                     m_CP09 = Trim(GRD1.TextMatrix(ii, 35))
                  Else
                  '2024/4/8 END
                     m_CP09 = Trim(GRD1.TextMatrix(ii, 6))
                  End If
                  
                  'Added by Morgan 2020/7/28 HP加密壓縮檔要先解密
                  strZipSrc = ""
                  If Right(UCase(stFileName), 14) = ".ENCRYPTED.ZIP" Then
                     strZipSrc = stFileName
                     stFileName = Left(stFileName, Len(stFileName) - 14)
                     
                     strSrcFile = App.path & "\" & stFileName
                     strZipFile = App.path & "\$ENCRYPTED.ZIP"
                     If Right(UCase(stFullName), 14) = ".ENCRYPTED.ZIP" Then
                        stFullName = Left(stFullName, Len(stFullName) - 14)
                     End If
                  End If
                  'end 2020/7/28
                  
                  If stFullName <> "" Then
                     If Dir(stFullName) <> "" Then
                        If MsgBox("檔案[ " & Trim(GRD1.TextMatrix(ii, 9)) & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                           stFullName = ""
                        End If
                     End If
                     If stFullName <> "" Then
                        'Added by Morgan 2020/7/28
                        If strZipSrc <> "" Then
                           If GetAttachFile(m_CP09, strZipSrc, strZipFile) = False Then
                              MsgBox "無法下載檔案[ " & stFileName & " ]！"
                              GoTo RunExit
                           End If
                        
                           If PUB_UnZipFile2(strZipFile, App.path, cHPZipPwd) = False Then
                              MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                              GoTo RunExit
                           End If
                           
                           'Added by Morgan 2025/3/28
                           '壓縮檔內的檔名可能會少0(因曾經有改過規則)
                           If Left(m_CP02, 1) = "0" Then
                              If Dir(strSrcFile) = "" Then
                                 strSrcFile = Replace(strSrcFile, m_CP02, Mid(m_CP02, 2))
                              End If
                           End If
                           'end 2025/3/28
                           
                           FileCopy strSrcFile, stFullName
                           Kill strZipFile
                           Kill strSrcFile
                        Else
                        'end 2020/7/28
                        
                           If GetAttachFile(m_CP09, stFileName, stFullName) = False Then
                              MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                              GoTo RunExit
                           End If
                           
                        End If 'Added by Morgan 2020/7/28
                     End If
                  End If
               End If
            Next ii
         End If
      Else
         If InStrRev(stFileName, " (") > 0 Then
            'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
            If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
            '2021/8/6 END
               stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            End If
         End If
         
         'Added by Morgan 2020/7/28 HP加密壓縮檔要先解密
         strZipSrc = ""
         If Right(UCase(stFileName), 14) = ".ENCRYPTED.ZIP" Then
            strZipSrc = stFileName
            stFileName = Left(stFileName, Len(stFileName) - 14)
            strSrcFile = App.path & "\" & stFileName
            strZipFile = App.path & "\$ENCRYPTED.ZIP"
         End If
         'end 2020/7/28
         
         'Modified by Lydia 2019/12/03 取得電子檔對應下載檔名
         'stFullName = GetSaveName(stFileName)
         'Modified by Lydia 2019/12/04 傳入本所案號
         'strName2 = ChkEfileMap(m_CP01, strCP10, stFileName)
         strName2 = ChkEfileMap(m_CP01, m_CP02, m_CP03, m_CP04, strCP10, stFileName)
         'stFullName = GetSaveName(strName2)
         stFullName = stFolderPath & strName2 'Modify By Sindy 2020/1/14
         'end 2019/12/03
         If stFullName <> "" Then
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案[ " & stFullName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
                  stFullName = ""
               End If
            End If
            If stFullName <> "" Then
               'Added by Morgan 2020/7/28
               If strZipSrc <> "" Then
                  If GetAttachFile(m_CP09, strZipSrc, strZipFile) = False Then
                     MsgBox "無法下載檔案[ " & stFileName & " ]！"
                     GoTo RunExit
                  End If
               
                  If PUB_UnZipFile2(strZipFile, App.path, cHPZipPwd) = False Then
                     MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                     GoTo RunExit
                  End If
                  
                  'Added by Morgan 2025/3/28
                  '壓縮檔內的檔名可能會少0(因曾經有改過規則)
                  If Left(m_CP02, 1) = "0" Then
                     If Dir(strSrcFile) = "" Then
                        strSrcFile = Replace(strSrcFile, m_CP02, Mid(m_CP02, 2))
                     End If
                  End If
                  'end 2025/3/28
                  
                  FileCopy strSrcFile, stFullName
                  Kill strZipFile
                  Kill strSrcFile
               Else
               'end 2020/7/28
               
                  If GetAttachFile(m_CP09, stFileName, stFullName) = False Then
                     MsgBox "無法儲存檔案[ " & stFileName & " ]！"
                     GoTo RunExit
                  End If
                  
               End If 'Added by Morgan 2020/7/28
            End If
         End If
      End If
      If stFullName <> "" Then
         MsgBox "下載完成！"
      End If
   End If
RunExit:
   Screen.MousePointer = vbDefault
End Sub

'新增
Private Sub cmdAddAtt_Click()
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim stReName As String
   Dim bolAdd As Boolean
   Dim intChkCnt As Integer
   Dim strFile As String
   Dim strCaseNoName As String
   Dim strCP82 As String 'Add By Sindy 2014/3/11
   
On Error GoTo ErrHnd
   
   intChkCnt = 0
   For ii = 1 To GRD1.Rows - 1
      'Modify By Sindy 2014/7/25
      GRD1.row = ii
      GRD1.col = 1
      'If GRD1.TextMatrix(ii, 0) = "V" Then
      If GRD1.CellBackColor = &HFFC0C0 Then
      '2014/7/25 END
         m_CP09 = Trim(GRD1.TextMatrix(ii, 6))
         m_CP10 = Trim(GRD1.TextMatrix(ii, 7))
         m_CP10Nm = Trim(GRD1.TextMatrix(ii, 3))
         strCP82 = Trim(GRD1.TextMatrix(ii, 11))
         intChkCnt = intChkCnt + 1
      End If
   Next ii
   If intChkCnt = 0 Then
      MsgBox "請選取一筆欲新增電子檔的總收文號！"
      Exit Sub
   ElseIf intChkCnt > 1 Then
      MsgBox "只可選取一筆總收文號做新增！"
      Exit Sub
   'Modify By Sindy 2014/6/24 Mark
'   'Add By Sindy 2014/3/11
'   ElseIf Trim(strCP82) = "" Then
'      MsgBox "此文尚未發文，不可異動附件！"
'      Exit Sub
   End If
   
   'Add By Sindy 2015/5/27 只能勾選是總收文號的資料列做新增
   'Modified by Morgan 2017/1/17 總收文號+D
   If InStr("A,B,C,D", Left(Trim(m_CP09), 1)) = 0 Then
      MsgBox "該筆資料並非總收文號，不可以做新增！"
      Exit Sub
   ElseIf Len(Trim(m_CP09)) <> 9 Then
      MsgBox "該筆資料並非總收文號，不可以做新增！"
      Exit Sub
   End If
   
   'Add By Sindy 2015/5/26
   Call frm100101_L_2.SetParent(Me)
   frm100101_L_2.m_CPP11 = m_CPP11 '電子表單單號 Add By Sindy 2023/2/18
   frm100101_L_2.m_identity = m_identity
   frm100101_L_2.m_CP09 = m_CP09
   frm100101_L_2.m_CP10 = m_CP10
   frm100101_L_2.m_CP10Nm = m_CP10Nm
   frm100101_L_2.m_Nation = m_Nation
   frm100101_L_2.Show vbModal
   Exit Sub
   '2015/5/26 END
     
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2013/10/30 檢查是否為電子送件,若是,電子檔是否全數歸檔
Private Sub ChkCP121()
   If m_CP01 = "P" And m_Nation = "000" Then
      'Modify By Sindy 2014/5/21 Mark 因UpdateCP121會檢查新案是否已歸足,若未歸足會重新檢核
'      '檢查新申請案
'      strExc(0) = "select cp09,cp10,cpm26 from caseprogress,casepropertymap" & _
'                  " where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
'                  " and cp10 in(" & NewCasePtyList & ")" & _
'                  " and cp57 is null and cp118 is not null and cp120='Y' and cp121 is null" & _
'                  " and cp01=cpm01(+) and cp10=cpm02(+)"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         Call UpdateCP121(RsTemp.Fields("cp09"), RsTemp.Fields("cp10"), "" & RsTemp.Fields("cpm26"), m_CP09)
'      End If
      '檢查此筆文號
      strExc(0) = "select cp09,cp10,cpm26 from caseprogress,casepropertymap" & _
                  " where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
                  " and cp57 is null and cp118 is not null and cp120='Y' and cp121 is null" & _
                  " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                  " and cp09='" & m_CP09 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Call UpdateCP121(m_CP09, RsTemp.Fields("cp10"), "" & RsTemp.Fields("cpm26"))
      End If
   End If
   '2013/10/30 END
End Sub

Private Function AddListX(stNewItem As String) As Boolean
   Dim idx As Integer, stFileName As String
   Dim stCP09 As String 'Add By Sindy 2015/3/6
   
   If stNewItem <> "" Then
      For idx = 1 To GRD1.Rows - 1
         stFileName = Trim(GetFileName(GRD1.TextMatrix(idx, 4)))
         stCP09 = Trim(GRD1.TextMatrix(idx, 6))
         'Modify By Sindy 2014/6/20
         'If GetFileName(stNewItem) = stFileName Then
         'Modify By Sindy 2015/3/6 +And stCP09 = m_CP09
         If UCase(GetFileName(stNewItem)) = UCase(stFileName) And stCP09 = m_CP09 Then
         '2014/6/20 END
            MsgBox "附件 " & stFileName & " 已存在！"
            AddListX = False
            'Exit For
            Exit Function
         End If
      Next idx
      
      AddListX = True
   End If
End Function

Private Function GetSaveName(ByVal pFileName As String) As String

On Error GoTo ErrHnd

   With CommonDialog1
      .CancelError = True
      .FileName = pFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowSave
      If .FileName <> "" Then
         GetSaveName = .FileName
      End If
   End With

   Exit Function

ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Function

'刪除
Private Sub cmdRemAtt_Click()
Dim bolDel As Boolean
Dim intChkCnt As Integer
Dim bolConn As Boolean
Dim rsTmp As New ADODB.Recordset
   
On Error GoTo ErrHnd
   
   intChkCnt = 0
   For ii = 1 To GRD1.Rows - 1
      If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v") And Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
         'Modify By Sindy 2015/9/15 + InStr(UCase(Trim(GRD1.TextMatrix(ii, 4))), UCase(EMP_多案承辦單)) > 0
         If InStr(UCase(Trim(GRD1.TextMatrix(ii, 4))), UCase(EMP_承辦單)) > 0 Or _
            InStr(UCase(Trim(GRD1.TextMatrix(ii, 4))), UCase(EMP_多案承辦單)) > 0 Then
            
            MsgBox "不可刪除承辦單(電子檔)！", vbExclamation
            Exit Sub
         End If
         '2015/9/15 END
         
         'Add By Sindy 2024/4/8
         If Trim(GRD1.TextMatrix(ii, 35)) <> "" Then
            MsgBox "此電子檔非此文號所有，無法刪除！", vbExclamation
            Exit Sub
         End If
         '2024/4/8 END
         
         'Add By Sindy 2015/5/26 檢查是否有刪除的權限
         'A.業務助理 S.智權人員 E.承辦人/工程師 (InStr("A,S,E", m_identity) > 0 And m_identity <> "")
         '+ 外專承辦
         If (m_identity <> "F" And m_identity <> "C") Or _
            Pub_StrUserSt03 = "F23" Then
            'Modify By Sindy 2017/6/21 外專承辦開放可以刪除國外部信件區匯入的郵件
            If Not (Pub_StrUserSt03 = "F23" And Trim(GRD1.TextMatrix(ii, 21)) = "F") Then
               '資料來源是卷宗區才能刪除
               'Modify By Sindy 2023/3/3 開放接洽單電子收文( Len(cpp11)=10 )時,存入的電子檔也可以再此處刪除
               If Trim(GRD1.TextMatrix(ii, 21)) <> "A" And Len(Trim(GRD1.TextMatrix(ii, 31))) <> 10 Then
                  MsgBox "電子檔（" & Trim(GRD1.TextMatrix(ii, 4)) & "）不是在卷宗區新增的,不可以刪除！"
                  Exit Sub
               'Modify By Sindy 2023/6/8 + Create ID
               ElseIf Trim(GRD1.TextMatrix(ii, 22)) <> strUserNum And Trim(GRD1.TextMatrix(ii, 32)) <> strUserNum Then '自己放的檔案才能刪除
                  MsgBox "電子檔（" & Trim(GRD1.TextMatrix(ii, 4)) & "）並非您新增的,不可以刪除！"
                  Exit Sub
               ElseIf DBDATE(DateAdd("d", 7, Format(Trim(GRD1.TextMatrix(ii, 23)), "####/##/##"))) < strSrvDate(1) Then '新增檔案日期在7天內的檔案才能刪除
                  MsgBox "電子檔（" & Trim(GRD1.TextMatrix(ii, 4)) & "）已超過7日可以刪除的期限,不可以刪除！"
                  Exit Sub
               End If
            End If
         End If
         '2015/5/26 END
         
'         'Add By Sindy 2018/6/5
'         If Trim(Grd1.TextMatrix(ii, 5)) = "官方來函" Then
'            MsgBox "不可刪除官方來函電子檔！", vbExclamation
'            Exit Sub
'         End If
'         '2018/6/5 END
         'Modify By Sindy 2018/6/15 官方來函使用者真的有時需要抽換電子檔,所以不用鎖使用者(刪除會留存一個月)
         '                          ,反而是怕電腦中心不知情誤刪
         If Trim(GRD1.TextMatrix(ii, 5)) = "官方來函" And Pub_StrUserSt03 = "M51" Then
            If MsgBox("確定要刪除官方來函電子檔嗎？" & vbCrLf & vbCrLf & _
                      "(若智慧局電子公文收錯案號, 刪除進度後人員即可重新收文)", vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            End If
         End If
         '2018/6/15 END
         
         intChkCnt = intChkCnt + 1
      End If
   Next ii
   'Add By Sindy 2019/2/19 檢查是否要刪除寄件備份
   For ii = 1 To GRD1.Rows - 1
      GRD1.col = 4
      GRD1.row = ii
      'Add By Sindy 2022/12/30
      If Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
         '0檔案非.menu的可以刪除
         strExc(10) = GetFileName(Trim(GRD1.TextMatrix(ii, 4)))
         If Val(GRD1.TextMatrix(ii, 16)) = 0 And UCase(Mid(strExc(10), InStrRev(strExc(10), "."))) <> UCase(".menu") Then
            GRD1.TextMatrix(ii, 0) = "v"
            intChkCnt = intChkCnt + 1
      '2022/12/30 END
         ElseIf GRD1.CellBackColor = &HFFC0C0 And _
            UCase(Right(GetFileName(Trim(GRD1.TextMatrix(ii, 4))), Len(".Email.menu"))) = UCase(".Email.menu") And _
            GRD1.TextMatrix(ii, 23) <> "" And GRD1.TextMatrix(ii, 29) <> "" And GRD1.TextMatrix(ii, 6) <> "" Then
            If MsgBox("確定要永久刪除[寄件備份]" & GetFileName(Trim(GRD1.TextMatrix(ii, 4))) & "電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
               Screen.MousePointer = vbDefault
               Exit Sub
            Else
               bolDel = True
               '寄件備份只有一筆資料
               'Modified by Morgan 2025/11/13 修正欄位
               'strExc(0) = "select smb01,smb11 from smailbackup where smb01='" & GRD1.TextMatrix(ii, 6) & "' and smb02=" & GRD1.TextMatrix(ii, 23) & " and smb03=" & GRD1.TextMatrix(ii, 29)
               strExc(0) = "select smb01,smb11 from smailbackup where smb01='" & GRD1.TextMatrix(ii, 6) & "' and smb02=" & GRD1.TextMatrix(ii, 24) & " and smb03=" & GRD1.TextMatrix(ii, 25)
               'end 2025/11/13
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If RsTemp.RecordCount > 0 Then
                  If RsTemp.RecordCount = 1 Then
                     '卷宗區只有一筆資料
                     'Modified by Morgan 2025/11/13 修正欄位
                     'strExc(0) = "select count(*) from casepaperpdf where cpp01='" & GRD1.TextMatrix(ii, 6) & "' and cpp06=" & GRD1.TextMatrix(ii, 23) & " and cpp07=" & GRD1.TextMatrix(ii, 29)
                     strExc(0) = "select count(*) from casepaperpdf where cpp01='" & GRD1.TextMatrix(ii, 6) & "' and cpp08=" & GRD1.TextMatrix(ii, 24) & " and cpp09=" & GRD1.TextMatrix(ii, 25)
                     'end 2025/11/13
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If RsTemp.Fields(0) > 0 Then
                        If RsTemp.Fields(0) = 1 Then
                           cnnConnection.BeginTrans: bolConn = True
                           '寄件備份
                           'Modified by Morgan 2025/11/13 修正欄位
                           'strSql = "delete smailbackup where smb01='" & GRD1.TextMatrix(ii, 6) & "' and smb02=" & GRD1.TextMatrix(ii, 23) & " and smb03=" & GRD1.TextMatrix(ii, 29)
                           strSql = "delete smailbackup where smb01='" & GRD1.TextMatrix(ii, 6) & "' and smb02=" & GRD1.TextMatrix(ii, 24) & " and smb03=" & GRD1.TextMatrix(ii, 25)
                           'end 2025/11/13
                           Pub_SeekTbLog strSql 'Add By Sindy 2019/5/16
                           cnnConnection.Execute strSql
                           '卷宗區
                           'Modified by Morgan 2025/11/13 修正欄位
                           'strSql = "delete casepaperpdf where cpp01='" & GRD1.TextMatrix(ii, 6) & "' and cpp06=" & GRD1.TextMatrix(ii, 23) & " and cpp07=" & GRD1.TextMatrix(ii, 29)
                           strSql = "delete casepaperpdf where cpp01='" & GRD1.TextMatrix(ii, 6) & "' and cpp08=" & GRD1.TextMatrix(ii, 24) & " and cpp09=" & GRD1.TextMatrix(ii, 25)
                           'end 2025/11/13
                           
                           Pub_SeekTbLog strSql 'Add By Sindy 2019/5/16
                           cnnConnection.Execute strSql
                           If bolConn = True Then cnnConnection.CommitTrans: bolConn = False
                           Call ReadAttachFile
                        Else
                           MsgBox GetFileName(Trim(GRD1.TextMatrix(ii, 4))) & vbCrLf & vbCrLf & "不可刪除，卷宗區不只一筆資料，請查明！", vbExclamation
                           Exit Sub
                        End If
                     Else
                        MsgBox GetFileName(Trim(GRD1.TextMatrix(ii, 4))) & vbCrLf & vbCrLf & "不可刪除，卷宗區查無資料，請查明！", vbExclamation
                        Exit Sub
                     End If
                  Else
                     MsgBox GetFileName(Trim(GRD1.TextMatrix(ii, 4))) & vbCrLf & vbCrLf & "不可刪除，寄件備份不只一筆資料，請查明！", vbExclamation
                     Exit Sub
                  End If
               Else
                  MsgBox GetFileName(Trim(GRD1.TextMatrix(ii, 4))) & vbCrLf & vbCrLf & "不可刪除，寄件備份查無資料，請查明！", vbExclamation
                  Exit Sub
               End If
            End If
         End If
      End If
   Next ii
   '2019/2/19 END
   If intChkCnt <= 0 Then
      'Modify By Sindy 2018/11/20 + If bolDel = False Then
      If bolDel = False Then MsgBox "請至少勾選一筆欲刪除電子檔的資料！"
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   'bolDel = False
   For ii = 1 To GRD1.Rows - 1
      If (GRD1.TextMatrix(ii, 0) = "V" Or GRD1.TextMatrix(ii, 0) = "v") And Trim(GRD1.TextMatrix(ii, 4)) <> "" Then
         If MsgBox("確定要永久刪除" & GetFileName(Trim(GRD1.TextMatrix(ii, 4))) & "電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         '直接從資料庫刪除檔案
         m_CP09 = Trim(GRD1.TextMatrix(ii, 6))
         If DeleteFile(m_CP09, GetFileName(Trim(GRD1.TextMatrix(ii, 4)))) = True Then
            bolDel = True
            'Added by Lydia 2019/03/06 FCP之公告公報1228增加判斷是否有公告本
            If m_CP01 = "FCP" And Trim(GRD1.TextMatrix(ii, 7)) = "1228" And InStr(UCase(GetFileName(Trim(GRD1.TextMatrix(ii, 4)))), ".GAZ.PDF") > 0 Then
                 Call UpdateCP121(m_CP09, Trim(GRD1.TextMatrix(ii, 7)), "GAZ")
            End If
            'end 2019/03/06
         End If
      End If
   Next ii
   
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   If bolDel = True Then Call ReadAttachFile
   Exit Sub
   
ErrHnd:
   Set rsTmp = Nothing
   If bolConn = True Then cnnConnection.RollbackTrans: bolConn = False
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Added by Lydia 2018/02/01 FCP含已發文之客戶提供文件
Private Sub ChkDelC_Click()
   Call ReadAttachFile
End Sub

'Added by Lydia 2018/09/06 不顯示郵件
Private Sub ChkDelMsg_Click()
   Call ReadAttachFile
End Sub

'Added by Lydia 2019/12/03 電子檔對應下載檔名
'Modified by Lydia 2019/12/04 傳入本所案號
'Private Function ChkEfileMap(ByVal pCP01 As String, pCP10 As String, pOldName As String) As String
Private Function ChkEfileMap(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, pCP10 As String, pOldName As String) As String
Dim intJ As Integer, strB As String
Dim rsB1 As New ADODB.Recordset
Dim m_stReName As String
          
    ChkEfileMap = pOldName
    
    If pCP01 <> "" And pCP10 <> "" Then
        strB = "select efm03, efm04 from EFILEMAP where EFM01='" & pCP01 & "' AND EFM02='" & pCP10 & "' order by efm03 "
        intJ = 1
        Set rsB1 = ClsLawReadRstMsg(intJ, strB)
        If intJ = 1 Then
            rsB1.MoveFirst
            Do While Not rsB1.EOF
                 If Right(UCase(pOldName), Len("" & rsB1.Fields("efm03"))) = UCase("" & rsB1.Fields("efm03")) Then
                     'Modified by Lydia 2019/12/04 統一卷宗區的檔名之本所案號 => 本所案號: 系統別+完整流水號
                     'ChkEfileMap = Left(pOldName, Len(pOldName) - Len("" & rsB1.Fields("efm03"))) & rsB1.Fields("efm04")
                     ChkEfileMap = pCP01 & pCP02 & IIf(pCP03 & pCP04 <> "000", pCP03 & pCP04, "") & rsB1.Fields("efm04")
                     Exit Do
                 End If
                 rsB1.MoveNext
            Loop
        'Added by Lydia 2019/12/04 統一卷宗區的檔名之本所案號 => 本所案號: 系統別+完整流水號
        Else
               If PUB_GetEmpFlowReNameFile(pCP01, pCP02, pCP03, pCP04, pCP10, pOldName, m_stReName, True, 1) = True Then
                   ChkEfileMap = m_stReName
               End If
        'end 2019/12/04
        End If
        Set rsB1 = Nothing
    End If
    
End Function

'Add by Amy 2022/06/17 取得T延展結案之總收文號及下一程序號
Private Function GetTI02(ByVal stTi01 As String, ByVal stCP09 As String, ByRef stNP22 As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    GetTI02 = "": stNP22 = ""
    'Memo 同一天(ti01)同一總收文號及下一程序號 只會有一筆資料
    'Modify by Amy 2022/06/22 拿掉And Nvl(ti06,' ')=' ' ex:T-177073 有取消延展,又做了閉卷
    strQ = "Select ti02,ti04 From T102Inform,CaseProgress " & _
              "Where ti01='" & stTi01 & "' And cp09='" & stCP09 & "' And ti02=cp43(+) And ti04=cp30(+) " & _
              "And cp43 is not null And cp30 is not null "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetTI02 = "" & RsQ.Fields("ti02") '總收文號
        stNP22 = "" & RsQ.Fields("ti04") '下一程序號
    End If
    Set RsQ = Nothing
End Function

