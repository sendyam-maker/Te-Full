VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210145 
   BorderStyle     =   1  '單線固定
   Caption         =   "寄件查詢"
   ClientHeight    =   5740
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5740
   ScaleWidth      =   8950
   Tag             =   "加班資料"
   Begin VB.TextBox systemkind 
      Height          =   270
      Left            =   945
      TabIndex        =   7
      Text            =   "ALL"
      Top             =   1170
      Width           =   1680
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Height          =   645
      Left            =   4455
      TabIndex        =   39
      Top             =   0
      Width           =   4515
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Caption         =   "點我展開"
         Height          =   345
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   41
         Top             =   0
         Width           =   4515
      End
      Begin VB.ComboBox cboAtt 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frm210145.frx":0000
         Left            =   900
         List            =   "frm210145.frx":0002
         Style           =   2  '單純下拉式
         TabIndex        =   40
         Top             =   330
         Width           =   3645
      End
      Begin VB.Label lblAttCnt 
         Alignment       =   2  '置中對齊
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  '單線固定
         Caption         =   " PDF:(0)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -15
         TabIndex        =   42
         Top             =   330
         Width           =   930
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5115
      Left            =   4440
      TabIndex        =   23
      Top             =   630
      Width           =   4515
      ExtentX         =   7964
      ExtentY         =   9022
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   375
      Left            =   30
      TabIndex        =   27
      Top             =   5430
      Width           =   4425
      Begin VB.CommandButton cmdEmail 
         Caption         =   "EMail"
         Height          =   315
         Left            =   2745
         TabIndex        =   32
         Top             =   0
         Width           =   600
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "預覽"
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdSaveAtt 
         Caption         =   "下載"
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   0
         Width           =   800
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   315
         Index           =   0
         Left            =   840
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   800
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2205
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label LblCnt 
         AutoSize        =   -1  'True
         Caption         =   "共 0 筆"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   3630
         TabIndex        =   38
         Top             =   60
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "收文"
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   12
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "卷宗區"
      Height          =   285
      Index           =   4
      Left            =   1170
      TabIndex        =   13
      Top             =   0
      Width           =   800
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   30
      TabIndex        =   28
      Top             =   300
      Width           =   2745
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Left            =   885
         TabIndex        =   11
         Top             =   0
         Width           =   1800
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3175;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員:"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   0
         TabIndex        =   29
         Top             =   60
         Width           =   765
      End
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   2130
      MaxLength       =   7
      TabIndex        =   1
      Top             =   570
      Width           =   825
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1230
      MaxLength       =   7
      TabIndex        =   0
      Top             =   570
      Width           =   825
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "進度(&C)"
      Height          =   285
      Index           =   0
      Left            =   2790
      TabIndex        =   15
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   285
      Index           =   3
      Left            =   3600
      TabIndex        =   16
      Top             =   0
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3735
      Left            =   30
      TabIndex        =   21
      Top             =   1650
      Width           =   4395
      _ExtentX        =   7743
      _ExtentY        =   6579
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|發文日|本所案號|案件性質|案件名稱|國家|申請人|pa09|pa26|cp09|CP10"
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
      _Band(0).Cols   =   11
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm210145.frx":0004
      Left            =   3120
      List            =   "frm210145.frx":0017
      TabIndex        =   2
      Top             =   570
      Width           =   690
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結案"
      Height          =   285
      Index           =   6
      Left            =   1980
      TabIndex        =   14
      Top             =   0
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   945
      MaxLength       =   3
      TabIndex        =   3
      Top             =   855
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   1455
      MaxLength       =   6
      TabIndex        =   4
      Top             =   855
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   5
      Top             =   855
      Width           =   276
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   6
      Top             =   855
      Width           =   330
   End
   Begin VB.CheckBox Check2 
      Caption         =   "剔除已收文,結案,閉卷,無期限,結案中"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1470
      Width           =   3225
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "關係企業"
      Height          =   285
      Index           =   2
      Left            =   2700
      TabIndex        =   10
      Top             =   1140
      Width           =   885
   End
   Begin VB.CheckBox Check1 
      Caption         =   "關閉預覽"
      Height          =   195
      Left            =   3330
      TabIndex        =   17
      Top             =   330
      Width           =   1065
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "查詢(&Q)"
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   9
      Top             =   1140
      Width           =   800
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   30
      TabIndex        =   43
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label lblColorDesc 
      AutoSize        =   -1  'True
      Caption         =   "E化"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   0
      Left            =   3285
      TabIndex        =   37
      Top             =   900
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label lblColor 
      Appearance      =   0  '平面
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   3060
      TabIndex        =   36
      Top             =   900
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblColor 
      Appearance      =   0  '平面
      BackColor       =   &H000080FF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   3585
      TabIndex        =   35
      Top             =   900
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label lblColorDesc 
      AutoSize        =   -1  'True
      Caption         =   "全E化"
      ForeColor       =   &H8000000D&
      Height          =   180
      Index           =   1
      Left            =   3810
      TabIndex        =   34
      Top             =   900
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "("
      Height          =   195
      Index           =   3
      Left            =   3000
      TabIndex        =   31
      Top             =   600
      Width           =   105
   End
   Begin VB.Label Label1 
      Caption         =   "個月)"
      Height          =   195
      Index           =   1
      Left            =   3840
      TabIndex        =   30
      Top             =   630
      Width           =   435
   End
   Begin VB.Label LblAppl 
      AutoSize        =   -1  'True
      Caption         =   "LblAppl"
      Height          =   180
      Left            =   690
      TabIndex        =   26
      Top             =   330
      Width           =   585
   End
   Begin VB.Label Label3 
      Caption         =   "申請人:"
      Height          =   225
      Left            =   30
      TabIndex        =   25
      Top             =   330
      Width           =   615
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2085
      X2              =   2205
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Label Label1 
      Caption         =   "發文室發文日:"
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   24
      Top             =   630
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(雙擊預覽)"
      ForeColor       =   &H00FF0000&
      Height          =   180
      Index           =   0
      Left            =   3360
      TabIndex        =   22
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Left            =   30
      TabIndex        =   33
      Top             =   900
      Width           =   900
   End
End
Attribute VB_Name = "frm210145"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/18 改成Form2.0 (GRD1,Combo1)
'Create by Sindy 2014/5/12
Option Explicit

' 變數宣告區
Dim m_CP09 As String
Dim m_CP10 As String 'Add By Sindy 2014/11/4
Dim ii As Integer, jj As Integer
'Dim m_bolDblClick As Boolean
Dim m_mouseRow As Integer
Dim Str01 As String

'附件宣告區
Dim m_AttachPath As String
Dim m_AttachPath2 As String 'Added by Morgan 2020/2/18
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
'Const GrdMaxW = 9474 'Removed by Moran 2015/9/1 為避免加欄位還要調寬度改變設定方法
Const GrdMinW = 4395
Public intWorkItem As Integer 'Add By Sindy 2014/5/22 0.以申請人查詢寄件資料 1.寄件查詢
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Public cmdState As Integer '紀錄作用按鍵
Dim bolChgSystemkind As Boolean, strOldSystemkind As String 'Add By Sindy 2020/9/4


'Added by Morgan 2019/3/4
Private Sub cboAtt_Click()
   Dim hLocalFile As Long
   Dim arrFileName() As String
   
   arrFileName = Split(cboAtt.List(cboAtt.ListIndex), Chr(9))
   'Modified by Morgan 2020/2/18
   'WebBrowser1.Navigate m_AttachPath & "\" & arrFileName(0): DoEvents
   'Modified by Morgan 2021/5/7
   'WebBrowser1.Navigate m_AttachPath2 & "\" & arrFileName(0): DoEvents
   strExc(1) = m_AttachPath2 & "\" & arrFileName(0)
   strExc(2) = strExc(1) & "_Copy"
   If Dir(strExc(2)) = "" Then
      FileCopy strExc(1), strExc(2)
   End If
   WebBrowser1.Navigate strExc(2): DoEvents
   'end 2021/5/7
End Sub

Private Sub SetGrd(Optional bolSetRow As Boolean = True)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   Dim iDep As String
   
   iDep = PUB_GetST06(strUserNum)
   'Modified by Morgan 2015/9/1 +報價
   '                        0    1         2           3           4           5       6         7           8           9           10      11          12      13      14      15      16       17      18      19      20      21      22      23      24      25
   arrGridHeadText = Array("V", "發文日", "本所案號", "案件性質", "案件名稱", "國家", "申請人", "本所期限", "寄件方式", "分所案號", "報價", "備註", "確認人員", "pa09", "pa26", "cp09", "CP10", "CP127", "cp43", "cp01", "cp02", "cp03", "cp04", "pa58", "lP26", "LP35")
   '開啟預覽
   If Check1.Value = 0 Then
      '電腦中心，跟分所才秀分所案號
      If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
         arrGridHeadWidth = Array(200, 800, 1000, 1200, 1000, 500, 800, 800, 450, 0, 860, 1100, 800, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      Else
         arrGridHeadWidth = Array(200, 800, 1000, 1200, 1000, 500, 800, 800, 450, 800, 860, 900, 800, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      End If
   '關閉預覽
   Else
      '電腦中心，跟分所才秀分所案號
      If GetStaffDepartment(strUserNum) <> "M51" And iDep = "1" Then
         arrGridHeadWidth = Array(200, 800, 1000, 1200, 1000, 500, 800, 800, 450, 0, 860, 1500, 800, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      Else
         arrGridHeadWidth = Array(200, 800, 1000, 1200, 1000, 500, 800, 800, 450, 800, 860, 1300, 800, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      End If
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
   GRD1.Visible = True
End Sub

'以申請人做查詢
Public Function QueryData(bolCmd As Boolean) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String
'Dim strTMSP As String
Dim strNotTM As String, strNotPA As String
Dim strP As String, strTM As String, strSP As String 'Add by Amy 2020/02/05
   
   Str01 = ""
   Str01 = Me.Tag
   If bolCmd = False Then
      '檢查國內外權限
      If Len(Str01) <> 9 Then Str01 = Str01 & "0"
      If CheckSR12(Str01) = False Then
         Screen.MousePointer = vbDefault
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Function
      End If
   End If
   
   QueryData = True
   '清空及預設欄位值
   GRD1.Clear
   SetGrd
   m_mouseRow = 0 'Add By Sindy 2014/11/21
   
   'Add By Sindy 2020/9/4
   bolChgSystemkind = False
   If Text1(0) = "" Or Text1(1) = "" Then
      If systemkind = "" Then
         systemkind = "ALL"
      End If
   Else
      bolChgSystemkind = True
      strOldSystemkind = systemkind
      systemkind = Text1(0).Text
   End If
   '2020/9/4 END
   
   strCon = ""
   'Add By Sindy 2015/1/12 不列出A,B類
   If Check2.Value = 1 Then '剔除已收文,結案,閉卷
      'Modified by Morgan 2017/1/17 總收文號+D
      'strCon = strCon & " and substr(cp09,1,1)='C'"
      strCon = strCon & " and substr(cp09,1,1)>='C'"
      'end 2017/1/17
      
      'Add By Sindy 2019/2/1 為使後續逐筆過濾資料會導致速度慢,儘量在SQL裡把資料過濾掉
      '剔除無期限(本所期限),抓有期限的
      strCon = strCon & " and cp06 is not null"
      '剔除已閉卷,抓未閉卷的
      'Modify by Amy 2020/02/05 原:strCon
      strP = strP & " and pa58 is null"
      strTM = strTM & " and tm30 is null"
      strSP = strSP & " and sp16 is null"
      'end 2020/02/05
      '2019/2/1 END
   End If
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   '案件資料
   'Modify By Sindy 2014/7/10 +,pa47 as 分所案號 及 ,pa47
   'Modify By Sindy 2014/8/6 +,cp43,cp01,cp02,cp03,cp04
   'Modify By Sindy 2014/10/2 +,pa58
   'Modify By Sindy 2014/10/9 +因案件會有多申請人狀況,但同文不可重覆出現,所以資料讀取出來再逐筆查詢申請人
'   strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04) as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58" & _
'            " from casepaperpdf,(" & _
'            "Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)='" & Left(Str01, 8) & "' and substr(pa26,9,1)='" & Mid(Str01, 9, 1) & "'" & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa27,pa47,cp43,pa58 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa27,1,8)='" & Left(Str01, 8) & "' and substr(pa27,9,1)='" & Mid(Str01, 9, 1) & "'" & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa28,pa47,cp43,pa58 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa28,1,8)='" & Left(Str01, 8) & "' and substr(pa28,9,1)='" & Mid(Str01, 9, 1) & "'" & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa29,pa47,cp43,pa58 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa29,1,8)='" & Left(Str01, 8) & "' and substr(pa29,9,1)='" & Mid(Str01, 9, 1) & "'" & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa30,pa47,cp43,pa58 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa30,1,8)='" & Left(Str01, 8) & "' and substr(pa30,9,1)='" & Mid(Str01, 9, 1) & "'" & _
'            "),casepropertymap,nation,customer,LetterProgress,staff" & _
'            " where cp09=cpp01(+) and cpp02 is not null" & _
'            " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+)" & _
'            " and cp09=LP01(+) and LP06=st01(+)" & _
'            " order by cp127 desc"
   'Modified by Morgan 2015/6/29 +LP26
   'Modify By Sindy 2015/6/24 以本所案號查詢
   'Modified by Morgan 2015/9/1 +報價
   'Modify By Sindy 2016/8/5 原:"Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,cp43 from caseprogress where cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "' and cp127=(select max(cp127) from caseprogress where cp127>0 and cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "'" & strCon & ")"
   '                         加:日期做條件
   'Modify by Amy 2020/02/05 Mark原程式 +Trademark,ServicePractice,整理成一句
'   If Text1(0) <> "" And Text1(1) <> "" Then
'      'Modify By Sindy 2016/8/5
'      If Text4 <> "" Then '有輸入發文日期
'         strCon = strCon & " and cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5)
'      Else
'         strCon = strCon & " and cp127>0"
'      End If
'      '2016/8/5 END
'      'Modified by Morgan 2019/1/23 +FDesc
'      'Modify By Sindy 2019/8/1 + ,cp30
'      strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58,LP26,LP35,'' FDesc" & _
'               " from casepaperpdf,patent,(" & _
'               "Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,cp43,cp30 from caseprogress where cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "'" & strCon & _
'               "),casepropertymap,nation,LetterProgress,staff,Customer" & _
'               " where cp09=cpp01(+) and cpp02 is not null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
'               " and ((substr(pa26,1,8)='" & Left(Str01, 8) & "' and substr(pa26,9,1)='" & Mid(Str01, 9, 1) & "') or (substr(pa27,1,8)='" & Left(Str01, 8) & "' and substr(pa27,9,1)='" & Mid(Str01, 9, 1) & "') or (substr(pa28,1,8)='" & Left(Str01, 8) & "' and substr(pa28,9,1)='" & Mid(Str01, 9, 1) & "') or (substr(pa29,1,8)='" & Left(Str01, 8) & "' and substr(pa29,9,1)='" & Mid(Str01, 9, 1) & "') or (substr(pa30,1,8)='" & Left(Str01, 8) & "' and substr(pa30,9,1)='" & Mid(Str01, 9, 1) & "'))" & _
'               " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
'               " and cp09=LP01(+) and LP06=st01(+)" & _
'               " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)"
'   Else
'   '2015/6/24 END
'      'Modify By Sindy 2019/8/1 + ,cp30
'      strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58,LP26,LP35,'' FDesc" & _
'               " from casepaperpdf,(" & _
'               "Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)='" & Left(Str01, 8) & "' and substr(pa26,9,1)='" & Mid(Str01, 9, 1) & "'" & strCon & _
'               " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa27,1,8)='" & Left(Str01, 8) & "' and substr(pa27,9,1)='" & Mid(Str01, 9, 1) & "'" & strCon & _
'               " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa28,1,8)='" & Left(Str01, 8) & "' and substr(pa28,9,1)='" & Mid(Str01, 9, 1) & "'" & strCon & _
'               " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa29,1,8)='" & Left(Str01, 8) & "' and substr(pa29,9,1)='" & Mid(Str01, 9, 1) & "'" & strCon & _
'               " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa30,1,8)='" & Left(Str01, 8) & "' and substr(pa30,9,1)='" & Mid(Str01, 9, 1) & "'" & strCon & _
'               "),casepropertymap,nation,LetterProgress,staff,Customer" & _
'               " where cp09=cpp01(+) and cpp02 is not null" & _
'               " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
'               " and cp09=LP01(+) and LP06=st01(+)" & _
'               " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)"
'   End If
   If Text1(0) <> MsgText(601) And Text1(1) <> MsgText(601) Then
        strCon = strCon & " And cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "'"
   End If
   If Text4 <> "" Then '有輸入發文日期
         strCon = strCon & " And cp127 between " & DBDATE(Text4) & " And " & DBDATE(Text5)
   ElseIf Text1(0) <> MsgText(601) And Text1(1) <> MsgText(601) And Text4 = MsgText(601) Then
        strCon = strCon & " And cp127>0"
   End If
   
   If Check2.Value = 1 Then
      '增加剔除的語法
      strNotPA = AddCheck2Sql(strNotTM)
   End If
   
   'Modified by Morgan 2020/8/3 ServicePractice移到下面
'   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
'        strTMSP = " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)='" & Left(Str01, 8) & "' and substr(tm23,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
'                         " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm78,1,8)='" & Left(Str01, 8) & "' and substr(tm78,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
'                         " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm79,1,8)='" & Left(Str01, 8) & "' and substr(tm79,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
'                         " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm80,1,8)='" & Left(Str01, 8) & "' and substr(tm80,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
'                         " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm81,1,8)='" & Left(Str01, 8) & "' and substr(tm81,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM
'   End If
   'Modified by Morgan 2022/8/12 +已判發LP05>0
   strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58,LP26,LP35,'' FDesc" & _
            " from casepaperpdf,(" & _
            "Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa26,1,8)='" & Left(Str01, 8) & "' and substr(pa26,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa27,1,8)='" & Left(Str01, 8) & "' and substr(pa27,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa28,1,8)='" & Left(Str01, 8) & "' and substr(pa28,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa29,1,8)='" & Left(Str01, 8) & "' and substr(pa29,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and substr(pa30,1,8)='" & Left(Str01, 8) & "' and substr(pa30,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp09 as pa09,sp08 as pa26,sp28 as pa47,cp43,sp16 as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)='" & Left(Str01, 8) & "' and substr(sp08,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp09 as pa09,sp08 as pa26,sp28 as pa47,cp43,sp16 as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp58,1,8)='" & Left(Str01, 8) & "' and substr(sp58,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp09 as pa09,sp08 as pa26,sp28 as pa47,cp43,sp16 as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp59,1,8)='" & Left(Str01, 8) & "' and substr(sp59,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp09 as pa09,sp08 as pa26,sp28 as pa47,cp43,sp16 as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp65,1,8)='" & Left(Str01, 8) & "' and substr(sp65,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp09 as pa09,sp08 as pa26,sp28 as pa47,cp43,sp16 as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp66,1,8)='" & Left(Str01, 8) & "' and substr(sp66,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            "),casepropertymap,nation,LetterProgress,staff,Customer" & _
            " where cp09=cpp01(+) and cpp02 is not null" & _
            " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
            " and cp09=LP01(+) and LP06=st01(+) and lp05>0" & _
            " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)" & strNotPA
   'Modify By Sindy 2021/4/13
   'Modified by Morgan 2022/8/12 +已判發LP05>0
   strSql = strSql & " union " & _
            "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58,LP26,LP35,'' FDesc" & _
            " from casepaperpdf,(" & _
            "Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm23,1,8)='" & Left(Str01, 8) & "' and substr(tm23,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm78,1,8)='" & Left(Str01, 8) & "' and substr(tm78,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm79,1,8)='" & Left(Str01, 8) & "' and substr(tm79,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm80,1,8)='" & Left(Str01, 8) & "' and substr(tm80,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and substr(tm81,1,8)='" & Left(Str01, 8) & "' and substr(tm81,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp09 as pa09,sp08 as pa26,sp28 as pa47,cp43,sp16 as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp08,1,8)='" & Left(Str01, 8) & "' and substr(sp08,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp09 as pa09,sp08 as pa26,sp28 as pa47,cp43,sp16 as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp58,1,8)='" & Left(Str01, 8) & "' and substr(sp58,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp09 as pa09,sp08 as pa26,sp28 as pa47,cp43,sp16 as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp59,1,8)='" & Left(Str01, 8) & "' and substr(sp59,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp09 as pa09,sp08 as pa26,sp28 as pa47,cp43,sp16 as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp65,1,8)='" & Left(Str01, 8) & "' and substr(sp65,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp09 as pa09,sp08 as pa26,sp28 as pa47,cp43,sp16 as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and substr(sp66,1,8)='" & Left(Str01, 8) & "' and substr(sp66,9,1)='" & Mid(Str01, 9, 1) & "' and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & strCon & strSP & _
            "),casepropertymap,nation,LetterProgress,staff,Customer" & _
            " where cp09=cpp01(+) and cpp02 is not null" & _
            " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
            " and cp09=LP01(+) and LP06=st01(+) and lp05>0" & _
            " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)" & strNotTM
   '2021/4/13 END
   'end 2020/8/3
   'end 2020/02/05
   strSql = strSql & " order by cp127 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
      LblCnt.Visible = True
      LblCnt.Caption = "共 0 筆" 'Add By Sindy 2019/2/1
   'End If
   If rsTmp.RecordCount > 0 Then
      'If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
         LblCnt.Visible = True
         LblCnt.Caption = "共 " & rsTmp.RecordCount & " 筆"  'Add By Sindy 2019/2/1
      'End If
      Set GRD1.Recordset = rsTmp
      Call RunLoopChkData 'Modify By Sindy 2014/8/6
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      QueryData = False
      rsTmp.Close
      Set rsTmp = Nothing
      Me.Enabled = True
      If bolCmd = False Then
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      End If
      Exit Function
   End If
   rsTmp.Close
   
   GRD1.col = 0
   GRD1.row = 1
   
   If bolChgSystemkind = True Then systemkind = strOldSystemkind 'Add By Sindy 2020/9/4
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   
EXITSUB:
   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Function

'以智權人員做查詢
Public Sub QueryData2()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
'Dim strTMSP As String
Dim strNotTM As String, strNotPA As String
Dim strP As String, strTM As String, strSP As String 'Add by Amy 2020/02/05
   
   '清空及預設欄位值
   GRD1.Clear
   SetGrd
   m_mouseRow = 0 'Add By Sindy 2014/11/21
   m_blnColOrderAsc = True
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
'SQLGrpStr(GetAllSysKind(systemkind), 1):專利'P','FCP','CFP',' '
'SQLGrpStr(GetAllSysKind(systemkind), 2):商標'TF','T','FCT','CFT',' '
'SQLGrpStr(GetAllSysKind(systemkind), 3):法務'LIN','L','FCL','CFL','ACS',' '
'SQLGrpStr(GetAllSysKind(systemkind), 4):顧問'LA',' '
'SQLGrpStr(GetAllSysKind(systemkind), 5):服務'TT','TS','TR','TM','TD','TC','TB','S','PS','FG','CPS','CFC',' '
   'Add By Sindy 2020/9/4
   bolChgSystemkind = False
   If Text1(0) = "" Or Text1(1) = "" Then
      If systemkind = "" Then
         systemkind = "ALL"
      End If
   Else
      bolChgSystemkind = True
      strOldSystemkind = systemkind
      systemkind = Text1(0).Text
   End If
   '2020/9/4 END
   
   strCon = ""
   If Trim(Combo1.Text) <> "全部" Then
      'Modify by Amy 2020/02/05 +MCTF0X
      If Mid(Trim(Combo1.Text), 1, 4) = "MCTF" Then
         'Modify By Sindy 2020/11/23
         strExc(10) = Pub_GetSpecMan(Left(Trim(Combo1.Text), 6))
         If strExc(10) <> "" Then
            strCon = " and cp13||'' in('" & Left(Trim(Combo1.Text), 6) & "','" & strExc(10) & "')"
         Else
         '2020/11/23 END
            strCon = " and cp13||''='" & Left(Trim(Combo1.Text), 6) & "'"
         End If
      Else
        'Modify By Sindy 2014/9/3 改抓CP13,因吳金龍客戶轉其他人時,就抓不到資料了
        'strCon = " and cu13='" & Left(Trim(Combo1.Text), 5) & "'"
        strCon = " and cp13||''='" & Left(Trim(Combo1.Text), 5) & "'"
        '2014/9/3 END
      End If
   End If
   'Add By Sindy 2015/1/12 不列出A,B類
   If Check2.Value = 1 Then '剔除已收文,結案,閉卷
      'Modified by Morgan 2017/1/17 總收文號+D
      'strCon = strCon & " and substr(cp09,1,1)='C'"
      strCon = strCon & " and substr(cp09,1,1)>='C'"
      'end 2017/1/17
      
      'Add By Sindy 2019/2/1 為使後續逐筆過濾資料會導致速度慢,儘量在SQL裡把資料過濾掉
      '剔除無期限(本所期限),抓有期限的
      strCon = strCon & " and cp06 is not null"
      '剔除已閉卷,抓未閉卷的
      'Modify by Amy 2020/02/05 原:strCon
      strP = strP & " and pa58 is null"
      strTM = strTM & " and tm30 is null"
      strSP = strSP & " and sp16 is null"
      'end 2020/02/05
      '2019/2/1 END
   End If
   
   '案件資料
   'Modify By Sindy 2014/7/10 +,pa47 as 分所案號 及 ,pa47
   'Modify By Sindy 2014/8/6 +,cp43,cp01,cp02,cp03,cp04
   'Modify By Sindy 2014/10/2 +,pa58
   'Modify By Sindy 2014/10/9 +因案件會有多申請人狀況,但同文不可重覆出現,所以資料讀取出來再逐筆查詢申請人
'   strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04) as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58" & _
'            " from casepaperpdf,(" & _
'            "Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,cu04,cu05,cu06,cu88,cu89,cu90,pa47,cp43,pa58 from caseprogress,patent,customer where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa26 is not null and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+)" & strCon & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa27,cu04,cu05,cu06,cu88,cu89,cu90,pa47,cp43,pa58 from caseprogress,patent,customer where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa27 is not null and substr(pa27,1,8)=cu01(+) and substr(pa27,9,1)=cu02(+)" & strCon & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa28,cu04,cu05,cu06,cu88,cu89,cu90,pa47,cp43,pa58 from caseprogress,patent,customer where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa28 is not null and substr(pa28,1,8)=cu01(+) and substr(pa28,9,1)=cu02(+)" & strCon & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa29,cu04,cu05,cu06,cu88,cu89,cu90,pa47,cp43,pa58 from caseprogress,patent,customer where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa29 is not null and substr(pa29,1,8)=cu01(+) and substr(pa29,9,1)=cu02(+)" & strCon & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa30,cu04,cu05,cu06,cu88,cu89,cu90,pa47,cp43,pa58 from caseprogress,patent,customer where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa30 is not null and substr(pa30,1,8)=cu01(+) and substr(pa30,9,1)=cu02(+)" & strCon & _
'            "),casepropertymap,nation,LetterProgress,staff" & _
'            " where cp09=cpp01(+) and cpp02 is not null" & _
'            " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
'            " and cp09=LP01(+) and LP06=st01(+)" & _
'            " order by cp127 desc"
'   strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04) as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,'' AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58" & _
'            " from casepaperpdf,(" & _
'            "Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa26 is not null" & strCon & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa27 is not null" & strCon & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa28 is not null" & strCon & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa29 is not null" & strCon & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa30 is not null" & strCon & _
'            "),casepropertymap,nation,LetterProgress,staff" & _
'            " where cp09=cpp01(+) and cpp02 is not null" & _
'            " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
'            " and cp09=LP01(+) and LP06=st01(+)" & _
'            " order by cp127 desc"
   'Modified by Morgan 2015/6/29 +LP26
   'Modify By Sindy 2015/6/24 以本所案號查詢
   'Modified by Morgan 2015/9/1 +報價
   'Modify By Sindy 2016/8/5 原:"Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,cp43 from caseprogress where cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "' and cp127=(select max(cp127) from caseprogress where cp127>0 and cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "'" & strCon & ")"
   '                         加:日期做條件
   'Modify by Amy 2020/02/05 Mark原程式 +Trademark,ServicePractice,整理成一句
'   If Text1(0) <> "" And Text1(1) <> "" Then
'      'Modify By Sindy 2016/8/5
'      If Text4 <> "" Then '有輸入發文日期
'         strCon = strCon & " and cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5)
'      Else
'         strCon = strCon & " and cp127>0"
'      End If
'      '2016/8/5 END
'      'Modify By Sindy 2019/8/1 + ,cp30
'      strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58,LP26,LP35" & _
'               " from casepaperpdf,patent,(" & _
'               "Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,cp43,cp30 from caseprogress where cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "'" & strCon & _
'               "),casepropertymap,nation,LetterProgress,staff,Customer" & _
'               " where cp09=cpp01(+) and cpp02 is not null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
'               " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
'               " and cp09=LP01(+) and LP06=st01(+)" & _
'               " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)"
'   Else
'   '2015/6/24 END
'      'Modify By Sindy 2019/8/1 + ,cp30
'      strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58,LP26,LP35" & _
'               " from casepaperpdf,(" & _
'               "Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & strCon & _
'               "),casepropertymap,nation,LetterProgress,staff,Customer" & _
'               " where cp09=cpp01(+) and cpp02 is not null" & _
'               " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
'               " and cp09=LP01(+) and LP06=st01(+)" & _
'               " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)"
'   End If
   If Text1(0) <> MsgText(601) And Text1(1) <> MsgText(601) Then
        strCon = strCon & " And cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "'"
   End If
   If Text4 <> "" Then '有輸入發文日期
        strCon = strCon & " and cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5)
   ElseIf Text1(0) <> MsgText(601) And Text1(1) <> MsgText(601) And Text4 = MsgText(601) Then
        strCon = strCon & " and cp127>0"
   End If
   
   If Check2.Value = 1 Then
      '增加剔除的語法
      'strSql = strSql & AddCheck2Sql()
      strNotPA = AddCheck2Sql(strNotTM)
'      '剔除已收文
'      strSql = strSql & " and not exists(select np01 from nextprogress where np01=decode(cp01||cp10,'P1913',cp43,'PS1913',cp43,'FCP1913',cp43,'FG1913',cp43,'CFP1913',cp43,'CPS1913',cp43,cp09)" & _
'                                       " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & _
'                                       " and np06 in('Y','N')" & strNpSqlOfNoSalesDuty & ")"
'
'      'Add By Sindy 2015/5/20 或用CP09串不到NP01,也是無期限資料 ex.P-103520
'      strSql = strSql & " and (exists(select np01 from nextprogress where np01=cp09" & _
'                                      " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & strNpSqlOfNoSalesDuty & ")" & _
'                             " or instr('P1913','PS1913','FCP1913','FG1913','CFP1913','CPS1913',cp01||cp10)>0)"
'
'      'Add By Sindy 2015/10/14 剔除結案中的
'      'Modify By Sindy 2017/7/25 8碼為結案電子表單編號 ==> and length(np24)=8
'      strSql = strSql & " and not exists(select np01 from nextprogress where np01=decode(cp01||cp10,'P1913',cp43,'PS1913',cp43,'FCP1913',cp43,'FG1913',cp43,'CFP1913',cp43,'CPS1913',cp43,cp09)" & _
'                                       " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & _
'                                       " and np24 is not null and length(np24)=8" & strNpSqlOfNoSalesDuty & ")"
   End If
   
   'Modified by Morgan 2020/8/3 ServicePractice移到下面
'   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
'        strTMSP = " Union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM
'   End If
   'Modified by Morgan 2022/8/12 +已判發LP05>0
   strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58,LP26,LP35" & _
            " from casepaperpdf,(" & _
            " Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,             pa05,             pa06,              pa07,             pa09,             pa26,             pa47,cp43,              pa58,cp30 from caseprogress,patent where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " Union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            "),casepropertymap,nation,LetterProgress,staff,Customer" & _
            " where cp09=cpp01(+) and cpp02 is not null" & _
            " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
            " and cp09=LP01(+) and LP06=st01(+) and LP05>0" & _
            " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)" & strNotPA
   'Modify By Sindy 2021/4/13
   'Modify By Sindy 2021/11/24 + IIf(Mid(Trim(Combo1.Text), 1, 4) = "MCTF" And Pub_StrUserSt03 = "F11", " and exists(select fa120 from fagent where fa01=substr(tm44,1,8) and fa02=substr(tm44,9,1) and fa120='" & Mid(Trim(Combo1.Text), 1, 4) & "')", "")
   'Modified by Morgan 2022/8/12 +已判發LP05>0
   strSql = strSql & " union " & _
            "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,decode(lp29||lp30,'','',decode(lp29,null,'',ltrim(to_char(lp29,'99,999,999'))||'('||lp30||')')) 報價,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58,LP26,LP35" & _
            " from casepaperpdf,(" & _
            " Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & IIf(Mid(Trim(Combo1.Text), 1, 4) = "MCTF" And Pub_StrUserSt03 = "F11", " and exists(select fa120 from fagent where fa01=substr(tm44,1,8) and fa02=substr(tm44,9,1) and fa120='" & Mid(Trim(Combo1.Text), 1, 4) & "')", "") & strCon & strTM & _
            " Union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & IIf(Mid(Trim(Combo1.Text), 1, 4) = "MCTF" And Pub_StrUserSt03 = "F11", " and exists(select fa120 from fagent where fa01=substr(sp26,1,8) and fa02=substr(sp26,9,1) and fa120='" & Mid(Trim(Combo1.Text), 1, 4) & "')", "") & strCon & strSP & _
            "),casepropertymap,nation,LetterProgress,staff,Customer" & _
            " where cp09=cpp01(+) and cpp02 is not null" & _
            " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
            " and cp09=LP01(+) and LP06=st01(+) and LP05>0" & _
            " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)" & strNotTM
   '2021/4/13 END
   'end 2020/8/3
   'end 2020/02/05
   strSql = strSql & " order by cp127 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
      LblCnt.Visible = True: LblCnt.Caption = "共 0 筆" 'Add By Sindy 2019/2/1
   'End If
   If rsTmp.RecordCount > 0 Then
      'If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
         LblCnt.Visible = True: LblCnt.Caption = "共 " & rsTmp.RecordCount & " 筆" 'Add By Sindy 2019/2/1
      'End If
      Set GRD1.Recordset = rsTmp
      Call RunLoopChkData 'Modify By Sindy 2014/8/6
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Me.Enabled = True
      Exit Sub
   End If
   rsTmp.Close
   
   GRD1.col = 0
   GRD1.row = 1
   
   If bolChgSystemkind = True Then systemkind = strOldSystemkind 'Add By Sindy 2020/9/4
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   
EXITSUB:
   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub

'Add By Sindy 2019/2/18 增加剔除的語法
Private Function AddCheck2Sql(ByRef AddCheck2TM As String) As String
   AddCheck2Sql = ""
   AddCheck2TM = "" 'Add By Sindy 2021/4/13
   If Check2.Value = 1 Then
      '剔除已收文
      'Modify By Sindy 2019/8/1 + and cp30=np22 ex:CFP-029571 年費(杜燕文)
      'Modify By Sindy 2019/8/2 + and cp30 is not null and cp30=np22 玲玲:P-118761
      AddCheck2Sql = " and not exists(select np01 from nextprogress where np01=decode(cp01||cp10,'P1913',cp43,'PS1913',cp43,'FCP1913',cp43,'FG1913',cp43,'CFP1913',cp43,'CPS1913',cp43,cp09)" & _
                                      " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & _
                                      " and cp30 is not null and cp30=np22 and np06 in('Y','N')" & strNpSqlOfNoSalesDuty & ")"
      'Modify By Sindy 2021/4/13
      AddCheck2TM = " and not exists(select np01 from nextprogress where np01=decode(substr(cp01,1,1)||cp10,'T1725',cp43,cp09)" & _
                                      " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & _
                                      " and cp30 is not null and cp30=np22 and np06 in('Y','N')" & strNpSqlOfNoSalesDuty & ")"
                                      
      'Modify By Sindy 2019/8/2 + 玲玲:P-118761 增加一段SQL + and cp30 is null
      AddCheck2Sql = AddCheck2Sql & " and not exists(select np01 from nextprogress where np01=decode(cp01||cp10,'P1913',cp43,'PS1913',cp43,'FCP1913',cp43,'FG1913',cp43,'CFP1913',cp43,'CPS1913',cp43,cp09)" & _
                                      " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & _
                                      " and cp30 is null and np06 in('Y','N')" & strNpSqlOfNoSalesDuty & ")"
      'Modify By Sindy 2021/4/13
      AddCheck2TM = AddCheck2TM & " and not exists(select np01 from nextprogress where np01=decode(substr(cp01,1,1)||cp10,'T1725',cp43,cp09)" & _
                                      " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & _
                                      " and cp30 is null and np06 in('Y','N')" & strNpSqlOfNoSalesDuty & ")"
      
      'Add By Sindy 2015/5/20 或用CP09串不到NP01,也是無期限資料 ex.P-103520
      AddCheck2Sql = AddCheck2Sql & " and (exists(select np01 from nextprogress where np01=cp09" & _
                                      " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & strNpSqlOfNoSalesDuty & ")" & _
                                      " or instr('P1913,PS1913,FCP1913,FG1913,CFP1913,CPS1913',cp01||cp10)>0)"
      'Modify By Sindy 2021/4/13
      AddCheck2TM = AddCheck2TM & " and (exists(select np01 from nextprogress where np01=cp09" & _
                                      " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & strNpSqlOfNoSalesDuty & ")" & _
                                      " or instr('T1725',substr(cp01,1,1)||cp10)>0)"
                                      
      'Add By Sindy 2015/10/14 剔除結案中的
      'Modify By Sindy 2017/7/25 8碼為結案電子表單編號 ==> and length(np24)=8
      'Modify By Sindy 2019/8/1 + and cp30=np22 ex:CFP-029571 年費(杜燕文)
      'Modify By Sindy 2019/8/2 + and cp30 is not null and cp30=np22 玲玲:P-118761
      AddCheck2Sql = AddCheck2Sql & " and not exists(select np01 from nextprogress where np01=decode(cp01||cp10,'P1913',cp43,'PS1913',cp43,'FCP1913',cp43,'FG1913',cp43,'CFP1913',cp43,'CPS1913',cp43,cp09)" & _
                                       " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & _
                                       " and cp30 is not null and cp30=np22 and np24 is not null and length(np24)=8" & strNpSqlOfNoSalesDuty & ")"
      'Modify By Sindy 2021/4/13
      AddCheck2TM = AddCheck2TM & " and not exists(select np01 from nextprogress where np01=decode(substr(cp01,1,1)||cp10,'T1725',cp43,cp09)" & _
                                       " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & _
                                       " and cp30 is not null and cp30=np22 and np24 is not null and length(np24)=8" & strNpSqlOfNoSalesDuty & ")"
                                       
      'Modify By Sindy 2019/8/2 + 玲玲:P-118761 增加一段SQL + and cp30 is null
      AddCheck2Sql = AddCheck2Sql & " and not exists(select np01 from nextprogress where np01=decode(cp01||cp10,'P1913',cp43,'PS1913',cp43,'FCP1913',cp43,'FG1913',cp43,'CFP1913',cp43,'CPS1913',cp43,cp09)" & _
                                       " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & _
                                       " and cp30 is null and np24 is not null and length(np24)=8" & strNpSqlOfNoSalesDuty & ")"
      'Modify By Sindy 2021/4/13
      AddCheck2TM = AddCheck2TM & " and not exists(select np01 from nextprogress where np01=decode(substr(cp01,1,1)||cp10,'T1725',cp43,cp09)" & _
                                       " and np01 is not null and np02=cp01 and np03=cp02 and np04=cp03 and np05=cp04" & _
                                       " and cp30 is null and np24 is not null and length(np24)=8" & strNpSqlOfNoSalesDuty & ")"
   End If
End Function

'Added by Morgan 2019/3/4
Private Sub lblAttCnt_Click()
   SendMessage cboAtt.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
End Sub

'Add By Sindy 2014/5/22
Private Sub Check1_Click()
   If Check1.Value = 1 Then
      WebBrowser1.Navigate "about:blank"
      'Modified by Mogrgan 2019/3/4 加pdf檔下拉選單
      'Command4.Visible = False
      Frame4.Visible = False
      'end 2019/3/4
      WebBrowser1.Visible = False
      'Modified by Moran 2015/9/1 為避免加欄位還要調寬度改變設定方法
      'GRD1.Width = GrdMaxW
      GRD1.Width = Me.Width - 200
      'end 2015/9/1
      Call SetGrd(False)
   Else
      'Modified by Mogrgan 2019/3/4 加pdf檔下拉選單
      'Command4.Visible = True
      Frame4.Visible = True
      'end 2019/3/4
      WebBrowser1.Visible = True
      GRD1.Width = GrdMinW
      Call SetGrd(False)
   End If
End Sub

'Add By Sindy 2014/8/6
Private Sub Check2_Click()
   Call cmdok1_Click(1)
End Sub

'Added by Morgan 2014/12/24
Private Sub cmdEmail_Click()
   cmdOpenAtt_Click 99
End Sub

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index '紀錄作用按鍵
   PubShowNextData
   Exit Sub
End Sub

Public Sub PubShowNextData()
Dim StrTag As String
Dim i As Integer, j As Integer
Dim bolSelisV As Boolean 'Add By Sindy 2014/8/4
Dim bolCancel As Boolean
Dim nFrm As Form 'Added by Lydia 2021/04/13
Dim oRunform As Form 'Add By Sindy 2022/9/16
   
   'Add By Sindy 2022/9/16
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      Set oRunform = frm090801_New
   Else
      Set oRunform = frm090801
   End If
   '2022/9/16 END
   
   Select Case cmdState
      Case 0 '進度
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            bolSelisV = False 'Add By Sindy 2014/8/4
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then 'And bolSelIsUseButton = False Then
               bolSelisV = True 'Add By Sindy 2014/8/4
               GRD1.col = 0
               GRD1.Text = ""
'               For j = 1 To GRD1.Cols - 1
'                  GRD1.col = j
'                  GRD1.CellBackColor = QBColor(15)
'               Next j
               GRD1.col = 2
               If Not IsNull(GRD1.Text) Then
                  If fnSaveParentForm(Me) = False Then
                     Me.Enabled = True
                     Exit Sub
                  End If
                  Screen.MousePointer = vbHourglass
                  StrTag = Pub_RplStr(GRD1.Text)
                  If UBound(Split(StrTag, "-")) = 1 Then
                     StrTag = StrTag & "-0-00"
                  End If
                  frm100101_2.Show
                  frm100101_2.Tag = StrTag
                  frm100101_2.cmdOK(6).Visible = False 'Add By Sindy 2014/6/17 結束按鈕隱藏
                  frm100101_2.StrMenu
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
            End If
         Next i
         Me.Enabled = True
      Case 1 '查詢
         'Modify By Sindy 2015/6/24
         Call Text1_LostFocus(3)
         If Not (Text1(0) <> "" And Text1(1) <> "") Then
            If Text4.Text = "" Then
               MsgBox "發文起始日不可空白！"
               Text4.SetFocus
               Exit Sub
            End If
            If Text5.Text = "" Then
               MsgBox "發文截止日不可空白！"
               Text5.SetFocus
               Exit Sub
            End If
            If Val(Text4.Text) > Val(Text5.Text) Then
               MsgBox "發文起始日不可大於發文截止日！"
               Text4.SetFocus
               Exit Sub
            End If
         End If
         '2015/6/24 END
         '寄件查詢
         If intWorkItem = 1 Then
            If Trim(Combo1.Text) = "" Then
               MsgBox "智權人員不可空白！"
               Combo1.SetFocus
               Exit Sub
            End If
            
            'Add By Sindy 2014/8/6
            bolCancel = False
            Call Combo1_Validate(bolCancel)
            If bolCancel = True Then Exit Sub
            '電腦中心才可輸入其他人員做查詢
            If Pub_StrUserSt03 <> "M51" Then
               bolCancel = False
               For i = 0 To Combo1.ListCount - 1
                  If Trim(Combo1.List(i)) = Trim(Combo1.Text) Then
                     bolCancel = True
                     Exit For
                  End If
               Next i
               If bolCancel = False Then
                  MsgBox "無權限查詢該人員資料！", vbExclamation
                  Combo1.SetFocus
                  Exit Sub
               End If
            End If
            '2014/8/6 END
            
            Call QueryData2
         Else
            If CmdOk1(2).Enabled = False Then
               Call StrMenu1
            Else
               Call QueryData(True)
            End If
         End If
      Case 2 '關係企業
         Call StrMenu(Str01)
         Call StrMenu1
         CmdOk1(2).Enabled = False
      Case 3 '結束
         If Val(intWorkItem) = 1 Then
            Unload Me
         Else
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
         End If
      Case 4 '卷宗區
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            bolSelisV = False 'Add By Sindy 2014/8/4
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then 'And bolSelIsUseButton = False Then
               bolSelisV = True 'Add By Sindy 2014/8/4
               GRD1.col = 0
'               GRD1.Text = ""
'               For j = 1 To GRD1.Cols - 1
'                  GRD1.col = j
'                  GRD1.CellBackColor = QBColor(15)
'               Next j
               GRD1.col = 2
               If Not IsNull(GRD1.Text) Then
'                     If fnSaveParentForm(Me) = False Then
'                         Me.Enabled = True
'                         Exit Sub
'                     End If
                  Screen.MousePointer = vbHourglass
                  StrTag = Pub_RplStr(GRD1.Text)
                  If UBound(Split(StrTag, "-")) = 1 Then
                     StrTag = StrTag & "-0-00"
                  End If
                  frm100101_L.m_strKey = StrTag
                  'frm100101_L.Hide
                  frm100101_L.SetParent Me
                  If frm100101_L.QueryData = True Then
                     frm100101_L.Show
                     Me.Hide
                  Else
                     Unload frm100101_L
                  End If
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
            End If
         Next i
         Me.Enabled = True
      Case 5 '收文
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            bolSelisV = False 'Add By Sindy 2014/8/4
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then 'And bolSelIsUseButton = False Then
               bolSelisV = True 'Add By Sindy 2014/8/4
               GRD1.col = 0
'               GRD1.Text = ""
'               For j = 1 To GRD1.Cols - 1
'                  GRD1.col = j
'                  GRD1.CellBackColor = QBColor(15)
'               Next j
               GRD1.col = 2
               If Not IsNull(GRD1.Text) Then
'                  If fnSaveParentForm(Me) = False Then
'                     Me.Enabled = True
'                     Exit Sub
'                  End If
                  Screen.MousePointer = vbHourglass
                  StrTag = Pub_RplStr(GRD1.Text)
                  If UBound(Split(StrTag, "-")) = 1 Then
                     StrTag = StrTag & "-0-00"
                  End If
                  'Modify By Sindy 2022/9/16 frm090801 改用 oRunform
                  'Added by Lydia 2021/04/13 查名單輸入的設定; (From 林青祺) 國內接洽單收文T台灣申請案，沒有顯示「查名單輸入」按鈕。
                  Set nFrm = Forms(0).GetForm("frm090126")
                  If Not nFrm Is Nothing Then
                     Set oRunform.Tmpfrm090126 = nFrm
                  End If
                  'end 2021/04/13
                  oRunform.bolExternalCall = True '記錄是外部程式呼叫使用
                  oRunform.Show
                  oRunform.Tag = StrTag
                  oRunform.Option1(1).Value = True
                  oRunform.Text1(6) = SystemNumber(StrTag, 1)
                  oRunform.Text1(7) = SystemNumber(StrTag, 2)
                  oRunform.Text1(8) = SystemNumber(StrTag, 3)
                  oRunform.Text1(9) = SystemNumber(StrTag, 4)
                  oRunform.Text1_LostFocus (9)
                  oRunform.bolExternalCall = False '還原預設值
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
            End If
         Next i
         Me.Enabled = True
      'Add By Sindy 2014/6/19
      Case 6 '結案
         Me.Enabled = False
         For i = 1 To GRD1.Rows - 1
            bolSelisV = False 'Add By Sindy 2014/8/4
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then 'And bolSelIsUseButton = False Then
               bolSelisV = True 'Add By Sindy 2014/8/4
               GRD1.col = 0
               GRD1.Text = ""
'               For j = 1 To GRD1.Cols - 1
'                  GRD1.col = j
'                  GRD1.CellBackColor = QBColor(15)
'               Next j
               GRD1.col = 2
               If Not IsNull(GRD1.Text) Then
'                     If fnSaveParentForm(Me) = False Then
'                         Me.Enabled = True
'                         Exit Sub
'                     End If
                  Screen.MousePointer = vbHourglass
                  StrTag = Pub_RplStr(GRD1.Text)
                  If UBound(Split(StrTag, "-")) = 1 Then
                     StrTag = StrTag & "-0-00"
                  End If
                  frm210133.Txt1(0) = SystemNumber(StrTag, 1)
                  frm210133.Txt1(1) = SystemNumber(StrTag, 2)
                  frm210133.Txt1(2) = SystemNumber(StrTag, 3)
                  frm210133.Txt1(3) = SystemNumber(StrTag, 4)
                  
                  frm210133.m_NP01 = GRD1.TextMatrix(i, 15) 'Add By Sindy 2015/1/20
                  frm210133.m_NP22 = "" 'Add By Sindy 2015/1/20
                  frm210133.SetParent Me
                  If frm210133.doQuery = True Then
                     frm210133.Show
                     Me.Hide
                  Else
                     Unload frm210133
                  End If
                  Screen.MousePointer = vbDefault
                  Me.Enabled = True
                  Exit Sub
               End If
            End If
         Next i
         Me.Enabled = True
      '2014/6/19 END
   End Select
   'Add By Sindy 2014/8/4
   Select Case cmdState
      Case 0, 6
         If bolSelisV = False Then
            Me.Enabled = False
            For i = 1 To GRD1.Rows - 1
               GRD1.col = 1
               GRD1.row = i
               If GRD1.CellBackColor = &HFFC0C0 Then
                  GRD1.TextMatrix(i, 0) = "V"
                  Exit For
               End If
            Next i
            Me.Enabled = True
         End If
   End Select
   '2014/8/4 END
End Sub

'列出關係企業
Sub StrMenu(StrToGrid)
   Screen.MousePointer = vbHourglass
   cnnConnection.Execute "DELETE FROM R100102 where id='" & strUserNum & "' "
   '已申請人查詢之資料庫
   '以編號 LIKE
   strSql = "SELECT CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●',''),NVL(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,CU80,CU79 FROM CUSTOMER,NATION WHERE CU10=NA01(+) AND CU01>='" & Left(StrToGrid, 6) & "00' AND CU01<='" & Left(StrToGrid, 6) & "zz'"
   '傳入R1時找出相關的X
   strSql = strSql & " union SELECT CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●',''),NVL(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,CU80,CU79 " & _
                              "From CUSTOMER, PotCustomer1, Nation " & _
                             "WHERE CU10=NA01(+) " & _
                               "AND POC01>='" & Left(StrToGrid, 6) & "00' AND POC01<='" & Left(StrToGrid, 6) & "zz' " & _
                               "AND CU01>=(substr(POC16,1,6)||'00') AND CU01<=(substr(POC16,1,6)||'zz') " & _
                               "AND POC16 is not null"
   '傳入R時找出相關的X
   strSql = strSql & " union SELECT CU01||CU02||Decode(CU02,'0','','＊')||Decode(cu111,'Y','$','')||Decode(cu121,'Y','●',''),NVL(CU04,Decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NA03,CU80,CU79 " & _
                              "From CUSTOMER, PotCustomer, Nation " & _
                             "WHERE CU10=NA01(+) " & _
                               "AND PCU01>='" & Left(StrToGrid, 6) & "00' AND PCU01<='" & Left(StrToGrid, 6) & "zz' " & _
                               "AND CU01>=(substr(PCU47,1,6)||'00') AND CU01<=(substr(PCU47,1,6)||'zz') " & _
                               "AND PCU47 is not null"
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount <> 0 Then
       adoRecordset.MoveFirst
       Do While adoRecordset.EOF = False
       strSql = "INSERT INTO R100102 values ('"
       If Not IsNull(adoRecordset.Fields(0)) Then
           strSql = strSql + ChgSQL(CheckStr(adoRecordset.Fields(0))) + "','"
       Else
           strSql = strSql + "','"
       End If
       If Not IsNull(adoRecordset.Fields(1)) Then
           strSql = strSql + ChgSQL(CheckStr(adoRecordset.Fields(1))) + "','"
       Else
           strSql = strSql + "','"
       End If
       If Not IsNull(adoRecordset.Fields(2)) Then
           strSql = strSql + ChgSQL(CheckStr(adoRecordset.Fields(2))) + "','" & strUserNum & "')"
       Else
           strSql = strSql + "','" & strUserNum & "')"
       End If
       cnnConnection.Execute strSql
       adoRecordset.MoveNext
       Loop
   Else
       ShowNoData
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   CheckOC
   Screen.MousePointer = vbDefault
End Sub

'關係企業案件資料
Sub StrMenu1()
Dim strCon As String, strMainCon As String
'Dim strTMSP As String
Dim strNotTM As String, strNotPA As String
Dim strP As String, strTM As String, strSP As String 'Add by Amy 2020/02/05
   
   'Add By Sindy 2020/9/4
   bolChgSystemkind = False
   If Text1(0) = "" Or Text1(1) = "" Then
      If systemkind = "" Then
         systemkind = "ALL"
      End If
   Else
      bolChgSystemkind = True
      strOldSystemkind = systemkind
      systemkind = Text1(0).Text
   End If
   '2020/9/4 END
   
   strCon = "": strMainCon = ""
   'Add By Sindy 2015/1/12 不列出A,B類
   If Check2.Value = 1 Then '剔除已收文,結案,閉卷
      'Modified by Morgan 2017/1/17 總收文號+D
      'strCon = strCon & " and substr(cp09,1,1)='C'"
      strCon = strCon & " and substr(cp09,1,1)>='C'"
      'end 2017/1/17
      
      'Add By Sindy 2019/2/1 為使後續逐筆過濾資料會導致速度慢,儘量在SQL裡把資料過濾掉
      '剔除無期限(本所期限),抓有期限的
      strCon = strCon & " and cp06 is not null"
      '剔除已閉卷,抓未閉卷的
      'Modify by Amy 2020/02/05 原:strCon
      strP = strP & " and pa58 is null"
      strTM = strTM & " and tm30 is null"
      strSP = strSP & " and sp16 is null"
      'end 2020/02/05
      '2019/2/1 END
   End If
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   GRD1.Clear
   SetGrd
   m_mouseRow = 0 'Add By Sindy 2014/11/21
   'Modify By Sindy 2014/8/6 +,pa47 as 分所案號 及 ,pa47
   'Modify By Sindy 2014/8/6 +,cp43,cp01,cp02,cp03,cp04
   'Modify By Sindy 2014/10/2 +,pa58
   'Modify By Sindy 2014/10/9 +因案件會有多申請人狀況,但同文不可重覆出現,所以資料讀取出來再逐筆查詢申請人
'   strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04) as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58" & _
'            " from casepaperpdf,(" & _
'            "Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58 from caseprogress,patent,R100102 where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=substr(pa26,1,8) AND SUBSTR(R06001,9,1)=substr(pa26,9,1)" & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa27,pa47,cp43,pa58 from caseprogress,patent,R100102 where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=substr(pa27,1,8) AND SUBSTR(R06001,9,1)=substr(pa27,9,1)" & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa28,pa47,cp43,pa58 from caseprogress,patent,R100102 where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=substr(pa28,1,8) AND SUBSTR(R06001,9,1)=substr(pa28,9,1)" & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa29,pa47,cp43,pa58 from caseprogress,patent,R100102 where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=substr(pa29,1,8) AND SUBSTR(R06001,9,1)=substr(pa29,9,1)" & _
'            " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa30,pa47,cp43,pa58 from caseprogress,patent,R100102 where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=substr(pa30,1,8) AND SUBSTR(R06001,9,1)=substr(pa30,9,1)" & _
'            "),casepropertymap,nation,customer,LetterProgress,staff" & _
'            " where cp09=cpp01(+) and cpp02 is not null" & _
'            " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+)" & _
'            " and cp09=LP01(+) and LP06=st01(+)" & _
'            " order by cp127 desc"
   'Modify By Sindy 2015/6/24 以本所案號查詢
   'Modify By Sindy 2016/8/5 原:"Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,cp43 from caseprogress where cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "' and cp127=(select max(cp127) from caseprogress where cp127>0 and cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "'" & strCon & ")"
   '                         加:日期做條件
   'Modify by Amy 2020/02/05 Mark原程式 +Trademark,ServicePractice,整理成一句
'   If Text1(0) <> "" And Text1(1) <> "" Then
'      'Modify By Sindy 2016/8/5
'      If Text4 <> "" Then '有輸入發文日期
'         strCon = strCon & " and cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5)
'      Else
'         strCon = strCon & " and cp127>0"
'      End If
'      '2016/8/5 END
'      'Modify By Sindy 2019/8/1 + ,cp30
'      strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58" & _
'               " from casepaperpdf,patent,R100102,(" & _
'               "Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,cp43,cp30 from caseprogress where cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "'" & strCon & _
'               "),casepropertymap,nation,LetterProgress,staff,Customer" & _
'               " where cp09=cpp01(+) and cpp02 is not null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X'" & _
'               " and ((SUBSTR(R06001,1,8)=substr(pa26,1,8) AND SUBSTR(R06001,9,1)=substr(pa26,9,1)) or (SUBSTR(R06001,1,8)=substr(pa27,1,8) AND SUBSTR(R06001,9,1)=substr(pa27,9,1)) or (SUBSTR(R06001,1,8)=substr(pa28,1,8) AND SUBSTR(R06001,9,1)=substr(pa28,9,1)) or (SUBSTR(R06001,1,8)=substr(pa29,1,8) AND SUBSTR(R06001,9,1)=substr(pa29,9,1)) or (SUBSTR(R06001,1,8)=substr(pa30,1,8) AND SUBSTR(R06001,9,1)=substr(pa30,9,1)))" & _
'               " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
'               " and cp09=LP01(+) and LP06=st01(+)" & _
'               " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)"
'   Else
'   '2015/6/24 END
'      'Modify By Sindy 2019/8/1 + ,cp30
'      strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58" & _
'               " from casepaperpdf,(" & _
'               "Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent,R100102 where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=substr(pa26,1,8) AND SUBSTR(R06001,9,1)=substr(pa26,9,1)" & strCon & _
'               " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent,R100102 where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=substr(pa27,1,8) AND SUBSTR(R06001,9,1)=substr(pa27,9,1)" & strCon & _
'               " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent,R100102 where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=substr(pa28,1,8) AND SUBSTR(R06001,9,1)=substr(pa28,9,1)" & strCon & _
'               " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent,R100102 where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=substr(pa29,1,8) AND SUBSTR(R06001,9,1)=substr(pa29,9,1)" & strCon & _
'               " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent,R100102 where cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' AND SUBSTR(R06001,1,8)=substr(pa30,1,8) AND SUBSTR(R06001,9,1)=substr(pa30,9,1)" & strCon & _
'               "),casepropertymap,nation,LetterProgress,staff,Customer" & _
'               " where cp09=cpp01(+) and cpp02 is not null" & _
'               " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
'               " and cp09=LP01(+) and LP06=st01(+)" & _
'               " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)"
'   End If
   strCon = strCon & " and id='" & strUserNum & "' AND SUBSTR(R06001,1,1)='X' "
   '案號
   If Text1(0) <> MsgText(601) And Text1(1) <> MsgText(601) Then
        strCon = strCon & " And cp01='" & Text1(0) & "' and cp02='" & Text1(1) & "' and cp03='" & Text1(2) & "' and cp04='" & Text1(3) & "'"
   End If
   If Text4 <> "" Then '有輸入發文日期
        strCon = strCon & " and cp127 between " & DBDATE(Text4) & " and " & DBDATE(Text5)
   ElseIf Text1(0) <> MsgText(601) And Text1(1) <> MsgText(601) And Text4 = MsgText(601) Then
        strCon = strCon & " and cp127>0"
   End If
   
   If Check2.Value = 1 Then
      '增加剔除的語法
      strNotPA = AddCheck2Sql(strNotTM)
   End If
   
   'Modified by Morgan 2020/8/3 ServicePractice移到下面
'   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
'        strTMSP = " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark,R100102 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(R06001,1,8)=substr(tm23,1,8) AND SUBSTR(R06001,9,1)=substr(tm23,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
'                         " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark,R100102 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(R06001,1,8)=substr(tm78,1,8) AND SUBSTR(R06001,9,1)=substr(tm78,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
'                         " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark,R100102 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(R06001,1,8)=substr(tm79,1,8) AND SUBSTR(R06001,9,1)=substr(tm79,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
'                         " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark,R100102 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(R06001,1,8)=substr(tm80,1,8) AND SUBSTR(R06001,9,1)=substr(tm80,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
'                         " union Select cp01,cp02,cp03,cp04,cp06,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark,R100102 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(R06001,1,8)=substr(tm81,1,8) AND SUBSTR(R06001,9,1)=substr(tm81,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM
'   End If
   strSql = "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58" & _
            " from casepaperpdf,(" & _
            "Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent,R100102 where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND SUBSTR(R06001,1,8)=substr(pa26,1,8) AND SUBSTR(R06001,9,1)=substr(pa26,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent,R100102 where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND SUBSTR(R06001,1,8)=substr(pa27,1,8) AND SUBSTR(R06001,9,1)=substr(pa27,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent,R100102 where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND SUBSTR(R06001,1,8)=substr(pa28,1,8) AND SUBSTR(R06001,9,1)=substr(pa28,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent,R100102 where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND SUBSTR(R06001,1,8)=substr(pa29,1,8) AND SUBSTR(R06001,9,1)=substr(pa29,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,pa05,pa06,pa07,pa09,pa26,pa47,cp43,pa58,cp30 from caseprogress,patent,R100102 where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) AND SUBSTR(R06001,1,8)=substr(pa30,1,8) AND SUBSTR(R06001,9,1)=substr(pa30,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 1) & ")" & strCon & strP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice,R100102 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(R06001,1,8)=substr(sp08,1,8) AND SUBSTR(R06001,9,1)=substr(sp08,9,1) and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice,R100102 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(R06001,1,8)=substr(sp58,1,8) AND SUBSTR(R06001,9,1)=substr(sp58,9,1) and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice,R100102 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(R06001,1,8)=substr(sp59,1,8) AND SUBSTR(R06001,9,1)=substr(sp59,9,1) and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice,R100102 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(R06001,1,8)=substr(sp65,1,8) AND SUBSTR(R06001,9,1)=substr(sp65,9,1) and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice,R100102 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(R06001,1,8)=substr(sp66,1,8) AND SUBSTR(R06001,9,1)=substr(sp66,9,1) and cp01 in ('PS','FG','CPS')" & strCon & strSP & _
            "),casepropertymap,nation,LetterProgress,staff,Customer" & _
            " where cp09=cpp01(+) and cpp02 is not null" & _
            " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
            " and cp09=LP01(+) and LP06=st01(+)" & _
            " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)" & strNotPA
   'Modify By Sindy 2021/4/13
   strSql = strSql & " union " & _
            "Select distinct '' as V,sqldatet(cp127) as 發文日,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) as 本所案號,decode(pa09,'000',CPM03,CPM04)||getrelatecasepropertyname(cp09,'1') as 案件性質,NVL(NVL(PA05,PA06),PA07) AS 案件名稱,na03 國家,Decode(CU04, Null, Decode(CU05, Null, CU06, CU05|| ' '||CU88||' '||CU89||' '||CU90), CU04) AS 申請人,sqldatet(cp06) as 本所期限,decode(LP11,'Y','直寄','0','親送','1','寄送','2','不寄',LP11) as 寄件方式,pa47 as 分所案號,LP12 as 備註,st02 as 確認人員,pa09,pa26,cp09,cp10,cp127,cp43,cp01,cp02,cp03,cp04,pa58" & _
            " from casepaperpdf,(" & _
            "Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark,R100102 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(R06001,1,8)=substr(tm23,1,8) AND SUBSTR(R06001,9,1)=substr(tm23,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark,R100102 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(R06001,1,8)=substr(tm78,1,8) AND SUBSTR(R06001,9,1)=substr(tm78,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark,R100102 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(R06001,1,8)=substr(tm79,1,8) AND SUBSTR(R06001,9,1)=substr(tm79,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark,R100102 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(R06001,1,8)=substr(tm80,1,8) AND SUBSTR(R06001,9,1)=substr(tm80,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm10 as pa09,tm23 as pa26,tm34 as pa47,cp43,tm30 as pa58,cp30 from caseprogress,TradeMark,R100102 where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) AND SUBSTR(R06001,1,8)=substr(tm81,1,8) AND SUBSTR(R06001,9,1)=substr(tm81,9,1) and cp01 in (" & SQLGrpStr(GetAllSysKind(systemkind), 2) & ")" & strCon & strTM & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice,R100102 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(R06001,1,8)=substr(sp08,1,8) AND SUBSTR(R06001,9,1)=substr(sp08,9,1) and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice,R100102 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(R06001,1,8)=substr(sp58,1,8) AND SUBSTR(R06001,9,1)=substr(sp58,9,1) and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice,R100102 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(R06001,1,8)=substr(sp59,1,8) AND SUBSTR(R06001,9,1)=substr(sp59,9,1) and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice,R100102 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(R06001,1,8)=substr(sp65,1,8) AND SUBSTR(R06001,9,1)=substr(sp65,9,1) and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & strCon & strSP & _
            " union Select cp01,cp02,cp03,cp04,cp06,cp07,cp09,cp10,cp127,sp05 as pa05,sp06  as pa06,sp07  as pa07,sp09 as pa09,sp08 as pa26,sp28  as pa47,cp43,sp16  as pa58,cp30 from caseprogress,ServicePractice,R100102 where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) AND SUBSTR(R06001,1,8)=substr(sp66,1,8) AND SUBSTR(R06001,9,1)=substr(sp66,9,1) and cp01 in ('TT','TS','TR','TM','TD','TC','TB','S','CFC')" & strCon & strSP & _
            "),casepropertymap,nation,LetterProgress,staff,Customer" & _
            " where cp09=cpp01(+) and cpp02 is not null" & _
            " and cp01=cpm01(+) and cp10=cpm02(+) and pa09=na01(+)" & _
            " and cp09=LP01(+) and LP06=st01(+)" & _
            " and CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1)" & strNotTM
   '2021/4/13 END
   'end 2020/8/3
   'end 2020/02/05
   strSql = strSql & " order by cp127 desc"
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
      LblCnt.Visible = True: LblCnt.Caption = "共 0 筆"  'Add By Sindy 2019/2/1
   'End If
   If adoRecordset.RecordCount <> 0 Then
      'If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
         LblCnt.Visible = True: LblCnt.Caption = "共 " & adoRecordset.RecordCount & " 筆"  'Add By Sindy 2019/2/1
      'End If
      Set GRD1.Recordset = adoRecordset
      Call RunLoopChkData 'Modify By Sindy 2014/8/6
   End If
   CheckOC
   GRD1.col = 0
   GRD1.row = 1
   
   If bolChgSystemkind = True Then systemkind = strOldSystemkind 'Add By Sindy 2020/9/4
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub

'Add By Sindy 2014/8/6
'逐筆檢查資料
Private Sub RunLoopChkData()
Dim iRow As Integer
Dim strKey As String
Dim strCP01 As String, strType As String 'Add By Sindy 2015/5/20
'Added by Morgan 2015/6/29
Dim idxLP26 As Integer
Dim idxColor As Integer
   
   idxLP26 = GetFieldId("LP26", GRD1) 'Added by Morgan 2015/6/29
   
   'Add By Sindy 2014/7/22 --經理:案件性質+相關總收文號的案件性質
   For iRow = 1 To GRD1.Rows - 1
      'Added by Morgan 2015/6/29
      With GRD1
      If .TextMatrix(iRow, idxLP26) <> "" Then
         'E化
         If .TextMatrix(iRow, idxLP26) = "Y" Then
            idxColor = 0
         '指定E化
         Else
            idxColor = 1
         End If
         .row = iRow
         .col = 2
         .CellBackColor = lblColor(idxColor).BackColor
         lblColor(idxColor).Visible = True
         lblColorDesc(idxColor).Visible = True
      End If
      End With
      'end 2015/6/29
      
      '加相關總收文號的案件性質
      'Modify By Sindy 2019/1/29 Mark,改用Sql抓相關案件性質
'      GRD1.TextMatrix(iRow, 3) = GRD1.TextMatrix(iRow, 3) & PUB_GetRelateCasePropertyName(GRD1.TextMatrix(iRow, 15), "1")
      
      'Modify By Sindy 2018/8/16 Mark:改在前頭組SQL語法裡抓申請人名稱,提升速度
      '申請人
'      GRD1.TextMatrix(iRow, 6) = PUB_GetCustName(GRD1.TextMatrix(iRow, 19) & GRD1.TextMatrix(iRow, 20) & GRD1.TextMatrix(iRow, 21) & GRD1.TextMatrix(iRow, 22), True)
      '2018/8/16 END
      
      'Modify By Sindy 2019/2/18 把逐筆剔除改到查詢SQL裡判斷
'      If Check2.Value = 1 Then '剔除已收文,結案,閉卷
'         strCP01 = GRD1.TextMatrix(iRow, 19)
'         If Trim(GRD1.TextMatrix(iRow, 7)) <> "" Then '有本所期限者才檢查
'            '案件性質為1913通知期限,用CP43抓NP檢查,其他用CP09抓NP檢查
'            If (strCP01 = "P" Or strCP01 = "PS" Or strCP01 = "FCP" Or strCP01 = "FG" Or strCP01 = "CFP" Or strCP01 = "CPS") _
'               And Val(GRD1.TextMatrix(iRow, 16)) = "1913" Then
'               strKey = GRD1.TextMatrix(iRow, 18)
'               strType = "CP43"
'            Else
'               strKey = GRD1.TextMatrix(iRow, 15)
'               strType = "CP09"
'            End If
'            '剔除程序管制的案件性質(strNpSqlOfNoSalesDuty:下一程序非智權人員掌控之案件性質)
'            strSql = "select np01 from nextprogress where np01='" & strKey & "'" & _
'                     " and np02='" & GRD1.TextMatrix(iRow, 19) & "'" & _
'                     " and np03='" & GRD1.TextMatrix(iRow, 20) & "'" & _
'                     " and np04='" & GRD1.TextMatrix(iRow, 21) & "'" & _
'                     " and np05='" & GRD1.TextMatrix(iRow, 22) & "'" & _
'                     " and np06 in('Y','N')" & strNpSqlOfNoSalesDuty
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'            '串到下一程序已收文
'            If intI = 1 Then
'               GRD1.RowHeight(iRow) = 0
'            Else
'               'Add By Sindy 2014/10/2 檢查是否已閉卷
'               If Val(Trim(GRD1.TextMatrix(iRow, 23))) > 0 Then
'                  GRD1.RowHeight(iRow) = 0
'               'Add By Sindy 2015/5/20 或用CP09串不到NP01,也是無期限資料 ex.P-103520
'               Else
'                  If strType = "CP09" Then
'                     strSql = "select np01 from nextprogress where np01='" & strKey & "'" & _
'                              " and np02='" & GRD1.TextMatrix(iRow, 19) & "'" & _
'                              " and np03='" & GRD1.TextMatrix(iRow, 20) & "'" & _
'                              " and np04='" & GRD1.TextMatrix(iRow, 21) & "'" & _
'                              " and np05='" & GRD1.TextMatrix(iRow, 22) & "'" & strNpSqlOfNoSalesDuty
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                     If intI = 0 Then
'                        GRD1.RowHeight(iRow) = 0
'                     End If
'                  End If
'               '2015/5/20 END
'               End If
'               '2014/10/2 END
'            End If
'
'            'Add By Sindy 2015/10/14 剔除結案中的
'            'Modify By Sindy 2017/7/25 8碼為結案電子表單編號 ==> and length(np24)=8
'            If GRD1.RowHeight(iRow) > 0 Then
'               strSql = "select np01 from nextprogress where np01='" & strKey & "'" & _
'                        " and np02='" & GRD1.TextMatrix(iRow, 19) & "'" & _
'                        " and np03='" & GRD1.TextMatrix(iRow, 20) & "'" & _
'                        " and np04='" & GRD1.TextMatrix(iRow, 21) & "'" & _
'                        " and np05='" & GRD1.TextMatrix(iRow, 22) & "'" & _
'                        " and np24 is not null and length(np24)=8" & strNpSqlOfNoSalesDuty
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'               '串到下一程序結案中
'               If intI = 1 Then
'                  GRD1.RowHeight(iRow) = 0
'               End If
'            End If
'            '2015/10/14 END
'
'         'Add By Sindy 2015/4/14
'         '剔除無期限
'         Else
'            GRD1.RowHeight(iRow) = 0
'         '2015/4/14 END
'         End If
'      End If
   Next
   '2014/7/22 END
End Sub

'Add By Sindy 2014/8/6
Private Sub Combo1_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo1_LostFocus()
   If Combo1 <> "" And Trim(Combo1) <> "全部" Then
      Combo1 = Trim(Left(Combo1, 6)) & " " & GetPrjSalesNM(Trim(Left(Combo1, 6)))
   End If
End Sub
Private Sub Combo1_Validate(Cancel As Boolean)
Dim strEmp As String
Dim arrID                'add by sonia 2016/7/1
   
   If Combo1 <> "" And Trim(Combo1) <> "全部" Then
      'modify by sonia 2016/7/1 因S29故改寫法
      'strEmp = GetStaffName(Trim(Left(Combo1, 6)))
      arrID = Split(Combo1.Text, " ")
      strEmp = GetStaffName(arrID(0), True)
      'end 2016/7/1
      If strEmp = "" Then
         MsgBox "智權人員輸入錯誤！", vbCritical
         Combo1.SetFocus
         Cancel = True
      End If
   End If
End Sub
'2014/8/6 END

'Add By Sindy 2014/6/9
Private Sub Combo2_Change()
   If Val(Combo2.Text) > 0 Then
      Text4.Text = Val(Format(DateAdd("M", -1 * Val(Combo2.Text), ChangeWStringToWDateString(strSrvDate(1))), "YYYYMMDD")) - 19110000
      Text5.Text = strSrvDate(2)
      'Call cmdok1_Click(1) '查詢
   End If
End Sub
Private Sub Combo2_Click()
   If Val(Combo2.Text) > 0 Then
      Text4.Text = Val(Format(DateAdd("M", -1 * Val(Combo2.Text), ChangeWStringToWDateString(strSrvDate(1))), "YYYYMMDD")) - 19110000
      Text5.Text = strSrvDate(2)
      'Call cmdok1_Click(1) '查詢
   End If
End Sub
'2014/6/9 END

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
      'Modified by Mogrgan 2019/3/4 加pdf檔下拉選單
      If lngLeft = 0 Then lngLeft = WebBrowser1.Left
      If pFull = True Then
         WebBrowser1.Left = 0
         Command4.Caption = "點我還原"
      Else
         WebBrowser1.Left = lngLeft
         Command4.Caption = "點我展開"
      End If
      WebBrowser1.Width = Me.Width - WebBrowser1.Left - 90
      WebBrowser1.Height = Me.Height - Frame4.Height - 390
      Frame4.Left = WebBrowser1.Left
      Frame4.Width = WebBrowser1.Width
      Command4.Width = Frame4.Width
      cboAtt.Width = Frame4.Width - cboAtt.Left
      'end 2019/3/4
      
      GRD1.Height = Me.Height - GRD1.Top - Frame1.Height - 300
      If Check1.Value = 1 Then
         GRD1.Width = Me.Width - 150
      Else
         GRD1.Width = GrdMinW
      End If
      Frame1.Top = GRD1.Top + GRD1.Height - 0 '50
   End If
End Sub

Private Sub SetCombo1()
   Dim ii As Integer
   Dim strMCTF As String 'Add by Amy 2020/02/05
   Dim stDef As String 'Add by Sindy 2020/11/20 預設
   
   'Modify By Sindy 2022/5/25 設定屬智權人員作業的下拉選單(共用模組)
   'Modify by Amy 2023/02/10 +Me.Name
   Call PUB_SetCombo1Sales(Combo1, , Me.Name)
   
'   Combo1.Clear
'   Combo1.AddItem strUserNum & " " & strUserName
   
   'Add by Amy 2020/02/05 P2部門開頭,若為MCTF人員或MCTF主管操作時,選單增加 MCTF01....
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
        strMCTF = Pub_GetSpecMan("MCTM", False)
        'Modify by Sindy 2021/11/24 + Or Pub_StrUserSt03 = "F11"
        If (Left(Pub_StrUserSt03, 2) = "P2" Or Pub_StrUserSt03 = "F11") Or _
           InStr(strMCTF, strUserNum) > 0 Then
             strSql = ""
             'MCTM
             If InStr(strMCTF, strUserNum) > 0 Then
                 strSql = "And SubStr(st01,1,4)='MCTF' "
             Else
                 strMCTF = GetMCTF0XAllCode(strUserNum)
                 strSql = "And st01 in ('" & strMCTF & "') "
             End If
          
             If strSql <> MsgText(601) Then
                strSql = "Select st01,st02 From Staff Where 1=1 " & strSql & " Order by st01 "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                If intI = 1 Then
                   RsTemp.MoveFirst
                   Do While Not RsTemp.EOF
                      For ii = 0 To Combo1.ListCount - 1
                         If InStr(Combo1.List(ii), RsTemp(0)) = 1 Then
                            'Add By Sindy 2020/11/20 抓第一筆為預設值
                            If stDef = "" Then
                              stDef = RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                            End If
                            '2020/11/20 END
                            Exit For
                         End If
                      Next
                      If ii = Combo1.ListCount Then
                        Combo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                        'Add By Sindy 2020/11/20 抓第一筆為預設值
                        If stDef = "" Then
                           stDef = RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
                        End If
                        '2020/11/20 END
                      End If
                      RsTemp.MoveNext
                   Loop
                End If
             End If
        End If
   End If
   'end 2020/02/05
   
'   '檢查當時是否需要為他人職代
'   Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)

'   'Modify By Sindy 2014/8/28 帶人的權限
'   'Modified by Lydia 2020/06/08 +增加特殊權限"AREA"
'   Call Pub_SetSAManageEmpCombo(strUserNum, Combo1, False, , , "AREA")
'   '2014/8/28 END
'
'   'Added by Morgan 2014/5/15
'   '專利處智權同仁代處理人
'   'Modify by Amy 2015/03/13 +特殊設定(總經理業務工作代理人員)
'   If InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Or InStr(Pub_GetSpecMan("總經理業務工作代理人員"), strUserNum) > 0 Then
'      If InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Then
'        strSql = "select st01,st02 from setSpecMan,staff where ocode='A7' and instr(oMan,st01)>0"
'      Else
'        strSql = "select st01,st02 from setSpecMan,staff where ocode='總經理員工編號' and instr(oMan,st01)>0"
'      End If
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         Do While Not RsTemp.EOF
'            For ii = 0 To Combo1.ListCount - 1
'               If InStr(Combo1.List(ii), RsTemp(0)) = 1 Then
'                  Exit For
'               End If
'            Next
'            If ii = Combo1.ListCount Then
'               Combo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
'            End If
'            RsTemp.MoveNext
'         Loop
'      End If
'   End If
'   'end 2014/5/15
   
   'Add By Sindy 2014/7/10 操作人員為3.南所之M71.管理部分所人員時,可查詢所有S31人員的寄件資料
   If Pub_StrUserSt03 = "M71" And PUB_GetST06(strUserNum) = "3" Then
      strSql = "select st01,st02 from staff where st15='S31' and st04='1' and st01 not in('00001','00002','00003','00004')"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            For ii = 0 To Combo1.ListCount - 1
               If InStr(Combo1.List(ii), RsTemp(0)) = 1 Then
                  Exit For
               End If
            Next
            If ii = Combo1.ListCount Then
               Combo1.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
            End If
            RsTemp.MoveNext
         Loop
      End If
   End If
   '2014/7/10 END
   
   If Pub_StrUserSt03 = "M51" Then
      Combo1.AddItem "      " & "全部"
   End If
   
   'Modify By Sindy 2020/11/20 有預設值就帶 ex:MCTF01
   If stDef <> "" Then
      Combo1 = stDef
   Else
      Combo1.ListIndex = 0
   End If
'cancel by sonia 2024/9/27
'   'Add By Sindy 2023/5/16
'   If InStr(Pub_GetSpecMan("P1004管理人員"), strUserNum) > 0 Then
'      Combo1 = "P1004 " & GetPrjSalesNM("P1004")
'   End If
'   '2023/5/16 END
'end 2024/9/27
End Sub

Private Sub Form_Activate()
   If Screen.ActiveForm.Name <> Me.Name Then Exit Sub 'Added by Morgan 2023/6/2
   If Me.WindowState = 0 Then Me.WindowState = 2 '最大化
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   WebBrowser1.Navigate "about:blank"
   cmdState = -1
   
   ReDim m_FilesRemoved(0)
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
   
   'Remove by Morgan 2016/3/2 目前沒有使用
   'If Pub_StrUserSt03 = "M51" And InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") Then
   '   cmdPrintAtt.Visible = True
   'Else
   '   cmdPrintAtt.Visible = False
   'End If
   'end 2016/3/2
   
   '1.寄件查詢(預設15天內的資料)
   If Val(intWorkItem) = 1 Then
      Text4.Text = Val(Format(DateAdd("d", -15, ChangeWStringToWDateString(strSrvDate(1))), "YYYYMMDD")) - 19110000
      Text5.Text = strSrvDate(2)
      CmdOk1(2).Visible = False '隱藏關係企業
      '不顯示申請人
      Label3.Visible = False
      LblAppl.Visible = False
      '顯示智權人員
      Frame2.Visible = True
      SetCombo1
      'SetGrd
      Call QueryData2
   '0.以申請人查詢寄件資料(預設一個月內的資料)
   Else
      Text4.Text = Val(Format(DateAdd("M", -1, ChangeWStringToWDateString(strSrvDate(1))), "YYYYMMDD")) - 19110000
      Text5.Text = strSrvDate(2)
      CmdOk1(2).Visible = True '開啟關係企業
      '顯示申請人
      Label3.Visible = True
      LblAppl.Visible = True
      '不顯示智權人員
      Frame2.Visible = False
   End If
   
'   'Add By Sindy 2019/2/1
'   If Pub_StrUserSt03 = "M51" Then
'      lblCnt.Visible = True
'   Else
'      lblCnt.Visible = False
'   End If
   
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

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   Set frm210145 = Nothing
End Sub

Private Sub KillAttach()
On Error Resume Next
   'Modified by Morgan 2020/2/18
   'If Dir(m_AttachPath & "\.") <> "" Then
   '   Kill m_AttachPath & "\*.pdf"
   'End If
   PUB_ClearTempFolder m_AttachPath
   'end 2020/2/18
End Sub

Private Function GetValue(pRow As Integer, pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As String
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetValue = .TextMatrix(pRow, iRow)
         Exit For
      End If
   Next
   End With
End Function

Private Sub SelectRow(ByRef pRow As Integer, ByRef FlexGrid As MSHFlexGrid, ByRef pPrevRow As Integer)
Dim nCol As Integer, iCol As Integer
Dim strCaseNo As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   
   With FlexGrid
   nCol = .col
   If pPrevRow > 0 Then
      If pPrevRow <> pRow Then
         .row = pPrevRow
         .TextMatrix(pPrevRow, 0) = ""
         If .FixedCols > 0 Then
            .col = .FixedCols - 1
            .CellBackColor = .BackColorFixed
            .CellForeColor = .ForeColor
         End If
         For iCol = .FixedCols To .Cols - 1
            .col = iCol
            If .CellBackColor = &HFFC0C0 Then 'Added by Morgan 2015/6/29
               .CellBackColor = .BackColor
            End If
         Next
      End If
   End If

   If pRow > 0 Then
      strCaseNo = .TextMatrix(pRow, 2)
      If UBound(Split(strCaseNo, "-")) = 1 Then
         strCaseNo = strCaseNo & "-0-00"
      End If
      strCP01 = SystemNumber(strCaseNo, 1)
      strCP02 = SystemNumber(strCaseNo, 2)
      strCP03 = SystemNumber(strCaseNo, 3)
      strCP04 = SystemNumber(strCaseNo, 4)
      '檢查權限
      If CheckSR09(strUserNum, strCP01, "Y", , strCP01, strCP02, strCP03, strCP04) = False Then
         Exit Sub
      End If
      
      .row = pRow
      .TextMatrix(pRow, 0) = "V"
      If .FixedCols > 0 Then
         .col = .FixedCols - 1
         .CellBackColor = .BackColorSel
         .CellForeColor = .ForeColorSel
      End If
      For iCol = .FixedCols To .Cols - 1
        .col = iCol
         If .CellBackColor = .BackColor Then 'Added by Morgan 2015/6/29
            .CellBackColor = &HFFC0C0
         End If
      Next
   End If
   .col = nCol
   pPrevRow = pRow
   End With
End Sub

Private Sub Grd1_Click()
Dim nCol As Integer, nRow As Integer, iRow As Integer, iCol As Integer
Dim stValue As String
Dim stCP09 As String
   
   With GRD1
   .Visible = False
   nCol = .MouseCol
   nRow = .MouseRow
   If nRow = 0 Then
      '紀錄前次點選的收文號
      If m_mouseRow > 0 Then
         stCP09 = GetValue(m_mouseRow, "cp09", GRD1)
      End If
      
      .col = nCol
      If m_blnColOrderAsc = False Then '字串降冪
         .Sort = 5 '字串昇冪
         m_blnColOrderAsc = True
      Else
         .Sort = 6 '字串降冪
         m_blnColOrderAsc = False
      End If
               
      '重設排序後前次點選的位置
      If m_mouseRow > 0 Then
         For iRow = 1 To .Rows - 1
            If stCP09 = GetValue(iRow, "cp09", GRD1) Then
               m_mouseRow = iRow
               Exit For
            End If
         Next
      End If
   ElseIf nRow > 0 And .TextMatrix(nRow, 2) <> "" Then
      SelectRow nRow, GRD1, m_mouseRow
   End If
   .Visible = True
   End With
End Sub

Private Sub GRD1_DblClick()
   If GRD1.MouseRow > 0 Then
'      Screen.MousePointer = vbHourglass
'      m_bolDblClick = True
      cmdOpenAtt_Click 1
      'Added by Morgan 2017/12/27
      'Removed by Morgan 2018/10/1 看進度備註,LP35先保留
      'strExc(1) = GetValue(GRD1.MouseRow, "LP35", GRD1)
      'If strExc(1) <> "" Then
      '   MsgBox strExc(1), , "報價備註"
      'End If
      'end 2018/10/1
      'end 2017/12/27
   End If
End Sub

'Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim nCol As Long, nRow As Long
'Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
'Dim bolClear As Boolean
'Dim strCaseNo As String
'
'   getGrdColRow GRD1, x, y, nCol, nRow
'   If nCol < 0 Or nRow < 0 Then Exit Sub
'   'GRD1.col = nCol
'   GRD1.row = nRow
'   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
'      If Me.GRD1.Text = "目次" Then
'         If m_blnColOrderAsc = True Then
'            Me.GRD1.Sort = 3  '數值昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.GRD1.Sort = 4 '數值降冪
'            m_blnColOrderAsc = True
'         End If
'      Else
'         If m_blnColOrderAsc = True Then
'            Me.GRD1.Sort = 5 '字串昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.GRD1.Sort = 6 '字串降冪
'            m_blnColOrderAsc = True
'         End If
'      End If
'   End If
'
'   GRD1.Visible = False
'   GRD1.row = nRow
'   If GRD1.row <> 0 And Trim(GRD1.TextMatrix(GRD1.row, 15)) <> "" Then
'      bolClear = False
'      If m_mouseRow > 0 Then
'         If GRD1.TextMatrix(m_mouseRow, 0) = "V" Then
'            '清除反白
'            GRD1.TextMatrix(m_mouseRow, 0) = ""
'            GRD1.row = m_mouseRow
'            For jj = 1 To GRD1.Cols - 1
'               GRD1.col = jj
'               GRD1.CellBackColor = QBColor(15)
'            Next jj
'            bolClear = True
'         End If
'      End If
'      m_mouseRow = GRD1.row
'      If m_mouseRow = GRD1.row And bolClear = True Then
'         GoTo gotoExit
'      End If
'      '資料列反白
'      GRD1.TextMatrix(GRD1.row, 0) = "V"
'      strCaseNo = GRD1.TextMatrix(GRD1.row, 2)
'      If UBound(Split(strCaseNo, "-")) = 1 Then
'         strCaseNo = strCaseNo & "-0-00"
'      End If
'      strCP01 = SystemNumber(strCaseNo, 1)
'      strCP02 = SystemNumber(strCaseNo, 2)
'      strCP03 = SystemNumber(strCaseNo, 3)
'      strCP04 = SystemNumber(strCaseNo, 4)
'      '檢查權限
'      If CheckSR09(strUserNum, strCP01, "Y", , strCP01, strCP02, strCP03, strCP04) = False Then
'         GoTo gotoExit
'      End If
'      For jj = 1 To GRD1.Cols - 1
'         GRD1.col = jj
'         GRD1.CellBackColor = &HFFC0C0
'      Next jj
'   End If
'
'gotoExit:
'   GRD1.Visible = True
'End Sub

Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
GRD1.ToolTipText = ""
If GRD1.MouseRow <> 0 And GRD1.MouseCol > 0 Then
   If GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
      GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
   End If
End If
End Sub

''Private Sub grd1_SelChange()
'Private Sub Grd1SelChange(intMouseRow As Integer)
'Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
'Dim bolClear As Boolean
'Dim strCaseNo As String
'
'GRD1.Visible = False
'If GRD1.MouseRow <> 0 And Trim(GRD1.TextMatrix(GRD1.MouseRow, 15)) <> "" Then
'   bolClear = False
'   If m_mouseRow > 0 Then
'      If GRD1.TextMatrix(m_mouseRow, 0) = "V" Then
'         '清除反白
'         GRD1.TextMatrix(m_mouseRow, 0) = ""
'         GRD1.row = m_mouseRow
'         For jj = 1 To GRD1.Cols - 1
'            GRD1.col = jj
'            GRD1.CellBackColor = QBColor(15)
'         Next jj
'         bolClear = True
'      End If
'   End If
'   If m_mouseRow = GRD1.MouseRow And bolClear = True Then
'      GoTo gotoExit
'   End If
'   m_mouseRow = GRD1.MouseRow
'   '資料列反白
'   GRD1.TextMatrix(GRD1.MouseRow, 0) = "V"
'   GRD1.row = GRD1.MouseRow
'   strCaseNo = GRD1.TextMatrix(GRD1.MouseRow, 2)
'   If UBound(Split(strCaseNo, "-")) = 1 Then
'      strCaseNo = strCaseNo & "-0-00"
'   End If
'   strCP01 = SystemNumber(strCaseNo, 1)
'   strCP02 = SystemNumber(strCaseNo, 2)
'   strCP03 = SystemNumber(strCaseNo, 3)
'   strCP04 = SystemNumber(strCaseNo, 4)
'   '檢查權限
'   If CheckSR09(strUserNum, strCP01, "Y", , strCP01, strCP02, strCP03, strCP04) = False Then
'      GoTo gotoExit
'   End If
'   For jj = 1 To GRD1.Cols - 1
'      GRD1.col = jj
'      GRD1.CellBackColor = &HFFC0C0
'   Next jj
'End If
'gotoExit:
'GRD1.Visible = True
'End Sub

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
'Removed by Morgan 2016/3/2 目前沒用使用
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
''         '清除反白
''         GRD1.col = 0
''         GRD1.row = ii
''         GRD1.TextMatrix(ii, 0) = ""
''         For jj = 1 To GRD1.Cols - 1
''            GRD1.col = jj
''            GRD1.CellBackColor = QBColor(15)
''         Next jj
'         If Trim(GRD1.TextMatrix(ii, 15)) <> "" Then
'            bolIsSelect = True
'            m_CP09 = Trim(GRD1.TextMatrix(ii, 15))
'            m_CP10 = Trim(GRD1.TextMatrix(ii, 16)) 'Add By Sindy 2014/11/4
'            'Modify By Sindy 2014/11/4 +, m_CP10
'            'Modified by Morgan 2016/3/2 改呼叫共用
'            'If GetAttachFile(m_CP09, stFileName, "", m_CP10) = False Then
'            If PUB_GetAttachFile4Cust(m_CP09, stFileName, "", True, m_CP10) = False Then
'               MsgBox "無法儲存檔案[ " & stFileName & " ]！"
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

'Removed by Morgan 2016/3/2 改寫共用 PUB_GetAttachFile4Cust
''Modified by Morgan 2014/12/24 +pJoinDoc 是否合併
'Private Function GetAttachFile(ByVal strCPP01 As String, ByRef pFileName As String, Optional pSavePath As String, Optional pCP10 As String, Optional pJoinDoc As Boolean = True) As Boolean
'   Dim stAttPath As String
'   Dim lngSize As Long
'   Dim iFileNo As Integer
'   Dim bytes() As Byte
'   Dim strCmd As String
'   Dim process_id As Long
'   Dim process_handle As Long
'   Dim strMergeFN As String, strFileName As String
'   Dim intFileCnt As Integer
'   'Added by Morgan 2014/9/29
'   Dim stExceptCon As String '過濾不要顯示的檔案
'   Dim stSort2 As String '順序
'
'On Error GoTo ErrHnd
'
'   GetAttachFile = False
'   pFileName = ""
'   strMergeFN = ""
'   intFileCnt = 0
'
'   If pSavePath = "" Then
'      If Dir(m_AttachPath, vbDirectory) = "" Then
'         MkDir m_AttachPath
'      End If
'      '切換至來源目錄
'      If m_AttachPath <> "." Then ChDir m_AttachPath
'      stAttPath = m_AttachPath
'   Else
'      If InStr(pSavePath, m_AttachPath) > 0 Then
'         If Dir(m_AttachPath, vbDirectory) = "" Then
'            MkDir m_AttachPath
'         End If
'      End If
'      stAttPath = pSavePath
'   End If
'
'   '客戶函在最上面
'   'strExc(0) = "select '1' Srt1,decode(instr(UPPER(cpp02),'.CUS.PDF'),0,2,1) Srt2,length(cpp02) Srt3,casepaperpdf.* from casepaperpdf where cpp01='" & strCPP01 & "' order by Srt1,Srt2,Srt3"
'   'Modify By Sindy 2014/11/4 修改需讀取出來的電子檔及排序(參考:frm210144 修改)
'   '排除條件
'   stExceptCon = ""
'   '客戶函在最上面
'   '來函 C 類
'   'Modified by Morgan 2015/3/23 讀取改呼叫共用函數(要改為FTP方式)
'   If strCPP01 > "C" Then
'      '文件順序1.客戶函(cus)  2.收據(receipt) 3.其他
'      stSort2 = "DECODE(SUBSTR(UPPER(cpp02),-8),'.CUS.PDF',1,DECODE(SUBSTR(UPPER(cpp02),-12),'.RECEIPT.PDF',2,3))"
'      strExc(0) = " select '1' Srt1," & stSort2 & " Srt2,length(cpp02) Srt3,cpp01,cpp02" & _
'         " from casepaperpdf c where cpp01='" & strCPP01 & "' and SUBSTR(UPPER(CPP02),-4)='.PDF' " & stExceptCon
'      '證書要一併抓通知公告的檔案
'      '抓1228公告公報才對
'      If pCP10 = "1603" Then
'         strExc(0) = strExc(0) & " union all"
'         strExc(0) = strExc(0) & " select '2' Srt,9 Srt2,length(cpp02) Srt3,cpp01,cpp02"
'         strExc(0) = strExc(0) & " from caseprogress a,caseprogress b,casepaperpdf c" & _
'            " where a.cp09='" & strCPP01 & "' and b.cp01(+)=a.cp01 and b.cp02(+)=a.cp02 and b.cp03(+)=a.cp03 and b.cp04(+)=a.cp04 and b.cp10='1228'" & _
'            " and cpp01(+)=b.cp09 and SUBSTR(UPPER(CPP02),-4)='.PDF' " & stExceptCon
'      '通知申請號一併抓申請程序的檔案
'      ElseIf pCP10 = "1101" Then
'         strExc(0) = strExc(0) & " union all"
'         '文件順序1.客戶函(cus) 2.收據(receipt) 3.申請書(data) 4.說明書(inv,utl,des...), 5.圖(dwg)
'         stSort2 = "DECODE(SUBSTR(UPPER(cpp02),-8),'.CUS.PDF',1,'.DWG.PDF',5,DECODE(SUBSTR(UPPER(cpp02),-12),'.RECEIPT.PDF',2,DECODE(SUBSTR(UPPER(cpp02),-9),'.DATA.PDF',3,4)))"
'         strExc(0) = strExc(0) & " select '1' Srt1," & stSort2 & " Srt2,length(cpp02) Srt3,cpp01,cpp02"
'         strExc(0) = strExc(0) & " from caseprogress a,caseprogress b,casepaperpdf c" & _
'            " where a.cp09='" & strCPP01 & "' and b.cp09(+)=a.cp43  and cpp01(+)=a.cp43 and SUBSTR(UPPER(CPP02),-4)='.PDF' " & stExceptCon & _
'            " and (instr(upper(cpp02),'.RECEIPT.PDF')>0 OR exists(select * from custletterrefext where cle01=b.cp10 and cle02='000'" & _
'            " and  instr(upper(cpp02),'.'||upper(cle03)||'.PDF')>0 AND ( CLE04 IS NULL OR CLE04=DECODE(b.CP118,NULL,2,1) ) ) )"
'      End If
'   '發文 A,B 類
'   Else
'      '領證年費不要抓申請書
'      If pCP10 = "601" Or pCP10 = "605" Then
'         stExceptCon = stExceptCon & " and INSTR(UPPER(CPP02),'.DATA.PDF')=0"
'      End If
'      '文件順序1.客戶函(cus) 2.收據(receipt) 3.申請書(data) 4.說明書(inv,utl,des...), 5.圖(dwg)
'      'Modified by Morgan 2014/11/20 收據要判斷有發文規費的才抓否便為回執無需給客戶
'      stSort2 = "DECODE(SUBSTR(UPPER(cpp02),-8),'.CUS.PDF',1,'.DWG.PDF',5,DECODE(SUBSTR(UPPER(cpp02),-12),'.RECEIPT.PDF',2,DECODE(SUBSTR(UPPER(cpp02),-9),'.DATA.PDF',3,4)))"
'      strExc(0) = "select '1' Srt1," & stSort2 & " Srt2,length(cpp02) Srt3,cpp01,cpp02"
'      strExc(0) = strExc(0) & " from casepaperpdf c,caseprogress where cpp01='" & strCPP01 & "' and SUBSTR(UPPER(CPP02),-4)='.PDF' and cp09(+)=cpp01 " & stExceptCon & _
'         " and (instr(upper(cpp02),'.CUS.PDF')>0 OR (instr(upper(cpp02),'.RECEIPT.PDF')>0 and cp84>0) OR exists(select * from custletterrefext where cle01=cp10 and cle02='000'" & _
'         " and  instr(upper(cpp02),'.'||upper(cle03)||'.PDF')>0 AND ( CLE04 IS NULL OR CLE04=DECODE(CP118,NULL,2,1))))"
'   End If
'   strExc(0) = strExc(0) & " order by Srt1,Srt2,Srt3"
'   '2014/11/4 END
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      intFileCnt = RsTemp.RecordCount
'      With RsTemp
'      .MoveFirst
'      Do While Not .EOF
'         'Modified by Morgan 2015/6/18 檔名含空白無法合併
'         'strFileName = stAttPath & IIf(Right(stAttPath, 1) = "\", "", "\") & .Fields("cpp02")
'         strFileName = stAttPath & IIf(Right(stAttPath, 1) = "\", "", "\") & Replace(.Fields("cpp02"), " ", "_")
'         'end 2015/6/18
'
'      'Modified by Morgan 2014/12/25 檔案存在時不必再下載
'         'If pSavePath = "" Then
'         '   If Dir(strFileName) <> "" Then Kill strFileName
'         'Else
'         '   If Dir(strFileName) <> "" Then
'         '      If MsgBox("檔案[ " & strFileName & " ]已存在是否要覆蓋??", vbYesNo + vbDefaultButton2) = vbNo Then
'         '         GoTo ExitEnd
'         '      Else
'         '         Kill strFileName
'         '      End If
'         '   End If
'         'End If
'      If Dir(strFileName) = "" Then
'      'end 2014/12/25
'
''Modified by Morgan 2015/3/23
''            lngSize = Val(.Fields("cpp03").Value)
''            ReDim bytes(lngSize)
''            If lngSize > 0 Then
''               bytes() = .Fields("cpp04").GetChunk(lngSize)
''            End If
''         iFileNo = FreeFile
''         Open strFileName For Binary Access Write As #iFileNo
''         If lngSize > 0 Then Put #iFileNo, , bytes()
''         Close #iFileNo
'         'Modified by Morgan 2015/6/18 檔名含空白無法合併
'         'If PUB_GetAttachFile_CPP(.Fields("cpp01"), .Fields("cpp02"), stAttPath) = False Then
'         If PUB_GetAttachFile_CPP(.Fields("cpp01"), .Fields("cpp02"), strFileName, True) = False Then
'         'end 2015/6/18
'            Exit Function
'         End If
''end 2015/3/23
'      End If 'Added by Morgan 2014/12/25
'
'      pFileName = pFileName & ";" & strFileName
'      'Modified by Morgan 2015/6/18 檔名含空白無法合併
'      'strMergeFN = strMergeFN & IIf(strMergeFN <> "", " ", "") & ".\" & .Fields("cpp02")
'      strMergeFN = strMergeFN & IIf(strMergeFN <> "", " ", "") & ".\" & Replace(.Fields("cpp02"), " ", "_")
'      'end 2015/6/18
'      .MoveNext
'      Loop
'      End With
'   End If
'
'   If pSavePath = "" Then
'      '合併
'      'Modified by Morgan 2014/12/24
'      'If intFileCnt > 1 Then
'      If intFileCnt > 1 And pJoinDoc = True Then
'      'end 2014/12/24
'
'         pFileName = "merge" & ServerTime & ".pdf"
'         strCmd = pub_PdftkEXE & " " & strMergeFN & " cat output .\" & pFileName
'         pFileName = stAttPath & "\" & pFileName
'         process_id = SHELL(strCmd, vbHide)
'         process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
'         If process_handle <> 0 Then
'            For intI = 1 To 10
'               If PUB_CheckIsRunning(pub_PdftkName) = True Then
'                  Sleep 1000
'               Else
'                  Exit For
'               End If
'            Next
'            If intI > 10 Then
'               TerminateProcess process_handle, 0&
'               CloseHandle process_handle
'               MsgBox "合併PDF失敗！"
'               GoTo ErrHnd
'            Else
'               CloseHandle process_handle
'            End If
'         Else
'            MsgBox "合併PDF失敗！"
'            GoTo ErrHnd
'         End If
'      Else
'         pFileName = Mid(pFileName, 2)
'      End If
'   End If
'
'ExitEnd:
'   GetAttachFile = True
'   ChDir App.path '目錄切回
'   Exit Function
'
'ErrHnd:
'   ChDir App.path '目錄切回
'   If Err.NUMBER = 70 Then
'      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
'   Else
'      MsgBox Err.Description, vbCritical
'   End If
'   If iFileNo > 0 Then Close #iFileNo
'End Function

'開啟附件
Private Sub cmdOpenAtt_Click(Index As Integer)
   Dim hLocalFile As Long
   Dim stFileName As String
   Dim bolIsSelect As Boolean
   Dim bolJoin As Boolean 'Added by Morgan 2014/12/24
   Dim stFileDescs As String 'Added by Morgan 2019/3/4
   Dim stFiles As String, idx As Integer 'Added by Morgan 2020/2/18
   Dim arrFileName() As String 'Added by Morgan 2020/2/18
   
   KillAttach
   If Index = 1 Then
      If Check1.Value = 1 Then
         Check1.Value = 0
      End If
      WebBrowser1.Navigate "about:blank"
   End If
   bolIsSelect = False
   
   'Added by Morgan 2014/12/24
   If Index = 99 Then
      bolJoin = False
   Else
      bolJoin = True
   End If
   'end 2014/12/24
   
   Screen.MousePointer = vbHourglass
   
'   If m_bolDblClick = True Then
'      If GRD1.TextMatrix(m_mouseRow, 0) = "" Then
'         GRD1.TextMatrix(m_mouseRow, 0) = "V"
'         For jj = 1 To GRD1.Cols - 1
'            GRD1.col = jj
'            GRD1.CellBackColor = &HFFC0C0
'         Next jj
'      End If
'   End If
   
   For ii = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(ii, 0) = "V" Then
'         '清除反白
'         GRD1.col = 0
'         GRD1.row = ii
'         GRD1.TextMatrix(ii, 0) = ""
'         For jj = 1 To GRD1.Cols - 1
'            GRD1.col = jj
'            GRD1.CellBackColor = QBColor(15)
'         Next jj
'         If m_bolDblClick = False Or (m_bolDblClick = True And ii = m_mouseRow) Then
            If Trim(GRD1.TextMatrix(ii, 15)) <> "" Then
               bolIsSelect = True
               m_CP09 = Trim(GRD1.TextMatrix(ii, 15))
               m_CP10 = Trim(GRD1.TextMatrix(ii, 16)) 'Add By Sindy 2014/11/4
               'Modify By Sindy 2014/11/4 +, m_CP10
               'Modified by Morgan 2014/12/24 +bolJoin
               'Modified by Morgan 2016/3/2 改呼叫共用
               'If GetAttachFile(m_CP09, stFileName, "", m_CP10, bolJoin) = False Then
               'Modified by Morgan 2019/3/4 +stFileDescs
               'Modified by Morgan 2020/2/18
               'If PUB_GetAttachFile4Cust(m_CP09, stFileName, "", bolJoin, m_CP10, stFileDescs) = False Then
               m_AttachPath2 = m_AttachPath & "\" & m_CP09
               If PUB_GetAttachFile4Cust(m_CP09, stFiles, m_AttachPath2, bolJoin, m_CP10, stFileDescs) = False Then
               'end 2020/2/18
                  'MsgBox "無法儲存檔案[ " & stFileName & " ]！"
               End If
               
               If stFileDescs <> "" Then
                  SetValue m_mouseRow, "FDesc", stFileDescs, GRD1
               Else
                  stFileDescs = GetValue(m_mouseRow, "FDesc", GRD1)
               End If
               
               'Added by Morgan 2020/2/18
               arrFileName = Split(stFiles, ";")
               For idx = UBound(arrFileName) To LBound(arrFileName) Step -1
                  If arrFileName(idx) <> "" Then
                     If bolJoin Then
                        stFileName = m_AttachPath2 & "\" & arrFileName(idx)
                        Exit For
                     Else
                        stFileName = m_AttachPath2 & "\" & arrFileName(idx) & ";" & stFileName
                     End If
                  End If
               Next
               'end 2020/2/18
            
               'stFileName
               '開啟檔案
               If Index = 0 Then
                  ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
               'Added by Morgan 2014/12/24
               'EMail
               ElseIf Index = 99 Then
                  'Modified by Morgan 2021/5/7
                  'If stFileName <> "" Then
                  If PUB_ChkAttFile(stFileName) = True Then
                  'end 2021/5/7
                     Screen.MousePointer = vbDefault
                     'Modified by Morgan 2021/8/19 Email內文要帶出法定期限及本所期限
                     'Modified by Morgan 2021/8/31 定稿太多，有些不該有期限如發文通知，改先做期限通知(D類)
                     'PUB_ShowMailForm m_CP09, stFileName, GetValue(ii, "案件性質", grd1)
                     PUB_SetDateAndFeeYear m_CP09, strExc(6), strExc(7), strExc(8)
                     strExc(9) = GetValue(ii, "案件性質", GRD1) & IIf(strExc(8) = "", "", " [ " & strExc(8) & " ] ")
                     'Modified by Lydia 2021/09/15 和寄發文件一樣，同時存寄件備份 by 周哲丞
                     'PUB_ShowMailForm m_CP09, stFileName, strExc(9), , , strExc(6), strExc(7)
                     'Add By Sindy 2024/10/16 + bolReadLP42=True
                     PUB_ShowMailForm m_CP09, stFileName, strExc(9), , , strExc(6), strExc(7), True, , , , , , , , , , , , , , , True
                     'end 2021/8/19
                     Screen.MousePointer = vbHourglass
                  End If
               'end 2014/12/24
               Else
                  WebBrowser1.Navigate stFileName
                  SetAttList stFileDescs
               End If
            End If
'         End If
      End If
   Next ii
   If bolIsSelect = False Then
      MsgBox "無檔案可開啟！"
   End If
   
ErrHnd:
'   m_bolDblClick = False
   Screen.MousePointer = vbDefault
End Sub

'下載
Private Sub cmdSaveAtt_Click()
   Dim stFolderPath As String, stFileName As String
   Dim bolIsSelect As Boolean
   Dim ii As Integer
   
   Screen.MousePointer = vbHourglass
   bolIsSelect = False
   
   For ii = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(ii, 0) = "V" Then
'         '清除反白
'         GRD1.col = 0
'         GRD1.row = ii
'         GRD1.TextMatrix(ii, 0) = ""
'         For jj = 1 To GRD1.Cols - 1
'            GRD1.col = jj
'            GRD1.CellBackColor = QBColor(15)
'         Next jj
         If Trim(GRD1.TextMatrix(ii, 15)) <> "" Then
            bolIsSelect = True
            m_CP09 = Trim(GRD1.TextMatrix(ii, 15))
            m_CP10 = Trim(GRD1.TextMatrix(ii, 16)) 'Add By Sindy 2014/11/4
            stFolderPath = BrowseForFolder()
            If stFolderPath <> "" Then
               'Modify By Sindy 2014/11/4 +, m_CP10
               'Modified by Morgan 2016/3/2 改呼叫共用
               'If GetAttachFile(m_CP09, stFileName, stFolderPath, m_CP10) = False Then
               If PUB_GetAttachFile4Cust(m_CP09, stFileName, stFolderPath, , m_CP10) = False Then
                  MsgBox "無法儲存檔案[ " & stFolderPath & " ]！"
                  GoTo RunExit
               Else
                  MsgBox "下載完成！"
               End If
            End If
         End If
      End If
   Next ii
   If bolIsSelect = False Then
      MsgBox "無檔案可開啟！"
   End If
   
RunExit:
   Screen.MousePointer = vbDefault
End Sub

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

'Add by Sindy 2020/9/4
Private Sub systemkind_GotFocus()
   TextInverse systemkind
   CloseIme
End Sub
Private Sub systemkind_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2015/6/24
Private Sub Text1_GotFocus(Index As Integer)
   Text1(Index).SelStart = 0
   Text1(Index).SelLength = Len(Text1(Index))
End Sub
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Text1_LostFocus(Index As Integer)
   If Index = 3 Then
      If Text1(0) <> "" And Text1(1) <> "" Then
         If Text1(2) = "" Then Text1(2) = "0"
         If Text1(3) = "" Then Text1(3) = "00"
         'Modify By Sindy 2016/8/5 Mark
'         Text4.Text = ""
'         Text5.Text = ""
      End If
   End If
End Sub
'2015/6/24 END

Private Sub Text4_GotFocus()
   Text4.SelStart = 0
   Text4.SelLength = Len(Text4)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_LostFocus()
   If PUB_CheckKeyInDate(Me.Text4) = -1 Then
      Me.Text4.SetFocus
      Text4_GotFocus
      Exit Sub
   End If
   'Modify By Sindy 2016/8/5 Mark
'   'Add By Sindy 2015/6/25
'   If Text4 <> "" Then
'      Text1(0).Text = ""
'      Text1(1).Text = ""
'      Text1(2).Text = ""
'      Text1(3).Text = ""
'   End If
'   '2015/6/25 END
End Sub

Private Sub Text5_GotFocus()
   Text5.SelStart = 0
   Text5.SelLength = Len(Text5)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_LostFocus()
   If PUB_CheckKeyInDate(Me.Text5) = -1 Then
      Me.Text5.SetFocus
      Text5_GotFocus
      Exit Sub
   End If
   If Not nickChgRan(Text4, Text5, "發文日期") Then
      Text4.SetFocus
      Text4_GotFocus
      Exit Sub
   End If
   'Modify By Sindy 2016/8/5 Mark
'   'Add By Sindy 2015/6/25
'   If Text5 <> "" Then
'      Text1(0).Text = ""
'      Text1(1).Text = ""
'      Text1(2).Text = ""
'      Text1(3).Text = ""
'   End If
'   '2015/6/25 END
End Sub

Private Function GetFieldId(pFieldName As String, ByRef FlexGrid As MSHFlexGrid) As Integer
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         GetFieldId = iRow
         Exit For
      End If
   Next
   End With
End Function
'Added by Morgan 2019/3/4
Private Sub SetAttList(Optional pItems As String)
   Dim arrItem() As String
   Dim ii As Integer, iAttCnt As Integer
   
   cboAtt.Clear: lblAttCnt = " PDF:(0)"
   If pItems <> "" Then
      arrItem = Split(pItems, ";")
      For ii = LBound(arrItem) To UBound(arrItem)
         If arrItem(ii) <> "" Then
            cboAtt.AddItem arrItem(ii)
            iAttCnt = iAttCnt + 1
         End If
      Next
      lblAttCnt = " PDF:(" & iAttCnt & ")"
   End If
End Sub

Private Function SetValue(pRow As Integer, pFieldName As String, pValue As String, ByRef FlexGrid As MSHFlexGrid) As Boolean
   Dim iRow As Integer
   With FlexGrid
   For iRow = 0 To .Cols - 1
      If UCase(.TextMatrix(0, iRow)) = UCase(pFieldName) Then
         .TextMatrix(pRow, iRow) = pValue
         SetValue = True
         Exit Function
      End If
   Next
   End With
End Function
