VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_M 
   BorderStyle     =   1  '單線固定
   Caption         =   "原始檔查詢"
   ClientHeight    =   6220
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9310
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6220
   ScaleWidth      =   9310
   Tag             =   "加班資料"
   Begin VB.CheckBox ChkDelMsg 
      Caption         =   "不顯示郵件"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   210
      TabIndex        =   19
      Top             =   1050
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      BackColor       =   &H0000C0C0&
      Caption         =   "複製到..."
      Height          =   315
      Left            =   6325
      Style           =   1  '圖片外觀
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   390
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4875
      Left            =   75
      TabIndex        =   14
      Top             =   1320
      Width           =   9195
      _ExtentX        =   16228
      _ExtentY        =   8608
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "原始檔區"
      TabPicture(0)   =   "frm100101_M.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GRD1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "承辦人暫存區"
      TabPicture(1)   =   "frm100101_M.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRD1(1)"
      Tab(1).Control(1)=   "Label2"
      Tab(1).ControlCount=   2
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   6120
         Locked          =   -1  'True
         MousePointer    =   1  '箭號形狀
         TabIndex        =   20
         Text            =   "承辦人暫存區"
         Top             =   30
         Width           =   1665
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   4410
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   8955
         _ExtentX        =   15804
         _ExtentY        =   7796
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Height          =   4185
         Index           =   1
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   8955
         _ExtentX        =   15804
         _ExtentY        =   7391
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
      Begin VB.Label Label2 
         Caption         =   "注意：案件發文或取消收文一個月後，會刪除暫存區"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   -74820
         TabIndex        =   17
         Top             =   4600
         Width           =   6375
      End
   End
   Begin VB.CommandButton cmdReName 
      Caption         =   "更名"
      Height          =   360
      Left            =   1170
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   800
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Word修改"
      Height          =   360
      Left            =   180
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CheckBox ChkDelF 
      Caption         =   "含已刪除檔 (顯示紅色列)"
      Height          =   195
      Left            =   3720
      TabIndex        =   11
      Top             =   480
      Width           =   2475
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "取消全選"
      Height          =   360
      Left            =   5400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdAddAtt 
      Caption         =   "新增"
      Height          =   360
      Left            =   6325
      TabIndex        =   3
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdRemAtt 
      Caption         =   "刪除"
      Height          =   360
      Left            =   7166
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdSaveAtt 
      Caption         =   "下載"
      Height          =   360
      Left            =   4560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   360
      Left            =   3720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8010
      TabIndex        =   6
      Top             =   30
      Width           =   800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8010
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Win10因為解析度,物件需要有間隔。"
      ForeColor       =   &H00C000C0&
      Height          =   405
      Left            =   2010
      TabIndex        =   18
      Top             =   30
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   19
      Left            =   210
      TabIndex        =   10
      Top             =   450
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      Height          =   180
      Index           =   18
      Left            =   210
      TabIndex        =   9
      Top             =   750
      Width           =   960
   End
   Begin VB.Label lblCaseNo 
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   450
      Width           =   1830
   End
   Begin MSForms.Label lblCaseName 
      Height          =   255
      Left            =   1185
      TabIndex        =   7
      Top             =   750
      Width           =   7575
      Size            =   "13361;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm100101_M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/01 Form2.0已修改 lblCaseName
'Create by Sindy 2013/6/11
Option Explicit

' 變數宣告區
Public m_strKey As String '本所案號 或 多筆總收文號

Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_CP09 As String
Dim m_CP10 As String
Dim m_Nation As String

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
'Added by Lydia 2020/02/06 欄位的位置
Dim colCp09 As Integer   'CP09
Dim colCP10 As Integer '案件性質代號
Dim colCPM As Integer '案件性質
Dim colCP43 As Integer '相關總收文號
Dim colCP82 As Integer '發文時間
Dim colCPF02 As Integer '檔案名稱
Dim colCPF08 As Integer '檔案最後日期+時間
Dim colFN As Integer '檔案名稱(size)
Dim colReFN As Integer  '檔案更名
Dim colCPF10 As Integer 'Flag註記
Dim colCPF05 As Integer 'Added by Lydia 2020/03/03 異動人員
Dim colCPF06 As Integer 'Added by Sindy 2020/3/18 異動日期
Dim colCPF11 As Integer 'Added by Sindy 2020/3/18 資料來源
Dim colCPF14 As Integer  'Added by Sindy 2023/6/8 建檔人員
Dim colCPF13 As Integer 'Added by Lydia 2023/08/24 FTP路徑
Dim bolFileOpen As Boolean '下載檔案是否開啟
'Added by Lydia 2020/03/03
Dim m_identity As String '身份
Dim m_CP13 As String '目前案件的智權人員
Dim m_bolAddTmpFile As Boolean
Dim m_LimitType As String, m_RecvNo As String 'Add By Sindy 2020/12/31
Dim strYYMMSql As String 'Added by Lydia 2023/12/22 限制年月
Dim m_strF23User As String 'Added by Lydia 2024/11/27 外專承辦：該案承辦＋案件職代

Private Sub SetGrd(Index As Integer)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modified by Lydia 2020/02/06
   'arrGridHeadText = Array("V", "總收文號", "收文日", "案件性質", "檔案名稱", "CP09", "CP10", "最後修改時間", "檔案更名", "副檔名", "CP82", "CPF10", "CP43")
   'arrGridHeadWidth = Array(200, 1000, 800, 1500, 3400, 0, 0, 1500, 0, 0, 0, 0, 0)
   'Modified by Lydia 2020/03/03 +異動人員CPF05
   'Modified by Sindy 2023/6/8 +, "CPF07", "CPF14", "CPF15", "CPF16"
   '                        0    1           2         3           4           5       6       7               8           9             10      11       12      13       14       15       16       17       18       19       20
   'Modified by Lydia 2023/08/24
   'arrGridHeadText = Array("V", "總收文號", "收文日", "案件性質", "檔案名稱", "CP09", "CP10", "最後修改時間", "檔案更名", "CPF0809", "CP82", "CPF10", "CP43", "CPF02", "CPF05", "CPF06", "CPF11", "CPF07", "CPF14", "CPF15", "CPF16")
   'arrGridHeadWidth = Array(200, 1000, 800, 1500, 3400, 0, 0, 1500, 0, 0, _
                                         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   'end 2020/02/06
   arrGridHeadText = Array("V", "總收文號", "收文日", "案件性質", "檔案名稱", "CP09", "CP10", "最後修改時間", "檔案更名", "CPF0809", "CP82", "CPF10", "CP43", "CPF02", "CPF05", "CPF06", "CPF11", "CPF07", "CPF14", "CPF15", "CPF16", _
                           "CP66", "CP67", "COLSORT", "CPF08", "CPF09", "CPF13")
   arrGridHeadWidth = Array(200, 1000, 800, 1500, 3400, 0, 0, 1500, 0, 0, _
                                         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                                         0, 0, 0, 0, 0, 0)
   GRD1(Index).Visible = False
   GRD1(Index).Cols = UBound(arrGridHeadText) + 1
   GRD1(Index).Rows = 2
   For iRow = 0 To GRD1(Index).Cols - 1
      GRD1(Index).row = 0
      GRD1(Index).col = iRow
      GRD1(Index).Text = arrGridHeadText(iRow)
      GRD1(Index).ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1(Index).CellAlignment = flexAlignCenterCenter
   Next
   'Added by Lydia 2020/02/06
   If colCPF02 = 0 Then
       colCp09 = PUB_MGridGetId("CP09", Me.GRD1(Index))
       colCP10 = PUB_MGridGetId("CP10", Me.GRD1(Index))
       colCPM = PUB_MGridGetId("案件性質", Me.GRD1(Index))
       colCP43 = PUB_MGridGetId("CP43", Me.GRD1(Index))
       colCP82 = PUB_MGridGetId("CP82", Me.GRD1(Index))
       colCPF02 = PUB_MGridGetId("CPF02", Me.GRD1(Index))
       colCPF08 = PUB_MGridGetId("CPF0809", Me.GRD1(Index))
       colFN = PUB_MGridGetId("檔案名稱", Me.GRD1(Index))
       colReFN = PUB_MGridGetId("檔案更名", Me.GRD1(Index))
       colCPF10 = PUB_MGridGetId("CPF10", Me.GRD1(Index))
       colCPF05 = PUB_MGridGetId("CPF05", Me.GRD1(Index)) 'Added by Lydia 2020/03/03
       colCPF06 = PUB_MGridGetId("CPF06", Me.GRD1(Index)) 'Added by Sindy 2020/3/18
       colCPF11 = PUB_MGridGetId("CPF11", Me.GRD1(Index)) 'Added by Sindy 2020/3/18
       colCPF14 = PUB_MGridGetId("CPF14", Me.GRD1(Index)) 'Added by Sindy 2023/6/8
       colCPF13 = PUB_MGridGetId("CPF13", Me.GRD1(Index)) 'Added by Lydia 2023/08/24 FTP路徑
   End If
   'end 2020/02/06
   GRD1(Index).Visible = True
End Sub

'檢查權限
Private Function CheckLimits() As Boolean
Dim strMsgTxt As String 'Add By Sindy 2022/7/18
Dim bolSpecCase As Boolean 'Add By Sindy 2023/4/27

   CheckLimits = False '無權限
'   m_QueryEfile = "" '鎖可查詢的系統別及副檔名
   
   'Modify By Sindy 2023/4/27 應該是先做CheckSR09檢查權限，有權限的人再檢查ACS特殊案的權限，非特殊案就不用再檢查了。
   bolSpecCase = PUB_ChkCPPAndCPFLimits_Spec(m_CP01, m_CP02, m_CP03, m_CP04, m_LimitType, m_RecvNo, strMsgTxt)
   'Modify By Sindy 2023/5/2
   If bolSpecCase = False Then
   '2023/5/2 END
      If CheckSR09(strUserNum, m_CP01, "Y", , m_CP01, m_CP02, m_CP03, m_CP04, Me.Name) = False Then
         Exit Function
      End If
   End If
   '2023/4/27 END
   
   'Add by Sindy 2020/12/31 原始檔及卷宗區特殊權限
   'If PUB_ChkCPPAndCPFLimits_Spec(m_CP01, m_CP02, m_CP03, m_CP04, m_LimitType, m_RecvNo, strMsgTxt) = True Then
   If bolSpecCase = True Then
      If m_LimitType = "" Then
         'Modify By Sindy 2021/8/19
         'MsgBox "您沒有查詢案件明細的權限", vbOKOnly, "檢核資料"
         MsgBox "此案與ACS" & strMsgTxt & "有關，您無權限查詢，您可請顧服組協助！", vbOKOnly, "檢核資料"
         '2021/8/19 END
         Exit Function
      End If
   End If
   CheckLimits = True '有權限
End Function

Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim varSplit As Variant
Dim pYYMM As String 'Added by Lydia 2023/12/22 限制收文年月

   '清空及預設欄位值
   GRD1(0).Clear
   Call SetGrd(0)
   GRD1(1).Clear
   Call SetGrd(1)
   lblCaseNo.Caption = Empty
   lblCaseName.Caption = Empty

   Screen.MousePointer = vbHourglass
   Me.Enabled = False

   If InStr(m_strKey, "-") = 0 Then '總收文號
      pub_QL05 = ";總收文號：" & m_strKey & "(原始檔)" 'Add By Sindy 2025/8/7
      varSplit = Split(m_strKey, ",")
      strSql = "Select cp01,cp02,cp03,cp04" & _
               " From CaseProgress" & _
               " Where CP09='" & varSplit(0) & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         m_CP01 = rsTmp.Fields("cp01")
         m_CP02 = rsTmp.Fields("cp02")
         m_CP03 = rsTmp.Fields("cp03")
         m_CP04 = rsTmp.Fields("cp04")
         
'         '檢查權限
'         'Modify by Sindy 2020/9/10 + , Me.Name
''         If CheckSR09(strUserNum, m_CP01, "Y", , m_CP01, m_CP02, m_CP03, m_CP04, Me.Name) = False Then
''            'tmpBol = fnCancelNowFormAndShowParentForm(Me)
''            Screen.MousePointer = vbDefault
''            Set rsTmp = Nothing
''            Call cmdExit_Click
''            Exit Function
''         End If
'         If CheckLimits() = False Then
'            'tmpBol = fnCancelNowFormAndShowParentForm(Me)
'            Screen.MousePointer = vbDefault
'            Set rsTmp = Nothing
'            Call cmdExit_Click
'            Exit Function
'         End If
         
         If InStr(m_strKey, ",") > 0 Then
            m_strKey = "'" & Replace(Trim(m_strKey), ",", "','") & "'"
         Else
            m_strKey = "'" & m_strKey & "'"
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
      pub_QL05 = ";本所案號：" & m_strKey & "(原始檔)" 'Add By Sindy 2025/8/7
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
'      'Modify by Sindy 2020/9/10 + , Me.Name
''      If CheckSR09(strUserNum, m_CP01, "Y", , m_CP01, m_CP02, m_CP03, m_CP04, Me.Name) = False Then
''         'tmpBol = fnCancelNowFormAndShowParentForm(Me)
''         Screen.MousePointer = vbDefault
''         Set rsTmp = Nothing
''         Call cmdExit_Click
''         Exit Function
''      End If
'      If CheckLimits() = False Then
'         'tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         Screen.MousePointer = vbDefault
'         Set rsTmp = Nothing
'         Call cmdExit_Click
'         Exit Function
'      End If
      'Modified by Lydia 2023/12/22 + strYYMMSql
      strSql = "Select distinct cp09" & _
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
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         m_strKey = ""
'         Do While Not rsTmp.EOF
'            m_strKey = m_strKey & "'" & Trim(rsTmp.Fields("cp09")) & "',"
'            rsTmp.MoveNext
'         Loop
         'Modify By Sindy 2021/10/15
         'm_strKey = Left(m_strKey, Len(m_strKey) - 1)
         'Modified by Lydia 2023/12/22 + strYYMMSql
         m_strKey = " select cp09 from caseprogress,Acc090 where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "' AND CP12=A0901(+)" & strYYMMSql
         'Add By Sindy 2020/5/12
         '以操作人員之st15再抓ACC090之A0911，再以進度CP12再抓ACC090之A0911，若與操作人員帶出之A0911相同才可顯示進度。卷宗區也是。
         'st05 in ('00',’01’)人員不受上述限制
         'Modify By Sindy 2022/5/23 再加入系統特殊設定「全所智權部主管」的人員也不限制。
         If InStr("00,01", Pub_strUserST05) = 0 And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) = 0 And m_CP01 = "TT" And m_CP02 = "999999" Then
            m_strKey = m_strKey + " AND A0911=(select a90.A0911 from acc090 a90 where a90.a0901='" & Pub_StrUserSt15 & "') "
         End If
         '2020/5/12 END
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
     
   '案件資料
   strSql = "Select PA01||'-'||PA02||'-'||PA03||'-'||PA04 as 本所案號,PA05||PA06||PA07 as 案件名稱,PA09 as 申請國家" & _
            " From Patent" & _
            " Where PA01='" & m_CP01 & "' And PA02='" & m_CP02 & "' And PA03='" & m_CP03 & "' And PA04='" & m_CP04 & "'"
   strSql = strSql & " union Select TM01||'-'||TM02||'-'||TM03||'-'||TM04 as 本所案號,TM05 as 案件名稱,TM10 as 申請國家" & _
            " From Trademark" & _
            " Where TM01='" & m_CP01 & "' And TM02='" & m_CP02 & "' And TM03='" & m_CP03 & "' And TM04='" & m_CP04 & "'"
   strSql = strSql & " union Select SP01||'-'||SP02||'-'||SP03||'-'||SP04 as 本所案號,SP05||SP06||SP07 as 案件名稱,SP09 as 申請國家" & _
            " From Servicepractice" & _
            " Where SP01='" & m_CP01 & "' And SP02='" & m_CP02 & "' And SP03='" & m_CP03 & "' And SP04='" & m_CP04 & "'"
   strSql = strSql & " union Select HC01||'-'||HC02||'-'||HC03||'-'||HC04 as 本所案號,HC06 as 案件名稱,'000' as 申請國家" & _
            " From Hirecase" & _
            " Where HC01='" & m_CP01 & "' And HC02='" & m_CP02 & "' And HC03='" & m_CP03 & "' And HC04='" & m_CP04 & "'"
   strSql = strSql & " union Select LC01||'-'||LC02||'-'||LC03||'-'||LC04 as 本所案號,LC05||LC06||LC07 as 案件名稱,LC15 as 申請國家" & _
            " From Lawcase" & _
            " Where LC01='" & m_CP01 & "' And LC02='" & m_CP02 & "' And LC03='" & m_CP03 & "' And LC04='" & m_CP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("本所案號")) Then lblCaseNo.Caption = rsTmp.Fields("本所案號")
      If Not IsNull(rsTmp.Fields("案件名稱")) Then lblCaseName.Caption = rsTmp.Fields("案件名稱")
      m_Nation = ""
      If Not IsNull(rsTmp.Fields("申請國家")) Then m_Nation = rsTmp.Fields("申請國家")
      'Added by Lydia 2020/02/06 如果是台灣各區
      If m_Nation <> "" And Val(m_Nation) <= 10 Then
          m_Nation = "000"
      End If
      'end 2020/02/06
      m_CP13 = ShowCurrCP13(m_CP01, m_CP02, m_CP03, m_CP04, m_Nation)  'Added by Lydia 2020/03/03
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Call cmdExit_Click
      Exit Function
   End If
   rsTmp.Close
   
   Call ChkModifyLimits
   'Modify By Sindy 2023/2/8 改檢查權限的位置,從上方Move至此處
   '檢查權限
   'Modify by Sindy 2020/9/10 + , Me.Name
'         If CheckSR09(strUserNum, m_CP01, "Y", , m_CP01, m_CP02, m_CP03, m_CP04, Me.Name) = False Then
'            'tmpBol = fnCancelNowFormAndShowParentForm(Me)
'            Screen.MousePointer = vbDefault
'            Set rsTmp = Nothing
'            Call cmdExit_Click
'            Exit Function
'         End If
   If CheckLimits() = False Then
      'tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Screen.MousePointer = vbDefault
      Set rsTmp = Nothing
      Call cmdExit_Click
      Exit Function
   End If
   
   Call ReadAttachFile(0) '原始檔區
   If pub_QL04 <> "" Then InsertQueryLog (GRD1(0).Rows - 1) 'Add By Sindy 2025/8/7
   Call ReadAttachFile(1) '暫存區
   
   'Add By Sindy 2020/9/24 增加開放核判人員也可以放未完稿暫存區附件
   m_bolAddTmpFile = False
   If UCase(TypeName(m_PrevForm)) = "FRM100101_2" Then
      If m_PrevForm.bolEmpFlow = True Then '從歷程維護作業點進來的
         strSql = "select cp09 from caseprogress,engineerprogress" & _
                  " where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
                  " and cp09=ep02 and (ep04='" & strUserNum & "' or ep40='" & strUserNum & "')"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            m_bolAddTmpFile = True
         End If
         rsTmp.Close
      End If
   End If
   '2020/9/24 END
   
   'Add By Sindy 2020/3/16
   'Modify By Sindy 2020/9/24 + bolAddTmpFile = True
   If UCase(TypeName(m_PrevForm)) = "FRM090201_2" Or _
      UCase(TypeName(m_PrevForm)) = "FRM090201_B" Or _
      m_bolAddTmpFile = True Then
      Me.Caption = "未完稿暫存區"
      
      'Add By Sindy 2020/9/28
      If m_bolAddTmpFile = False Then
      '2020/9/28 END
         SSTab1.TabVisible(0) = False: Text2.Visible = False
      End If
      'Add By Sindy 2020/10/21
      If m_bolAddTmpFile = True Then
         SSTab1.Tab = 0 '原始檔區
      Else
      '2020/10/21 END
         SSTab1.Tab = 1 '暫存區
      End If
      '可新增刪除暫存區
      cmdAddAtt.Enabled = True
      cmdRemAtt.Enabled = True
      cmdModify.Visible = False
      cmdReName.Visible = False
      GRD1(1).col = 0
      GRD1(1).row = 1
      '資料列反白
      GRD1(1).TextMatrix(1, 0) = "V"
      For jj = 0 To GRD1(1).Cols - 1
         GRD1(1).col = jj
         GRD1(1).CellBackColor = &HFFC0C0
      Next jj
   Else
      SSTab1.Tab = 0
      'Add By Sindy 2023/11/23
      If Left(PUB_GetST03(m_CP13), 2) = "F2" Then
         SSTab1.TabVisible(1) = False: Text2.Visible = False
      End If
      '2023/11/23 END
   End If
   '2020/3/16 END
   
   QueryData = True
   Screen.MousePointer = vbDefault
   Me.Enabled = True

EXITSUB:
   Set rsTmp = Nothing
End Function

'檢查可新增刪除的權限
Private Sub ChkModifyLimits()
   'm_identity.身份：A.業務助理 C.電腦中心 S.智權人員 E.承辦人/工程師 F.程序人員 W.檔案室 P.繪圖人員 T.中打室
   'Modify By Sindy 2020/3/16 改成共用函數
   m_identity = PUB_ChkCPPAndCPFLimits(m_CP01, m_CP02, m_CP03, m_CP04, m_CP13)
   'Added by Lydia 2022/04/07 (蔡亮丞A3033,許廷璋A3032)以上二位，請開系統與打字室一樣的權限（除限閱）
                    '1.比對打字室的權限，增加”案件基本資料-打字室”的權限，其餘權限大致相同。
                    '2.原始檔區的新增檔案權限。
   If InStr("A3032,A3033", strUserNum) > 0 Then
       m_identity = "T"
   End If
   'end 2022/04/07
   
   cmdAddAtt.Enabled = False
   cmdRemAtt.Enabled = False
   cmdCopy.Enabled = False '複製 'Add By Sindy 2021/10/21
   If (m_identity <> "" And InStr("C,F,P,T", m_identity) > 0) Or _
      Pub_StrUserSt03 = "F10" Or Pub_StrUserSt03 = "F11" Or _
      Pub_StrUserSt03 = "P10" Or _
      Pub_StrUserSt03 = "P20" Or Pub_StrUserSt03 = "P21" Then

        cmdAddAtt.Enabled = True
        cmdRemAtt.Enabled = True
        cmdCopy.Enabled = True '複製 'Add By Sindy 2021/10/21
   End If
   
   Call CmdLimits
   
   'Added by Lydia 2024/11/27 外專承辦：該案承辦＋(若承辦請假)案件職代----Anny
   m_strF23User = ""
   If Pub_StrUserSt03 = "F23" Then
      strExc(0) = PUB_GetFCPSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)
      If strExc(0) <> "" Then
         m_strF23User = m_strF23User & "," & strExc(0)
         strExc(1) = GetCaseDutyAgent(strExc(0), "", False, , True, "A")
         If strExc(1) <> "" Then
            m_strF23User = m_strF23User & ";" & strExc(1)
         End If
         m_strF23User = Mid(m_strF23User, 2)
      End If
   End If
   'end 2024/11/278
End Sub

'檔案更名和Word維護-修改的顯示／隱藏
Private Sub CmdLimits()

   cmdModify.Visible = False 'Word維護-修改
   cmdReName.Visible = False  '檔案更名
   If Pub_StrUserSt03 = "M51" Then
      If SSTab1.Tab = 0 Then
         cmdModify.Visible = True
         cmdReName.Visible = True
      End If
   'Added by Lydia 2020/03/31 開放人員，可以更名
   ElseIf (m_identity <> "" And InStr("C,F,P,T", m_identity) > 0) Or _
        Pub_StrUserSt03 = "F10" Or Pub_StrUserSt03 = "F11" Or _
        Pub_StrUserSt03 = "P10" Or _
        Pub_StrUserSt03 = "P20" Or Pub_StrUserSt03 = "P21" Then
        If SSTab1.Tab = 0 Then
           cmdReName.Visible = True
        End If
   'end 2020/03/31
   End If
End Sub

'查詢附件檔
'Modified by Lydia 2020/02/06 改成共用模組
'Private Function ReadAttachFile() As Boolean
Public Function ReadAttachFile(Index As Integer) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strConSql As String, strVal As String
   
   ReadAttachFile = True
   
   Screen.MousePointer = vbHourglass
   KillAttach
   GRD1(Index).Clear
   Call SetGrd(Index)
   
   'Modify By Sindy 2020/3/16
   If Not (ChkDelF.Visible = True And ChkDelF.Value = 1) Then
      strConSql = " AND (cpf10<>'D' or cpf10 is null)"
   End If
   If Index = 0 Then '原始檔區
      strConSql = strConSql & " And (cpf11<>'Z' or cpf11 is null)"
   Else '暫存區
      strConSql = strConSql & " And cpf11='Z'"
   End If
   
   'Added by Sindy 2022/5/4
   strCon2 = ""
   If ChkDelMsg.Value = 1 Then '排除郵件檔
       strCon2 = strCon2 & " and Upper(cpf02) not like '%.MSG' "
   End If
   '2022/5/4 END
   
   '有電子檔的文號
   strVal = "Select distinct cpf01 From Casepaperfile" & _
               " Where cpf01 in(" & m_strKey & ")" & strConSql '& " group by cpf01"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While Not rsTmp.EOF
'         strExistsFRev = strExistsFRev & "'" & Trim(rsTmp.Fields(0)) & "',"
'         rsTmp.MoveNext
'      Loop
'      strExistsFRev = Left(strExistsFRev, Len(strExistsFRev) - 1)
'   End If
'   rsTmp.Close
   '2020/3/16 END
   
   'Modified by Lydia 2020/02/06
   'strSql = "Select ' ' as V,cp09 as 總收文號,sqldatet(cp05) as 收文日,Decode('" & m_Nation & "','000',CPM03,CPM04) as 案件性質,decode(cpf02,null,'',cpf02||' ('||Round(cpf03 / 1024, 2)||' KB)') as 檔案名稱,cp09,CP10,sqldatet(cpf08)||' '||sqltime(cpf09)||decode(' ('||cpf05||cpf11||')',' ()','',' ('||cpf05||cpf11||')') as 最後修改時間,' ',cpf08||cpf09 as 副檔名,CP82,CPF10,CP43" & _
            " From Casepaperfile,caseprogress,Casepropertymap" & _
            " Where cp09 in(" & m_strKey & ")" & _
            " And cp01='" & m_CP01 & "' And cp02='" & m_CP02 & "' And cp03='" & m_CP03 & "' And cp04='" & m_CP04 & "'" & _
            " And cp09=cpf01(+) and (cpf10 is null or cpf10='X'" & IIf(ChkDelF.Visible = True And ChkDelF.Value = 1, " or cpf10='D'", "") & ")" & _
            " And cp01=cpm01(+) And cp10=cpm02(+)" & strConSql & _
            " order by SQLDatet2(CP05) desc,CP66 desc,CP67 desc,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3') desc,CP09 desc,cpf08 desc,cpf09 desc"
   'Modified by Lydial 2020/03/03 +異動人員CPF05
   'Modify By Sindy 2022/5/4 + & strCon2
   'Modified by Lydia 2023/08/24 +CPF13
   'Modified by Lydia 2023/12/22 + strYYMMSql
   strSql = "Select ' ' as V,cp09 as 總收文號,sqldatet(cp05) as 收文日,Decode('" & m_Nation & "','000',CPM03,CPM04) as 案件性質,decode(cpf02,null,'',cpf02||' ('||Round(cpf03 / 1024, 2)||' KB)') as 檔案名稱,cp09,CP10," & _
            "sqldatet(cpf08)||' '||sqltime(cpf09)||decode(' ('||cpf05||cpf11||')',' ()','',' ('||cpf05||cpf11||')') as 最後修改時間,' ' as 檔案更名,cpf08||cpf09 as CPF0809,CP82,CPF10,CP43,CPF02,CPF05,CPF06,CPF11,CPF07,CPF14,CPF15,CPF16,CP66,CP67,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3') ColSort,cpf08,cpf09,CPF13" & _
            " From Casepaperfile,caseprogress,Casepropertymap" & _
            " Where cp09 in(" & m_strKey & ") And cpf01 is not null" & _
            " And cp01='" & m_CP01 & "' And cp02='" & m_CP02 & "' And cp03='" & m_CP03 & "' And cp04='" & m_CP04 & "'" & _
            " And cp09=cpf01(+)" & _
            " And cp01=cpm01(+) And cp10=cpm02(+)" & strConSql & strCon2 & strYYMMSql
   'Modify By Sindy 2020/3/16 要抓無電子檔的進度資料
   'IIf(strExistsFRev = "", "", " and cp09 not in(" & strExistsFRev & ")")
   'Modified by Lydia 2023/08/24 +'' AS CPF13
   'Modified by Lydia 2023/12/22 + strYYMMSql
   strSql = strSql & " union " & _
            "Select ' ' as V,cp09 as 總收文號,sqldatet(cp05) as 收文日,Decode('" & m_Nation & "','000',CPM03,CPM04) as 案件性質,' ' as 檔案名稱,cp09,CP10," & _
            "' ' as 最後修改時間,' ' as 檔案更名,' ' as CPF0809,CP82,' 'CPF10,CP43,' ' CPF02,' ' CPF05,0 CPF06,' ' CPF11,0 CPF07,'' CPF14,0 CPF15,0 CPF16,CP66,CP67,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3') ColSort,0 cpf08,0 cpf09,'' AS CPF13" & _
            " From caseprogress,Casepropertymap,Acc090" & _
            " Where cp01='" & m_CP01 & "' And cp02='" & m_CP02 & "' And cp03='" & m_CP03 & "' And cp04='" & m_CP04 & "'" & strYYMMSql & _
            " And cp01=cpm01(+) And cp10=cpm02(+) AND CP12=A0901(+)" & _
            IIf(InStr(UCase(m_strKey), "FROM") = 0, " And cp09 in(" & m_strKey & ")", "") & _
            " And cp09 not in(" & strVal & ")"
            'Add By Sindy 2020/5/12
            '以操作人員之st15再抓ACC090之A0911，再以進度CP12再抓ACC090之A0911，若與操作人員帶出之A0911相同才可顯示進度。卷宗區也是。
            'st05 in ('00',’01’)人員不受上述限制
            'Modify By Sindy 2022/5/23 再加入系統特殊設定「全所智權部主管」的人員也不限制。
            If InStr("00,01", Pub_strUserST05) = 0 And InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) = 0 And m_CP01 = "TT" And m_CP02 = "999999" Then
               strSql = strSql + " AND A0911=(select a90.A0911 from acc090 a90 where a90.a0901='" & Pub_StrUserSt15 & "') "
            End If
            '2020/5/12 END
   '2020/3/16 END
   'strSql = strSql & " order by SQLDatet2(CP05) desc,CP66 desc,CP67 desc,DECODE(SUBSTR(CP09,1,1),'A','1','C','2','3') desc,CP09 desc,cpf08 desc,cpf09 desc"
   strSql = strSql & " order by 收文日 desc,CP66 desc,CP67 desc,ColSort desc,CP09 desc,cpf08 desc,cpf09 desc"
   'end 2020/02/06
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1(Index).Recordset = rsTmp
      
      Call QueryDelData(Index, rsTmp.RecordCount)
   Else
'      If QueryDelData(Index, 0) = False Then
         rsTmp.Close
         Set rsTmp = Nothing
         ReadAttachFile = False
         Exit Function
'      End If
   End If
   rsTmp.Close
   
EXITSUB:
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
End Function

Private Sub recovercolor(Index As Integer, intRow As Integer)
Dim j As Integer
   
   GRD1(Index).row = intRow
   If Trim(GRD1(Index).TextMatrix(intRow, colCPF10)) = "D" Then
      For j = 0 To GRD1(Index).Cols - 1
         GRD1(Index).col = j
         GRD1(Index).CellBackColor = &H8080FF
      Next j
   End If
End Sub

Private Function QueryDelData(Index As Integer, intQuyCnt As Integer) As Boolean
Dim Rs As New ADODB.Recordset
Dim strFileName As String, strFileName_jj As String, intLen As Integer 'Add By Sindy 2013/9/24
Dim strCP09 As String
Dim bolInsData As Boolean, bolhave As Boolean
Dim intRow As Integer
Dim strSpecTxt As String, strCP64 As String 'Add By Sindy 2024/9/30
   
   QueryDelData = False
   
   GRD1(Index).Visible = False
   strCP09 = ""
   For ii = 1 To GRD1(Index).Rows - 1
      strSpecTxt = "": strCP64 = ""
      If strCP09 = Trim(GRD1(Index).TextMatrix(ii, colCp09)) Then
         GRD1(Index).TextMatrix(ii, 1) = "" '總收文號
         GRD1(Index).TextMatrix(ii, 2) = "" '收文日
         GRD1(Index).TextMatrix(ii, colCPM) = "" '案件性質
      Else
         'Add By Sindy 2015/10/2 判斷有相關總收文號才做較快
         If GRD1(Index).TextMatrix(ii, colCP43) <> "" Then
            GRD1(Index).TextMatrix(ii, colCPM) = GRD1(Index).TextMatrix(ii, colCPM) & PUB_GetRelateCasePropertyName(GRD1(Index).TextMatrix(ii, colCp09), "1")
         End If
         '2015/10/2 END
         'Add By Sindy 2024/9/27 依系統別調整案件性質的顯示內容
         Select Case m_CP01
            '專利
            Case "CFP", "FCP", "P"
               strExc(0) = "select * from caseprogress where cp09='" & GRD1(Index).TextMatrix(ii, colCp09) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strCP64 = "" & RsTemp.Fields("cp64")
               End If
               If m_CP01 = "P" And GRD1(Index).TextMatrix(ii, colCP10) = "1604" Then
                  If InStr(strCP64, "專利權消滅") > 0 Then
                     strSpecTxt = Mid(strCP64, InStr(strCP64, "專利權消滅"), 15)
                  ElseIf InStr(strCP64, "消滅") > 0 Then
                     strSpecTxt = Mid(strCP64, InStr(strCP64, "消滅"), 12)
                  End If
               ElseIf m_CP01 = "CFP" And GRD1(Index).TextMatrix(ii, colCP10) = "1604" Then
                  If InStr(strCP64, "消滅") > 0 Then
                     strSpecTxt = Mid(strCP64, InStr(strCP64, "消滅"), 12)
                  End If
               End If
               If strSpecTxt <> "" Then
                  GRD1(Index).TextMatrix(ii, colCPM) = strSpecTxt
               End If
            '法務
            Case "CFL", "FCL", "L", "LIN"
               strExc(0) = "select * from caseprogress where cp09='" & GRD1(Index).TextMatrix(ii, colCp09) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If "" & RsTemp.Fields("cp46") = "19221111" Then
                     strSpecTxt = "回執退件日:" & ChangeWStringToTDateString("" & RsTemp.Fields("cp47")) & ";"
                  End If
               End If
               strSpecTxt = strSpecTxt & "" & RsTemp.Fields("CP64") & "(" & GRD1(Index).TextMatrix(ii, colCPM) & ")"
               If strSpecTxt <> "" Then
                  strSpecTxt = Mid(Trim(strSpecTxt), 1, 500)
                  GRD1(Index).TextMatrix(ii, colCPM) = strSpecTxt
               End If
            '顧問
            Case "LA"
               strExc(0) = "select * from caseprogress where cp09='" & GRD1(Index).TextMatrix(ii, colCp09) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If "" & RsTemp.Fields("cp46") = "19221111" Then
                     strSpecTxt = "回執退件日:" & ChangeWStringToTDateString("" & RsTemp.Fields("cp47")) & ";"
                  End If
               End If
               If GRD1(Index).TextMatrix(ii, colCP10) = "0" Then
                  strSpecTxt = strSpecTxt & ChangeWStringToTDateString("" & RsTemp.Fields("cp53")) & "--" & ChangeWStringToTDateString("" & RsTemp.Fields("cp54"))
               Else
                  strSpecTxt = strSpecTxt & "" & RsTemp.Fields("CP64") & "(" & GRD1(Index).TextMatrix(ii, colCPM) & ")"
               End If
               If strSpecTxt <> "" Then
                  strSpecTxt = Mid(Trim(strSpecTxt), 1, 50)
                  GRD1(Index).TextMatrix(ii, colCPM) = strSpecTxt
               End If
         End Select
         '2024/9/27 END
      End If
      strCP09 = Trim(GRD1(Index).TextMatrix(ii, colCp09))
      '檢查檔案重覆另外命名
      If Trim(GRD1(Index).TextMatrix(ii, colFN)) <> "" Then
         'Modified by Lydia 2020/02/06
         'If InStrRev(Trim(GRD1(Index).TextMatrix(ii, colFN)), " (") > 0 Then
         '   strFileName = Left(Trim(GRD1(Index).TextMatrix(ii, colFN)), InStrRev(Trim(GRD1(Index).TextMatrix(ii, colFN)), " (") - 1)
        ' Else
        '    strFileName = Trim(GRD1(Index).TextMatrix(ii, colFN))
        ' End If
         strFileName = GetFileName(Trim(GRD1(Index).TextMatrix(ii, colFN)))
         'end 2020/02/06
         
         GRD1(Index).TextMatrix(ii, colReFN) = strFileName  'Memo by Lydia 2020/02/06 檔案更名:預設原檔名
         For jj = 1 To ii - 1
             'Modified by Lydia 2020/02/06
            'If InStrRev(Trim(GRD1(Index).TextMatrix(jj, colFN)), " (") > 0 Then
            '   strFileName_jj = Left(Trim(GRD1(Index).TextMatrix(jj, colFN)), InStrRev(Trim(GRD1(Index).TextMatrix(jj, colFN)), " (") - 1)
            'Else
            '   strFileName_jj = Trim(GRD1(Index).TextMatrix(jj, colFN))
            'End If
            strFileName_jj = GetFileName(Trim(GRD1(Index).TextMatrix(jj, colFN)))
            'end 2020/02/06
            If strFileName_jj = strFileName Then
               If InStrRev(UCase(strFileName), "DWG.") > 0 Then
                  intLen = InStrRev(UCase(strFileName), "DWG.")
               Else
                  intLen = InStrRev(strFileName, ".")
               End If
               'Memo by Lydia 2020/02/06 檔名後面+檔案最後日期+時間
               GRD1(Index).TextMatrix(ii, colReFN) = Replace(Left(strFileName, intLen - 1) & "_" & Trim(GRD1(Index).TextMatrix(ii, colCPF08)) & Right(strFileName, Len(strFileName) - (intLen - 1)), "._", ".")
               Exit For
            End If
         Next jj
      End If
      Call recovercolor(Index, ii) 'Add By Sindy 2014/6/26
   Next ii
   GRD1(Index).col = 0
   GRD1(Index).row = 1
   
   'Add By Sindy 2020/3/17
   For intRow = 8 To GRD1(Index).Cols - 1
      GRD1(Index).ColWidth(intRow) = 0
   Next intRow
   '2020/3/17 END
   
   GRD1(Index).Visible = True
End Function

Private Sub ChkDelF_Click()
   Call ReadAttachFile(SSTab1.Tab)
End Sub

'Added by Sindy 2022/5/4 不顯示郵件
Private Sub ChkDelMsg_Click()
   Call ReadAttachFile(SSTab1.Tab)
End Sub

'Add By Sindy 2021/10/21 複製到...
Private Sub cmdCopy_Click()
Dim strSaveFiles As String
Dim strRecvNo As String
   
   m_CP09 = "" 'Add By Sindy 2025/7/21
   For ii = 1 To GRD1(0).Rows - 1
      GRD1(0).row = ii
      GRD1(0).col = 1
      If GRD1(0).TextMatrix(ii, 0) = "V" And Trim(GRD1(0).TextMatrix(ii, 4)) <> "" Then
         
         '只能勾選是總收文號的資料列做新增
         If InStr("A,B,C,D", Left(Trim(GRD1(0).TextMatrix(ii, 5)), 1)) = 0 Then
            MsgBox "該筆資料並非總收文號，不可複製！"
            Exit Sub
         ElseIf Len(Trim(Trim(GRD1(0).TextMatrix(ii, 5)))) <> 9 Then '總收文號+D
            MsgBox "該筆資料並非總收文號，不可複製！"
            Exit Sub
         End If
         
         If m_CP09 <> "" And m_CP09 <> Trim(GRD1(0).TextMatrix(ii, 5)) Then
            MsgBox "可點選多筆附件做複製，但須同一筆總收文號！"
            Exit Sub
         End If
         
         m_CP09 = Trim(GRD1(0).TextMatrix(ii, 5))
         m_CP10 = Trim(GRD1(0).TextMatrix(ii, 6))
         strSaveFiles = strSaveFiles & "&" & Trim(GRD1(0).TextMatrix(ii, 5))
         If InStr(strRecvNo, Trim(GRD1(0).TextMatrix(ii, 5))) = 0 Then
            strRecvNo = strRecvNo & ",'" & Trim(GRD1(0).TextMatrix(ii, 5)) & "'"
         End If
         strSaveFiles = strSaveFiles & "  " & GetFileName(Trim(GRD1(0).TextMatrix(ii, 4)))
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
   frm100101_L_4.Show vbModal
   
   'Add By Sindy 2025/7/21
   Call ReadAttachFile(0) '原始檔區
   Call ReadAttachFile(1) '暫存區
   '2025/7/21 END
End Sub

'結束
Private Sub cmdExit_Click()
   'Modified by Lydia 2020/02/20 改到Form_Unload
   'm_PrevForm.Show
   
   'Modify By Sindy 2025/6/9
   'Unload Me
   If Not m_PrevForm Is Nothing Then
      If Left(UCase(m_PrevForm.Name), 6) = "FRM100" Then
         tmpBol = fnCancelNowFormAndShowParentForm(Me) '下一筆
      Else
          Unload Me
      End If
   Else
   '2025/6/9 END
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

'Added by Lydia 2020/02/06 Word維護-修改
Private Sub cmdModify_Click()
Dim hLocalFile As Long
Dim stFileName As String
Dim bolIsSelect As Boolean
Dim bolAsked As Boolean
Dim Index As Integer 'Add By Sindy 2020/3/16
   
   KillAttach 'Added by Lydia 2020/02/20
   Index = SSTab1.Tab 'Add By Sindy 2020/3/16
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   For ii = 1 To Me.GRD1(Index).Rows - 1
      If Me.GRD1(Index).TextMatrix(ii, 0) = "V" Then
         If Trim(Me.GRD1(Index).TextMatrix(ii, colFN)) <> "" Then
            '讀取檔案名稱
            stFileName = Trim(Me.GRD1(Index).TextMatrix(ii, colCPF02))
            If Right(UCase(stFileName), 4) = ".DOC" Or Right(UCase(stFileName), 5) = ".DOCX" Then
                bolIsSelect = True
                '清除反白
                Me.GRD1(Index).col = 0
                Me.GRD1(Index).row = ii
                Me.GRD1(Index).TextMatrix(ii, 0) = ""
                For jj = 0 To Me.GRD1(Index).Cols - 1
                   Me.GRD1(Index).col = jj
                   Me.GRD1(Index).CellBackColor = QBColor(15)
                Next jj
                If stFileName <> "" Then
                   m_CP09 = Trim(Me.GRD1(Index).TextMatrix(ii, colCp09))
                   'Modified by Lydia 2020/02/20 改放在使用者\原檔名
                   'If GetAttachFile(m_CP09, stFileName, App.path & "\$TEMP", True) = False Then
                   If GetAttachFile(m_CP09, stFileName, m_AttachPath & "\" & Trim(GRD1(Index).TextMatrix(ii, colReFN)), True) = False Then
                        MsgBox "無法開啟檔案[ " & stFileName & " ]！"
                        GoTo EXITSUB
                   End If
                   '檢查若已定稿維護畫面已開啟時確認是否只開Word並提醒無法直接上傳
                   If PUB_CheckFormExist("frm100101_M_1") Then
                      If Not bolAsked Then
                         MsgBox "原始檔Word維護畫面已開啟，此次修改將以 Word 開啟且無法直接上傳！", vbExclamation
                         bolAsked = True
                      End If
                   End If
                   If bolFileOpen = True Then
                       Call frm100101_M_1.SetFormParent(Me, Me.lblCaseNo.Caption, _
                                                   Me.GRD1(Index).TextMatrix(ii, colCp09), Me.GRD1(Index).TextMatrix(ii, colCPF02), "D", Me.GRD1(Index).TextMatrix(ii, colCP10))
                       frm100101_M_1.Show
                       Exit For 'Added by Lydia 2020/02/20 同時只開啟一個檔案
                   End If
                End If
            End If
         End If
      End If
   Next ii
   If bolIsSelect = False Then
      MsgBox "無Word檔案可開啟！"
   End If
   
EXITSUB:

   Screen.MousePointer = vbDefault

End Sub

'Added by Lydia 2020/02/06 更名
Private Sub cmdReName_Click()
Dim strReName As String
Dim intChkCnt As Integer
Dim strNewFile As String
Dim strCPF02 As String, strChkCPF02 As String
Dim Index As Integer 'Add By Sindy 2020/3/16
   
On Error GoTo ErrHnd
   
   Index = SSTab1.Tab 'Add By Sindy 2020/3/16
   
   intChkCnt = 0
   For ii = 1 To Me.GRD1(Index).Rows - 1
      Me.GRD1(Index).row = ii
      Me.GRD1(Index).col = 1
      If (Me.GRD1(Index).TextMatrix(ii, 0) = "V" Or Me.GRD1(Index).TextMatrix(ii, 0) = "v") And Trim(Me.GRD1(Index).TextMatrix(ii, colFN)) <> "" Then
         m_CP09 = Trim(Me.GRD1(Index).TextMatrix(ii, colCp09))
         m_CP10 = Trim(Me.GRD1(Index).TextMatrix(ii, colCP10))
         strCPF02 = Trim(Me.GRD1(Index).TextMatrix(ii, colCPF02))
         If m_CP09 <> "" And strCPF02 <> "" Then
            intChkCnt = intChkCnt + 1
         End If
      End If
   Next ii
   strCPF02 = GetFileName(strCPF02)
   strChkCPF02 = strCPF02
   If intChkCnt = 0 Or strCPF02 = "" Then
      MsgBox "請勾選一筆欲更名的電子檔！"
      Exit Sub
   ElseIf intChkCnt > 1 Then
      MsgBox "只可勾選一筆資料做更名！"
      Exit Sub
   End If
   
   'Added by Morgan 2020/7/28
   If Right(UCase(strCPF02), 14) = ".ENCRYPTED.ZIP" Then
      MsgBox "加密壓縮檔不可更名！", vbExclamation
      Exit Sub
   End If
   'end 2020/7/28
   
ShowInput:
   strNewFile = InputBox("確定是否「更名」？", "更名！", strCPF02)
   If UCase(strNewFile) = UCase(strChkCPF02) Then
      MsgBox "請輸入欲更改的電子檔名！"
      strCPF02 = strNewFile
      GoTo ShowInput
   End If
   
   If Trim(strNewFile) = "" Then
      Exit Sub
   Else
      strNewFile = PUB_GetReNameMax(strNewFile, m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, 75)
      'intQueryKind= 0 , 原始檔除了English_Vers外,不可上傳PDF
      If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, strNewFile, strReName, True, 0) = False Then
         strCPF02 = strNewFile
         GoTo ShowInput
      End If
      
      If UCase(strReName) = UCase(strChkCPF02) Then
         MsgBox "請輸入欲更改的電子檔名！"
         strCPF02 = strReName
         GoTo ShowInput
      End If
      If Right(UCase(strReName), 4) = UCase(".del") Then
         MsgBox "新的電子檔名最後面不能是(.del)！"
         strCPF02 = strReName
         GoTo ShowInput
      End If

      For ii = 1 To Me.GRD1(Index).Rows - 1
         Me.GRD1(Index).row = ii
         Me.GRD1(Index).col = 1
         If Me.GRD1(Index).TextMatrix(ii, 1) <> "" Then Me.GRD1(Index).Tag = Me.GRD1(Index).TextMatrix(ii, 1)
         If Me.GRD1(Index).Tag = m_CP09 Then
            If UCase(strReName) = UCase(GetFileName(Trim(Me.GRD1(Index).TextMatrix(ii, colFN)))) Then
               MsgBox "檔名重覆，請輸入欲更改的電子檔名！"
               strCPF02 = strReName
               GoTo ShowInput
            End If
         End If
      Next ii
      
      strSql = "update casepaperfile set CPF02='" & strReName & "' where cpf01='" & m_CP09 & "' and CPF02='" & strChkCPF02 & "'"
      Pub_SaveLog strUserNum, strSql
      cnnConnection.Execute strSql
     
      Call ReadAttachFile(Index)
   End If
   
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Activate()
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
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   ReDim m_FilesRemoved(0)
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   Call ChangSelect
   cmdCopy.Visible = False
   'Add By Sindy 2014/6/25
   If Pub_StrUserSt03 = "M51" Then
      ChkDelF.Visible = True
      cmdCopy.Visible = True
   Else
      ChkDelF.Visible = False
   End If
   '2014/6/25 END
   Text2.ZOrder '暫存區文字框
End Sub

Private Sub Form_Unload(Cancel As Integer)
   KillAttach
   'Added by Lydia 2020/02/20
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" Then
      m_PrevForm.Show
   End If
   'end 2020/02/20
   Set m_PrevForm = Nothing
   Set frm100101_M = Nothing
End Sub

Private Sub KillAttach()
'Modified by Lydia 2020/02/20 改用模組
'On Error Resume Next
'   If Dir(m_AttachPath & "\.") <> "" Then
'      Kill m_AttachPath & "\*.*"
'   End If
    PUB_KillTempFile strUserNum & "\*.*"
End Sub

Private Sub grd1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1(Index), x, y, nCol, nRow
If nCol < 0 Then nCol = 0 'Add By Sindy 2020/1/14
If nRow < 0 Then nRow = 0 'Add By Sindy 2020/1/14
GRD1(Index).col = nCol
GRD1(Index).row = nRow
End Sub

Private Sub grd1_SelChange(Index As Integer)
GRD1(Index).Visible = False
If GRD1(Index).MouseRow <> 0 And Trim(GRD1(Index).TextMatrix(GRD1(Index).MouseRow, colCp09)) <> "" Then
   If GRD1(Index).TextMatrix(GRD1(Index).MouseRow, 0) = "V" Then
      '清除反白
      GRD1(Index).TextMatrix(GRD1(Index).MouseRow, 0) = ""
      GRD1(Index).row = GRD1(Index).MouseRow
      For jj = 0 To GRD1(Index).Cols - 1
         GRD1(Index).col = jj
         GRD1(Index).CellBackColor = QBColor(15)
      Next jj
      Call recovercolor(Index, GRD1(Index).MouseRow) 'Add By Sindy 2014/6/26
   Else
      '資料列反白
      GRD1(Index).TextMatrix(GRD1(Index).MouseRow, 0) = "V"
      GRD1(Index).row = GRD1(Index).MouseRow
      For jj = 0 To GRD1(Index).Cols - 1
         GRD1(Index).col = jj
         GRD1(Index).CellBackColor = &HFFC0C0
      Next jj
   End If
End If
GRD1(Index).Visible = True
End Sub

'Modified by Lydia 2023/08/24 + bolDel As Boolean
Private Function DeleteFile(strCP09 As String, strFileName As String, bolDel As Boolean) As Boolean
   
On Error GoTo ErrHand
   
   DeleteFile = True
   Screen.MousePointer = vbHourglass
   'Modified by Lydia 2023/08/24
   'If DelAttFile_File(lblCaseNo.Caption, strCP09, strFileName) = False Then GoTo ErrHand
   If DelAttFile_File(lblCaseNo.Caption, strCP09, strFileName, , , , bolDel) = False Then GoTo ErrHand
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

'Modified by Lydia 2020/02/06 + bolOpen
Private Function GetAttachFile(ByVal strCP09 As String, ByRef pFileName As String, Optional pSavePath As String, Optional bolOpen As Boolean = False) As Boolean
   Dim stAttPath As String
   Dim lngSize As Long
   Dim iFileNo As Integer
   Dim bytes() As Byte

On Error GoTo ErrHnd

   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
   Else
      'Add By Sindy 2013/12/27
      If InStr(pSavePath, m_AttachPath) > 0 Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
      End If
      '2013/12/27 END
      stAttPath = pSavePath
   End If
   
   bolFileOpen = False 'Added by Lydia 2020/02/06
   
   'Added by Morgan 2015/4/28
   'Modified by Morgan 2015/5/22 FTP上線
   strExc(0) = "select cpf13 from casepaperfile where cpf01='" & strCP09 & "' and cpf02='" & ChgSQL(pFileName) & "'"
   'Modified by Lydia 2020/02/06 彈訊息
   'intI = 1
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp(0)) Then
         pFileName = stAttPath
         GetAttachFile = PUB_GetFtpFile("" & RsTemp(0), stAttPath, "CASEPAPERFILE", True)
         'Added by Lydia 2020/02/06 下載檔案並且開啟
         If GetAttachFile = True And bolOpen = True Then
             If PUB_OpenWord(stAttPath) = True Then
                 bolFileOpen = True
             End If
         End If
         'end 2020/02/06
      End If
   End If
   Exit Function
   'end 2015/4/28
   
'Removed by Morgan 2015/5/22 不再存DB
'   strExc(0) = "select * from casepaperfile where cpf01='" & strCP09 & "' and cpf02='" & ChgSQL(pFileName) & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If Dir(stAttPath) <> "" Then Kill stAttPath
'      With RsTemp
'         lngSize = Val(.Fields("cpf03").Value)
'         ReDim bytes(lngSize)
'         If lngSize > 0 Then
'            bytes() = .Fields("cpf04").GetChunk(lngSize)
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
'   Exit Function
'end 2015/5/22

ErrHnd:
   If Err.Number = 70 Then
      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
   If iFileNo > 0 Then Close #iFileNo
End Function

'開啟附件
Private Sub cmdOpenAtt_Click()
Dim hLocalFile As Long
Dim stFileName As String
Dim bolIsSelect As Boolean
Dim Index As Integer 'Add By Sindy 2020/3/16
Dim strZipSrc As String, strZipFile As String, strSrcFile As String 'Added by Morgan 2020/7/28
   
   Index = SSTab1.Tab 'Add By Sindy 2020/3/16
   
   KillAttach
   
   bolIsSelect = False
   Screen.MousePointer = vbHourglass
   For ii = 1 To GRD1(Index).Rows - 1
      If GRD1(Index).TextMatrix(ii, 0) = "V" Then
         If Trim(GRD1(Index).TextMatrix(ii, colFN)) <> "" Then
            '讀取檔案名稱
            'Modified by Lydia 2020/02/06
            'stFileName = Trim(GRD1(Index).TextMatrix(ii, colFN))
            'If InStrRev(stFileName, " (") > 0 Then
            '   stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
            'End If
            stFileName = Trim(GRD1(Index).TextMatrix(ii, colCPF02))
            'end 2020/02/06
            bolIsSelect = True
            '清除反白
            GRD1(Index).col = 0
            GRD1(Index).row = ii
            GRD1(Index).TextMatrix(ii, 0) = ""
            For jj = 0 To GRD1(Index).Cols - 1
               GRD1(Index).col = jj
               GRD1(Index).CellBackColor = QBColor(15)
            Next jj
            'Modified by Lydia 2020/02/06
            'If InStr(stFileName, "\") = 0 Then
            If stFileName <> "" Then
               m_CP09 = Trim(GRD1(Index).TextMatrix(ii, colCp09))
               m_CP10 = Trim(GRD1(Index).TextMatrix(ii, colCP10))  'Added by Lydia 2020/03/03
               
               'Added by Morgan 2020/7/28 HP加密壓縮檔要先解密
               strZipSrc = ""
               If Right(UCase(stFileName), 14) = ".ENCRYPTED.ZIP" Then
                  strZipSrc = stFileName
                  stFileName = Left(stFileName, Len(stFileName) - 14)
                  strSrcFile = App.path & "\" & stFileName
                  strZipFile = App.path & "\$ENCRYPTED.ZIP"
                  
                  If GetAttachFile(m_CP09, strZipSrc, strZipFile) = False Then
                     GoTo EXITSUB
                  End If
                  
                  If PUB_UnZipFile2(strZipFile, App.path, cHPZipPwd) = False Then
                     GoTo EXITSUB
                  End If
                  stFileName = m_AttachPath & "\" & Trim(GRD1(Index).TextMatrix(ii, colReFN))
                  If Right(UCase(stFileName), 14) = ".ENCRYPTED.ZIP" Then
                     stFileName = Left(stFileName, Len(stFileName) - 14)
                  End If
                  
                  FileCopy strSrcFile, stFileName
                  Kill strZipFile
                  Kill strSrcFile
               Else
               'end 2020/7/28
                     
                  If GetAttachFile(m_CP09, stFileName, m_AttachPath & "\" & Trim(GRD1(Index).TextMatrix(ii, colReFN))) = False Then
                     MsgBox "無法開啟檔案[ " & stFileName & " ]！"
                     GoTo EXITSUB
                  End If
                  
               End If 'Added by Morgan 2020/7/28
            End If
            'Added by Lydia 2020/03/03 FCP之〔專利案件〕991的中說，舊案用PE2存檔，原檔案無副檔名可以用Notepad打開
            If m_CP10 = cnt專利案件 And InStr(UCase(stFileName), ".DOC") = 0 And InStr(UCase(stFileName), ".RTF") = 0 And InStr(UCase(stFileName), ".PDF") = 0 _
               And InStr(UCase(stFileName), ".JPG") = 0 And InStr(UCase(stFileName), ".PNG") = 0 And InStr(UCase(stFileName), ".BMP") = 0 _
               And InStr(UCase(stFileName), ".TIF") = 0 And InStr(UCase(stFileName), ".WMF") = 0 And InStr(UCase(stFileName), ".EMF") = 0 _
               And InStr(UCase(stFileName), ".ZIP") = 0 And InStr(UCase(stFileName), ".7Z") = 0 And InStr(UCase(stFileName), ".TXT") = 0 _
               And InStr(UCase(stFileName), ".PPT") = 0 And InStr(UCase(stFileName), ".XLS") = 0 And InStr(UCase(stFileName), ".CSV") = 0 _
               And InStr(UCase(stFileName), ".MSG") = 0 And InStr(UCase(stFileName), ".HTM") = 0 And InStr(UCase(stFileName), ".XML") = 0 _
               And InStr(UCase(stFileName), ".MENU") = 0 And InStr(UCase(stFileName), ".DEL") = 0 Then
               SHELL "Notepad.exe " & stFileName, vbNormalFocus
            Else
            'end 2020/03/03
               'end 2020/03/03
               SetAttr stFileName, vbReadOnly 'Add By Sindy 2020/3/17 檔案設定成唯讀屬性,防止直接修改儲存,以為就是上傳了
               '開啟檔案
               ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            End If 'Added by Lydia 2020/03/03
         End If
      End If
   Next ii
   If bolIsSelect = False Then
      MsgBox "無檔案可開啟！"
   End If
   
EXITSUB:
   Screen.MousePointer = vbDefault
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
Dim Index As Integer 'Add By Sindy 2020/3/16
   
   Index = SSTab1.Tab 'Add By Sindy 2020/3/16
      
   If cmdSelect.Caption = "全選" Then
      GRD1(Index).Visible = False
      For k = 1 To GRD1(Index).Rows - 1
         GRD1(Index).col = 0
         GRD1(Index).row = k
         If Trim(GRD1(Index).Text) = "" Then
            GRD1(Index).Text = "V"
            For i = 0 To GRD1(Index).Cols - 1
               GRD1(Index).col = i
               GRD1(Index).CellBackColor = &HFFC0C0
            Next i
         End If
      Next k
      GRD1(Index).Visible = True
   ElseIf cmdSelect.Caption = "取消全選" Then
      GRD1(Index).Visible = False
      For k = 1 To GRD1(Index).Rows - 1
         GRD1(Index).col = 0
         GRD1(Index).row = k
         If Trim(GRD1(Index).Text) = "V" Then
            GRD1(Index).Text = ""
            For i = 0 To GRD1(Index).Cols - 1
               GRD1(Index).col = i
               GRD1(Index).CellBackColor = QBColor(15)
            Next i
         End If
      Next k
      GRD1(Index).Visible = True
   End If
   Call ChangSelect
End Sub

'下載
Private Sub cmdSaveAtt_Click()
Dim stFileName As String, stFolderPath As String, stFullName As String
Dim bMultiFile As Boolean
Dim ii As Integer
Dim Index As Integer 'Add By Sindy 2020/3/16
Dim strZipSrc As String, strZipFile As String, strSrcFile As String 'Added by Morgan 2020/7/28
   
   Index = SSTab1.Tab 'Add By Sindy 2020/3/16
      
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
   For ii = 1 To GRD1(Index).Rows - 1
      If GRD1(Index).TextMatrix(ii, 0) = "V" And Trim(GRD1(Index).TextMatrix(ii, colFN)) <> "" Then
         If stFileName <> "" Then
            bMultiFile = True
            Exit For
         Else
            stFileName = Trim(GRD1(Index).TextMatrix(ii, colFN))
            m_CP09 = Trim(GRD1(Index).TextMatrix(ii, colCp09))
            m_CP10 = Trim(GRD1(Index).TextMatrix(ii, colCP10)) 'Add By Sindy 2020/3/17
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
            For ii = 1 To GRD1(Index).Rows - 1
               If GRD1(Index).TextMatrix(ii, 0) = "V" And Trim(GRD1(Index).TextMatrix(ii, colFN)) <> "" Then
                  m_CP09 = Trim(GRD1(Index).TextMatrix(ii, colCp09))
                  m_CP10 = Trim(GRD1(Index).TextMatrix(ii, colCP10)) 'Add By Sindy 2020/3/17
                  
                  'Modified by Lydia 2020/02/06
                  'stFileName = Trim(GRD1(Index).TextMatrix(ii, colFN))
                  'If InStrRev(stFileName, " (") > 0 Then
                  '   stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
                  'End If
                  stFileName = GetFileName(Trim(GRD1(Index).TextMatrix(ii, colFN)))
                  'end 2020/02/06
                  
                  'Modify By Sindy 2020/3/17
                  If Index = 1 Then
                     '暫存區,為防止易造成和原始檔區同檔名,自動加上.TMP.
                     '下載時要把.TMP.拿掉
                     stFullName = Trim(GRD1(Index).TextMatrix(ii, colReFN))
                     If InStr(UCase(stFullName), "." & m_CP10 & ".TMP.") > 0 Then
                        stFullName = stFolderPath & Replace(stFullName, "." & m_CP10 & ".TMP.", "." & m_CP10 & ".")
                     Else
                        stFullName = stFolderPath & stFullName
                     End If
                  Else
                  '2020/3/17 END
                     stFullName = stFolderPath & Trim(GRD1(Index).TextMatrix(ii, colReFN))
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
         'Modified by Lydia 2020/02/06
         'stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
         stFileName = GetFileName(stFileName)
         
         'Added by Morgan 2020/7/28 HP加密壓縮檔要先解密
         strZipSrc = ""
         If Right(UCase(stFileName), 14) = ".ENCRYPTED.ZIP" Then
            strZipSrc = stFileName
            stFileName = Left(stFileName, Len(stFileName) - 14)
            strSrcFile = App.path & "\" & stFileName
            strZipFile = App.path & "\$ENCRYPTED.ZIP"
         End If
         'end 2020/7/28
         
         'Modify By Sindy 2020/3/17
         If Index = 1 Then
            '暫存區,為防止易造成和原始檔區同檔名,自動加上.TMP.
            '下載時要把.TMP.拿掉
            If InStr(UCase(stFileName), "." & m_CP10 & ".TMP.") > 0 Then
               stFullName = stFolderPath & Replace(stFileName, "." & m_CP10 & ".TMP.", "." & m_CP10 & ".")
            Else
               stFullName = stFolderPath & stFileName
            End If
         Else
         '2020/3/17 END
            'Modify By Sindy 2020/1/14
            'stFullName = GetSaveName(stFileName)
            stFullName = stFolderPath & stFileName
            '2020/1/14 END
         End If
         
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
Dim strCPF02Name As String 'Added by Lydia 2020/02/06
Dim Index As Integer 'Add By Sindy 2020/3/16
         
On Error GoTo ErrHnd
   
   Index = SSTab1.Tab 'Add By Sindy 2020/3/16
      
Star_Run:
   intChkCnt = 0
   For ii = 1 To GRD1(Index).Rows - 1
      If GRD1(Index).TextMatrix(ii, 0) = "V" Then
         m_CP09 = Trim(GRD1(Index).TextMatrix(ii, colCp09))
         m_CP10 = Trim(GRD1(Index).TextMatrix(ii, colCP10))
         strCP82 = Trim(GRD1(Index).TextMatrix(ii, colCP82))
         intChkCnt = intChkCnt + 1
         
        'Added by Lydia 2020/03/26 檢查是否有新增的權限
        If (InStr("P", m_identity) > 0 And m_identity <> "") Then
            '繪圖人員：限系統別、案件性質
            If m_CP01 = "P" Or m_CP01 = "CFP" Or m_CP01 = "PS" Or m_CP01 = "CPS" Then
                If m_CP01 = "P" And (GRD1(Index).TextMatrix(ii, colCP10) = cntEnglish_Vers Or GRD1(Index).TextMatrix(ii, colCP10) = cnt專利案件) Then
                 MsgBox "屬於〔" & IIf(GRD1(Index).TextMatrix(ii, colCP10) = cnt專利案件, "專利案件", "English_Vers") & "〕的電子檔,不可以新增！"
                 Exit Sub
                End If
            Else
                MsgBox "非P、CFP、PS、CPS案件的電子檔,不可以新增！"
                Exit Sub
            End If
        End If
        'English_Vers和專利案件的讀寫權限 , 比照Typing2
        '專利案件 991: 中打室和FCP程序F22可讀寫 , DomainUser只可讀(含工程師F21)
        'English_Vers 992: 中打室和FCP承辦F23可讀寫 , DomainUser只可讀(含工程師F21)
        If (m_CP01 = "P" Or m_CP01 = "FCP") And (GRD1(Index).TextMatrix(ii, colCP10) = cnt專利案件 Or GRD1(Index).TextMatrix(ii, colCP10) = cntEnglish_Vers) Then
               If m_identity = "T" Or Pub_StrUserSt03 = "M51" Then '中打室和電腦中心
               Else
                   If GRD1(Index).TextMatrix(ii, colCP10) = cnt專利案件 And Pub_StrUserSt03 <> "F22" Then
                       MsgBox "屬於〔" & IIf(GRD1(Index).TextMatrix(ii, colCP10) = cnt專利案件, "專利案件", "English_Vers") & "〕的電子檔,不可以新增！"
                       Exit Sub
                   End If
                   If GRD1(Index).TextMatrix(ii, colCP10) = cntEnglish_Vers And Pub_StrUserSt03 <> "F23" Then
                       MsgBox "屬於〔" & IIf(GRD1(Index).TextMatrix(ii, colCP10) = cnt專利案件, "專利案件", "English_Vers") & "〕的電子檔,不可以新增！"
                       Exit Sub
                   End If
               End If
        End If
        'end 2020/03/26
      End If
   Next ii
   
   If intChkCnt = 0 Then
      'Modify By Sindy 2020/3/19
      '暫存區沒有點選欲新增那一筆資料列時,自動選取第一筆
      If Index = 1 Then
         GRD1(1).col = 0
         GRD1(1).row = 1
         '資料列反白
         GRD1(1).TextMatrix(1, 0) = "V"
         For jj = 0 To GRD1(1).Cols - 1
            GRD1(1).col = jj
            GRD1(1).CellBackColor = &HFFC0C0
         Next jj
         GoTo Star_Run
      Else
      '2020/3/19 END
         MsgBox "請勾選一筆欲新增電子檔的總收文號！"
         Exit Sub
      End If
   ElseIf intChkCnt > 1 Then
      MsgBox "只可勾選一筆總收文號做新增！"
      Exit Sub
   'Modify By Sindy 2014/6/24 Mark
'   'Add By Sindy 2014/3/11
'   ElseIf Trim(strCP82) = "" Then
'      MsgBox "此文尚未發文，不可異動附件！"
'      Exit Sub
   End If
   
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
               'Add By Sindy 2013/10/9
               If InStr(CStr(sFile(ii)), "#") > 0 Then
                  MsgBox CStr(sFile(ii)) & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
                  GoTo EXITSUB
               End If
               '2013/10/9 END
               
               '檢查檔名規則
               If PUB_ChkEmpFlowFNMRule(lblCaseNo, CStr(sFile(ii)), "Y", m_CP10, strCaseNoName, , False) = False Then
                  GoTo EXITSUB
               End If
               'Modify By Sindy 2020/3/17
               If Index = 1 Then
                  If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, CStr(sFile(ii)), stReName, True, 1) = False Then GoTo EXITSUB
                  '暫存區,為防止易造成和原始檔區同檔名,自動加上.TMP.
                  If InStr(UCase(stReName), "." & m_CP10 & ".TMP.") = 0 Then
                     stReName = Replace(stReName, "." & m_CP10 & ".", "." & m_CP10 & ".TMP.")
                  End If
               Else
               '2020/3/17 END
                  If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, CStr(sFile(ii)), stReName, True, 0) = False Then GoTo EXITSUB
                  strCPF02Name = PUB_GetReNameCPF02(CStr(sFile(ii)), m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, "M") 'Added by Lydia 2020/02/06 CPF02檔名處理
               End If
               
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               'Modify By Sindy 2013/9/6 檔案大小為 0 KB 有誤
               If f.Size = 0 Then
                  ShowMsg sFile(ii) & MsgText(9221)
                  GoTo EXITSUB
               'Add By Sindy 2014/3/11
               ElseIf f.Size > 5242880 Then
                  'If Pub_StrUserSt15 = "P13" Then
                     If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                        GoTo EXITSUB
                     End If
                  'End If
               '2014/3/11 END
               End If
               '2013/9/6 END
               If AddListX(Index, stFileName & " (" & Round(f.Size / 1024, 2) & " KB)") = True Then
                  '存檔
                  'Modified by Lydia 2020/02/06 指定檔名
                  'If SaveAttFile_Org(m_CP09, stFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), "A") = True Then
                  'Modify By Sindy 2020/3/17 判斷要存放那一區 IIf(Index = 0, "A", "Z")
                  If SaveAttFile_Org(m_CP09, stFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), IIf(Index = 0, "A", "Z"), strCPF02Name) = True Then
                     bolAdd = True
                  Else
                     GoTo EXITSUB
                  End If
'                  Pub_SaveLog strUserNum, "新增原始檔區附件：" & sFile(ii), m_CP01, m_CP02, m_CP03, m_CP04, m_CP09
               End If
            Next ii
         Else
            'stFileName = GetFileName(.FileName)
            'Modify By Sindy 2013/10/9
            'strFile = GetFileName(.FileName)
            strFile = Mid(.FileName, InStrRev(.FileName, "\") + 1)
            If InStr(strFile, "#") > 0 Then
               MsgBox strFile & vbCrLf & vbCrLf & "【#】符號為系統保留字，不可使用於檔案命名"
               GoTo EXITSUB
            End If
            '2013/10/9 END
            
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
            If PUB_ChkEmpFlowFNMRule(lblCaseNo, strFile, "Y", m_CP10, strCaseNoName, , False) = False Then
               GoTo EXITSUB
            End If
            'Modify By Sindy 2020/3/17
            If Index = 1 Then
               If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, strFile, stReName, True, 1) = False Then GoTo EXITSUB
               '暫存區,為防止易造成和原始檔區同檔名,自動加上.TMP.
               If InStr(UCase(stReName), "." & m_CP10 & ".TMP.") = 0 Then
                  stReName = Replace(stReName, "." & m_CP10 & ".", "." & m_CP10 & ".TMP.")
               End If
            Else
            '2020/3/17 END
               If PUB_GetEmpFlowReNameFile(m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, strFile, stReName, True, 0) = False Then GoTo EXITSUB
               strCPF02Name = PUB_GetReNameCPF02(strFile, m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, "M") 'Added by Lydia 2020/02/06 CPF02檔名處理
            End If
            
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            'Modify By Sindy 2013/9/6 檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               ShowMsg strFile & MsgText(9221)
               GoTo EXITSUB
            'Add By Sindy 2014/3/11
            ElseIf f.Size > 5242880 Then
               'If Pub_StrUserSt15 = "P13" Then
                  If MsgBox("檔案過大（容量超過5MB），確認是否要傳送？", vbYesNo, "警告") = vbNo Then
                     GoTo EXITSUB
                  End If
               'End If
            '2014/3/11 END
            End If
            '2013/9/6 END
            If AddListX(Index, stFileName & " (" & Round(f.Size / 1024, 2) & " KB)") = True Then
               '存檔
               'Modified by Lydia 2020/02/06 指定檔名
               'If SaveAttFile_Org(m_CP09, stFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), "A") = True Then
               'Modify By Sindy 2020/3/17 判斷要存放那一區 IIf(Index = 0, "A", "Z")
               If SaveAttFile_Org(m_CP09, stFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), IIf(Index = 0, "A", "Z"), strCPF02Name) = True Then
                  bolAdd = True
               Else
                  GoTo EXITSUB
               End If
'               Pub_SaveLog strUserNum, "新增原始檔區附件：" & strFile, m_CP01, m_CP02, m_CP03, m_CP04, m_CP09
            End If
         End If
EXITSUB:
         If bolAdd = True Then
            Call ReadAttachFile(Index)
         End If
      End If
      ChDir App.path 'Add By Sindy 2020/1/13 釋放資料夾權限
   End With
   Exit Sub
   
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Function AddListX(Index As Integer, stNewItem As String) As Boolean
   Dim idx As Integer, stFileName As String
   Dim stCP09 As String 'Add By Sindy 2015/3/6
   
   If stNewItem <> "" Then
      For idx = 1 To GRD1(Index).Rows - 1
         stFileName = Trim(GetFileName(GRD1(Index).TextMatrix(idx, colFN)))
         stCP09 = Trim(GRD1(Index).TextMatrix(idx, colCp09))
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
Dim Index As Integer 'Add By Sindy 2020/3/16
   
   Index = SSTab1.Tab 'Add By Sindy 2020/3/16
   
   intChkCnt = 0
   For ii = 1 To GRD1(Index).Rows - 1
      If GRD1(Index).TextMatrix(ii, 0) = "V" And Trim(GRD1(Index).TextMatrix(ii, colFN)) <> "" Then
         intChkCnt = intChkCnt + 1
      End If
      'Modify By Sindy 2020/3/17
      'Modified by Lydia 2020/03/26 + GRD1(Index).TextMatrix(ii, 0) = "V"
      If Trim(GRD1(Index).TextMatrix(ii, colFN)) <> "" And GRD1(Index).TextMatrix(ii, 0) = "V" Then
         'Added by Lydia 2020/03/26 English_Vers和專利案件的讀寫權限，比照Typing2
         '專利案件 991: 中打室和FCP程序F22可讀寫 , DomainUser只可讀(含工程師F21)
         'English_Vers 992: 中打室和FCP承辦F23可讀寫 , DomainUser只可讀(含工程師F21)
         If (m_CP01 = "P" Or m_CP01 = "FCP") And (GRD1(Index).TextMatrix(ii, colCP10) = cnt專利案件 Or GRD1(Index).TextMatrix(ii, colCP10) = cntEnglish_Vers) Then
               If m_identity = "T" Or Pub_StrUserSt03 = "M51" Then '中打室和電腦中心
               Else
                   If GRD1(Index).TextMatrix(ii, colCP10) = cnt專利案件 And Pub_StrUserSt03 <> "F22" Then
                       MsgBox "電子檔（" & Trim(GRD1(Index).TextMatrix(ii, colCPF02)) & "），屬於〔專利案件〕的電子檔,不可以刪除！"
                       Exit Sub
                   End If
                   If GRD1(Index).TextMatrix(ii, colCP10) = cntEnglish_Vers And Pub_StrUserSt03 <> "F23" Then
                       MsgBox "電子檔（" & Trim(GRD1(Index).TextMatrix(ii, colCPF02)) & "），屬於〔English_Vers〕的電子檔,不可以刪除！"
                       Exit Sub
                   End If
               End If
         Else
         'end 2020/03/26
                '檢查是否有刪除的權限
                If (m_identity <> "F" And m_identity <> "C") Or _
                   Pub_StrUserSt03 = "F23" Then '程序人員(F):開放全部權限
                   '外專承辦開放可以刪除國外部信件區匯入的郵件
                   If Not (Pub_StrUserSt03 = "F23" And Trim(GRD1(Index).TextMatrix(ii, colCPF11)) = "F") Then
                      If Trim(GRD1(Index).TextMatrix(ii, colCPF11)) <> "A" And _
                         Trim(GRD1(Index).TextMatrix(ii, colCPF11)) <> "Z" Then
                         '資料來源是原始檔區,暫存區才能刪除
                         MsgBox "電子檔（" & Trim(GRD1(Index).TextMatrix(ii, colCPF02)) & "）不是在原始檔區,暫存區新增的,不可以刪除！"
                         Exit Sub
                      'Added by Lydia 2024/11/27 調整權限為：上傳檔案本人外，應包含該案原智權人員（即該案承辦），以及該案承辦之案件職代。---Anny
                      ElseIf m_strF23User <> "" And InStr(m_strF23User, strUserNum) > 0 Then
                          '-----開放權限
                      'end 2024/11/27
                      'Modify By Sindy 2023/6/8 + Create ID
                      ElseIf Trim(GRD1(Index).TextMatrix(ii, colCPF05)) <> strUserNum And Trim(GRD1(Index).TextMatrix(ii, colCPF14)) <> strUserNum Then
                         '自己放的檔案才能刪除
                         MsgBox "電子檔（" & Trim(GRD1(Index).TextMatrix(ii, colCPF02)) & "）並非您新增的,不可以刪除！"
                         Exit Sub
                      ElseIf Index = 0 And DBDATE(DateAdd("d", 7, Format(Trim(GRD1(Index).TextMatrix(ii, colCPF06)), "####/##/##"))) < strSrvDate(1) Then
                         '管制原始檔區 : 新增檔案日期在7天內的檔案才能刪除
                         MsgBox "電子檔（" & Trim(GRD1(Index).TextMatrix(ii, colCPF02)) & "）已超過7日可以刪除的期限,不可以刪除！"
                         Exit Sub
                      End If
                   End If
                End If
         End If 'Added by Lydia 2020/03/26
         
      '2020/3/17 END
      End If
   Next ii
   If intChkCnt <= 0 Then
      MsgBox "請至少勾選一筆欲刪除電子檔的資料！"
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   bolDel = False
   For ii = 1 To GRD1(Index).Rows - 1
      If GRD1(Index).TextMatrix(ii, 0) = "V" And Trim(GRD1(Index).TextMatrix(ii, colFN)) <> "" Then
         If MsgBox("確定要永久刪除" & GetFileName(Trim(GRD1(Index).TextMatrix(ii, colFN))) & "電子檔？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         '直接從資料庫刪除檔案
         m_CP09 = Trim(GRD1(Index).TextMatrix(ii, colCp09))
         'Modified by Lydia 2023/08/24 判斷外專新案若FTP路徑尚未修正，直接刪除檔案，避免殘留FTP/TEMP目錄
         'If DeleteFile(m_CP09, GetFileName(Trim(GRD1(Index).TextMatrix(ii, colFN)))) = True Then
         strExc(1) = ""
         If Left(m_CP09, 1) = "D" And UCase(Left(Trim(GRD1(Index).TextMatrix(ii, colCPF13)), 5)) = "TEMP/" Then
            strExc(1) = "Y"
         End If
         If DeleteFile(m_CP09, GetFileName(Trim(GRD1(Index).TextMatrix(ii, colFN))), IIf(strExc(1) = "Y", True, False)) = True Then
         'end 2023/08/24
            bolDel = True
         End If
         
      End If
   Next ii
   
   Screen.MousePointer = vbDefault
   
   If bolDel = True Then Call ReadAttachFile(Index)
End Sub

'Added by Lydia 2020/02/06
Private Sub GRD1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
GRD1(Index).ToolTipText = ""
If GRD1(Index).MouseRow <> 0 And GRD1(Index).MouseCol > 0 Then
   If GRD1(Index).MouseCol = colFN Or GRD1(Index).MouseCol = colCPM Then   '檔案名稱(size)、案件性質
      If GRD1(Index).TextMatrix(GRD1(Index).MouseRow, GRD1(Index).MouseCol) <> "" Then
         GRD1(Index).ToolTipText = GRD1(Index).TextMatrix(GRD1(Index).MouseRow, GRD1(Index).MouseCol)
      End If
   End If
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   Call CmdLimits
   
   'Add By Sindy 2020/9/28
   If m_bolAddTmpFile = False Then
   '2020/9/28 END
      'Add By Sindy 2020/3/17
      If SSTab1.TabVisible(0) = True Then '有原始檔區時
         If SSTab1.Tab = 1 Then
            '不可新增刪除暫存區,隱藏
            cmdAddAtt.Visible = False
            cmdRemAtt.Visible = False
            cmdCopy.Visible = False 'Add By Sindy 2021/11/1
         Else
            cmdAddAtt.Visible = True
            cmdRemAtt.Visible = True
            cmdCopy.Visible = True 'Add By Sindy 2021/11/1
         End If
      End If
   End If
End Sub

'Add By Sindy 2020/3/18
Private Sub Text2_Click()
   SSTab1.Tab = 1
End Sub

'Added by Lydia 2020/07/02
Private Sub GRD1_DblClick(Index As Integer)
   '雙擊直接開啟檔案by經理
   If GRD1(Index).row > 0 Then
'      Screen.MousePointer = vbHourglass
'      If PubShowNextData(GRD1(Index).row) = False Then
'         m_bolDblClick = True
'
'         cmdOpenAtt_Click 1
'      End If
'      Screen.MousePointer = vbDefault
      Call cmdOpenAtt_Click
   End If
End Sub
