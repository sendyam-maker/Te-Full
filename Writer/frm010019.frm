VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm010019 
   BorderStyle     =   1  '單線固定
   Caption         =   "總務處工作報告列印"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   7995
   Begin VB.OptionButton opt1 
      Caption         =   "日統計"
      Height          =   225
      Index           =   1
      Left            =   2580
      TabIndex        =   3
      Top             =   720
      Width           =   915
   End
   Begin VB.OptionButton opt1 
      Caption         =   "月統計"
      Height          =   225
      Index           =   0
      Left            =   1590
      TabIndex        =   2
      Top             =   720
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.TextBox txt2 
      Height          =   270
      Left            =   2610
      MaxLength       =   7
      TabIndex        =   1
      Top             =   960
      Width           =   945
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "預覽(&V)"
      Default         =   -1  'True
      Height          =   435
      Index           =   2
      Left            =   5310
      TabIndex        =   4
      Top             =   30
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Left            =   1560
      MaxLength       =   7
      TabIndex        =   0
      Top             =   960
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4035
      Left            =   30
      TabIndex        =   12
      Top             =   1740
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   7117
      _Version        =   393216
      Rows            =   26
      Cols            =   9
      FixedRows       =   0
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      MergeCells      =   1
      AllowUserResizing=   3
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.Timer Timer1 
      Left            =   4680
      Top             =   0
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   435
      Index           =   1
      Left            =   7095
      TabIndex        =   6
      Top             =   30
      Width           =   885
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Height          =   435
      Index           =   0
      Left            =   6210
      TabIndex        =   5
      Top             =   30
      Width           =   885
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   5640
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   960
      Width           =   2340
   End
   Begin VB.Label Label3 
      Caption         =   $"frm010019.frx":0000
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   120
      TabIndex        =   14
      Top             =   80
      Width           =   4095
   End
   Begin VB.Line Line1 
      X1              =   2250
      X2              =   2880
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作報告西元月："
      Height          =   180
      Left            =   90
      TabIndex        =   13
      Top             =   1020
      Width           =   1440
   End
   Begin VB.Label lblCnt 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   1380
      Width           =   7815
   End
   Begin VB.Label lblPB 
      BackColor       =   &H00C00000&
      Height          =   255
      Left            =   90
      TabIndex        =   10
      Top             =   1350
      Width           =   15
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      Height          =   180
      Index           =   1
      Left            =   4920
      TabIndex        =   8
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lblPBBox 
      BackColor       =   &H00C000C0&
      BorderStyle     =   1  '單線固定
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   1320
      Width           =   7875
   End
End
Attribute VB_Name = "frm010019"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/08/23 Form2.0已修改(無需修改)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
Option Explicit

Dim SeekPrint As Integer, SeekPrintL As Integer
Dim iLine As Integer
Dim MaxLine As Integer
Dim PLeft(9) As Integer
Dim strSql As String
Dim AllStates As Integer
Dim strStates() As String
Dim NowState As Integer
Dim i As Integer
Dim BoxTop As Integer
Dim BoxBotton As Integer
Dim BoxLeft As Integer
Dim BoxRight  As Integer
Dim TheLBoxH As Integer
Dim oLines As Integer
Dim isPrt As Boolean
Dim tmpD1 As String
Dim tmpD2 As String
'add by nickc 2007/10/17
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_bPrint As Boolean

Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean
Select Case Index
Case 0
    If Trim(txt1) = "" Then
        MsgBox "請輸入日期!!", vbInformation, "無法統計!!"
        txt1.SetFocus
        Exit Sub
    Else
        Cancel = False
        txt1_Validate Cancel
        If Cancel = True Then
            Exit Sub
        End If
    End If
    If Trim(txt2) = "" Then
        MsgBox "請輸入日期!!", vbInformation, "無法統計!!"
        txt2.SetFocus
        Exit Sub
    Else
        Cancel = False
        txt2_Validate Cancel
        If Cancel = True Then
            Exit Sub
        End If
    End If
    If Val(txt1) > Val(txt2) Then
        MsgBox "日期區間錯誤!!", vbInformation, "無法統計!!"
        txt1.SetFocus
        Exit Sub
    End If
    'Added by Lydia 2021/03/10 檢查當月發文數輸入
    If opt1(0).Value = True Then
        If CheckGeneralDispatch = False Then
             Exit Sub
        End If
    End If
    'end 2021/03/10
    
    isPrt = True
    Screen.MousePointer = vbHourglass
    Set Printer = Printers(Combo1.ListIndex)
    NowState = 0
    DoEvents
    'edit by nickc 2007/08/15 改成區間
    'GetAllData Trim(txt1) & "01"
    If Val(tmpD1) <> Val(txt1) Or Val(tmpD2) <> Val(txt2) Then
        'Added by Lydia 2021/03/10 因為PrintData要用變數日期存記錄
        tmpD1 = txt1
        tmpD2 = txt2
        'end 2021/03/10
        GetAllData ChangeTStringToWString(Trim(txt1)), ChangeTStringToWString(Trim(txt2))
    End If
    NowState = 63
    PrintData
    Screen.MousePointer = vbDefault
Case 1
    Unload Me
Case 2
    If Trim(txt1) = "" Then
        MsgBox "請輸入日期!!", vbInformation, "無法統計!!"
        txt1.SetFocus
        Exit Sub
    Else
        Cancel = False
        txt1_Validate Cancel
        If Cancel = True Then
            Exit Sub
        End If
    End If
    If Trim(txt2) = "" Then
        MsgBox "請輸入日期!!", vbInformation, "無法統計!!"
        txt2.SetFocus
        Exit Sub
    Else
        Cancel = False
        txt2_Validate Cancel
        If Cancel = True Then
            Exit Sub
        End If
    End If
    If Val(txt1) > Val(txt2) Then
        MsgBox "日期區間錯誤!!", vbInformation, "無法統計!!"
        txt1.SetFocus
        Exit Sub
    End If
    'Added by Lydia 2021/03/10 檢查當月發文數輸入
    If opt1(0).Value = True Then
        If CheckGeneralDispatch = False Then
             Exit Sub
        End If
    End If
    'end 2021/03/10
    
    isPrt = False
    tmpD1 = txt1
    tmpD2 = txt2
    'Modified by Lydia 2018/10/15 加備註
    'If Height = 5850 Then
    If Height = 6400 Then
        'Height = 1710
        MoveFormToCenter Me
        GetAllData ChangeTStringToWString(Trim(txt1)), ChangeTStringToWString(Trim(txt2))
        NowState = 0
    Else
        SetGrd
        'Modified by Lydia 2018/10/15
        'Height = 5850
        Height = 6400
        MoveFormToCenter Me
        'edit by nickc 2007/08/15 改成區間
        'GetAllData Trim(txt1) & "01"
        GetAllData ChangeTStringToWString(Trim(txt1)), ChangeTStringToWString(Trim(txt2))
        NowState = 0
    End If
End Select
End Sub

Sub InitialALL()
AllStates = 66
ReDim strStates(AllStates) As String
strStates(0) = "按下列印後開始列印！！"
strStates(1) = "計算  [收文]總數"
strStates(2) = "計算  [業務收文]數"
strStates(3) = "計算  [發文總數、標準局發文]數"
strStates(4) = "計算  新增[國內、國外]客戶數"
strStates(5) = "計算  [客戶來訪]數 "
strStates(6) = "計算  新卷"
strStates(7) = "計算  新卷[商標]數"
strStates(8) = "計算  新卷[專利]數"
strStates(9) = "計算  新卷[著作權] 數"
strStates(10) = "計算  新卷[顧問]數"
strStates(11) = "計算  新卷[法務]數"
strStates(12) = "計算  新卷[CFT]數"
strStates(13) = "計算  新卷[CFP]數"
strStates(14) = "計算  新卷[CFC]數"
strStates(15) = "計算  新卷[FCT]數"
strStates(16) = "計算  新卷[FCP]數"
strStates(17) = "計算  新卷[FCL]數"
strStates(18) = "計算  新卷[CFL]數"
strStates(19) = "計算  新卷[B、D、S、M、F]數"
strStates(20) = "讀取  上月資料"
strStates(21) = "計算  上月[收文]總數"
strStates(22) = "計算  上月[業務收文]數"
strStates(23) = "計算  上月[發文總數、標準局發文]數"
strStates(24) = "計算  上月新增[國內、國外]客戶數"
strStates(25) = "計算  上月[客戶來訪]數 "
strStates(26) = "計算  上月新卷"
strStates(27) = "計算  上月新卷[商標]數"
strStates(28) = "計算  上月新卷[專利]數"
strStates(29) = "計算  上月新卷[著作權] 數"
strStates(30) = "計算  上月新卷[顧問]數"
strStates(31) = "計算  上月新卷[法務]數"
strStates(32) = "計算  上月新卷[CFT]數"
strStates(33) = "計算  上月新卷[CFP]數"
strStates(34) = "計算  上月新卷[CFC]數"
strStates(35) = "計算  上月新卷[FCT]數"
strStates(36) = "計算  上月新卷[FCP]數"
strStates(37) = "計算  上月新卷[FCL]數"
strStates(38) = "計算  上月新卷[CFL]數"
strStates(39) = "計算  上月新卷[B、D、S、M、F]數"
strStates(40) = "計算  上月成長數"
strStates(41) = "計算  上月成長率"
strStates(42) = "讀取  去年同期資料"
strStates(43) = "計算  去年同期[收文]總數"
strStates(44) = "計算  去年同期[業務收文]數"
strStates(45) = "計算  去年同期[發文總數、標準局發文]數"
strStates(46) = "計算  去年同期新增[國內、國外]客戶數"
strStates(47) = "計算  去年同期[客戶來訪]數 "
strStates(48) = "計算  去年同期新卷"
strStates(49) = "計算  去年同期新卷[商標]數"
strStates(50) = "計算  去年同期新卷[專利]數"
strStates(51) = "計算  去年同期新卷[著作權] 數"
strStates(52) = "計算  去年同期新卷[顧問]數"
strStates(53) = "計算  去年同期新卷[法務]數"
strStates(54) = "計算  去年同期新卷[CFT]數"
strStates(55) = "計算  去年同期新卷[CFP]數"
strStates(56) = "計算  去年同期新卷[CFC]數"
strStates(57) = "計算  去年同期新卷[FCT]數"
strStates(58) = "計算  去年同期新卷[FCP]數"
strStates(59) = "計算  去年同期新卷[FCL]數"
strStates(60) = "計算  去年同期新卷[CFL]數"
strStates(61) = "計算  去年同期新卷[B、D、S、M、F]數"
strStates(62) = "計算  去年同期成長數"
strStates(63) = "計算  去年同期成長率"
strStates(64) = "列印中......."
strStates(65) = "列印完成 ^ ^ "
lblCnt.Caption = ""
'Modified by Lydia 2018/10/15
'lblCnt.Top = 1020
lblCnt.Top = 1350
lblCnt.Left = 90
lblCnt.Height = 195
lblCnt.Width = 7815
lblPB.Width = 15
'Modified by Lydia 2018/10/15
'lblPB.Top = 990
lblPB.Top = 1350
lblPB.Left = 90
lblPBBox.Width = 7875
'Modified by Lydia 2018/10/15
'lblPBBox.Top = 960
lblPBBox.Top = 1320
lblPBBox.Left = 60
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
'add by nickc 2007/10/17 增加權限控管
m_bInsert = IsUserHasRightOfFunction("frm010019", strAdd, False)
m_bUpdate = IsUserHasRightOfFunction("frm010019", strEdit, False)
m_bDelete = IsUserHasRightOfFunction("frm010019", strDel, False)
m_bQuery = IsUserHasRightOfFunction("frm010019", strFind, False)
m_bPrint = IsUserHasRightOfFunction("frm010019", strPrint, False)
If m_bPrint Then
    cmdok(0).Enabled = True
Else
    cmdok(0).Enabled = False
End If
If m_bQuery Then
    cmdok(2).Enabled = True
Else
    cmdok(2).Enabled = False
End If

MoveFormToCenter Me
txt1 = ChangeWStringToTString(Mid(strSrvDate(1), 1, 6) & "01")
txt2 = ChangeWStringToTString(ChangeWDateStringToWString(DateAdd("d", -1, ChangeWStringToWDateString(Mid(ChangeWDateStringToWString(DateAdd("m", 1, ChangeWStringToWDateString(strSrvDate(1)))), 1, 6) & "01"))))
InitialALL
strSql = Printer.DeviceName
SeekPrintL = Printer.Orientation
For i = 0 To Printers.Count - 1
    Set Printer = Printers(i)
    Combo1.AddItem Printer.DeviceName, j
     j = j + 1
    If Printer.DeviceName = strSql Then
        SeekPrint = i
    End If
Next i
Set Printer = Printers(SeekPrint)
Combo1.Text = Combo1.List(SeekPrint)
NowState = 0
Timer1.Interval = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm010019 = Nothing
End Sub

Private Sub Timer1_Timer()
lblPB.Width = Fix(7815 / AllStates * NowState)
lblCnt.Caption = strStates(NowState)
DoEvents
End Sub

'edit by nickc 2007/08/15 改成區間
'Sub GetAllData(oDate As String)
Sub GetAllData(oDate1 As String, oDate2 As String)
Screen.MousePointer = vbHourglass
grd1.MousePointer = flexArrowHourGlass
DoEvents
grd1.Clear
grd1.TextMatrix(0, 2) = "本　月"
grd1.TextMatrix(0, 3) = "上　月"
grd1.TextMatrix(0, 4) = "成長數"
grd1.TextMatrix(0, 5) = "成長率"
grd1.TextMatrix(0, 6) = "去年同期"
grd1.TextMatrix(0, 7) = "成長數"
grd1.TextMatrix(0, 8) = "成長率"
grd1.TextMatrix(1, 0) = "收發文數"
grd1.TextMatrix(1, 1) = "收文總數"
grd1.TextMatrix(2, 0) = "收發文數"
grd1.TextMatrix(2, 1) = "業務收文"
grd1.TextMatrix(3, 0) = "收發文數"
grd1.TextMatrix(3, 1) = "發文總數"
grd1.TextMatrix(4, 0) = "收發文數"
grd1.TextMatrix(4, 1) = "標準局發文"
grd1.TextMatrix(5, 0) = "新增客戶"
grd1.TextMatrix(5, 1) = "國　　內"
grd1.TextMatrix(6, 0) = "新增客戶"
grd1.TextMatrix(6, 1) = "國　　外"
grd1.TextMatrix(7, 0) = "新增客戶"
grd1.TextMatrix(7, 1) = "合　　計"
grd1.TextMatrix(8, 0) = "客戶來訪"
grd1.TextMatrix(8, 1) = "智　　權"
grd1.TextMatrix(9, 0) = "客戶來訪"   'add by nickc 2007/11/01 以下都往下 + 1
grd1.TextMatrix(9, 1) = "臺一投資"     'modify by sonia 2021/2/25 改名稱
grd1.TextMatrix(10, 0) = "客戶來訪"
grd1.TextMatrix(10, 1) = "非  智  權"
grd1.TextMatrix(11, 0) = "客戶來訪"
grd1.TextMatrix(11, 1) = "合　　計"
grd1.TextMatrix(12, 0) = "檔案新卷"
grd1.TextMatrix(12, 1) = "商　　標"
grd1.TextMatrix(13, 0) = "檔案新卷"
grd1.TextMatrix(13, 1) = "專　　利"
grd1.TextMatrix(14, 0) = "檔案新卷"
grd1.TextMatrix(14, 1) = "著  作  權"
grd1.TextMatrix(15, 0) = "檔案新卷"
grd1.TextMatrix(15, 1) = "顧　　問"
grd1.TextMatrix(16, 0) = "檔案新卷"
grd1.TextMatrix(16, 1) = "法　　務"
grd1.TextMatrix(17, 0) = "檔案新卷"
grd1.TextMatrix(17, 1) = "CFT"
grd1.TextMatrix(18, 0) = "檔案新卷"
grd1.TextMatrix(18, 1) = "CFP"
grd1.TextMatrix(19, 0) = "檔案新卷"
grd1.TextMatrix(19, 1) = "CFC"
grd1.TextMatrix(20, 0) = "檔案新卷"
grd1.TextMatrix(20, 1) = "FCT"
grd1.TextMatrix(21, 0) = "檔案新卷"
grd1.TextMatrix(21, 1) = "FCP"
grd1.TextMatrix(22, 0) = "檔案新卷"
grd1.TextMatrix(22, 1) = "FCL"
grd1.TextMatrix(23, 0) = "檔案新卷"
grd1.TextMatrix(23, 1) = "CFL"
grd1.TextMatrix(24, 0) = "檔案新卷"
grd1.TextMatrix(24, 1) = "B、D、S、M、F"
grd1.TextMatrix(25, 0) = "檔案新卷"
grd1.TextMatrix(25, 1) = "合　　計"
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
'計算  [收文]總數
    '案件
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    strSql = "select nvl(count(cp09),0) from ("
    'edit by nickc 2007/08/15 改成日期區間
    'strSQL = strSQL & " select cp09 from caseprogress where substr(rtrim(ltrim(to_char(cp66))),1,6)='" & Mid(oDate, 1, 6) & "' and substr(cp09,1,1)='A'  "
    'strSQL = strSQL & " union select dd14 from datadeleterecord where substr(rtrim(ltrim(to_char(dd25))),1,6)='" & Mid(oDate, 1, 6) & "' and substr(dd14,1,1)='A' and dd18 is not null  "
    'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
    strSql = strSql & " select cp09 from caseprogress where cp66>=" & oDate1 & " and cp66<=" & oDate2 & " and substr(cp09,1,1)='A' and cp01||cp02<>'TT999999' "
    strSql = strSql & " union select dd14 from datadeleterecord where dd25>=" & oDate1 & " and dd25<=" & oDate2 & " and substr(dd14,1,1)='A' and dd18 is not null  "
    'add by sonia 2016/5/23 剔除已刪除但又救回來的進度T-203085之申請
    'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
    strSql = strSql & "          and dd14 not in (select cp09 from caseprogress where cp66>=" & oDate1 & " and cp66<=" & oDate2 & " and substr(cp09,1,1)='A' and cp01||cp02<>'TT999999' ) "
    'end 2016/5/23
    strSql = strSql & " ) AAAAA "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(1, 2) = CheckStr(rsTmp.Fields(0))
        grd1.TextMatrix(2, 2) = CheckStr(rsTmp.Fields(0))
    End If
    '政府機關        因為有些是沒有收進去  例如法院來的信件，但是會登記在簿子上   或是昨天漏輸，今天補上，或是漏登記卻有輸入電腦，所以誤差不管多或是少  都是對的
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/08/15 改成區間
    'strSQL = "select nvl(count(*),0) from mailrec where substr(rtrim(ltrim(to_char(mr02))),1,6)='" & Mid(oDate, 1, 6) & "' "
    strSql = "select nvl(count(*),0) from mailrec where mr02>=" & oDate1 & " and mr02<=" & oDate2 & "  "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(1, 2) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(1, 2)))
    End If
    '信件
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/08/15
    'strSQL = "select nvl(count(*),0) from letterinput where substr(rtrim(ltrim(to_char(li01))),1,6)='" & Mid(oDate, 1, 6) & "'    "
    strSql = "select nvl(count(*),0) from letterinput where li01>=" & oDate1 & "  and li01<=" & oDate2 & " "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(1, 2) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(1, 2)))
    End If
'計算  [業務收文]數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'strSQL = "select count(*) from caseprogress where substr(rtrim(ltrim(to_char(cp05))),1,6)='200705' and substr(cp09,1,1)='A'  "
'計算  [發文總數、標準局發文]數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/08/15
    'strSQL = "select nvl(sum(nvl(gd02,0)),0),nvl(sum(nvl(gd03,0)),0) from GeneralDispatch where substr(rtrim(ltrim(to_char(gd01))),1,6)='" & Mid(oDate, 1, 6) & "' "
    strSql = "select nvl(sum(nvl(gd02,0)),0),nvl(sum(nvl(gd03,0)),0) from GeneralDispatch where gd01>=" & oDate1 & " and gd01<=" & oDate2 & " "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(3, 2) = CheckStr(rsTmp.Fields(0))
        grd1.TextMatrix(4, 2) = CheckStr(rsTmp.Fields(1))
    End If
'計算  新增[國內、國外]客戶數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    '國內客戶
    'edit by nickc 2007/08/15
    'strSQL = "select nvl(count(*),0) from customer where substr(rtrim(ltrim(to_char(cu14))),1,6)='" & Mid(oDate, 1, 6) & "' and ((cu10 in ('020','013')) or cu10<='010')  and cu02='0' "
    strSql = "select nvl(count(*),0) from customer where cu14>=" & oDate1 & " and cu14<=" & oDate2 & " and ((cu10 in ('020','013')) or cu10<='010')  and cu02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(5, 2) = CheckStr(rsTmp.Fields(0))
    End If
    '國外客戶
    'edit by nickc 2007/08/15
    'strSQL = "select nvl(count(*),0) from customer where substr(rtrim(ltrim(to_char(cu14))),1,6)='" & Mid(oDate, 1, 6) & "' and not ((cu10 in ('020','013')) or cu10<='010') and cu02='0' "
    strSql = "select nvl(count(*),0) from customer where cu14>=" & oDate1 & " and cu14<=" & oDate2 & " and not ((cu10 in ('020','013')) or cu10<='010') and cu02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(6, 2) = CheckStr(rsTmp.Fields(0))
    End If
    '國內代理人
    'edit by nickc 2007/08/15
    'strSQL = "select nvl(count(*),0) from fagent where substr(rtrim(ltrim(to_char(fa11))),1,6)='" & Mid(oDate, 1, 6) & "' and ((fa10 in ('020','013')) or fa10<='010') and fa02='0' "
    strSql = "select nvl(count(*),0) from fagent where fa11>=" & oDate1 & " and fa11<=" & oDate2 & " and ((fa10 in ('020','013')) or fa10<='010') and fa02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(5, 2) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(5, 2)))
    End If
    '國內代理人
    'edit by nickc 2007/08/15
    'strSQL = "select nvl(count(*),0) from fagent where substr(rtrim(ltrim(to_char(fa11))),1,6)='" & Mid(oDate, 1, 6) & "' and not ((fa10 in ('020','013')) or fa10<='010') and fa02='0' "
    strSql = "select nvl(count(*),0) from fagent where fa11>=" & oDate1 & " and fa11<=" & oDate2 & " and not ((fa10 in ('020','013')) or fa10<='010') and fa02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(6, 2) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(6, 2)))
    End If
    grd1.TextMatrix(7, 2) = Trim(Val(grd1.TextMatrix(5, 2)) + Val(grd1.TextMatrix(6, 2)))
'計算  [客戶來訪]數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/08/15
    'strSQL = "select nvl(sum(nvl(oi06,0)),0),nvl(sum(nvl(oi07,0)),0) from otherinput where substr(rtrim(ltrim(to_char(oi01))),1,6)='" & Mid(oDate, 1, 6) & "'  "
    'edit by nickc 2007/11/01 加一個欄位
    'strSQL = "select nvl(sum(nvl(oi06,0)),0),nvl(sum(nvl(oi07,0)),0) from otherinput where oi01>=" & oDate1 & " and oi01<=" & oDate2 & "  "
    strSql = "select nvl(sum(nvl(oi06,0)),0),nvl(sum(nvl(oi14,0)),0),nvl(sum(nvl(oi07,0)),0) from otherinput where oi01>=" & oDate1 & " and oi01<=" & oDate2 & "  "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(8, 2) = CheckStr(rsTmp.Fields(0))
        grd1.TextMatrix(9, 2) = CheckStr(rsTmp.Fields(1))
        grd1.TextMatrix(10, 2) = CheckStr(rsTmp.Fields(2))
    End If
    grd1.TextMatrix(11, 2) = Trim(Val(grd1.TextMatrix(8, 2)) + Val(grd1.TextMatrix(9, 2)) + Val(grd1.TextMatrix(10, 2)))
'計算新卷
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/08/06 原先自動抓，現在抓輸入的 CreateFiles Table
    'strSQL = " select decode(cp01,'TD','TB','TS','TB','TM','TB','TF','TB','S','CFC','FG','FCP',cp01) as cp01,nvl(count(*),0) from caseprogress where substr(rtrim(ltrim(to_char(cp27))),1,6)='" & Mid(oDate, 1, 6) & "' and cp31='Y' and cp01 not in ('CPS','PS') group by decode(cp01,'TD','TB','TS','TB','TM','TB','TF','TB','S','CFC','FG','FCP',cp01) "
    'CFT 只有 101 、FCT、FCP 101-105 與收文量查詢相同
    'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
    strSql = "SELECT CP01,COUNT(CP09) FROM CASEPROGRESS " & _
                    " Where cp05 >= " & oDate1 & " And cp05 <= " & oDate2 & " And cp26 Is Null And cp21 Is Null and cp01||cp02<>'TT999999' " & _
                    "  AND CP09< 'B'  AND CP01||cp10 IN ('FCP101','CFP101','FCP102','FCP103','FCP104','FCP105','CFP102','CFP103','CFP104','CFP105','CFT101')  GROUP BY CP01 "
    strSql = strSql & " union select 'T',sum(nvl(cf02,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'P',sum(nvl(cf03,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'TC',sum(nvl(cf04,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'LA',sum(nvl(cf05,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'L',sum(nvl(cf06,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'CFC',sum(nvl(cf07,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'FCT',sum(nvl(cf08,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'FCL',sum(nvl(cf09,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'CFL',sum(nvl(cf10,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'TB',sum(nvl(cf11,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
   
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        '計算  新卷[商標]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(12, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "T"
                    grd1.TextMatrix(12, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[專利]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(13, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "P"
                    grd1.TextMatrix(13, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[著作權] 數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(14, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "TC"
                    grd1.TextMatrix(14, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[顧問]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(15, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "LA"
                    grd1.TextMatrix(15, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[法務]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(16, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "L"
                    grd1.TextMatrix(16, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFT]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(17, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFT"
                    grd1.TextMatrix(17, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFP]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(18, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFP"
                    grd1.TextMatrix(18, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFC]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(19, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFC"
                    grd1.TextMatrix(19, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[FCT]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(20, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "FCT"
                    grd1.TextMatrix(20, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[FCP]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(21, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "FCP"
                    grd1.TextMatrix(21, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[FCL]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(22, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "FCL"
                    grd1.TextMatrix(22, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFL]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(23, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFL"
                    grd1.TextMatrix(23, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[B、D、S、M、F]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(24, 2) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "TB"
                    grd1.TextMatrix(24, 2) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        grd1.TextMatrix(25, 2) = Trim(Val(grd1.TextMatrix(12, 2)) + Val(grd1.TextMatrix(13, 2)) + Val(grd1.TextMatrix(14, 2)) + Val(grd1.TextMatrix(15, 2)) + Val(grd1.TextMatrix(16, 2)) + Val(grd1.TextMatrix(17, 2)) + Val(grd1.TextMatrix(18, 2)) + Val(grd1.TextMatrix(19, 2)) + Val(grd1.TextMatrix(20, 2)) + Val(grd1.TextMatrix(21, 2)) + Val(grd1.TextMatrix(22, 2)) + Val(grd1.TextMatrix(23, 2)) + Val(grd1.TextMatrix(24, 2)))
    Else
        grd1.TextMatrix(11, 2) = "0"
        grd1.TextMatrix(12, 2) = "0"
        grd1.TextMatrix(13, 2) = "0"
        grd1.TextMatrix(14, 2) = "0"
        grd1.TextMatrix(15, 2) = "0"
        grd1.TextMatrix(16, 2) = "0"
        grd1.TextMatrix(17, 2) = "0"
        grd1.TextMatrix(18, 2) = "0"
        grd1.TextMatrix(19, 2) = "0"
        grd1.TextMatrix(20, 2) = "0"
        grd1.TextMatrix(21, 2) = "0"
        grd1.TextMatrix(22, 2) = "0"
        grd1.TextMatrix(23, 2) = "0"
        grd1.TextMatrix(24, 2) = "0"
        grd1.TextMatrix(25, 2) = "0"
    End If
'讀取  上月資料
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/08/15
    'strSQL = " select * from GeneralWork where substr(rtrim(ltrim(to_char(gw01))),1,6)='" & Mid(ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(oDate))), 1, 6) & "' "
    If opt1(0).Value = True Then
        strSql = " select * from GeneralWork where gw01>=" & ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Mid(oDate1, 1, 6) & "01"))) & " and gw01<=" & ChangeWDateStringToWString(DateAdd("d", -1, ChangeWStringToWDateString(Mid(oDate2, 1, 6) & "01"))) & " "
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rsTmp.RecordCount <> 0 Then
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(1, 3) = CheckStr(rsTmp.Fields("GW02"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(2, 3) = CheckStr(rsTmp.Fields("GW03"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(3, 3) = CheckStr(rsTmp.Fields("GW04"))
            grd1.TextMatrix(4, 3) = CheckStr(rsTmp.Fields("GW05"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(5, 3) = CheckStr(rsTmp.Fields("GW06"))
            grd1.TextMatrix(6, 3) = CheckStr(rsTmp.Fields("GW07"))
            grd1.TextMatrix(7, 3) = CheckStr(rsTmp.Fields("GW08"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(8, 3) = CheckStr(rsTmp.Fields("GW09"))
            grd1.TextMatrix(9, 3) = CheckStr(rsTmp.Fields("GW26"))
            grd1.TextMatrix(10, 3) = CheckStr(rsTmp.Fields("GW10"))
            grd1.TextMatrix(11, 3) = CheckStr(rsTmp.Fields("GW11"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(12, 3) = CheckStr(rsTmp.Fields("GW12"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(13, 3) = CheckStr(rsTmp.Fields("GW13"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(14, 3) = CheckStr(rsTmp.Fields("GW14"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(15, 3) = CheckStr(rsTmp.Fields("GW15"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(16, 3) = CheckStr(rsTmp.Fields("GW16"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(17, 3) = CheckStr(rsTmp.Fields("GW17"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(18, 3) = CheckStr(rsTmp.Fields("GW18"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(19, 3) = CheckStr(rsTmp.Fields("GW19"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(20, 3) = CheckStr(rsTmp.Fields("GW20"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(21, 3) = CheckStr(rsTmp.Fields("GW21"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(22, 3) = CheckStr(rsTmp.Fields("GW22"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(23, 3) = CheckStr(rsTmp.Fields("GW23"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(24, 3) = CheckStr(rsTmp.Fields("GW24"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(25, 3) = CheckStr(rsTmp.Fields("GW25"))
        Else
            GetUPMonthData ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Mid(oDate1, 1, 6) & "01"))), ChangeWDateStringToWString(DateAdd("d", -1, ChangeWStringToWDateString(Mid(oDate2, 1, 6) & "01")))
        End If
    Else
        GetUPMonthData ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(oDate1))), ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(oDate2)))
    End If
'計算  上月成長數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    grd1.TextMatrix(1, 4) = CheckStr(Val(grd1.TextMatrix(1, 2)) - Val(grd1.TextMatrix(1, 3)))
    grd1.TextMatrix(2, 4) = CheckStr(Val(grd1.TextMatrix(2, 2)) - Val(grd1.TextMatrix(2, 3)))
    grd1.TextMatrix(3, 4) = CheckStr(Val(grd1.TextMatrix(3, 2)) - Val(grd1.TextMatrix(3, 3)))
    grd1.TextMatrix(4, 4) = CheckStr(Val(grd1.TextMatrix(4, 2)) - Val(grd1.TextMatrix(4, 3)))
    grd1.TextMatrix(5, 4) = CheckStr(Val(grd1.TextMatrix(5, 2)) - Val(grd1.TextMatrix(5, 3)))
    grd1.TextMatrix(6, 4) = CheckStr(Val(grd1.TextMatrix(6, 2)) - Val(grd1.TextMatrix(6, 3)))
    grd1.TextMatrix(7, 4) = CheckStr(Val(grd1.TextMatrix(7, 2)) - Val(grd1.TextMatrix(7, 3)))
    grd1.TextMatrix(8, 4) = CheckStr(Val(grd1.TextMatrix(8, 2)) - Val(grd1.TextMatrix(8, 3)))
    grd1.TextMatrix(9, 4) = CheckStr(Val(grd1.TextMatrix(9, 2)) - Val(grd1.TextMatrix(9, 3)))
    grd1.TextMatrix(10, 4) = CheckStr(Val(grd1.TextMatrix(10, 2)) - Val(grd1.TextMatrix(10, 3)))
    grd1.TextMatrix(11, 4) = CheckStr(Val(grd1.TextMatrix(11, 2)) - Val(grd1.TextMatrix(11, 3)))
    grd1.TextMatrix(12, 4) = CheckStr(Val(grd1.TextMatrix(12, 2)) - Val(grd1.TextMatrix(12, 3)))
    grd1.TextMatrix(13, 4) = CheckStr(Val(grd1.TextMatrix(13, 2)) - Val(grd1.TextMatrix(13, 3)))
    grd1.TextMatrix(14, 4) = CheckStr(Val(grd1.TextMatrix(14, 2)) - Val(grd1.TextMatrix(14, 3)))
    grd1.TextMatrix(15, 4) = CheckStr(Val(grd1.TextMatrix(15, 2)) - Val(grd1.TextMatrix(15, 3)))
    grd1.TextMatrix(16, 4) = CheckStr(Val(grd1.TextMatrix(16, 2)) - Val(grd1.TextMatrix(16, 3)))
    grd1.TextMatrix(17, 4) = CheckStr(Val(grd1.TextMatrix(17, 2)) - Val(grd1.TextMatrix(17, 3)))
    grd1.TextMatrix(18, 4) = CheckStr(Val(grd1.TextMatrix(18, 2)) - Val(grd1.TextMatrix(18, 3)))
    grd1.TextMatrix(19, 4) = CheckStr(Val(grd1.TextMatrix(19, 2)) - Val(grd1.TextMatrix(19, 3)))
    grd1.TextMatrix(20, 4) = CheckStr(Val(grd1.TextMatrix(20, 2)) - Val(grd1.TextMatrix(20, 3)))
    grd1.TextMatrix(21, 4) = CheckStr(Val(grd1.TextMatrix(21, 2)) - Val(grd1.TextMatrix(21, 3)))
    grd1.TextMatrix(22, 4) = CheckStr(Val(grd1.TextMatrix(22, 2)) - Val(grd1.TextMatrix(22, 3)))
    grd1.TextMatrix(23, 4) = CheckStr(Val(grd1.TextMatrix(23, 2)) - Val(grd1.TextMatrix(23, 3)))
    grd1.TextMatrix(24, 4) = CheckStr(Val(grd1.TextMatrix(24, 2)) - Val(grd1.TextMatrix(24, 3)))
    grd1.TextMatrix(25, 4) = CheckStr(Val(grd1.TextMatrix(25, 2)) - Val(grd1.TextMatrix(25, 3)))
'計算  上月成長率
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    If Val(grd1.TextMatrix(1, 3)) <> 0 Then
        grd1.TextMatrix(1, 5) = Format(CheckStr(Val(grd1.TextMatrix(1, 4)) / Val(grd1.TextMatrix(1, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(1, 4)) <> 0 Then
            grd1.TextMatrix(1, 5) = "100 %"
        Else
            grd1.TextMatrix(1, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(2, 3)) <> 0 Then
        grd1.TextMatrix(2, 5) = Format(CheckStr(Val(grd1.TextMatrix(2, 4)) / Val(grd1.TextMatrix(2, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(2, 4)) <> 0 Then
            grd1.TextMatrix(2, 5) = "100 %"
        Else
            grd1.TextMatrix(2, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(3, 3)) <> 0 Then
        grd1.TextMatrix(3, 5) = Format(CheckStr(Val(grd1.TextMatrix(3, 4)) / Val(grd1.TextMatrix(3, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(3, 4)) <> 0 Then
            grd1.TextMatrix(3, 5) = "100 %"
        Else
            grd1.TextMatrix(3, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(4, 3)) <> 0 Then
        grd1.TextMatrix(4, 5) = Format(CheckStr(Val(grd1.TextMatrix(4, 4)) / Val(grd1.TextMatrix(4, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(4, 4)) <> 0 Then
            grd1.TextMatrix(4, 5) = "100 %"
        Else
            grd1.TextMatrix(4, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(5, 3)) <> 0 Then
        grd1.TextMatrix(5, 5) = Format(CheckStr(Val(grd1.TextMatrix(5, 4)) / Val(grd1.TextMatrix(5, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(5, 4)) Then
            grd1.TextMatrix(5, 5) = "100 %"
        Else
            grd1.TextMatrix(5, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(6, 3)) <> 0 Then
        grd1.TextMatrix(6, 5) = Format(CheckStr(Val(grd1.TextMatrix(6, 4)) / Val(grd1.TextMatrix(6, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(6, 4)) Then
            grd1.TextMatrix(6, 5) = "100 %"
        Else
            grd1.TextMatrix(6, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(7, 3)) <> 0 Then
        grd1.TextMatrix(7, 5) = Format(CheckStr(Val(grd1.TextMatrix(7, 4)) / Val(grd1.TextMatrix(7, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(7, 4)) <> 0 Then
            grd1.TextMatrix(7, 5) = "100 %"
        Else
            grd1.TextMatrix(7, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(8, 3)) <> 0 Then
        grd1.TextMatrix(8, 5) = Format(CheckStr(Val(grd1.TextMatrix(8, 4)) / Val(grd1.TextMatrix(8, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(8, 4)) <> 0 Then
            grd1.TextMatrix(8, 5) = "100 %"
        Else
            grd1.TextMatrix(8, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(9, 3)) <> 0 Then
        grd1.TextMatrix(9, 5) = Format(CheckStr(Val(grd1.TextMatrix(9, 4)) / Val(grd1.TextMatrix(9, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(9, 4)) <> 0 Then
            grd1.TextMatrix(9, 5) = "100 %"
        Else
            grd1.TextMatrix(9, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(10, 3)) <> 0 Then
        grd1.TextMatrix(10, 5) = Format(CheckStr(Val(grd1.TextMatrix(10, 4)) / Val(grd1.TextMatrix(10, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(10, 4)) <> 0 Then
            grd1.TextMatrix(10, 5) = "100 %"
        Else
            grd1.TextMatrix(10, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(11, 3)) <> 0 Then
        grd1.TextMatrix(11, 5) = Format(CheckStr(Val(grd1.TextMatrix(11, 4)) / Val(grd1.TextMatrix(11, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(11, 4)) <> 0 Then
            grd1.TextMatrix(11, 5) = "100 %"
        Else
            grd1.TextMatrix(11, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(12, 3)) <> 0 Then
        grd1.TextMatrix(12, 5) = Format(CheckStr(Val(grd1.TextMatrix(12, 4)) / Val(grd1.TextMatrix(12, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(12, 4)) <> 0 Then
            grd1.TextMatrix(12, 5) = "100 %"
        Else
            grd1.TextMatrix(12, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(13, 3)) <> 0 Then
        grd1.TextMatrix(13, 5) = Format(CheckStr(Val(grd1.TextMatrix(13, 4)) / Val(grd1.TextMatrix(13, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(13, 4)) <> 0 Then
            grd1.TextMatrix(13, 5) = "100 %"
        Else
            grd1.TextMatrix(13, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(14, 3)) <> 0 Then
        grd1.TextMatrix(14, 5) = Format(CheckStr(Val(grd1.TextMatrix(14, 4)) / Val(grd1.TextMatrix(14, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(14, 4)) <> 0 Then
            grd1.TextMatrix(14, 5) = "100 %"
        Else
            grd1.TextMatrix(14, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(15, 3)) <> 0 Then
        grd1.TextMatrix(15, 5) = Format(CheckStr(Val(grd1.TextMatrix(15, 4)) / Val(grd1.TextMatrix(15, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(15, 4)) <> 0 Then
            grd1.TextMatrix(15, 5) = "100 %"
        Else
            grd1.TextMatrix(15, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(16, 3)) <> 0 Then
        grd1.TextMatrix(16, 5) = Format(CheckStr(Val(grd1.TextMatrix(16, 4)) / Val(grd1.TextMatrix(16, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(16, 4)) <> 0 Then
            grd1.TextMatrix(16, 5) = "100 %"
        Else
            grd1.TextMatrix(16, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(17, 3)) <> 0 Then
        grd1.TextMatrix(17, 5) = Format(CheckStr(Val(grd1.TextMatrix(17, 4)) / Val(grd1.TextMatrix(17, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(17, 4)) <> 0 Then
            grd1.TextMatrix(17, 5) = "100 %"
        Else
            grd1.TextMatrix(17, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(18, 3)) <> 0 Then
        grd1.TextMatrix(18, 5) = Format(CheckStr(Val(grd1.TextMatrix(18, 4)) / Val(grd1.TextMatrix(18, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(18, 4)) <> 0 Then
            grd1.TextMatrix(18, 5) = "100 %"
        Else
            grd1.TextMatrix(18, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(19, 3)) <> 0 Then
        grd1.TextMatrix(19, 5) = Format(CheckStr(Val(grd1.TextMatrix(19, 4)) / Val(grd1.TextMatrix(19, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(19, 4)) <> 0 Then
            grd1.TextMatrix(19, 5) = "100 %"
        Else
            grd1.TextMatrix(19, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(20, 3)) <> 0 Then
        grd1.TextMatrix(20, 5) = Format(CheckStr(Val(grd1.TextMatrix(20, 4)) / Val(grd1.TextMatrix(20, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(20, 4)) <> 0 Then
            grd1.TextMatrix(20, 5) = "100 %"
        Else
            grd1.TextMatrix(20, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(21, 3)) <> 0 Then
        grd1.TextMatrix(21, 5) = Format(CheckStr(Val(grd1.TextMatrix(21, 4)) / Val(grd1.TextMatrix(21, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(21, 4)) <> 0 Then
            grd1.TextMatrix(21, 5) = "100 %"
        Else
            grd1.TextMatrix(21, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(22, 3)) <> 0 Then
        grd1.TextMatrix(22, 5) = Format(CheckStr(Val(grd1.TextMatrix(22, 4)) / Val(grd1.TextMatrix(22, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(22, 4)) <> 0 Then
            grd1.TextMatrix(22, 5) = "100 %"
        Else
            grd1.TextMatrix(22, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(23, 3)) <> 0 Then
        grd1.TextMatrix(23, 5) = Format(CheckStr(Val(grd1.TextMatrix(23, 4)) / Val(grd1.TextMatrix(23, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(23, 4)) <> 0 Then
            grd1.TextMatrix(23, 5) = "100 %"
        Else
            grd1.TextMatrix(23, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(24, 3)) <> 0 Then
        grd1.TextMatrix(24, 5) = Format(CheckStr(Val(grd1.TextMatrix(24, 4)) / Val(grd1.TextMatrix(24, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(24, 4)) <> 0 Then
            grd1.TextMatrix(24, 5) = "100 %"
        Else
            grd1.TextMatrix(24, 5) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(25, 3)) <> 0 Then
        grd1.TextMatrix(25, 5) = Format(CheckStr(Val(grd1.TextMatrix(25, 4)) / Val(grd1.TextMatrix(25, 3)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(24, 4)) <> 0 Then
            grd1.TextMatrix(25, 5) = "100 %"
        Else
            grd1.TextMatrix(25, 5) = "0 %"
        End If
    End If
'讀取  去年同期資料
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/08/15
    'strSQL = " select * from GeneralWork where substr(rtrim(ltrim(to_char(gw01))),1,6)='" & Mid(ChangeWDateStringToWString(DateAdd("yyyy", -1, ChangeWStringToWDateString(oDate))), 1, 6) & "' "
    If opt1(0).Value = True Then
        strSql = " select * from GeneralWork where gw01>=" & ChangeWDateStringToWString(DateAdd("yyyy", -1, ChangeWStringToWDateString(Mid(oDate1, 1, 6) & "01"))) & " and gw01<=" & ChangeWDateStringToWString(DateAdd("d", -1, DateAdd("yyyy", -1, DateAdd("m", 1, ChangeWStringToWDateString(Mid(oDate2, 1, 6) & "01"))))) & " "
        If rsTmp.State = 1 Then rsTmp.Close
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rsTmp.RecordCount <> 0 Then
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(1, 6) = CheckStr(rsTmp.Fields("GW02"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(2, 6) = CheckStr(rsTmp.Fields("GW03"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(3, 6) = CheckStr(rsTmp.Fields("GW04"))
            grd1.TextMatrix(4, 6) = CheckStr(rsTmp.Fields("GW05"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(5, 6) = CheckStr(rsTmp.Fields("GW06"))
            grd1.TextMatrix(6, 6) = CheckStr(rsTmp.Fields("GW07"))
            grd1.TextMatrix(7, 6) = CheckStr(rsTmp.Fields("GW08"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(8, 6) = CheckStr(rsTmp.Fields("GW09"))
            grd1.TextMatrix(9, 6) = CheckStr(rsTmp.Fields("GW26"))
            grd1.TextMatrix(10, 6) = CheckStr(rsTmp.Fields("GW10"))
            grd1.TextMatrix(11, 6) = CheckStr(rsTmp.Fields("GW11"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(12, 6) = CheckStr(rsTmp.Fields("GW12"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(13, 6) = CheckStr(rsTmp.Fields("GW13"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(14, 6) = CheckStr(rsTmp.Fields("GW14"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(15, 6) = CheckStr(rsTmp.Fields("GW15"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(16, 6) = CheckStr(rsTmp.Fields("GW16"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(17, 6) = CheckStr(rsTmp.Fields("GW17"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(18, 6) = CheckStr(rsTmp.Fields("GW18"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(19, 6) = CheckStr(rsTmp.Fields("GW19"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(20, 6) = CheckStr(rsTmp.Fields("GW20"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(21, 6) = CheckStr(rsTmp.Fields("GW21"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(22, 6) = CheckStr(rsTmp.Fields("GW22"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(23, 6) = CheckStr(rsTmp.Fields("GW23"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(24, 6) = CheckStr(rsTmp.Fields("GW24"))
            NowState = NowState + 1
            DoEvents
            Timer1_Timer
            grd1.TextMatrix(25, 6) = CheckStr(rsTmp.Fields("GW25"))
        Else
            GetUPYearData ChangeWDateStringToWString(DateAdd("yyyy", -1, ChangeWStringToWDateString(Mid(oDate1, 1, 6) & "01"))), ChangeWDateStringToWString(DateAdd("d", -1, DateAdd("yyyy", -1, DateAdd("m", 1, ChangeWStringToWDateString(Mid(oDate2, 1, 6) & "01")))))
        End If
    Else
        GetUPYearData ChangeWDateStringToWString(DateAdd("yyyy", -1, ChangeWStringToWDateString(oDate1))), ChangeWDateStringToWString(DateAdd("yyyy", -1, ChangeWStringToWDateString(oDate2)))
    End If
'計算  去年同期成長數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    grd1.TextMatrix(1, 7) = CheckStr(Val(grd1.TextMatrix(1, 2)) - Val(grd1.TextMatrix(1, 6)))
    grd1.TextMatrix(2, 7) = CheckStr(Val(grd1.TextMatrix(2, 2)) - Val(grd1.TextMatrix(2, 6)))
    grd1.TextMatrix(3, 7) = CheckStr(Val(grd1.TextMatrix(3, 2)) - Val(grd1.TextMatrix(3, 6)))
    grd1.TextMatrix(4, 7) = CheckStr(Val(grd1.TextMatrix(4, 2)) - Val(grd1.TextMatrix(4, 6)))
    grd1.TextMatrix(5, 7) = CheckStr(Val(grd1.TextMatrix(5, 2)) - Val(grd1.TextMatrix(5, 6)))
    grd1.TextMatrix(6, 7) = CheckStr(Val(grd1.TextMatrix(6, 2)) - Val(grd1.TextMatrix(6, 6)))
    grd1.TextMatrix(7, 7) = CheckStr(Val(grd1.TextMatrix(7, 2)) - Val(grd1.TextMatrix(7, 6)))
    grd1.TextMatrix(8, 7) = CheckStr(Val(grd1.TextMatrix(8, 2)) - Val(grd1.TextMatrix(8, 6)))
    grd1.TextMatrix(9, 7) = CheckStr(Val(grd1.TextMatrix(9, 2)) - Val(grd1.TextMatrix(9, 6)))
    grd1.TextMatrix(10, 7) = CheckStr(Val(grd1.TextMatrix(10, 2)) - Val(grd1.TextMatrix(10, 6)))
    grd1.TextMatrix(11, 7) = CheckStr(Val(grd1.TextMatrix(11, 2)) - Val(grd1.TextMatrix(11, 6)))
    grd1.TextMatrix(12, 7) = CheckStr(Val(grd1.TextMatrix(12, 2)) - Val(grd1.TextMatrix(12, 6)))
    grd1.TextMatrix(13, 7) = CheckStr(Val(grd1.TextMatrix(13, 2)) - Val(grd1.TextMatrix(13, 6)))
    grd1.TextMatrix(14, 7) = CheckStr(Val(grd1.TextMatrix(14, 2)) - Val(grd1.TextMatrix(14, 6)))
    grd1.TextMatrix(15, 7) = CheckStr(Val(grd1.TextMatrix(15, 2)) - Val(grd1.TextMatrix(15, 6)))
    grd1.TextMatrix(16, 7) = CheckStr(Val(grd1.TextMatrix(16, 2)) - Val(grd1.TextMatrix(16, 6)))
    grd1.TextMatrix(17, 7) = CheckStr(Val(grd1.TextMatrix(17, 2)) - Val(grd1.TextMatrix(17, 6)))
    grd1.TextMatrix(18, 7) = CheckStr(Val(grd1.TextMatrix(18, 2)) - Val(grd1.TextMatrix(18, 6)))
    grd1.TextMatrix(19, 7) = CheckStr(Val(grd1.TextMatrix(19, 2)) - Val(grd1.TextMatrix(19, 6)))
    grd1.TextMatrix(20, 7) = CheckStr(Val(grd1.TextMatrix(20, 2)) - Val(grd1.TextMatrix(20, 6)))
    grd1.TextMatrix(21, 7) = CheckStr(Val(grd1.TextMatrix(21, 2)) - Val(grd1.TextMatrix(21, 6)))
    grd1.TextMatrix(22, 7) = CheckStr(Val(grd1.TextMatrix(22, 2)) - Val(grd1.TextMatrix(22, 6)))
    grd1.TextMatrix(23, 7) = CheckStr(Val(grd1.TextMatrix(23, 2)) - Val(grd1.TextMatrix(23, 6)))
    grd1.TextMatrix(24, 7) = CheckStr(Val(grd1.TextMatrix(24, 2)) - Val(grd1.TextMatrix(24, 6)))
    grd1.TextMatrix(25, 7) = CheckStr(Val(grd1.TextMatrix(25, 2)) - Val(grd1.TextMatrix(25, 6)))
'計算  去年同期成長率
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    If Val(grd1.TextMatrix(1, 6)) <> 0 Then
        grd1.TextMatrix(1, 8) = Format(CheckStr(Val(grd1.TextMatrix(1, 7)) / Val(grd1.TextMatrix(1, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(1, 7)) <> 0 Then
            grd1.TextMatrix(1, 8) = "100 %"
        Else
            grd1.TextMatrix(1, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(2, 6)) <> 0 Then
        grd1.TextMatrix(2, 8) = Format(CheckStr(Val(grd1.TextMatrix(2, 7)) / Val(grd1.TextMatrix(2, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(2, 7)) <> 0 Then
            grd1.TextMatrix(2, 8) = "100 %"
        Else
            grd1.TextMatrix(2, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(3, 6)) <> 0 Then
        grd1.TextMatrix(3, 8) = Format(CheckStr(Val(grd1.TextMatrix(3, 7)) / Val(grd1.TextMatrix(3, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(3, 7)) <> 0 Then
            grd1.TextMatrix(3, 8) = "100 %"
        Else
            grd1.TextMatrix(3, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(4, 6)) <> 0 Then
        grd1.TextMatrix(4, 8) = Format(CheckStr(Val(grd1.TextMatrix(4, 7)) / Val(grd1.TextMatrix(4, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(4, 7)) <> 0 Then
            grd1.TextMatrix(4, 8) = "100 %"
        Else
            grd1.TextMatrix(4, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(5, 6)) <> 0 Then
        grd1.TextMatrix(5, 8) = Format(CheckStr(Val(grd1.TextMatrix(5, 7)) / Val(grd1.TextMatrix(5, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(5, 7)) <> 0 Then
            grd1.TextMatrix(5, 8) = "100 %"
        Else
            grd1.TextMatrix(5, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(6, 6)) <> 0 Then
        grd1.TextMatrix(6, 8) = Format(CheckStr(Val(grd1.TextMatrix(6, 7)) / Val(grd1.TextMatrix(6, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(6, 7)) <> 0 Then
            grd1.TextMatrix(6, 8) = "100 %"
        Else
            grd1.TextMatrix(6, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(7, 6)) <> 0 Then
        grd1.TextMatrix(7, 8) = Format(CheckStr(Val(grd1.TextMatrix(7, 7)) / Val(grd1.TextMatrix(7, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(7, 7)) <> 0 Then
            grd1.TextMatrix(7, 8) = "100 %"
        Else
            grd1.TextMatrix(7, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(8, 6)) <> 0 Then
        grd1.TextMatrix(8, 8) = Format(CheckStr(Val(grd1.TextMatrix(8, 7)) / Val(grd1.TextMatrix(8, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(8, 7)) <> 0 Then
            grd1.TextMatrix(8, 8) = "100 %"
        Else
            grd1.TextMatrix(8, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(9, 6)) <> 0 Then
        grd1.TextMatrix(9, 8) = Format(CheckStr(Val(grd1.TextMatrix(9, 7)) / Val(grd1.TextMatrix(9, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(9, 7)) <> 0 Then
            grd1.TextMatrix(9, 8) = "100 %"
        Else
            grd1.TextMatrix(9, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(10, 6)) <> 0 Then
        grd1.TextMatrix(10, 8) = Format(CheckStr(Val(grd1.TextMatrix(10, 7)) / Val(grd1.TextMatrix(10, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(10, 7)) <> 0 Then
            grd1.TextMatrix(10, 8) = "100 %"
        Else
            grd1.TextMatrix(10, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(11, 6)) <> 0 Then
        grd1.TextMatrix(11, 8) = Format(CheckStr(Val(grd1.TextMatrix(11, 7)) / Val(grd1.TextMatrix(11, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(11, 7)) <> 0 Then
            grd1.TextMatrix(11, 8) = "100 %"
        Else
            grd1.TextMatrix(11, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(12, 6)) <> 0 Then
        grd1.TextMatrix(12, 8) = Format(CheckStr(Val(grd1.TextMatrix(12, 7)) / Val(grd1.TextMatrix(12, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(12, 7)) <> 0 Then
            grd1.TextMatrix(12, 8) = "100 %"
        Else
            grd1.TextMatrix(12, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(13, 6)) <> 0 Then
        grd1.TextMatrix(13, 8) = Format(CheckStr(Val(grd1.TextMatrix(13, 7)) / Val(grd1.TextMatrix(13, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(13, 7)) <> 0 Then
            grd1.TextMatrix(13, 8) = "100 %"
        Else
            grd1.TextMatrix(13, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(14, 6)) <> 0 Then
        grd1.TextMatrix(14, 8) = Format(CheckStr(Val(grd1.TextMatrix(14, 7)) / Val(grd1.TextMatrix(14, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(14, 7)) <> 0 Then
            grd1.TextMatrix(14, 8) = "100 %"
        Else
            grd1.TextMatrix(14, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(15, 6)) <> 0 Then
        grd1.TextMatrix(15, 8) = Format(CheckStr(Val(grd1.TextMatrix(15, 7)) / Val(grd1.TextMatrix(15, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(15, 7)) <> 0 Then
            grd1.TextMatrix(15, 8) = "100 %"
        Else
            grd1.TextMatrix(15, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(16, 6)) <> 0 Then
        grd1.TextMatrix(16, 8) = Format(CheckStr(Val(grd1.TextMatrix(16, 7)) / Val(grd1.TextMatrix(16, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(16, 7)) <> 0 Then
            grd1.TextMatrix(16, 8) = "100 %"
        Else
            grd1.TextMatrix(16, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(17, 6)) <> 0 Then
        grd1.TextMatrix(17, 8) = Format(CheckStr(Val(grd1.TextMatrix(17, 7)) / Val(grd1.TextMatrix(17, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(17, 7)) <> 0 Then
            grd1.TextMatrix(17, 8) = "100 %"
        Else
            grd1.TextMatrix(17, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(18, 6)) <> 0 Then
        grd1.TextMatrix(18, 8) = Format(CheckStr(Val(grd1.TextMatrix(18, 7)) / Val(grd1.TextMatrix(18, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(18, 7)) <> 0 Then
            grd1.TextMatrix(18, 8) = "100 %"
        Else
            grd1.TextMatrix(18, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(19, 6)) <> 0 Then
        grd1.TextMatrix(19, 8) = Format(CheckStr(Val(grd1.TextMatrix(19, 7)) / Val(grd1.TextMatrix(19, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(19, 7)) <> 0 Then
            grd1.TextMatrix(19, 8) = "100 %"
        Else
            grd1.TextMatrix(19, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(20, 6)) <> 0 Then
        grd1.TextMatrix(20, 8) = Format(CheckStr(Val(grd1.TextMatrix(20, 7)) / Val(grd1.TextMatrix(20, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(20, 7)) <> 0 Then
            grd1.TextMatrix(20, 8) = "100 %"
        Else
            grd1.TextMatrix(20, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(21, 6)) <> 0 Then
        grd1.TextMatrix(21, 8) = Format(CheckStr(Val(grd1.TextMatrix(21, 7)) / Val(grd1.TextMatrix(21, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(21, 7)) <> 0 Then
            grd1.TextMatrix(21, 8) = "100 %"
        Else
            grd1.TextMatrix(21, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(22, 6)) <> 0 Then
        grd1.TextMatrix(22, 8) = Format(CheckStr(Val(grd1.TextMatrix(22, 7)) / Val(grd1.TextMatrix(22, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(22, 7)) <> 0 Then
            grd1.TextMatrix(22, 8) = "100 %"
        Else
            grd1.TextMatrix(22, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(23, 6)) <> 0 Then
        grd1.TextMatrix(23, 8) = Format(CheckStr(Val(grd1.TextMatrix(23, 7)) / Val(grd1.TextMatrix(23, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(23, 7)) <> 0 Then
            grd1.TextMatrix(23, 8) = "100 %"
        Else
            grd1.TextMatrix(23, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(24, 6)) <> 0 Then
        grd1.TextMatrix(24, 8) = Format(CheckStr(Val(grd1.TextMatrix(24, 7)) / Val(grd1.TextMatrix(24, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(24, 7)) <> 0 Then
            grd1.TextMatrix(24, 8) = "100 %"
        Else
            grd1.TextMatrix(24, 8) = "0 %"
        End If
    End If
    If Val(grd1.TextMatrix(25, 6)) <> 0 Then
        grd1.TextMatrix(25, 8) = Format(CheckStr(Val(grd1.TextMatrix(25, 7)) / Val(grd1.TextMatrix(25, 6)) * 100), "##0.00") & " %"
    Else
        If Val(grd1.TextMatrix(24, 7)) <> 0 Then
            grd1.TextMatrix(25, 8) = "100 %"
        Else
            grd1.TextMatrix(25, 8) = "0 %"
        End If
    End If
grd1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
On Error GoTo 0
On Error GoTo MsgErr
'存檔中.......  ' edit by nickc 2007/08/15 秀玲說不用存，之前的舊資料存就好，新的一律重算，若是資料不一致，以新系統運算為主
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    Dim oDate1, oDate2
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    oDate1 = ChangeTStringToWString(Trim(tmpD1))
    oDate2 = ChangeTStringToWString(Trim(tmpD2))
    'edit by nickc 2007/10/09 整月的才紀錄
    'If isPrt = True Then
    If isPrt = True And opt1(0).Value = True Then
        '查看有無本月資料
         With grd1
             'strSQL = " select * from GeneralWork where substr(rtrim(ltrim(to_char(gw01))),1,6)='" & Mid(oDate, 1, 6) & "' "
             strSql = " select * from GeneralWork where gw01>=" & oDate1 & " and gw01<=" & oDate2 & " "
             If rsTmp.State = 1 Then rsTmp.Close
             rsTmp.CursorLocation = adUseClient
             rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
             If rsTmp.RecordCount <> 0 Then
                 If Val(oDate1) >= 20071001 Then
                     strSql = "update GeneralWork set gw02=" & Val(.TextMatrix(1, 2)) & ",gw03=" & Val(.TextMatrix(2, 2)) & ",gw04=" & Val(.TextMatrix(3, 2)) & ",gw05=" & Val(.TextMatrix(4, 2)) & ",gw06=" & Val(.TextMatrix(5, 2)) & ",gw07=" & Val(.TextMatrix(6, 2)) & ",gw08=" & _
                                    Val(.TextMatrix(7, 2)) & ",gw09=" & Val(.TextMatrix(8, 2)) & ",gw10=" & Val(.TextMatrix(10, 2)) & ",gw11=" & Val(.TextMatrix(11, 2)) & ",gw12=" & Val(.TextMatrix(12, 2)) & ",gw13=" & Val(.TextMatrix(13, 2)) & ",gw14=" & Val(.TextMatrix(14, 2)) & ",gw15=" & _
                                    Val(.TextMatrix(15, 2)) & ",gw16=" & Val(.TextMatrix(16, 2)) & ",gw17=" & Val(.TextMatrix(17, 2)) & ",gw18=" & Val(.TextMatrix(18, 2)) & ",gw19=" & Val(.TextMatrix(19, 2)) & ",gw20=" & Val(.TextMatrix(20, 2)) & ",gw21=" & Val(.TextMatrix(21, 2)) & ",gw22=" & _
                                    Val(.TextMatrix(22, 2)) & ",gw23=" & Val(.TextMatrix(23, 2)) & ",gw24=" & Val(.TextMatrix(24, 2)) & ",gw25=" & Val(.TextMatrix(25, 2)) & ",gw26=" & Val(.TextMatrix(9, 2)) & " where gw01=" & oDate1 & " "
                     cnnConnection.Execute strSql
                 End If
             Else
                If Val(oDate1) >= 20071001 Then
                   strSql = "insert into GeneralWork (gw01,gw02,gw03,gw04,gw05,gw06,gw07,gw08,gw09,gw10,gw11,gw12,gw13,gw14,gw15,gw16,gw17,gw18,gw19,gw20,gw21,gw22,gw23,gw24,gw25,gw26) " & _
                                " values (" & oDate1 & "," & Val(.TextMatrix(1, 2)) & "," & Val(.TextMatrix(2, 2)) & "," & Val(.TextMatrix(3, 2)) & "," & Val(.TextMatrix(4, 2)) & "," & Val(.TextMatrix(5, 2)) & "," & Val(.TextMatrix(6, 2)) & "," & _
                                 Val(.TextMatrix(7, 2)) & "," & Val(.TextMatrix(8, 2)) & "," & Val(.TextMatrix(10, 2)) & "," & Val(.TextMatrix(11, 2)) & "," & Val(.TextMatrix(12, 2)) & "," & Val(.TextMatrix(13, 2)) & "," & Val(.TextMatrix(14, 2)) & "," & _
                                 Val(.TextMatrix(15, 2)) & "," & Val(.TextMatrix(16, 2)) & "," & Val(.TextMatrix(17, 2)) & "," & Val(.TextMatrix(18, 2)) & "," & Val(.TextMatrix(19, 2)) & "," & Val(.TextMatrix(20, 2)) & "," & Val(.TextMatrix(21, 2)) & "," & _
                                 Val(.TextMatrix(22, 2)) & "," & Val(.TextMatrix(23, 2)) & "," & Val(.TextMatrix(24, 2)) & "," & Val(.TextMatrix(25, 2)) & "," & Val(.TextMatrix(9, 2)) & ") "
                    cnnConnection.Execute strSql
                 End If
             End If
         End With
    End If
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    PrintTitle
    PrintDetil
    Printer.EndDoc
    NowState = NowState + 1
    Timer1_Timer
    DoEvents
    ShowPrintOk
    NowState = 0
    Timer1_Timer
    DoEvents
Exit Sub
MsgErr:
    MsgBox "列印發生錯誤！" & vbCrLf & Err.Description, vbInformation, "工作報告列印！"
End Sub

Private Sub txt1_GotFocus()
InverseTextBox txt1
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub txt1_Validate(Cancel As Boolean)
If txt1 <> "" Then
    If CheckIsTaiwanDate(txt1, False) = False Then
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
End Sub

Sub PrintTitle()
'畫框
If Printer.Orientation <> 1 Then
    Printer.Orientation = 1
End If
BoxTop = 1300 - 300 'edit by nickc 2007/11/01
BoxBotton = 16151 - 300 'edit by nickc 2007/11/01
BoxLeft = 300
BoxRight = 10800
oLines = 26
TheLBoxH = (BoxBotton - BoxTop) \ oLines
Printer.DrawWidth = 20
Printer.Line (BoxLeft, BoxTop)-(BoxRight, BoxBotton), , B
PLeft(1) = 350
PLeft(2) = 950
PLeft(3) = 2400
PLeft(4) = 3600
PLeft(5) = 4800
PLeft(6) = 6000
PLeft(7) = 7200
PLeft(8) = 8400
PLeft(9) = 9600

For i = 1 To oLines
    Select Case i
    Case 14, 18, 21, 25
        Printer.DrawWidth = 20
        Printer.Line (PLeft(1) - 50, BoxBotton - (TheLBoxH * i))-(BoxRight, BoxBotton - (TheLBoxH * i))
    Case Else
        Printer.DrawWidth = 1
        Printer.Line (PLeft(2) - 50, BoxBotton - (TheLBoxH * i))-(BoxRight, BoxBotton - (TheLBoxH * i))
    End Select
Next i

Printer.Line (PLeft(2) - 50, BoxTop + TheLBoxH)-(PLeft(2) - 50, BoxBotton)
Printer.Line (PLeft(3) - 50, BoxTop)-(PLeft(3) - 50, BoxBotton)
Printer.Line (PLeft(4) - 50, BoxTop)-(PLeft(4) - 50, BoxBotton)
Printer.Line (PLeft(5) - 50, BoxTop)-(PLeft(5) - 50, BoxBotton)
Printer.Line (PLeft(6) - 50, BoxTop)-(PLeft(6) - 50, BoxBotton)
Printer.Line (PLeft(7) - 50, BoxTop)-(PLeft(7) - 50, BoxBotton)
Printer.Line (PLeft(8) - 50, BoxTop)-(PLeft(8) - 50, BoxBotton)
Printer.Line (PLeft(9) - 50, BoxTop)-(PLeft(9) - 50, BoxBotton)
Printer.Font.Size = 20
'edit by nickc 2007/11/01  移掉
'Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("台一國際專利商標事務所") / 2
'Printer.CurrentY = 300
'Printer.Print "台一國際專利商標事務所"
Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("總務處工作報告") / 2
Printer.CurrentY = 700 - 300 'edit by nickc 2007/11/01
Printer.Print "總務處工作報告" '
Printer.Font.Size = 14
Printer.CurrentX = 0 + 300
Printer.CurrentY = 1000 - 300 'edit by nickc 2007/11/01
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(Trim(Val(Mid(ChangeTStringToWString(txt1), 1, 4)) - 1911) & " 年 " & Mid(ChangeTStringToWString(txt1), 5, 2) & " 月") - 600
Printer.CurrentY = 1000 - 300 'edit by nickc 2007/11/01
Printer.Print Trim(Val(Mid(ChangeTStringToWString(txt1), 1, 4)) - 1911) & " 年 " & Mid(ChangeTStringToWString(txt1), 5, 2) & " 月"
With grd1
    Printer.CurrentX = PLeft(3) + ((PLeft(4) - PLeft(3) - 50) / 2) - (Printer.TextWidth(.TextMatrix(0, 2)) / 2)
    Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(0, 2)) / 2)
    Printer.Print .TextMatrix(0, 2)
    Printer.CurrentX = PLeft(4) + ((PLeft(5) - PLeft(4) - 50) / 2) - (Printer.TextWidth(.TextMatrix(0, 3)) / 2)
    Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(0, 3)) / 2)
    Printer.Print .TextMatrix(0, 3)
    Printer.CurrentX = PLeft(5) + ((PLeft(6) - PLeft(5) - 50) / 2) - (Printer.TextWidth(.TextMatrix(0, 4)) / 2)
    Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(0, 4)) / 2)
    Printer.Print .TextMatrix(0, 4)
    Printer.CurrentX = PLeft(6) + ((PLeft(7) - PLeft(6) - 50) / 2) - (Printer.TextWidth(.TextMatrix(0, 5)) / 2)
    Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(0, 5)) / 2)
    Printer.Print .TextMatrix(0, 5)
    Printer.CurrentX = PLeft(7) + ((PLeft(8) - PLeft(7) - 50) / 2) - (Printer.TextWidth(.TextMatrix(0, 6)) / 2)
    Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(0, 6)) / 2)
    Printer.Print .TextMatrix(0, 6)
    Printer.CurrentX = PLeft(8) + ((PLeft(9) - PLeft(8) - 50) / 2) - (Printer.TextWidth(.TextMatrix(0, 7)) / 2)
    Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(0, 7)) / 2)
    Printer.Print .TextMatrix(0, 7)
    Printer.CurrentX = PLeft(9) + ((BoxRight - PLeft(9) - 50) / 2) - (Printer.TextWidth(.TextMatrix(0, 8)) / 2)
    Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(0, 8)) / 2)
    Printer.Print .TextMatrix(0, 8)
    For i = 1 To oLines - 1
        If i = oLines - 2 Then
            Printer.CurrentX = PLeft(2) + ((PLeft(3) - PLeft(2) - 50) / 2) - (Printer.TextWidth(Mid(.TextMatrix(i, 1), 1, 5)) / 2)
            Printer.CurrentY = BoxTop + (i * TheLBoxH)
            Printer.Print Mid(.TextMatrix(i, 1), 1, 5)
            Printer.CurrentX = PLeft(2) + ((PLeft(3) - PLeft(2) - 50) / 2) - (Printer.TextWidth(Mid(.TextMatrix(i, 1), 6)) / 2)
            Printer.CurrentY = BoxTop + (i * TheLBoxH) + (TheLBoxH / 2)
            Printer.Print Mid(.TextMatrix(i, 1), 6)
        Else
            Printer.CurrentX = PLeft(2) + ((PLeft(3) - PLeft(2) - 50) / 2) - (Printer.TextWidth(.TextMatrix(i, 1)) / 2)
            Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(i, 1)) / 2) + (i * TheLBoxH)
            Printer.Print .TextMatrix(i, 1)
            If i = 1 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("收") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("收") / 2) + (i * TheLBoxH)
                Printer.Print "收"
            ElseIf i = 2 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("發") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("發") / 2) + (i * TheLBoxH)
                Printer.Print "發"
            ElseIf i = 3 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("文") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("文") / 2) + (i * TheLBoxH)
                Printer.Print "文"
            ElseIf i = 4 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("數") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("數") / 2) + (i * TheLBoxH)
                Printer.Print "數"
            ElseIf i = 5 Then
                Printer.CurrentX = PLeft(1)
                Printer.CurrentY = BoxTop + (i * TheLBoxH) + Printer.TextHeight("新") + Printer.TextHeight("增")
                Printer.Print "新"
                Printer.CurrentX = PLeft(1)
                Printer.CurrentY = BoxTop + (i * TheLBoxH) + Printer.TextHeight("新") + Printer.TextHeight("新") + Printer.TextHeight("增")
                Printer.Print "增"
                Printer.CurrentX = PLeft(1) + Printer.TextWidth("新")
                Printer.CurrentY = BoxTop + (i * TheLBoxH)
                Printer.Print "客"
                Printer.CurrentX = PLeft(1) + Printer.TextWidth("新")
                Printer.CurrentY = BoxTop + (i * TheLBoxH) + Printer.TextHeight("客")
                Printer.Print "戶"
                Printer.CurrentX = PLeft(1) + Printer.TextWidth("新")
                Printer.CurrentY = BoxTop + (i * TheLBoxH) + Printer.TextHeight("客") + Printer.TextHeight("戶")
                Printer.Print "／"
                Printer.CurrentX = PLeft(1) + Printer.TextWidth("新")
                Printer.CurrentY = BoxTop + (i * TheLBoxH) + Printer.TextHeight("客") + Printer.TextHeight("戶") + Printer.TextHeight("／")
                Printer.Print "代"
                Printer.CurrentX = PLeft(1) + Printer.TextWidth("新")
                Printer.CurrentY = BoxTop + (i * TheLBoxH) + Printer.TextHeight("客") + Printer.TextHeight("戶") + Printer.TextHeight("／") + Printer.TextHeight("代")
                Printer.Print "理"
                Printer.CurrentX = PLeft(1) + Printer.TextWidth("新")
                Printer.CurrentY = BoxTop + (i * TheLBoxH) + Printer.TextHeight("客") + Printer.TextHeight("戶") + Printer.TextHeight("／") + Printer.TextHeight("代") + Printer.TextHeight("理")
                Printer.Print "人"
            ElseIf i = 8 Then
'                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("客") / 2)
'                Printer.CurrentY = BoxTop + Printer.TextHeight("客") + (i * TheLBoxH)
'                Printer.Print "客"
'                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("戶") / 2)
'                Printer.CurrentY = BoxTop + Printer.TextHeight("客") + (i * TheLBoxH) + Printer.TextHeight("客")
'                Printer.Print "戶"
'                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("來") / 2)
'                Printer.CurrentY = BoxTop + Printer.TextHeight("客") + (i * TheLBoxH) + Printer.TextHeight("客") + Printer.TextHeight("戶")
'                Printer.Print "來"
'                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("訪") / 2)
'                Printer.CurrentY = BoxTop + Printer.TextHeight("客") + (i * TheLBoxH) + Printer.TextHeight("客") + Printer.TextHeight("戶") + Printer.TextWidth("來")
'                Printer.Print "訪"
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("客") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("客") / 2) + (i * TheLBoxH)
                Printer.Print "客"
            ElseIf i = 9 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("戶") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("戶") / 2) + (i * TheLBoxH)
                Printer.Print "戶"
            ElseIf i = 10 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("來") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("來") / 2) + (i * TheLBoxH)
                Printer.Print "來"
            ElseIf i = 11 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("訪") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("訪") / 2) + (i * TheLBoxH)
                Printer.Print "訪"
            ElseIf i = 14 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("檔") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("檔") / 2) + (i * TheLBoxH)
                Printer.Print "檔"
            ElseIf i = 17 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("案") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("案") / 2) + (i * TheLBoxH)
                Printer.Print "案"
            ElseIf i = 20 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("新") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("新") / 2) + (i * TheLBoxH)
                Printer.Print "新"
            ElseIf i = 23 Then
                Printer.CurrentX = PLeft(1) + ((PLeft(2) - PLeft(1) - 50) / 2) - (Printer.TextWidth("卷") / 2)
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight("卷") / 2) + (i * TheLBoxH)
                Printer.Print "卷"
            End If
        End If
    Next i
End With
End Sub

Sub PrintDetil()
With grd1
        For i = 1 To 25
                Printer.CurrentX = PLeft(4) - 100 - Printer.TextWidth(Format(Val(.TextMatrix(i, 2)), "###,###,##0"))
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(i, 2)) / 2) + (i * TheLBoxH)
                Printer.Print Format(Val(.TextMatrix(i, 2)), "###,###,##0")
                Printer.CurrentX = PLeft(5) - 100 - Printer.TextWidth(Format(Val(.TextMatrix(i, 3)), "###,###,##0"))
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(i, 3)) / 2) + (i * TheLBoxH)
                Printer.Print Format(Val(.TextMatrix(i, 3)), "###,###,##0")
                Printer.CurrentX = PLeft(6) - 100 - Printer.TextWidth(Format(Val(.TextMatrix(i, 4)), "###,###,##0"))
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(i, 4)) / 2) + (i * TheLBoxH)
                Printer.Print Format(Val(.TextMatrix(i, 4)), "###,###,##0")
                Printer.CurrentX = PLeft(7) - 100 - Printer.TextWidth(.TextMatrix(i, 5))
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(i, 5)) / 2) + (i * TheLBoxH)
                Printer.Print .TextMatrix(i, 5)
                Printer.CurrentX = PLeft(8) - 100 - Printer.TextWidth(Format(Val(.TextMatrix(i, 6)), "###,###,##0"))
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(i, 6)) / 2) + (i * TheLBoxH)
                Printer.Print Format(Val(.TextMatrix(i, 6)), "###,###,##0")
                Printer.CurrentX = PLeft(9) - 100 - Printer.TextWidth(Format(Val(.TextMatrix(i, 7)), "###,###,##0"))
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(i, 7)) / 2) + (i * TheLBoxH)
                Printer.Print Format(Val(.TextMatrix(i, 7)), "###,###,##0")
                Printer.CurrentX = BoxRight - 100 - Printer.TextWidth(.TextMatrix(i, 8))
                Printer.CurrentY = BoxTop + (TheLBoxH / 2) - (Printer.TextHeight(.TextMatrix(i, 8)) / 2) + (i * TheLBoxH)
                Printer.Print .TextMatrix(i, 8)
        Next i
End With
End Sub

Private Sub SetGrd()
   Dim arrGridHeadWidth
   Dim iRow As Integer
   arrGridHeadWidth = Array(800, 800, 850, 850, 850, 850, 850, 850, 850)
   grd1.Cols = UBound(arrGridHeadWidth) + 1
   For iRow = 0 To grd1.Cols - 1
      grd1.row = 0
      grd1.col = iRow
      grd1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd1.CellAlignment = flexAlignCenterCenter
   Next
    If grd1.Rows >= 2 Then
        grd1.Visible = False
        For iRow = 1 To grd1.Rows - 1
           grd1.row = iRow
           grd1.col = 0
           grd1.CellAlignment = flexAlignCenterCenter
           grd1.col = 1
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 2
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 3
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 4
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 5
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 6
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 7
           grd1.CellAlignment = flexAlignRightCenter
           grd1.col = 8
           grd1.CellAlignment = flexAlignRightCenter
        Next
        grd1.Visible = True
    End If
End Sub

Private Sub txt2_GotFocus()
InverseTextBox txt2
End Sub

Private Sub txt2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 27 Or KeyAscii = 113 Or KeyAscii = 114 Or KeyAscii = 120 Or KeyAscii = 121 Then
Else
    KeyAscii = 0
End If
End Sub

Private Sub txt2_Validate(Cancel As Boolean)
If txt2 <> "" Then
    If CheckIsTaiwanDate(txt2, False) = False Then
        Cancel = True
    End If
End If
If Cancel = True Then MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
End Sub

'add by nickc 2007/08/15 計算上月資料
Sub GetUPMonthData(oDate1 As String, oDate2 As String)
DoEvents
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
'計算  [收文]總數
    '案件
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    strSql = "select nvl(count(cp09),0) from ("
    'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
    strSql = strSql & " select cp09 from caseprogress where cp66>=" & oDate1 & " and cp66<=" & oDate2 & " and substr(cp09,1,1)='A' and cp01||cp02<>'TT999999' "
    strSql = strSql & " union select dd14 from datadeleterecord where dd25>=" & oDate1 & " and dd25<=" & oDate2 & " and substr(dd14,1,1)='A' and dd18 is not null  "
    'add by sonia 2016/5/23 剔除已刪除但又救回來的進度T-203085之申請
    'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
    strSql = strSql & "          and dd14 not in (select cp09 from caseprogress where cp66>=" & oDate1 & " and cp66<=" & oDate2 & " and substr(cp09,1,1)='A' and cp01||cp02<>'TT999999') "
    'end 2016/5/23
    strSql = strSql & " ) AAAAA "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(1, 3) = CheckStr(rsTmp.Fields(0))
        grd1.TextMatrix(2, 3) = CheckStr(rsTmp.Fields(0))
    End If
    '政府機關        因為有些是沒有收進去  例如法院來的信件，但是會登記在簿子上   或是昨天漏輸，今天補上，或是漏登記卻有輸入電腦，所以誤差不管多或是少  都是對的
    DoEvents
    Timer1_Timer
    strSql = "select nvl(count(*),0) from mailrec where mr02>=" & oDate1 & " and mr02<=" & oDate2 & "  "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(1, 3) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(1, 3)))
    End If
    '信件
    DoEvents
    Timer1_Timer
    strSql = "select nvl(count(*),0) from letterinput where li01>=" & oDate1 & "  and li01<=" & oDate2 & " "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(1, 3) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(1, 3)))
    End If
'計算  [業務收文]數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'strSQL = "select count(*) from caseprogress where substr(rtrim(ltrim(to_char(cp05))),1,6)='200705' and substr(cp09,1,1)='A'  "
'計算  [發文總數、標準局發文]數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    strSql = "select nvl(sum(nvl(gd02,0)),0),nvl(sum(nvl(gd03,0)),0) from GeneralDispatch where gd01>=" & oDate1 & " and gd01<=" & oDate2 & " "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(3, 3) = CheckStr(rsTmp.Fields(0))
        grd1.TextMatrix(4, 3) = CheckStr(rsTmp.Fields(1))
    End If
'計算  新增[國內、國外]客戶數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    '國內客戶
    strSql = "select nvl(count(*),0) from customer where cu14>=" & oDate1 & " and cu14<=" & oDate2 & " and ((cu10 in ('020','013')) or cu10<='010')  and cu02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(5, 3) = CheckStr(rsTmp.Fields(0))
    End If
    '國外客戶
    'edit by nickc 2007/08/15
    strSql = "select nvl(count(*),0) from customer where cu14>=" & oDate1 & " and cu14<=" & oDate2 & " and not ((cu10 in ('020','013')) or cu10<='010') and cu02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(6, 3) = CheckStr(rsTmp.Fields(0))
    End If
    '國內代理人
    strSql = "select nvl(count(*),0) from fagent where fa11>=" & oDate1 & " and fa11<=" & oDate2 & " and ((fa10 in ('020','013')) or fa10<='010') and fa02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(5, 3) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(5, 3)))
    End If
    '國內代理人
    strSql = "select nvl(count(*),0) from fagent where fa11>=" & oDate1 & " and fa11<=" & oDate2 & " and not ((fa10 in ('020','013')) or fa10<='010') and fa02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(6, 3) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(6, 3)))
    End If
    grd1.TextMatrix(7, 3) = Trim(Val(grd1.TextMatrix(5, 3)) + Val(grd1.TextMatrix(6, 3)))
'計算  [客戶來訪]數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/11/01
    'strSQL = "select nvl(sum(nvl(oi06,0)),0),nvl(sum(nvl(oi07,0)),0) from otherinput where oi01>=" & oDate1 & " and oi01<=" & oDate2 & "  "
    strSql = "select nvl(sum(nvl(oi06,0)),0),nvl(sum(nvl(oi14,0)),0),nvl(sum(nvl(oi07,0)),0) from otherinput where oi01>=" & oDate1 & " and oi01<=" & oDate2 & "  "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(8, 3) = CheckStr(rsTmp.Fields(0))
        grd1.TextMatrix(9, 3) = CheckStr(rsTmp.Fields(1))
        grd1.TextMatrix(10, 3) = CheckStr(rsTmp.Fields(2))
    End If
    grd1.TextMatrix(11, 3) = Trim(Val(grd1.TextMatrix(8, 3)) + Val(grd1.TextMatrix(9, 3)) + Val(grd1.TextMatrix(10, 3)))
'計算新卷
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/08/06 原先自動抓，現在抓輸入的 CreateFiles Table
    'strSQL = " select decode(cp01,'TD','TB','TS','TB','TM','TB','TF','TB','S','CFC','FG','FCP',cp01) as cp01,nvl(count(*),0) from caseprogress where substr(rtrim(ltrim(to_char(cp27))),1,6)='" & Mid(oDate, 1, 6) & "' and cp31='Y' and cp01 not in ('CPS','PS') group by decode(cp01,'TD','TB','TS','TB','TM','TB','TF','TB','S','CFC','FG','FCP',cp01) "
    'CFT 只有 101 、FCT、FCP 101-105 與收文量查詢相同
    'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
    strSql = "SELECT CP01,COUNT(CP09) FROM CASEPROGRESS " & _
                    " Where cp05 >= " & oDate1 & " And cp05 <= " & oDate2 & " And cp26 Is Null And cp21 Is Null and cp01||cp02<>'TT999999' " & _
                    "  AND CP09< 'B'  AND CP01||cp10 IN ('FCP101','CFP101','FCP102','FCP103','FCP104','FCP105','CFP102','CFP103','CFP104','CFP105','CFT101')  GROUP BY CP01 "
    strSql = strSql & " union select 'T',sum(nvl(cf02,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'P',sum(nvl(cf03,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'TC',sum(nvl(cf04,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'LA',sum(nvl(cf05,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'L',sum(nvl(cf06,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'CFC',sum(nvl(cf07,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'FCT',sum(nvl(cf08,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'FCL',sum(nvl(cf09,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'CFL',sum(nvl(cf10,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'TB',sum(nvl(cf11,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
   
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        '計算  新卷[商標]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(12, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "T"
                    grd1.TextMatrix(12, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[專利]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(13, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "P"
                    grd1.TextMatrix(13, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[著作權] 數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(14, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "TC"
                    grd1.TextMatrix(14, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[顧問]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(15, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "LA"
                    grd1.TextMatrix(15, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[法務]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(16, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "L"
                    grd1.TextMatrix(16, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFT]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(17, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFT"
                    grd1.TextMatrix(17, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFP]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(18, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFP"
                    grd1.TextMatrix(18, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFC]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(19, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFC"
                    grd1.TextMatrix(19, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[FCT]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(20, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "FCT"
                    grd1.TextMatrix(20, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[FCP]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(21, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "FCP"
                    grd1.TextMatrix(21, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[FCL]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(22, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "FCL"
                    grd1.TextMatrix(22, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFL]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(23, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFL"
                    grd1.TextMatrix(23, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[B、D、S、M、F]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(24, 3) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "TB"
                    grd1.TextMatrix(24, 3) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        grd1.TextMatrix(25, 3) = Trim(Val(grd1.TextMatrix(12, 3)) + Val(grd1.TextMatrix(13, 3)) + Val(grd1.TextMatrix(14, 3)) + Val(grd1.TextMatrix(15, 3)) + Val(grd1.TextMatrix(16, 3)) + Val(grd1.TextMatrix(17, 3)) + Val(grd1.TextMatrix(18, 3)) + Val(grd1.TextMatrix(19, 3)) + Val(grd1.TextMatrix(20, 3)) + Val(grd1.TextMatrix(21, 3)) + Val(grd1.TextMatrix(22, 3)) + Val(grd1.TextMatrix(23, 3)) + Val(grd1.TextMatrix(24, 3)))
    Else
        grd1.TextMatrix(11, 3) = "0"
        grd1.TextMatrix(12, 3) = "0"
        grd1.TextMatrix(13, 3) = "0"
        grd1.TextMatrix(14, 3) = "0"
        grd1.TextMatrix(15, 3) = "0"
        grd1.TextMatrix(16, 3) = "0"
        grd1.TextMatrix(17, 3) = "0"
        grd1.TextMatrix(18, 3) = "0"
        grd1.TextMatrix(19, 3) = "0"
        grd1.TextMatrix(20, 3) = "0"
        grd1.TextMatrix(21, 3) = "0"
        grd1.TextMatrix(22, 3) = "0"
        grd1.TextMatrix(23, 3) = "0"
        grd1.TextMatrix(24, 3) = "0"
        grd1.TextMatrix(25, 3) = "0"
    End If
End Sub

'add by nickc 2007/08/15 計算去年同期資料
Sub GetUPYearData(oDate1 As String, oDate2 As String)
DoEvents
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
'計算  [收文]總數
    '案件
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    strSql = "select nvl(count(cp09),0) from ("
    'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
    strSql = strSql & " select cp09 from caseprogress where cp66>=" & oDate1 & " and cp66<=" & oDate2 & " and substr(cp09,1,1)='A' and cp01||cp02<>'TT999999' "
    strSql = strSql & " union select dd14 from datadeleterecord where dd25>=" & oDate1 & " and dd25<=" & oDate2 & " and substr(dd14,1,1)='A' and dd18 is not null  "
    'add by sonia 2016/5/23 剔除已刪除但又救回來的進度T-203085之申請
    'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
    strSql = strSql & "          and dd14 not in (select cp09 from caseprogress where cp66>=" & oDate1 & " and cp66<=" & oDate2 & " and substr(cp09,1,1)='A' and cp01||cp02<>'TT999999') "
    'end 2016/5/23
    strSql = strSql & " ) AAAAA "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(1, 6) = CheckStr(rsTmp.Fields(0))
        grd1.TextMatrix(2, 6) = CheckStr(rsTmp.Fields(0))
    End If
    '政府機關        因為有些是沒有收進去  例如法院來的信件，但是會登記在簿子上   或是昨天漏輸，今天補上，或是漏登記卻有輸入電腦，所以誤差不管多或是少  都是對的
    DoEvents
    Timer1_Timer
    strSql = "select nvl(count(*),0) from mailrec where mr02>=" & oDate1 & " and mr02<=" & oDate2 & "  "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(1, 6) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(1, 6)))
    End If
    '信件
    DoEvents
    Timer1_Timer
    strSql = "select nvl(count(*),0) from letterinput where li01>=" & oDate1 & "  and li01<=" & oDate2 & " "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(1, 6) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(1, 6)))
    End If
'計算  [業務收文]數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'strSQL = "select count(*) from caseprogress where substr(rtrim(ltrim(to_char(cp05))),1,6)='200705' and substr(cp09,1,1)='A'  "
'計算  [發文總數、標準局發文]數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    strSql = "select nvl(sum(nvl(gd02,0)),0),nvl(sum(nvl(gd03,0)),0) from GeneralDispatch where gd01>=" & oDate1 & " and gd01<=" & oDate2 & " "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(3, 6) = CheckStr(rsTmp.Fields(0))
        grd1.TextMatrix(4, 6) = CheckStr(rsTmp.Fields(1))
    End If
'計算  新增[國內、國外]客戶數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    '國內客戶
    strSql = "select nvl(count(*),0) from customer where cu14>=" & oDate1 & " and cu14<=" & oDate2 & " and ((cu10 in ('020','013')) or cu10<='010')  and cu02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(5, 6) = CheckStr(rsTmp.Fields(0))
    End If
    '國外客戶
    'edit by nickc 2007/08/15
    strSql = "select nvl(count(*),0) from customer where cu14>=" & oDate1 & " and cu14<=" & oDate2 & " and not ((cu10 in ('020','013')) or cu10<='010') and cu02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(6, 6) = CheckStr(rsTmp.Fields(0))
    End If
    '國內代理人
    strSql = "select nvl(count(*),0) from fagent where fa11>=" & oDate1 & " and fa11<=" & oDate2 & " and ((fa10 in ('020','013')) or fa10<='010') and fa02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(5, 6) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(5, 6)))
    End If
    '國內代理人
    strSql = "select nvl(count(*),0) from fagent where fa11>=" & oDate1 & " and fa11<=" & oDate2 & " and not ((fa10 in ('020','013')) or fa10<='010') and fa02='0' "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(6, 6) = Trim(Val(CheckStr(rsTmp.Fields(0))) + Val(grd1.TextMatrix(6, 6)))
    End If
    grd1.TextMatrix(7, 6) = Trim(Val(grd1.TextMatrix(5, 6)) + Val(grd1.TextMatrix(6, 6)))
'計算  [客戶來訪]數
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/11/01
    'strSQL = "select nvl(sum(nvl(oi06,0)),0),nvl(sum(nvl(oi07,0)),0) from otherinput where oi01>=" & oDate1 & " and oi01<=" & oDate2 & "  "
    strSql = "select nvl(sum(nvl(oi06,0)),0),nvl(sum(nvl(oi14,0)),0),nvl(sum(nvl(oi07,0)),0) from otherinput where oi01>=" & oDate1 & " and oi01<=" & oDate2 & "  "
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        grd1.TextMatrix(8, 6) = CheckStr(rsTmp.Fields(0))
        grd1.TextMatrix(9, 6) = CheckStr(rsTmp.Fields(1))
        grd1.TextMatrix(10, 6) = CheckStr(rsTmp.Fields(2))
    End If
    grd1.TextMatrix(11, 6) = Trim(Val(grd1.TextMatrix(8, 6)) + Val(grd1.TextMatrix(9, 6)) + Val(grd1.TextMatrix(10, 6)))
'計算新卷
    NowState = NowState + 1
    DoEvents
    Timer1_Timer
    'edit by nickc 2007/08/06 原先自動抓，現在抓輸入的 CreateFiles Table
    'strSQL = " select decode(cp01,'TD','TB','TS','TB','TM','TB','TF','TB','S','CFC','FG','FCP',cp01) as cp01,nvl(count(*),0) from caseprogress where substr(rtrim(ltrim(to_char(cp27))),1,6)='" & Mid(oDate, 1, 6) & "' and cp31='Y' and cp01 not in ('CPS','PS') group by decode(cp01,'TD','TB','TS','TB','TM','TB','TF','TB','S','CFC','FG','FCP',cp01) "
    'CFT 只有 101 、FCT、FCP 101-105 與收文量查詢相同
    'Modified by Lydia 2022/09/27 排除TT-999999案號 and cp01||cp02<>'TT999999'
    strSql = "SELECT CP01,COUNT(CP09) FROM CASEPROGRESS " & _
                    " Where cp05 >= " & oDate1 & " And cp05 <= " & oDate2 & " And cp26 Is Null And cp21 Is Null and cp01||cp02<>'TT999999' " & _
                    "  AND CP09< 'B'  AND CP01||cp10 IN ('FCP101','CFP101','FCP102','FCP103','FCP104','FCP105','CFP102','CFP103','CFP104','CFP105','CFT101')  GROUP BY CP01 "
    strSql = strSql & " union select 'T',sum(nvl(cf02,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'P',sum(nvl(cf03,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'TC',sum(nvl(cf04,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'LA',sum(nvl(cf05,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'L',sum(nvl(cf06,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'CFC',sum(nvl(cf07,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'FCT',sum(nvl(cf08,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'FCL',sum(nvl(cf09,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'CFL',sum(nvl(cf10,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
    strSql = strSql & " union select 'TB',sum(nvl(cf11,0)) from createfiles where cf01>=" & oDate1 & " and cf01<=" & oDate2 & " "
   
    If rsTmp.State = 1 Then rsTmp.Close
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <> 0 Then
        '計算  新卷[商標]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(12, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "T"
                    grd1.TextMatrix(12, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[專利]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(13, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "P"
                    grd1.TextMatrix(13, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[著作權] 數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(14, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "TC"
                    grd1.TextMatrix(14, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[顧問]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(15, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "LA"
                    grd1.TextMatrix(15, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[法務]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(16, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "L"
                    grd1.TextMatrix(16, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFT]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(17, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFT"
                    grd1.TextMatrix(17, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFP]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(18, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFP"
                    grd1.TextMatrix(18, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFC]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(19, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFC"
                    grd1.TextMatrix(19, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[FCT]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(20, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "FCT"
                    grd1.TextMatrix(20, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[FCP]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(21, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "FCP"
                    grd1.TextMatrix(21, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[FCL]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(22, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "FCL"
                    grd1.TextMatrix(22, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[CFL]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(23, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "CFL"
                    grd1.TextMatrix(23, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        '計算  新卷[B、D、S、M、F]數
        NowState = NowState + 1
        DoEvents
        Timer1_Timer
        grd1.TextMatrix(24, 6) = "0"
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Select Case CheckStr(rsTmp.Fields(0))
            Case "TB"
                    grd1.TextMatrix(24, 6) = Val(CheckStr(rsTmp.Fields(1)))
                    Exit Do
            End Select
            rsTmp.MoveNext
        Loop
        grd1.TextMatrix(25, 6) = Trim(Val(grd1.TextMatrix(12, 6)) + Val(grd1.TextMatrix(13, 6)) + Val(grd1.TextMatrix(14, 6)) + Val(grd1.TextMatrix(15, 6)) + Val(grd1.TextMatrix(16, 6)) + Val(grd1.TextMatrix(17, 6)) + Val(grd1.TextMatrix(18, 6)) + Val(grd1.TextMatrix(19, 6)) + Val(grd1.TextMatrix(20, 6)) + Val(grd1.TextMatrix(21, 6)) + Val(grd1.TextMatrix(22, 6)) + Val(grd1.TextMatrix(23, 6)) + Val(grd1.TextMatrix(24, 6)))
    Else
        grd1.TextMatrix(11, 6) = "0"
        grd1.TextMatrix(12, 6) = "0"
        grd1.TextMatrix(13, 6) = "0"
        grd1.TextMatrix(14, 6) = "0"
        grd1.TextMatrix(15, 6) = "0"
        grd1.TextMatrix(16, 6) = "0"
        grd1.TextMatrix(17, 6) = "0"
        grd1.TextMatrix(18, 6) = "0"
        grd1.TextMatrix(19, 6) = "0"
        grd1.TextMatrix(20, 6) = "0"
        grd1.TextMatrix(21, 6) = "0"
        grd1.TextMatrix(22, 6) = "0"
        grd1.TextMatrix(23, 6) = "0"
        grd1.TextMatrix(24, 6) = "0"
        grd1.TextMatrix(25, 6) = "0"
    End If
End Sub

'Added by Lydia 2021/03/10 檢查當月發文數輸入
Private Function CheckGeneralDispatch() As Boolean

    CheckGeneralDispatch = False
    
    strSql = "select * from GeneralDispatch where substr(gd01,1,6)=" & CNULL(Left(DBDATE(txt1), 6))
    intI = 1
    strExc(0) = "0"
    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
        strExc(0) = Val("" & RsTemp.Fields("gd03"))
    End If
    If strExc(0) = "0" Then
         If MsgBox(Left(txt1, Len(txt1) - 2) & "尚未輸入發文數，是否繼續作業？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
             Exit Function
         End If
    End If
    
    CheckGeneralDispatch = True
End Function
