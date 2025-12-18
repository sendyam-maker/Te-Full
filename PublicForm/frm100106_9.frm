VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm100106_9 
   BorderStyle     =   1  '單線固定
   Caption         =   "信件未處理查詢"
   ClientHeight    =   6260
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   9390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6260
   ScaleWidth      =   9390
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   60
      TabIndex        =   5
      Top             =   450
      Width           =   9255
      _ExtentX        =   16334
      _ExtentY        =   10178
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "未處理"
      TabPicture(0)   =   "frm100106_9.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "grdDataList(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "待歸檔"
      TabPicture(1)   =   "frm100106_9.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grdDataList(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   5205
         Index           =   0
         Left            =   -74940
         TabIndex        =   8
         Top             =   540
         Width           =   9090
         _ExtentX        =   16051
         _ExtentY        =   9172
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   4
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   5205
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   540
         Width           =   9090
         _ExtentX        =   16051
         _ExtentY        =   9172
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         HighLight       =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   4
      End
      Begin VB.Label Label2 
         Caption         =   "共   筆"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   330
         Width           =   2025
      End
      Begin VB.Label Label1 
         Caption         =   "共   筆"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74880
         TabIndex        =   6
         Top             =   330
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Height          =   315
      Left            =   6360
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdMail 
      Caption         =   "E-Mail催處理"
      Height          =   315
      Left            =   3570
      TabIndex        =   3
      Top             =   60
      Width           =   1365
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "信件狀況"
      Height          =   315
      Left            =   5010
      TabIndex        =   2
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   315
      Left            =   7350
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "繼續(&C)"
      Height          =   315
      Left            =   8340
      TabIndex        =   0
      Top             =   60
      Width           =   915
   End
   Begin VB.Label LblNote 
      AutoSize        =   -1  'True
      Caption         =   "信件處理管制期限為系統日起算2個工作天"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   60
      TabIndex        =   10
      Top             =   120
      Width           =   3330
   End
End
Attribute VB_Name = "frm100106_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2021/05/24 Form2.0已修改: grdDataList(index)改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/14 日期欄已修改
Option Explicit

'Add By Sindy 2019/5/30
'm_WorkType=0.專利處
'           1.商標處
'           2.外專
'           3.外商 Add By Sindy 2023/5/8
Public m_WorkType As Integer
Dim arrGridHeadText, arrGridHeadWidth
Dim PLeft(1 To 10) As Integer, iPrint As Integer, iPage As Integer, strTemp(1 To 10) As String
Dim m_iTitleFontSize As Single, m_iFontSize As Single
Dim m_iStartX As Integer, m_iStartY As Integer
Dim m_iPageHeight As Integer, m_iLineHeight As Integer
Dim dblPrevRow As Double
Dim m_PrevForm As Form 'Add By Sindy 2019/5/30
Dim m_AttachPath As String
Dim nCol As Long, nRow As Long
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件


'Add By Sindy 2019/5/30
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub SetDataListWidth(Index As Integer, Optional ByVal p_bol1st As Boolean = False)
Dim iCol As Integer
   
   If p_bol1st Then Call SetGridHead(Index)

   With grdDataList(Index)
      .Visible = False
      .row = 0
      For iCol = 0 To .Cols - 1
         .col = iCol
         .Text = arrGridHeadText(iCol)
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .CellAlignment = flexAlignCenterCenter
      Next
      .Refresh
      .Visible = True
   End With
End Sub

'Add By Sindy 2017/12/25
Private Sub cmdDetail_Click()
   Call PubShowNextData
End Sub

Public Function PubShowNextData() As Boolean
Dim i As Integer
Dim nFrm As Form
Dim Index As Integer
   
   Index = SSTab1.Tab
   PubShowNextData = False
   'If dblPrevRow > 0 Then
   For i = 1 To grdDataList(Index).Rows - 1
      grdDataList(Index).row = i
      grdDataList(Index).col = 2
      If grdDataList(Index).CellBackColor = &HFFC0C0 And grdDataList(Index).TextMatrix(i, 9) <> "" Then
         PubShowNextData = True
         '檢查表單是否已開啟，若是，則關閉
         For Each nFrm In Forms
            If StrComp(nFrm.Name, "frm06010613_1", vbTextCompare) = 0 Then
               Unload frm06010613_1
            End If
         Next
         
         '明細資料
         frm06010613_1.m_II01 = grdDataList(Index).TextMatrix(i, 7)
         frm06010613_1.m_II02 = grdDataList(Index).TextMatrix(i, 8)
         frm06010613_1.m_II03 = grdDataList(Index).TextMatrix(i, 9)
         'frm06010613_1.m_II19 = grdDataList(index).TextMatrix(i, 11)
         Call CancelRowColor(Index, i)
         frm06010613_1.cmdNext.Enabled = True 'False
'         For j = i To grdDataList(index).Rows - 1
'            If grdDataList(index).TextMatrix(j, 0) = "V" And grdDataList(index).TextMatrix(j, 9) <> "" Then
'               frm06010613_1.cmdNext.Enabled = True
'               Exit For
'            End If
'         Next j
         Call frm06010613_1.SetParent(Me)
         frm06010613_1.Show
         frm06010613_1.QueryData
         'Me.Hide
         Exit Function
      End If
   Next i
   'End If
End Function

'Add By Sindy 2019/5/30
Private Sub CancelRowColor(Index As Integer, intRow As Integer)
Dim j As Integer
   
   '清除反白
   'grdDataList(index).TextMatrix(intRow, 0) = ""
   grdDataList(Index).col = 0
   grdDataList(Index).row = intRow
   For j = 0 To grdDataList(Index).Cols - 1
      grdDataList(Index).col = j
      grdDataList(Index).CellBackColor = QBColor(15)
   Next j
   'Call SetColor(CDbl(intRow))
End Sub

'Add By Sindy 2019/5/30
Private Sub cmdMail_Click()
Dim i As Integer, jj As Integer
Dim bolHavdSel As Boolean
Dim strEmp As String
Dim ff1 As Integer
Dim strFileName As String
Dim Index As Integer
   
   Index = SSTab1.Tab
   
   bolHavdSel = False
   '檢查資料
   If grdDataList(Index).Rows - 1 < 1 Then Exit Sub
   If grdDataList(Index).Rows - 1 >= 1 And grdDataList(Index).TextMatrix(1, 9) = "" Then Exit Sub
   For i = 1 To grdDataList(Index).Rows - 1
      grdDataList(Index).row = i
      grdDataList(Index).col = 2
      If grdDataList(Index).CellBackColor = &HFFC0C0 Then
         bolHavdSel = True
         Exit For
      End If
   Next i
   If bolHavdSel = False Then
      'MsgBox "請至少點選一筆資料！", vbExclamation
      'Exit Sub
      If MsgBox("無取選資料列，是要全部發Mail催處理通知信嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Exit Sub
      Else
         '資料列全部反白
         For i = 1 To grdDataList(Index).Rows - 1
            grdDataList(Index).row = i
            For jj = 0 To grdDataList(Index).Cols - 1
               grdDataList(Index).col = jj
               grdDataList(Index).CellBackColor = &HFFC0C0
            Next jj
         Next i
      End If
   Else
      If MsgBox("確定要發E-Mail催處理信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         Exit Sub
      End If
   End If
   
Continue_Mail:
   strEmp = ""
   For i = 1 To grdDataList(Index).Rows - 1
      grdDataList(Index).row = i
      grdDataList(Index).col = 2
      If grdDataList(Index).CellBackColor = &HFFC0C0 And _
         grdDataList(Index).TextMatrix(i, 9) <> "" And _
         grdDataList(Index).TextMatrix(i, 10) <> "" Then
         If strEmp = "" Or strEmp = grdDataList(Index).TextMatrix(i, 10) Then
            If strEmp = "" Then
               strFileName = m_AttachPath & "\" & "信件未處理清單(" & Val(Me.Tag) - 19110000 & ")_" & GetPrjSalesNM(grdDataList(Index).TextMatrix(i, 10)) & ".txt"
               If Dir(strFileName) <> "" Then Kill strFileName
               If ff1 > 0 Then Close #ff1
               ff1 = FreeFile
               Open strFileName For Output As ff1
               Print #ff1, "備註：改字型Fixedsys標準11號字以橫式上下左右各10MM列印"
               If m_WorkType = 1 Then '商標處
                  Print #ff1, "轉寄日期       處理人員        主旨"
                  Print #ff1, "============== =============== ===================================================================================================="
               Else
                  Print #ff1, "收信/轉寄日期  本所案號        申請國家   申請人/處理人員 申請案號     案件名稱/主旨"
                  Print #ff1, "============== =============== ========== =============== ============ ===================================================================================================="
               End If
            End If
            For jj = 1 To 6
               strTemp(jj) = ""
            Next jj
            strTemp(1) = Trim(grdDataList(Index).TextMatrix(i, 0)) '收信(轉寄)日期
            strTemp(2) = Trim(grdDataList(Index).TextMatrix(i, 1)) '本所案號
            strTemp(3) = Trim(grdDataList(Index).TextMatrix(i, 2)) '申請國家
            strTemp(4) = Trim(grdDataList(Index).TextMatrix(i, 4)) '申請人/處理人
            strTemp(5) = Trim(grdDataList(Index).TextMatrix(i, 5)) '申請案號
            strTemp(6) = Trim(grdDataList(Index).TextMatrix(i, 3)) '案件名稱/主旨
            
            strTemp(1) = convForm(CheckStr(strTemp(1)), 14)
            strTemp(2) = convForm(CheckStr(strTemp(2)), 15)
            strTemp(3) = convForm(CheckStr(strTemp(3)), 10)
            strTemp(4) = convForm(CheckStr(strTemp(4)), 15)
            strTemp(5) = convForm(CheckStr(strTemp(5)), 12)
            strTemp(6) = convForm(CheckStr(strTemp(6)), 100)
            
            If m_WorkType = 1 Then '商標處
               Print #ff1, strTemp(1) & " " & strTemp(4) & " " & strTemp(6)
            Else
               Print #ff1, strTemp(1) & " " & strTemp(2) & " " & strTemp(3) & " " & strTemp(4) & " " & strTemp(5) & " " & strTemp(6)
            End If
            
            strEmp = grdDataList(Index).TextMatrix(i, 10)
            Call CancelRowColor(Index, i)
         End If
      End If
   Next i
   If strEmp <> "" Then
      Close ff1
      'Modify By Sindy 2024/3/6 David要求調整主旨
      PUB_SendMail strUserNum, strEmp, "", "信件已逾管制期限，尚未沖銷，請儘速至系統處理上個月未沖銷郵件！", _
         "信件已逾管制期限，尚未沖銷，資料如附件" & vbCrLf & "請儘速至系統處理上個月未沖銷郵件！" & vbCrLf & vbCrLf & vbCrLf & "          請橫印！", , strFileName, , , , , , , , , False
      GoTo Continue_Mail
   End If
   Exit Sub
End Sub

Private Sub cmdok_Click()
   'Modify By Sindy 2019/5/30
   'Modify By Sindy 2022/7/12
   'Modify By Sindy 2023/5/8 + Or m_WorkType = 3.外商
   If m_WorkType = 1 Or m_WorkType = 2 Or m_WorkType = 3 Then '1.商標處 2.外專
      Unload Me
   Else
   '2019/5/30 END
      Me.Hide
   End If
End Sub

Private Sub cmdPrint_Click()
Dim iRow As Integer
Dim Index As Integer
   
   Index = SSTab1.Tab
   
   GetPleft
   With grdDataList(Index)
      If .TextMatrix(1, 1) <> "" Then
         iPage = 1
         PrintPageHeader
         PrintPageHeader1
         For iRow = 1 To .Rows - 1
            strTemp(1) = .TextMatrix(iRow, 0)
            'Modify by Morgan 2009/2/9 要考慮追加或聯合案
            'strTemp(2) = Left(.TextMatrix(iRow, 1), 10)
            If Right(.TextMatrix(iRow, 1), 5) = "-0-00" Then
               strTemp(2) = Left(.TextMatrix(iRow, 1), 10)
            Else
               strTemp(2) = .TextMatrix(iRow, 1)
            End If
            strTemp(3) = .TextMatrix(iRow, 2)
            strTemp(4) = convForm(CheckStr(.TextMatrix(iRow, 3)), 88) 'Modify By Sindy 2021/10/27
            strTemp(5) = .TextMatrix(iRow, 4)
            strTemp(6) = convForm(CheckStr(.TextMatrix(iRow, 5)), 20)
            PrintDetail
         Next
         Call PrintReportFooter(iRow - 1)
      End If
   End With
End Sub

Sub GetPleft()
   m_iTitleFontSize = 22
   m_iFontSize = 12
   m_iStartX = 500
   m_iStartY = 500
   m_iPageHeight = 10300
   m_iLineHeight = 300

   Erase PLeft
   
   '收信日期
   PLeft(1) = 500
   '本所案號
   PLeft(2) = 1800
'   '申請國家
'   PLeft(3) = 3200
   '案件名稱
   PLeft(4) = 3300
   '申請人
   PLeft(5) = 14000
'   '申請案號
'   PLeft(6) = 15000
End Sub

Private Sub PrintNewLine(Optional ByVal p_bolHeader1 As Boolean = True, Optional ByVal p_iExtraLines As Integer = 1)
   iPrint = iPrint + m_iLineHeight
   If iPrint >= (m_iPageHeight - m_iLineHeight - p_iExtraLines * m_iLineHeight) Then
      Printer.CurrentX = m_iStartX
      Printer.CurrentY = iPrint
      Printer.Print String(135, "-")
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If p_bolHeader1 Then
         PrintPageHeader1
      End If
      iPrint = iPrint + m_iLineHeight
    End If
End Sub

Sub PrintDetail()
Dim iCol As Integer

    PrintNewLine
    For iCol = LBound(strTemp) To UBound(strTemp)
        Printer.CurrentX = PLeft(iCol)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(iCol)
    Next
End Sub

Sub PrintPageHeader()
    iPrint = m_iStartY
    Printer.Orientation = 2
    Printer.FontName = "細明體"
    Printer.Font.Size = m_iTitleFontSize
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    Printer.CurrentX = 5800
    Printer.CurrentY = iPrint
    Printer.Print Me.Caption & "清單"
    iPrint = iPrint + 500
    Printer.Font.Size = m_iFontSize
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    PrintNewLine
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print "列印人：" & strUserName
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
    PrintNewLine
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "頁    次：" & str(iPage)
    PrintNewLine
End Sub

Sub PrintPageHeader1()
    Call PrintNewLine(False, 1)
    
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print "收信日期"
    
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    Printer.Print "本所案號"
    
'    Printer.CurrentX = PLeft(3)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請國家"
    
    Printer.CurrentX = PLeft(4)
    Printer.CurrentY = iPrint
    Printer.Print "案件名稱/主旨"
    
    Printer.CurrentX = PLeft(5)
    Printer.CurrentY = iPrint
    Printer.Print "申請人/處理人員"
    
'    Printer.CurrentX = PLeft(6)
'    Printer.CurrentY = iPrint
'    Printer.Print "申請案號"
    
    PrintNewLine
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print String(135, "-")
End Sub

'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)
    Call PrintNewLine(True, 2)
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print String(135, "-")
    PrintNewLine
    Printer.CurrentX = m_iStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

'Add By Sindy 2019/5/31
Private Sub cmdQuery_Click()
   Call Process
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   
   Call SetDataListWidth(0, True)
   Call SetDataListWidth(1, True)
   
   'Add By Sindy 2019/5/30
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   cmdQuery.Visible = False
   'Modify By Sindy 2022/7/12
   'Modify By Sindy 2023/5/8 + Or m_WorkType = 3.外商
   If m_WorkType = 1 Or m_WorkType = 2 Or m_WorkType = 3 Then '1.商標處 2.外專
      cmdok.Caption = "結束(&C)"
      cmdQuery.Visible = True
      Call cmdQuery_Click
   End If
   
   SSTab1.Tab = 0 'Add By Sindy 2019/6/21
   
   '從 Unload 移來(因畫面沒離開時沒寫Log會造成逾時重新登入後重複執行)
   PUB_AddExcuteLog Me.Name
End Sub

Private Sub SetGridHead(Index As Integer)
   
   'Modify By Sindy 2017/12/25 + , "信件編號", "pi01", "pi02", "pi03"
   If m_WorkType = 1 Then '商標處
      '                        0           1           2           3       4           5           6           7       8       9       10              11      12      13      14      15      16
      arrGridHeadText = Array("轉寄日期", "本所案號", "申請國家", "主旨", "處理人員", "申請案號", "信件編號", "pi01", "pi02", "pi03", "處理人員代碼", "PA26", "PA27", "PA28", "PA29", "PA30", "PA75")
      'Modify By Sindy 2017/12/25 + 0
      arrGridHeadWidth = Array(1000, 1500, 0, 5200, 1000, 0, 1500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   Else
      'Modified by Lydia 2019/11/01 +PA26~PA30, PA75
      'arrGridHeadText = Array("收信(轉寄)日期", "本所案號", "申請國家", "案件名稱/主旨", "申請人/處理人員", "申請案號", "信件編號", "pi01", "pi02", "pi03", "處理人員代碼")
      'Modify By Sindy 2017/12/25 + 0
      'arrGridHeadWidth = Array(1000, 1500, 1000, 3000, 2000, 1500, 0, 0, 0, 0, 0)
      '                        0                 1           2           3                4                  5           6           7       8       9       10              11      12      13      14      15      16
      arrGridHeadText = Array("收信(轉寄)日期", "本所案號", "申請國家", "案件名稱/主旨", "申請人/處理人員", "申請案號", "信件編號", "pi01", "pi02", "pi03", "處理人員代碼", "PA26", "PA27", "PA28", "PA29", "PA30", "PA75")
      arrGridHeadWidth = Array(1000, 1500, 0, 5200, 1000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   End If
   grdDataList(Index).Cols = UBound(arrGridHeadText)
End Sub

Public Function Process(Optional ByVal p_Sys As String = "") As Boolean
Dim StrFa As String, stCon As String, stFDate As String, stTDate As String
Dim strSql As String
Dim rsA As New ADODB.Recordset
Dim rsB As New ADODB.Recordset
Dim dblRow As Double 'Add By Sindy 2025/9/3
   
   'Added by Lydia 2019/11/01
   m_AllSys = IIf(p_Sys <> "", p_Sys, GetAllSysKind(, "ALL"))
   intCufaCnt = 0
   
   Process = False
   m_blnColOrderAsc = True
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/3 清除查詢印表記錄檔欄位
   
   'Modify By Sindy 2019/5/30
   If m_WorkType = 1 Then '1.商標處
      '信件處理管制期限為2個工作天(ex:系統日2/25,抓的日期是2/23)
      stFDate = TransDate(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)), -2), 2)
      Me.Tag = stFDate
      pub_QL05 = pub_QL05 & ";信件日期<=" & stFDate
      LblNote.Caption = "信件日期<=" & ChangeTStringToTDateString(stFDate - 19110000) & " (系統日起算2個工作天)"
      
      '商標處收件夾資料查詢
      'Modify by Sindy 2021/2/25 and ti08<=" & stFDate => and ti12<=" & stFDate
      strSql = "select substr(' '||sqldatet(ti08),-9),ti18||'-'||ti19||'-'||ti20||'-'||ti21,'',ti17,s1.st02,'',ti01||'-'||ti03,ti01,ti02,ti03,s1.st01,TM23 PA26,TM78 PA27,TM79 PA28,TM80 PA29,TM81 PA30,TM44 PA75 " & _
               "From inputrecord, TMinput, staff s1, staff s2, trademark " & _
               "where length(ir03)=5 and substr(ir03,1,1)='T' and (ir16 not in('6') or ir16 is null) " & _
               "and ir08=0 and ir01=ti01(+) and ir03=ti03(+) " & _
               "and ir04=s1.st01(+) and ir22=s2.st01(+) " & _
               "and ti12<=" & stFDate & _
               " and ti18=tm01(+) and ti19=tm02(+) and ti20=tm03(+) and ti21=tm04(+)" & _
               " ORDER BY ti01||'-'||ti03 desc"
               
   ElseIf m_WorkType = 0 Then '專利處 Memo by Lydia 2019/11/14 共同查詢->以期限管制日查詢
      SSTab1.TabVisible(1) = False
      '2012/4/24 ADD BY SONIA 加入系統類別
      If p_Sys <> "" Then
         stCon = " AND FM03||'' IN (" & p_Sys & ") "
         pub_QL05 = pub_QL05 & ";" & Left(m_PrevForm.Label1(4), 5) & m_PrevForm.txt5(0)
      End If
      '2012/4/24 END
      '2008/8/27 modify by sonia 不限制C類來函故取消CP09>'C'
      'strSQL = "update Fagentmail set FM13='Y' where (FM01,FM02) in (select FM01,FM02 from FagentMail ,CaseProgress " & _
      '               "where  FM13 is null and FM01=CP119(+) and FM03=CP01(+) and FM04=CP02(+) and FM05=CP03(+) and FM06=CP04(+) and CP09 > 'C')"
      strSql = "update Fagentmail set FM13='Y' where (FM01,FM02) in (select FM01,FM02 from FagentMail ,CaseProgress " & _
                     "where FM13 is null" & stCon & "and FM01=CP119(+) and FM03=CP01(+) and FM04=CP02(+) and FM05=CP03(+) and FM06=CP04(+) and CP09 IS NOT NULL)"
      '2008/8/27 END
      If rsA.State <> adStateClosed Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      'Add by Toni 2008/8/18 代理人信件未處理查詢,收信日期<=系統日-3個工作天
      '統計起始日期抓系統前推3個工作天
      'stFDate = TransDate(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)), -3), 2)
      'Modify By Sindy 2018/1/11 雅娟:有關CFP程序人員的信件處理管制期限原為3個工作天,麻煩請改為5個工作天
      LblNote.Caption = "信件處理管制期限為系統日起算5個工作天"
      stFDate = TransDate(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)), -5), 2)
      Me.Tag = stFDate
      pub_QL05 = pub_QL05 & ";轉寄日期<=" & stFDate  'Add By Sindy 2010/11/3
      '2009/3/10 modify by sonia 加申請國家欄
      '2010/9/14 MODIFY BY SONIA 日期欄改百年日期排序問題
      'Added by Lydia 2015/04/28 剔除CREATE ID的部門為'Fxx'者的資料
   '   strSql = "select SUBSTR(' '||sqldatet(FM01),-9) AS FM01,FM03||'-'||FM04||'-'||FM05||'-'||FM06,substr(na03,1,4),substr(PA05,1,26),substr(CU04,1,10),PA11 from FagentMail ,CaseProgress,Patent ,Customer, nation " & _
               "where FM01<='" & stFDate & "' AND  FM13 is null" & stCon & " and FM01=CP119(+) and FM03=CP01(+) and FM04=CP02(+) and FM05=CP03(+) and FM06=CP04(+) and  " & _
               " CP09 is null and FM03=PA01(+) and FM04=PA02(+) and FM05=PA03(+) and FM06=PA04(+) and CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) and pa09=na01(+) "
      'Modified by Lydia 2019/11/01 利益衝突案件：增加申請人1~5,FC代理人
      'strSql = "select SUBSTR(' '||sqldatet(FM01),-9) AS FM01,FM03||'-'||FM04||'-'||FM05||'-'||FM06,'*'||substr(na03,1,4),substr(PA05,1,26),substr(CU04,1,10),PA11,'',0,0,'','' " & _
               "from FagentMail ,CaseProgress,Patent ,Customer, nation,staff " & _
               "where FM01<='" & stFDate & "' AND  FM13 is null" & stCon & " and FM01=CP119(+) and FM03=CP01(+) and FM04=CP02(+) and FM05=CP03(+) and FM06=CP04(+) and " & _
               "CP09 is null and FM03=PA01(+) and FM04=PA02(+) and FM05=PA03(+) and FM06=PA04(+) and CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) and pa09=na01(+) " & _
               "and FM07=ST01(+) AND substr(ST03,1,1)<>'F'"
      strSql = "select SUBSTR(' '||sqldatet(FM01),-9) AS FM01,FM03||'-'||FM04||'-'||FM05||'-'||FM06,'*'||substr(na03,1,4),substr(PA05,1,26),SUBSTRB(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),1,10) ,PA11,'',0,0,'','' " & _
               ", PA26 , PA27 , PA28 , PA29 , PA30 , PA75 " & _
               "from FagentMail, CaseProgress, Patent, Customer, nation, staff " & _
               "where FM01<='" & stFDate & "' AND  FM13 is null" & stCon & " and FM01=CP119(+) and FM03=CP01(+) and FM04=CP02(+) and FM05=CP03(+) and FM06=CP04(+) and " & _
               "CP09 is null and FM03=PA01(+) and FM04=PA02(+) and FM05=PA03(+) and FM06=PA04(+) and CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) and pa09=na01(+) " & _
               "and FM07=ST01(+) AND substr(ST03,1,1)<>'F'"
      'Add By Sindy 2016/10/17 + 專利處收件夾資料查詢
      'Modify By Sindy 2018/6/29 s1.st02 ==> decode(ir16,2,s2.st02,4,s2.st02,s1.st02)顯示目前處理人員
   '   strSql = strSql & " union all " & _
   '            "select SUBSTR(' '||sqldatet(pi12),-9),pi18||'-'||pi19||'-'||pi20||'-'||pi21,substr(na03,1,4),substr(PA05,1,26),substr(CU04,1,10),PA11,Pi01||'-'||pi03,pi01,pi02,pi03 " & _
   '            "from Patent,Customer,nation," & _
   '            "(select pi01,pi02,pi03,pi12,pi18,pi19,pi20,pi21 From inputrecord,patentinput " & _
   '            "where length(ir03)=5 and substr(ir03,1,1)='P' " & _
   '            "and ir08=0 " & _
   '            "and ir01=pi01(+) and ir03=pi03(+) " & _
   '            "and pi12<='" & stFDate & "' AND pi18 in (" & p_Sys & ") " & _
   '            "group by pi01,pi02,pi03,pi12,pi18,pi19,pi20,pi21) " & _
   '            "where pi18=PA01(+) and pi19=PA02(+) and pi20=PA03(+) and pi21=PA04(+) " & _
   '            "and CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) " & _
   '            "and pa09=na01(+)"
      '2016/10/17 END
      'Modified by Lydia 2019/11/01 利益衝突案件：增加申請人1~5,FC代理人
      'strSql = strSql & " union all " & _
               "select substr(' '||sqldatet(pi08),-9),pi18||'-'||pi19||'-'||pi20||'-'||pi21,'',pi17,decode(ir16,2,s2.st02,4,s2.st02,s1.st02),'',pi01||'-'||pi03,pi01,pi02,pi03,decode(ir16,2,s2.st01,4,s2.st01,s1.st01) " & _
               "From inputrecord, patentinput, staff s1, staff s2 " & _
               "where length(ir03)=5 and substr(ir03,1,1)='P' and ir08=0 and ir01=pi01(+) and ir03=pi03(+) " & _
               "and ir04=s1.st01(+) and ir22=s2.st01(+) " & _
               "and pi08<=" & stFDate
      strSql = strSql & " union all " & _
               "select substr(' '||sqldatet(pi08),-9),pi18||'-'||pi19||'-'||pi20||'-'||pi21,'',pi17,decode(ir16,2,s2.st02,4,s2.st02,s1.st02),'',pi01||'-'||pi03,pi01,pi02,pi03,decode(ir16,2,s2.st01,4,s2.st01,s1.st01) " & _
               ", PA26 , PA27 , PA28 , PA29 , PA30 , PA75 " & _
               "From inputrecord, patentinput, staff s1, staff s2, patent " & _
               "where length(ir03)=5 and substr(ir03,1,1)='P' and ir08=0 and ir01=pi01(+) and ir03=pi03(+) " & _
               "and ir04=s1.st01(+) and ir22=s2.st01(+) " & _
               "and pi08<=" & stFDate & _
               " and pi18=pa01(+) and pi19=pa02(+) and pi20=pa03(+) and pi21=pa04(+)"
      strSql = strSql & " ORDER BY 1,2 " '2010/9/14 ADD BY SONIA
      
   'Add By Sindy 2022/7/12
   Else '外專, 外商
      SSTab1.TabVisible(1) = False
      '加入系統類別
      If p_Sys <> "" Then
         stCon = " AND FM03||'' IN (" & p_Sys & ") "
         pub_QL05 = pub_QL05 & ";" & Left(m_PrevForm.Label1(4), 5) & m_PrevForm.txt5(0)
      End If
      
      '信件處理管制期限為2個工作天(ex:系統日2/25,抓的日期是2/23)
      stFDate = TransDate(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)), -2), 2)
      Me.Tag = stFDate
      pub_QL05 = pub_QL05 & ";信件日期<=" & stFDate
      LblNote.Caption = "信件日期<=" & ChangeTStringToTDateString(stFDate - 19110000) & " (系統日起算2個工作天)"
      
      'Add By Sindy 2023/5/8 + 外商
      If m_WorkType = 3 Then '外商
         strSql = "select substr(' '||sqldatet(ii08),-9),ii23||'-'||ii24||'-'||ii25||'-'||ii26,'',ii17,decode(ir16,null,s1.st02,4,s1.st02,9,s1.st02,s2.st02),'',ii01||'-'||ii03,ii01,ii02,ii03,decode(ir16,null,s1.st01,4,s1.st01,9,s1.st01,s2.st01) " & _
                  ",TM23 PA26,TM78 PA27,TM79 PA28,TM80 PA29,TM81 PA30,TM44 PA75 " & _
                  "From inputrecord, ipdeptinput, staff s1, staff s2, trademark " & _
                  "where length(ir03)=5 and substr(ir03,1,1)='F' " & _
                  "and ir08=0 and ir01=ii01(+) and ir03=ii03(+) " & _
                  "and ir04=s1.st01(+) and ir22=s2.st01(+) and substr(s1.st03,1,2)='" & Left(Pub_StrUserSt03, 2) & "' " & _
                  "and ii12<=" & stFDate & _
                  " and ii23=TM01(+) and ii24=TM02(+) and ii25=TM03(+) and ii26=TM04(+)"
         strSql = strSql & " ORDER BY 1,2 "
      Else
      '2023/5/8 END
         strSql = "select SUBSTR(' '||sqldatet(FM01),-9) AS FM01,FM03||'-'||FM04||'-'||FM05||'-'||FM06,'*'||substr(na03,1,4),substr(PA05,1,26),SUBSTRB(NVL(CU04,DECODE(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)),1,10) ,PA11,'',0,0,'','' " & _
                  ", PA26 , PA27 , PA28 , PA29 , PA30 , PA75 " & _
                  "from FagentMail, CaseProgress, Patent, Customer, nation, staff " & _
                  "where FM01<='" & stFDate & "' AND  FM13 is null" & stCon & " and FM01=CP119(+) and FM03=CP01(+) and FM04=CP02(+) and FM05=CP03(+) and FM06=CP04(+) and " & _
                  "CP09 is null and FM03=PA01(+) and FM04=PA02(+) and FM05=PA03(+) and FM06=PA04(+) and CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) and pa09=na01(+) " & _
                  "and FM07=ST01(+) AND substr(ST03,1,1)='F'"
         strSql = strSql & " union all " & _
                  "select substr(' '||sqldatet(ii08),-9),ii23||'-'||ii24||'-'||ii25||'-'||ii26,'',ii17,decode(ir16,null,s1.st02,4,s1.st02,9,s1.st02,s2.st02),'',ii01||'-'||ii03,ii01,ii02,ii03,decode(ir16,null,s1.st01,4,s1.st01,9,s1.st01,s2.st01) " & _
                  ", PA26 , PA27 , PA28 , PA29 , PA30 , PA75 " & _
                  "From inputrecord, ipdeptinput, staff s1, staff s2, patent " & _
                  "where length(ir03)=5 and substr(ir03,1,1)='F' " & _
                  "and ir08=0 and ir01=ii01(+) and ir03=ii03(+) " & _
                  "and ir04=s1.st01(+) and ir22=s2.st01(+) and s1.st03='" & Pub_StrUserSt03 & "' " & _
                  "and ii12<=" & stFDate & _
                  " and ii23=pa01(+) and ii24=pa02(+) and ii25=pa03(+) and ii26=pa04(+)"
         strSql = strSql & " ORDER BY 1,2 "
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   rsA.CursorLocation = adUseClient
   'Modified by Lydia 2019/11/01 改變型態
   'rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   rsA.Open strSql, cnnConnection, adOpenDynamic, adLockBatchOptimistic
   
   If rsA.RecordCount > 0 Then
      dblRow = rsA.RecordCount 'Add By Sindy 2025/9/3
      Process = True
      'Modify By Sindy 2022/7/12 任何單位都適用XY特殊權限的限閱機制, 所以做此段程式的調整
      'Added by Lydia 2019/11/01 逐案號判斷
      If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
         rsA.MoveFirst
         Do While rsA.EOF = False
             If "" & rsA.Fields("pa26") & rsA.Fields("pa27") & rsA.Fields("pa28") & rsA.Fields("pa29") & rsA.Fields("pa30") & rsA.Fields("pa75") <> "" Then  'Added by Lydia 2019/12/26 判斷有客戶編號或FC代理人編號
                 '利益衝突案件：逐案號判斷
                 If PUB_ChkCufaByCase(Me.Name, m_AllSys, "" & rsA.Fields(1), "" & rsA.Fields("pa26") & "," & rsA.Fields("pa27") & "," & rsA.Fields("pa28") & "," & rsA.Fields("pa29") & "," & rsA.Fields("pa30"), "" & rsA.Fields("pa75")) = False Then
                     intCufaCnt = intCufaCnt + 1
                     rsA.Delete
                 End If
             End If
             rsA.MoveNext
         Loop
         '利益衝突案件：限閱案件
         If intCufaCnt > 0 Then
            pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
            MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
         End If
         InsertQueryLog (dblRow)
         If rsA.RecordCount = 0 Then
            Label1.Caption = "共 " & rsA.RecordCount & " 筆"
            GoTo JumpToNoData
         'Added by Lydia 2019/12/26 重新讀取
         Else
            Set grdDataList(0).Recordset = rsA
            Call SetDataListWidth(0)
            Label1.Caption = "共 " & rsA.RecordCount & " 筆"
         'end 2019/12/26
         End If
      'end 2019/11/01
      Else
         InsertQueryLog (rsA.RecordCount)
         Set grdDataList(0).Recordset = rsA
         Call SetDataListWidth(0)
         Label1.Caption = "共 " & rsA.RecordCount & " 筆"
      End If
      
      If m_WorkType = 1 Then
         '商標處收件夾資料查詢
         strSql = "select count(*) " & _
                  "From inputrecord, TMinput " & _
                  "where length(ir03)=5 and substr(ir03,1,1)='T' and (ir16 not in('6') or ir16 is null) and ir08=0 and ir01=ti01(+) and ir03=ti03(+) " & _
                  "and ti12<=" & stFDate & _
                  " Group BY ti01||'-'||ti03"
         If rsB.State <> adStateClosed Then rsB.Close
         rsB.CursorLocation = adUseClient
         rsB.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         Label1.Caption = "共 " & rsB.RecordCount & " 封信件"
      End If
      '2022/7/12 END
'   Else
'      InsertQueryLog (0) 'Add By Sindy 2010/11/3
'      If m_WorkType = 1 Then
'         ShowNoData
'      Else
'         Me.Hide
'      End If
   End If
   
   'Add By Sindy 2019/6/21
   '待歸檔查詢
   If SSTab1.TabVisible(1) = True Then
      '商標處收件夾資料查詢
      strSql = "select substr(' '||sqldatet(ti08),-9),ti18||'-'||ti19||'-'||ti20||'-'||ti21,'',ti17,s1.st02,'',ti01||'-'||ti03,ti01,ti02,ti03,s1.st01 " & _
               "From inputrecord, TMinput, staff s1, staff s2 " & _
               "where length(ir03)=5 and substr(ir03,1,1)='T' and ir16='6' " & _
               "and ir08=0 and ir01=ti01(+) and ir03=ti03(+) " & _
               "and ir04=s1.st01(+) and ir22=s2.st01(+)" & _
               " ORDER BY ti01||'-'||ti03 desc"
      If rsA.State <> adStateClosed Then rsA.Close
      rsA.CursorLocation = adUseClient
      rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         Set grdDataList(1).Recordset = rsA
         Call SetDataListWidth(1)
         Process = True
         Label2.Caption = "共 " & rsA.RecordCount & " 筆"
'         'Add By Sindy 2019/6/4
'         If m_WorkType = 1 Then '1.商標處
            '商標處收件夾資料查詢
            strSql = "select count(*) " & _
                     "From inputrecord, TMinput " & _
                     "where length(ir03)=5 and substr(ir03,1,1)='T' and ir16='6' and ir08=0 and ir01=ti01(+) and ir03=ti03(+) " & _
                     " Group BY ti01||'-'||ti03"
            If rsB.State <> adStateClosed Then rsB.Close
            rsB.CursorLocation = adUseClient
            rsB.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            Label2.Caption = "共 " & rsB.RecordCount & " 封信件"
'         Else
'            Label2.Caption = "共 " & rsA.RecordCount & " 筆"
'         End If
      End If
   End If
   
   If Process = True Then
      Me.Show
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/11/3
JumpToNoData: 'Added by Lydia 2019/11/01
      'Modify By Sindy 2023/5/8 + Or m_WorkType = 3.外商
      If m_WorkType = 1 Or m_WorkType = 2 Or m_WorkType = 3 Then
         ShowNoData
      Else
         Me.Hide
      End If
   End If
   
   dblPrevRow = 0 'Add By Sindy 2017/12/25
   grdDataList(0).row = 0
   grdDataList(0).col = 0
   grdDataList(1).row = 0
   grdDataList(1).col = 0
   
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   If rsB.State <> adStateClosed Then rsB.Close
   Set rsB = Nothing
   
On Error GoTo ErrHnd

ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
End Function

Private Sub Form_Unload(Cancel As Integer)
Dim nFrm As Form
   
   '檢查表單是否已開啟，若是，則關閉
   For Each nFrm In Forms
      If StrComp(nFrm.Name, "frm06010613_1", vbTextCompare) = 0 Then
         Unload frm06010613_1
      End If
   Next
   
   'Add By Sindy 2019/5/30
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   Set frm100106_9 = Nothing
End Sub

Private Sub GrdDataList_Click(Index As Integer)
grdDataList(Index).Visible = False
grdDataList(Index).row = grdDataList(Index).MouseRow
grdDataList(Index).col = grdDataList(Index).MouseCol
nRow = grdDataList(Index).row
nCol = grdDataList(Index).col
If nRow = 0 Then
   If grdDataList(Index).Text <> "V" Then
      If grdDataList(Index).Text = "無" Then
         If m_blnColOrderAsc = True Then
            grdDataList(Index).Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            grdDataList(Index).Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            grdDataList(Index).Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            grdDataList(Index).Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End If
grdDataList(Index).Visible = True
End Sub

'Add By Sindy 2017/12/25
Private Sub grdDataList_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   grdDataList(Index).ToolTipText = ""
   If grdDataList(Index).MouseRow > 0 Then
      If grdDataList(Index).MouseCol = 3 Then
         If grdDataList(Index).TextMatrix(grdDataList(Index).MouseRow, 6) <> "" Then
            grdDataList(Index).ToolTipText = "信件編號:" & grdDataList(Index).TextMatrix(grdDataList(Index).MouseRow, 6)
         End If
      End If
   End If
End Sub

Private Sub grdDataList_SelChange(Index As Integer)
Dim i As Integer, intRow As Integer

grdDataList(Index).Visible = False
intRow = grdDataList(Index).MouseRow
'grdDataList(index).row = intRow
'grdDataList(index).col = 0
'If grdDataList(index).row <> 0 Then
'cmdDetail.Enabled = False
If intRow > 0 Then
'   If dblPrevRow > 0 Then
'      grdDataList(index).row = dblPrevRow
'      grdDataList(index).col = 0
'      For i = 0 To grdDataList(index).Cols - 1
'         grdDataList(index).col = i
'         grdDataList(index).CellBackColor = QBColor(15)
'      Next i
'   End If
   grdDataList(Index).row = intRow
   grdDataList(Index).col = 0
   dblPrevRow = intRow '記錄目前筆數
   If grdDataList(Index).CellBackColor = &HFFC0C0 Then
      'grdDataList(index).Text = ""
      For i = 0 To grdDataList(Index).Cols - 1
         grdDataList(Index).col = i
         grdDataList(Index).CellBackColor = QBColor(15)
      Next i
   Else
      'grdDataList(index).Text = "V"
      For i = 0 To grdDataList(Index).Cols - 1
         grdDataList(Index).col = i
         grdDataList(Index).CellBackColor = &HFFC0C0
      Next i
      If grdDataList(Index).TextMatrix(intRow, 6) <> "" Then
'         cmdDetail.Enabled = True
      End If
   End If
End If
grdDataList(Index).Visible = True
End Sub

'Add By Sindy 2019/6/21
Private Sub SSTab1_Click(PreviousTab As Integer)
   dblPrevRow = 0
   grdDataList(SSTab1.Tab).row = 0
   grdDataList(SSTab1.Tab).col = 0
End Sub
