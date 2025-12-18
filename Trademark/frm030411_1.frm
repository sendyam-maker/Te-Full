VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm030411_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "延遲承辦案件明細查詢"
   ClientHeight    =   5520
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7896
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7896
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   6750
      TabIndex        =   1
      Top             =   60
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5520
      TabIndex        =   0
      Top             =   60
      Width           =   1185
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4905
      Left            =   30
      TabIndex        =   2
      Top             =   570
      Width           =   7815
      _ExtentX        =   13801
      _ExtentY        =   8657
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
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
      _Band(0).Cols   =   1
   End
End
Attribute VB_Name = "frm030411_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/2/25 Form2.0已修改(grd1改Fonts)
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
'create by nickc 2008/01/10 陳經理有請作單
Option Explicit
Dim i As Integer, Page As Integer, iPrint As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
        frm030411.Show
        Unload Me
Case 1
        Unload frm030411
        Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm030411_1 = Nothing
End Sub

Sub Process()
Dim m_str As String
Dim m_str2 As String
Dim m_str6 As String
Dim m_rs As New ADODB.Recordset
Screen.MousePointer = vbHourglass
Me.grd1.MousePointer = flexArrowHourGlass
DoEvents
m_str2 = ""
m_str6 = ""
With frm030411
    If Trim(.txt1(0)) <> "" Then
        m_str2 = m_str2 + " AND CP01 IN (" & SQLGrpStr(.txt1(0), 2) & ") "
        m_str6 = m_str6 + " AND CP01 IN (" & SQLGrpStr(.txt1(0), 5) & ") "
        pub_QL05 = pub_QL05 & ";" & .Label1(0) & .txt1(0)  'Add By Sindy 2010/10/22
    End If
    m_str2 = m_str2 + " AND ((CP27>=" & Val(ChangeTStringToWString(.txt1(1))) & " AND CP27<=" & Val(ChangeTStringToWString(.txt1(2))) & ") or cp27 is null ) "
    m_str6 = m_str6 + " AND ((CP27>=" & Val(ChangeTStringToWString(.txt1(1))) & " AND CP27<=" & Val(ChangeTStringToWString(.txt1(2))) & ") or cp27 is null ) "
    pub_QL05 = pub_QL05 & ";" & .Label1(2) & .txt1(1) & "-" & .txt1(2) 'Add By Sindy 2010/10/22
    If Trim(.txt1(3)) <> "" Then
        m_str2 = m_str2 + " AND CP14='" & .txt1(3) & "' "
        m_str6 = m_str6 + " AND CP14='" & .txt1(3) & "' "
        pub_QL05 = pub_QL05 & ";" & .Label1(6) & .txt1(3) & .lbl1(1) 'Add By Sindy 2010/10/22
    End If
    If .txt1(4) = "Y" Then
        pub_QL05 = pub_QL05 & ";" & .Label1(8) & .txt1(4) 'Add By Sindy 2010/10/22
        'edit by nickc 2008/04/03 陳經理加欄位
        'm_str = "select st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as sorta,tm05,decode(tm10,'000',cpm03,cpm04),cp48,cp14,decode(cp27,null,to_char(sysdate, 'YYYYMMDD'),0,to_char(sysdate, 'YYYYMMDD'),to_char(cp27)) from caseprogress,trademark,staff ,casepropertymap where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp48<decode(cp27,null,to_number(to_char(sysdate,'YYYYMMDD')),0,to_number(to_char(sysdate,'YYYYMMDD')),cp27) and cp57 is null " & m_str2
        'm_str = m_str & " union select st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as sorta,nvl(sp05,nvl(sp06,sp07)),decode(sp09,'000',cpm03,cpm04),cp48,cp14,decode(cp27,null,to_char(sysdate, 'YYYYMMDD'),0,to_char(sysdate, 'YYYYMMDD'),to_char(cp27)) from caseprogress,servicepractice,staff ,casepropertymap where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp48<decode(cp27,null,to_number(to_char(sysdate,'YYYYMMDD')),0,to_number(to_char(sysdate,'YYYYMMDD')),cp27) and cp57 is null " & m_str6
        m_str = "select st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as sorta,tm05,decode(tm10,'000',cpm03,cpm04),sqldatet(cp05),cp48,nvl(decode(fa05,null,null,fa05||' '||fa63||' '||fa64||' '||fa65),nvl(fa04,fa06)),cp14,decode(cp27,null,to_char(sysdate, 'YYYYMMDD'),0,to_char(sysdate, 'YYYYMMDD'),to_char(cp27)) from caseprogress,trademark,staff ,casepropertymap,fagent where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp48<decode(cp27,null,to_number(to_char(sysdate,'YYYYMMDD')),0,to_number(to_char(sysdate,'YYYYMMDD')),cp27) and cp57 is null and substr(tm44,1,8)=fa01(+) and nvl(substr(tm44,9,1),'0')=fa02(+) " & m_str2
        m_str = m_str & " union select st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as sorta,nvl(sp05,nvl(sp06,sp07)),decode(sp09,'000',cpm03,cpm04),sqldatet(cp05),cp48,CP64,cp14,decode(cp27,null,to_char(sysdate, 'YYYYMMDD'),0,to_char(sysdate, 'YYYYMMDD'),to_char(cp27)) from caseprogress,servicepractice,staff ,casepropertymap where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp48<decode(cp27,null,to_number(to_char(sysdate,'YYYYMMDD')),0,to_number(to_char(sysdate,'YYYYMMDD')),cp27) and cp57 is null " & m_str6
        m_str = m_str & " order by cp14,sorta "
    Else
        m_str = "select st02,count(cp09),'','','',cp14,'' from ( select st02,cp09,'','','',cp14 from caseprogress,trademark,staff ,casepropertymap where cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp48<decode(cp27,null,to_number(to_char(sysdate,'YYYYMMDD')),0,to_number(to_char(sysdate,'YYYYMMDD')),cp27) and cp57 is null " & m_str2
        m_str = m_str & " union select st02,cp09,'','','',cp14 from caseprogress,servicepractice,staff ,casepropertymap where cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp14=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp48<decode(cp27,null,to_number(to_char(sysdate,'YYYYMMDD')),0,to_number(to_char(sysdate,'YYYYMMDD')),cp27) and cp57 is null " & m_str6
        m_str = m_str & ") AA  group by st02,cp14 order by cp14 "
    End If
    If m_rs.State = 1 Then m_rs.Close
    Set m_rs = New ADODB.Recordset
    m_rs.CursorLocation = adUseClient
    m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
    If m_rs.RecordCount <> 0 Then
        InsertQueryLog (m_rs.RecordCount) 'Add By Sindy 2010/10/22
        Set grd1.Recordset = m_rs
        SetGrd
        If .txt1(4) = "Y" Then
            For i = 1 To grd1.Rows - 1
                 'edit by nickc 2008/04/03 陳經理加欄位
                 'If Val(grd1.TextMatrix(i, 4)) > 0 Then
                 If Val(grd1.TextMatrix(i, 5)) > 0 Then
                    'edit by nickc 2008/04/03 陳經理加欄位
                    'grd1.TextMatrix(i, 4) = GetWorkDay(grd1.TextMatrix(i, 6), grd1.TextMatrix(i, 4))
                    grd1.TextMatrix(i, 5) = GetWorkDay(grd1.TextMatrix(i, 8), grd1.TextMatrix(i, 5))
                 Else
                    'edit by nickc 2008/04/03 陳經理加欄位
                    'grd1.TextMatrix(i, 4) = ""
                    grd1.TextMatrix(i, 5) = ""
                 End If
            Next i
        End If
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/10/22
        If .txt1(5) = "1" Then
            ShowNoData
        End If
        Exit Sub
    End If
End With
Me.grd1.MousePointer = flexDefault
Screen.MousePointer = vbDefault
End Sub

Sub SetGrd()
grd1.Visible = False
grd1.Cols = 9
grd1.row = 0
grd1.col = 0: grd1.Text = "承辦人"
grd1.ColWidth(0) = 800
grd1.col = 1: grd1.Text = IIf(frm030411.txt1(4) = "Y", "本所案號", "延遲件數")
grd1.ColWidth(1) = IIf(frm030411.txt1(4) = "Y", 1550, 800)
If frm030411.txt1(4) = "Y" Then
    grd1.CellAlignment = flexAlignCenterCenter
Else
    grd1.ColAlignment(1) = flexAlignRightCenter
End If
grd1.col = 2: grd1.Text = IIf(frm030411.txt1(4) = "Y", "案件名稱", "")
grd1.ColWidth(2) = IIf(frm030411.txt1(4) = "Y", 2600, 0)
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 3: grd1.Text = IIf(frm030411.txt1(4) = "Y", "案件性質", "")
grd1.ColWidth(3) = IIf(frm030411.txt1(4) = "Y", 1000, 0)
grd1.CellAlignment = flexAlignCenterCenter
'edit by nickc 2008/04/03 陳經理加欄位
'grd1.col = 4: grd1.Text = IIf(frm030411.TXT1(4) = "Y", "延遲天數", "")
'grd1.ColWidth(4) = IIf(frm030411.TXT1(4) = "Y", 1000, 0)
'If frm030411.TXT1(4) = "Y" Then
'    grd1.ColAlignment(4) = flexAlignRightCenter
'Else
'    grd1.CellAlignment = flexAlignCenterCenter
'End If
'grd1.col = 5: grd1.Text = ""
'grd1.ColWidth(5) = 0
'grd1.CellAlignment = flexAlignCenterCenter
'grd1.col = 6: grd1.Text = ""
'grd1.ColWidth(6) = 0
'grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 4: grd1.Text = IIf(frm030411.txt1(4) = "Y", "收文日", "")
grd1.ColWidth(4) = IIf(frm030411.txt1(4) = "Y", 800, 0)
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 5: grd1.Text = IIf(frm030411.txt1(4) = "Y", "延遲天數", "")
grd1.ColWidth(5) = IIf(frm030411.txt1(4) = "Y", 800, 0)
If frm030411.txt1(4) = "Y" Then
    grd1.ColAlignment(5) = flexAlignRightCenter
Else
    grd1.CellAlignment = flexAlignCenterCenter
End If
grd1.col = 6: grd1.Text = IIf(frm030411.txt1(4) = "Y", "代理人", "")
grd1.ColWidth(6) = IIf(frm030411.txt1(4) = "Y", 3000, 0)
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 7: grd1.Text = ""
grd1.ColWidth(7) = 0
grd1.CellAlignment = flexAlignCenterCenter
grd1.col = 8: grd1.Text = ""
grd1.ColWidth(8) = 0
grd1.CellAlignment = flexAlignCenterCenter

grd1.Visible = True
End Sub

Public Sub PrintData()
Dim MyMsg As String
Page = 1
'edit by nickc 2008/04/03 陳經理加欄位
'Printer.Orientation = 1
'MyMsg = PrintGrdData(grd1, 42)
Printer.Orientation = 2
MyMsg = PrintGrdData(grd1, 25)
If MyMsg = "" Then
    ShowPrintOk
Else
    MsgBox MyMsg, vbExclamation, "發生錯誤！"
End If
End Sub
'列印GRID 的資料 add by nickc 2008/01/11
'm_Grd   欲列印之GRID
'm_Line   列印區行數
'm_Left   定義欄寬，可不傳，將以GRID 欄寬為標準
'須自備 PrintTitle sub or function
Public Function PrintGrdData(m_Grd As MSHFlexGrid, m_Lines As Integer, ParamArray m_Left()) As String
Dim mm_Left() As Integer
Dim m_HStrTemp() As String
Dim m_i As Integer
Dim m_j As Integer
Dim m_k As Integer
Dim m_CalFeilds As Integer
Dim m_NowLine As Integer
Dim m_NowCol As Integer
PrintGrdData = ""
If m_Grd.DataSource Is Nothing Then
    PrintGrdData = "沒有可列印的資料！"
    Exit Function
ElseIf m_Grd.Recordset.RecordCount = 0 Then
    PrintGrdData = "沒有可列印的資料！"
    Exit Function
End If
'定義列印的左邊界
If UBound(m_Left) <> -1 Then
    ReDim mm_Left(UBound(m_Left)) As Integer
    For m_i = 0 To UBound(m_Left)
        mm_Left(m_i + 1) = m_Left(m_i)
        If m_i <> 0 Then
            For m_j = m_i To 0 Step -1
                mm_Left(m_i + 1) = mm_Left(m_i + 1) + mm_Left(m_j)
            Next m_j
        End If
    Next m_i
Else
    m_CalFeilds = 0
    For m_i = 0 To m_Grd.Cols - 1
        If m_Grd.ColWidth(m_i) <> 0 Then
            m_CalFeilds = m_CalFeilds + 1
            ReDim Preserve mm_Left(m_CalFeilds) As Integer
            If m_CalFeilds = 1 Then mm_Left(m_CalFeilds) = 0 Else mm_Left(m_CalFeilds) = (m_Grd.ColWidth(m_i - 1) / 3 * 4) + 200
            If m_CalFeilds <> 1 Then
                mm_Left(m_CalFeilds) = mm_Left(m_CalFeilds) + mm_Left(m_CalFeilds - 1)
            End If
        End If
    Next m_i
End If
'抓報表欄位
m_CalFeilds = 0
For m_i = 0 To m_Grd.Cols - 1
    If m_Grd.ColWidth(m_i) <> 0 Then
        m_CalFeilds = m_CalFeilds + 1
        ReDim Preserve m_HStrTemp(m_CalFeilds) As String
        m_HStrTemp(m_CalFeilds) = m_Grd.TextMatrix(0, m_i)
    End If
Next m_i
PrintTitle
For m_i = 1 To UBound(m_HStrTemp())
    Printer.CurrentX = mm_Left(m_i)
    Printer.CurrentY = iPrint
    Printer.Print m_HStrTemp(m_i)
Next m_i
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
m_NowLine = 0
For m_i = 0 + grd1.FixedRows To grd1.Rows - 1
    m_NowCol = 0
    For m_j = 0 + grd1.FixedCols To grd1.Cols - 1
        If m_Grd.ColWidth(m_j) <> 0 Then
            m_NowCol = m_NowCol + 1
            'Modified by Lydia 2016/01/06
            'Printer.CurrentX = mm_Left(m_NowCol) + IIf(m_Grd.ColAlignment(m_j) = flexAlignRightCenter, IIf(m_NowCol >= UBound(mm_Left), 1000 - Printer.TextWidth(m_Grd.TextMatrix(m_i, m_j)), mm_Left(m_j + 1) - 500 - Printer.TextWidth(m_Grd.TextMatrix(m_i, m_j))), 0)
            Printer.CurrentX = mm_Left(m_NowCol) + IIf(m_Grd.ColAlignment(m_j) = flexAlignRightCenter, IIf(m_NowCol >= UBound(mm_Left), 1000 - Printer.TextWidth(m_Grd.TextMatrix(m_i, m_j)), mm_Left(m_j + 1) - mm_Left(m_j) - 500 - Printer.TextWidth(m_Grd.TextMatrix(m_i, m_j))), 0)
            Printer.CurrentY = iPrint
            Select Case m_j
            Case 2
                Printer.Print StrToStr(m_Grd.TextMatrix(m_i, m_j), 15)
            Case 3
                Printer.Print StrToStr(m_Grd.TextMatrix(m_i, m_j), 6)
            Case 6
                Printer.Print StrToStr(m_Grd.TextMatrix(m_i, m_j), 18)
            Case Else
                 Printer.Print m_Grd.TextMatrix(m_i, m_j)
            End Select
        End If
    Next m_j
    iPrint = iPrint + 300
    m_NowLine = m_NowLine + 1
    If m_NowLine >= m_Lines Then
        Printer.NewPage
        Page = Page + 1
        PrintTitle
        For m_k = 1 To UBound(m_HStrTemp())
            Printer.CurrentX = mm_Left(m_k)
            Printer.CurrentY = iPrint
            Printer.Print m_HStrTemp(m_k)
        Next m_k
        iPrint = iPrint + 300
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
        iPrint = iPrint + 300
        m_NowLine = 0
    End If
Next m_i
Printer.EndDoc
End Function

Sub PrintTitle()
iPrint = 500
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("延遲承辦案件明細表") / 2
Printer.CurrentY = iPrint
Printer.Print "延遲承辦案件明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("發文日：" & Format(ChangeTStringToTDateString(frm030411.txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(frm030411.txt1(2))) / 2
Printer.CurrentY = iPrint
Printer.Print "發文日：" & Format(ChangeTStringToTDateString(frm030411.txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(frm030411.txt1(2))
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - 1000 - Printer.TextWidth("列印日期：" & Format(strSrvDate(2), "##/##/##"))
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = Printer.ScaleWidth - 1000 - Printer.TextWidth("列印日期：" & Format(strSrvDate(2), "##/##/##"))
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(19000, iPrint + 150)
iPrint = iPrint + 300
End Sub
