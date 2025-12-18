VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040207a 
   BorderStyle     =   1  '單線固定
   Caption         =   "延展前無效商標管制表"
   ClientHeight    =   4824
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7428
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4824
   ScaleWidth      =   7428
   Begin VB.CommandButton cmdok 
      Caption         =   "閉卷(&C)"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   3900
      TabIndex        =   3
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   5070
      TabIndex        =   1
      Top             =   120
      Width           =   1155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4275
      Left            =   30
      TabIndex        =   0
      Top             =   510
      Width           =   7365
      _ExtentX        =   12996
      _ExtentY        =   7535
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      SelectionMode   =   2
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
Attribute VB_Name = "frm040207a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/22 改成Form2.0 ; grd1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/12 日期欄已修改
Option Explicit

Dim strSql As String
Dim NowRow As Integer

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
        frm040207.Show
        Unload Me
Case 1
        Unload frm040207
        Unload Me
Case 2
        Screen.MousePointer = vbHourglass
        grd1.MousePointer = flexArrowHourGlass
        Me.Enabled = False
        CheckOC
        strSql = "select * from trademark where tm01='" & SystemNumber(grd1.TextMatrix(NowRow, 1), 1) & "' and tm02='" & SystemNumber(grd1.TextMatrix(NowRow, 1), 2) & "' and tm03='" & SystemNumber(grd1.TextMatrix(NowRow, 1), 3) & "' and tm04='" & SystemNumber(grd1.TextMatrix(NowRow, 1), 4) & "' and tm29 is null and tm30 is null and tm31 is null "
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount = 0 Then
            MsgBox "此案號已閉卷，請檢查！", vbInformation, "閉卷錯誤！"
        Else
            If MsgBox("確定要將 " & grd1.TextMatrix(NowRow, 1) & " 閉卷？", vbExclamation + vbYesNo, "警告！") = vbYes Then
                SaveClose
            End If
        End If
        Me.Enabled = True
        grd1.MousePointer = flexDefault
        Screen.MousePointer = vbDefault
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
NowRow = 0
CheckOC
strSql = "select decode(id,'" & strUserNum & "1','爭議無效','第二期未委託') as 種類,r116001||'-'||r116002||'-'||r116003||'-'||r116004 as 本所案號,r116005 as 註冊號,r116006 as 申請案號,r116007 as 商標名稱,r116008||' '||NVL(CU04,NVL(CU05||CU88||CU89||CU90,CU06)) as 申請人 from r040207,customer where (id='" & strUserNum & "1' or id='" & strUserNum & "2') and substr(r116008,1,8)=cu01(+) and decode(SUBSTR(r116008,9,1),'','0',SUBSTR(r116008,9,1))=CU02(+) order by 1,2 "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/30
    Set grd1.Recordset = adoRecordset
Else
    InsertQueryLog (0) 'Add By Sindy 2010/9/30
End If
SetGrid
CheckOC
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040207a = Nothing
End Sub
Sub SetGrid()
With grd1
    .row = 0
    .col = 0
    .ColWidth(0) = 1200
    .Text = "種類"
    .col = 1
    .ColWidth(1) = 1200
    .Text = "本所案號"
    .col = 2
    .ColWidth(2) = 900
    .Text = "註冊號"
    .col = 3
    .ColWidth(3) = 900
    .Text = "申請案號"
    .col = 4
    .ColWidth(4) = 2000
    .Text = "商標名稱"
    .col = 5
    .ColWidth(5) = 2000
    .Text = "申請人"
End With
End Sub

Private Sub grd1_SelChange()
Dim i As Integer
Dim SeekNewRow As Integer
grd1.Visible = False
SeekNewRow = grd1.MouseRow
If NowRow <> 0 Then
    grd1.row = NowRow
    For i = 0 To grd1.Cols - 1
        grd1.col = i
        grd1.CellBackColor = QBColor(15)
    Next i
End If
If NowRow <> SeekNewRow And SeekNewRow <> 0 Then
    grd1.row = SeekNewRow
     For i = 0 To grd1.Cols - 1
         grd1.col = i
         grd1.CellBackColor = &HFFC0C0
     Next i
     NowRow = SeekNewRow
     cmdok(2).Enabled = True
Else
     NowRow = 0
     cmdok(2).Enabled = False
End If
grd1.Visible = True
End Sub

Sub SaveClose()
Dim oCP01 As String
Dim oCP02 As String
Dim oCP03 As String
Dim oCP04 As String
Dim strSql As String, SCp(1 To 79) As String, i As Integer, oErrMsg As String
oCP01 = SystemNumber(grd1.TextMatrix(NowRow, 1), 1)
oCP02 = SystemNumber(grd1.TextMatrix(NowRow, 1), 2)
oCP03 = SystemNumber(grd1.TextMatrix(NowRow, 1), 3)
oCP04 = SystemNumber(grd1.TextMatrix(NowRow, 1), 4)
On Error GoTo CheckingErr
cnnConnection.BeginTrans

strSql = "UPDATE TRADEMARK SET TM29='Y',TM30=" & strSrvDate(1) & ",TM31='86'  WHERE TM01='" & oCP01 & "' AND TM02='" & oCP02 & "' AND TM03='" & oCP03 & "' AND TM04='" & oCP04 & "' "
cnnConnection.Execute strSql
strSql = "UPDATE CASEPROGRESS SET CP26='N',CP57=" & strSrvDate(1) & ",CP58='86' WHERE CP01='" & oCP01 & "' AND CP02='" & oCP02 & "' AND CP03='" & oCP03 & "' AND CP04='" & oCP04 & "' AND CP57 IS NULL AND CP27 IS NULL "
cnnConnection.Execute strSql
'add by nickc 2007/08/08 加入  t102inform
'Modify by Amy 2024/09/06 未加欄位會error
strSql = "insert into t102inform (ti01,ti02,ti03,ti04) select " & strSrvDate(1) & ",np01,'" & strUserNum & "',np22 from nextprogress where NP02='" & oCP01 & "' AND NP03='" & oCP02 & "' AND NP04='" & oCP03 & "' AND NP05='" & oCP04 & "' AND NP11 IS NULL AND NP06 IS NULL"
cnnConnection.Execute strSql
strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP11=" & strSrvDate(1) & ",NP12='86' WHERE NP02='" & oCP01 & "' AND NP03='" & oCP02 & "' AND NP04='" & oCP03 & "' AND NP05='" & oCP04 & "' AND NP11 IS NULL AND NP06 IS NULL"
cnnConnection.Execute strSql
' ADD 到案件進度檔
Dim strAutoNum As String
If ClsPDGetAutoNumber("B", strAutoNum, True, True) Then
    CheckOC
    strSql = "select au01||(au02-1911) from autonumber where au01='B'"
    adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If Not adoRecordset.BOF Then adoRecordset.MoveFirst
    If adoRecordset.BOF And adoRecordset.EOF Then MsgBox "自動編號錯誤", vbInformation: Exit Sub
    'Modify By Sindy 2010/8/18 比對自動編號年度
    'strAutoNum = CheckStr(adoRecordset.Fields(0).Value) & strAutoNum
    strAutoNum = "B" + CompAutoNumberYear(CStr(Val(Mid(strSrvDate(1), 1, 4)) - 1911)) + strAutoNum
    CheckOC
    strSql = "insert into caseprogress ( " & _
            "cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10," & _
            "cp11,cp12,cp13,cp14,cp15,cp16,cp17,cp18,cp19,cp20," & _
            "cp21,cp22,cp23,cp24,cp25,cp26,cp27,cp28,cp29,cp30," & _
            "cp31,cp32,cp33,cp34,cp35,cp36,cp37,cp38,cp39,cp40," & _
            "cp41,cp42,cp43,cp44,cp45,cp46,cp47,cp48,cp49,cp50," & _
            "cp51,cp52,cp53,cp54,cp55,cp56,cp57,cp58,cp59,cp60," & _
            "cp61,cp62,cp63,cp64,cp71,cp72,cp73,cp74,cp75,cp76," & _
            "cp77,cp78,cp79) values "
        For i = 1 To 79
            Select Case i
            Case 8, 11, 21, 22, 23, 24, 28, 29, 30, 31, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 50, 51, 52, 55, 56, 59, 60, 61, 62, 63
                SCp(i) = "null "
            Case 13 '智權人員
                SCp(i) = "'" & strUserNum & "'"
            Case 12 '業務區
                SCp(i) = "'" & GetST15(strUserNum) & "'"
            Case 14
                SCp(i) = "'" & strUserNum & "'"
            Case 1
                SCp(i) = "'" & Trim(ChgSQL(oCP01)) & "'"
            Case 2
                SCp(i) = "'" & Trim(ChgSQL(oCP02)) & "'"
            Case 3
                SCp(i) = "'" & Trim(ChgSQL(oCP03)) & "'"
            Case 4
                SCp(i) = "'" & Trim(ChgSQL(oCP04)) & "'"
            Case 5, 27
                SCp(i) = GetTodayDate
            Case 9
                SCp(i) = "'" & strAutoNum & "'"
            Case 20
                SCp(i) = "'N'"
            Case 26, 32
                SCp(i) = "'N'"
            Case 10
                SCp(i) = "'704'"
            Case 64
                SCp(i) = "null "
            Case 65, 66, 67, 68, 69, 70
                SCp(i) = ""
            Case 57
                SCp(i) = strSrvDate(1)
            Case 58
                SCp(i) = "'86'"
            Case Else
                SCp(i) = "null "
            End Select
        Next i
        strSql = strSql & " ("
        For i = 1 To 79
            Select Case i
            Case 65, 66, 67, 68, 69, 70
            Case Else
                strSql = strSql & SCp(i)
                If i <> 79 Then
                strSql = strSql & ","
                End If
            End Select
        Next i
        strSql = strSql & ") "
        cnnConnection.Execute strSql
    Else
        Screen.MousePointer = vbDefault
        oErrMsg = "自動給號錯誤"
        GoTo CheckingErr
    End If
    MsgBox "閉卷成功！", vbInformation, "訊息！"
    cnnConnection.CommitTrans
Exit Sub
CheckingErr:
    If oErrMsg = "" Then
         MsgBox Err.Description
    Else
         MsgBox oErrMsg
    End If
    cnnConnection.RollbackTrans
End Sub
