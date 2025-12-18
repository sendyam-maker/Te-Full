VERSION 5.00
Begin VB.Form frm090706_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員工作進度資料查詢"
   ClientHeight    =   2715
   ClientLeft      =   1920
   ClientTop       =   1335
   ClientWidth     =   6045
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6045
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4080
      TabIndex        =   1
      Top             =   10
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   4860
      TabIndex        =   0
      Top             =   10
      Width           =   1092
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   7
      Left            =   60
      TabIndex        =   9
      Top             =   2472
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   6
      Left            =   60
      TabIndex        =   8
      Top             =   2184
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   5
      Left            =   60
      TabIndex        =   7
      Top             =   1896
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   6
      Top             =   1620
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   3
      Left            =   60
      TabIndex        =   5
      Top             =   1344
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   1056
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   5952
   End
   Begin VB.Label lbl1 
      Caption         =   "可辦草圖："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   504
      Width           =   5952
   End
End
Attribute VB_Name = "frm090706_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/28 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 22) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String
Dim PLeft(0 To 22) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, TempSeekNick As String

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Screen.MousePointer = vbHourglass
     Me.Enabled = False
     PrintData
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1
     frm090706_1.Show
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Process
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090706_2 = Nothing
End Sub

Sub Process()
lbl1(0).Caption = "可辦草圖：    0   件"
lbl1(1).Caption = "可辦墨圖：    0   件"
lbl1(2).Caption = "達成草圖：    0   件"
lbl1(3).Caption = "達成墨圖：    0   件    0   張"
lbl1(4).Caption = "其他新案：    0   件    0   張    0   點"
lbl1(5).Caption = "其他舊案：    0   件    0   點"
lbl1(6).Caption = "逾時草圖：    0   件"
lbl1(7).Caption = "逾時墨圖：    0   件"
strSql = "SELECT r111002,SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)) FROM R090706_1 WHERE ID='" & strUserNum & "' AND R111001='" & frm090706_1.Combo1.Text & "' group by r111002 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
        Select Case Val(CheckStr(.Fields(0)))
        Case 1
             lbl1(0).Caption = "可辦草圖：  " & CheckStr(.Fields(1)) & "  件"
        Case 2
             lbl1(1).Caption = "可辦墨圖：  " & CheckStr(.Fields(1)) & "  件"
        Case 3
             lbl1(2).Caption = "達成草圖：  " & CheckStr(.Fields(1)) & "  件"
        Case 4
             lbl1(3).Caption = "達成墨圖：  " & CheckStr(.Fields(1)) & "  件  " & CheckStr(.Fields(2)) & "  張"
        Case 5
             lbl1(4).Caption = "其他新案：  " & CheckStr(.Fields(1)) & "  件  " & CheckStr(.Fields(2)) & "  張  " & CheckStr(.Fields(3)) & "  點"
        Case 6
             lbl1(5).Caption = "其他舊案：  " & CheckStr(.Fields(1)) & "  件  " & CheckStr(.Fields(3)) & "  點"
        Case 7
             lbl1(6).Caption = "逾時草圖：  " & CheckStr(.Fields(1)) & "  件"
        Case 8
             lbl1(7).Caption = "逾時墨圖：  " & CheckStr(.Fields(1)) & "  件"
        Case Else
        End Select
        .MoveNext
        Loop
    End If
End With
CheckOC
End Sub

Sub PrintData()
strSql = "SELECT DISTINCT R110001 FROM R090706 WHERE ID='" & strUserNum & "' "
CheckOC2
Page = 1
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    adoRecordset1.MoveFirst
    Do While adoRecordset1.EOF = False
        strTemp3 = CheckStr(adoRecordset1.Fields(0))
        PrintData1 (CheckStr(adoRecordset1.Fields(0)))
        PrintEnd1 (CheckStr(adoRecordset1.Fields(0)))
        Page = Page + 1
        Printer.NewPage
        adoRecordset1.MoveNext
    Loop
End If
CheckOC2
Printer.EndDoc
End Sub

Sub PrintData1(Strindex As String)
If Len(Strindex) = 0 Then
    strSql = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') "
Else
    strSql = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 22
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(3) = StrToStr(strTemp(3), 3)
            strTemp(6) = StrToStr(strTemp(6), 7)
            strTemp(7) = StrToStr(strTemp(7), 4)
            strTemp(22) = StrToStr(strTemp(20), 8)
            PrintDatil
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC
End Sub

Sub PrintEnd1(Strindex As String)
'列印結尾
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
If Len(Strindex) = 0 Then
    strSql = "SELECT r111002,SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)) FROM R090706_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') group by r111002 order by r111002 "
Else
    strSql = "SELECT r111002,SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)) FROM R090706_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' group by r111002 order by r111002 "
End If
CheckOC
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "當月統計："
iPrint = iPrint + 300
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            Select Case Val(CheckStr(.Fields(0)))
            Case 1
                 Printer.CurrentX = 0
                 Printer.CurrentY = iPrint
                 Printer.Print "可辦草圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件"
            Case 2
                 Printer.CurrentX = 4000
                 Printer.CurrentY = iPrint
                 Printer.Print "可辦墨圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件"
            Case 3
                 Printer.CurrentX = 8000
                 Printer.CurrentY = iPrint
                 Printer.Print "達成草圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件 " & Format(CheckStr(.Fields(2)), "###,###,###,###,##0") & " 張"
            Case 4
                 Printer.CurrentX = 12000
                 Printer.CurrentY = iPrint
                 Printer.Print "達成墨圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件 " & Format(CheckStr(.Fields(2)), "###,###,###,###,##0") & " 張 " & Format(CheckStr(.Fields(3)), "###,###,###,###,##0.00") & " 點 "
            Case 5
                 Printer.CurrentX = 0
                 Printer.CurrentY = iPrint + 300
                 Printer.Print "其他新案:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件 " & Format(CheckStr(.Fields(3)), "###,###,###,###,##0.00") & " 點 "
            Case 6
                 Printer.CurrentX = 4000
                 Printer.CurrentY = iPrint + 300
                 Printer.Print "其他舊案:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件 " & Format(CheckStr(.Fields(3)), "###,###,###,###,##0.00") & " 點"
            Case 7
                 Printer.CurrentX = 8000
                 Printer.CurrentY = iPrint + 300
                 Printer.Print "逾時草圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件"
            Case 8
                 Printer.CurrentX = 12000
                 Printer.CurrentY = iPrint + 300
                 Printer.Print "逾時墨圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0") & " 件"
            Case Else
            End Select
            .MoveNext
        Loop
    End If
End With
CheckOC
End Sub

Sub PrintTitle() '列印抬頭
iPrint = 0
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "繪圖人員工作進度資料表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "年月：" & frm090706.txt1(3) & "/" & frm090706.txt1(4)
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "繪圖人員：" & strTemp3
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
GetPleft
Printer.Font.Size = 9
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "收文"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "草  圖"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "草  圖"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "草圖"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "草圖作"
Printer.CurrentX = PLeft(12) + 100
Printer.CurrentY = iPrint
Printer.Print "墨  圖"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "墨  圖"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "墨圖"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "墨圖作"
Printer.CurrentX = PLeft(16) + 100
Printer.CurrentY = iPrint
Printer.Print "承辦時段"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "複雜"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "修  改  時  數"
Printer.CurrentX = PLeft(22)
Printer.CurrentY = iPrint
Printer.Print "備註"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "類別"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "張數"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "業天數"
Printer.CurrentX = PLeft(12) + 100
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "張數"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "業天數"
Printer.CurrentX = PLeft(16) + 100
Printer.CurrentY = iPrint
Printer.Print "草圖"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "墨圖"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "1"
Printer.CurrentX = PLeft(20)
Printer.CurrentY = iPrint
Printer.Print "2"
Printer.CurrentX = PLeft(21)
Printer.CurrentY = iPrint
Printer.Print "3"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.Font.Size = 9
End Sub

Sub PrintDatil() '列印資料
For i = 1 To 22
    Select Case i
    Case 4
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0.00"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "####0.00")
    Case 10, 11, 14, 15, 16, 17, 18, 19, 20, 21
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "####0")
    Case Else
        Printer.CurrentX = PLeft(i)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(i)
    End Select
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
'定陣列
'字 SIZE = 9
'1 WORD = 180 PIX
'0.5 WORD = 90 PIX
'SPACE = 90 PIX
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = PLeft(1) + (2.5 * 180)
PLeft(3) = PLeft(2) + (4.5 * 180)
PLeft(4) = PLeft(3) + (5.5 * 180)
PLeft(5) = PLeft(4) + (2.5 * 180)
PLeft(6) = PLeft(5) + (8 * 180)
PLeft(7) = PLeft(6) + (8 * 180)
PLeft(8) = PLeft(7) + (4.5 * 180)
PLeft(9) = PLeft(8) + (4.5 * 180)
PLeft(10) = PLeft(9) + (4 * 180)
PLeft(11) = PLeft(10) + (3 * 180)
PLeft(12) = PLeft(11) + (3 * 180)
PLeft(13) = PLeft(12) + (4.5 * 180)
PLeft(14) = PLeft(13) + (4.5 * 180)
PLeft(15) = PLeft(14) + (3 * 180)
PLeft(16) = PLeft(15) + (3 * 180)
PLeft(17) = PLeft(16) + (3 * 180)
PLeft(18) = PLeft(17) + (3 * 180)
PLeft(19) = PLeft(18) + (3 * 180)
PLeft(20) = PLeft(19) + (3 * 180)
PLeft(21) = PLeft(20) + (3 * 180)
PLeft(22) = PLeft(21) + (3 * 180)
End Sub

Sub ShowLine()
Printer.Line (0, iPrint + 150)-(16500, iPrint + 150)
iPrint = iPrint + 300
End Sub

