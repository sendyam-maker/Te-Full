VERSION 5.00
Begin VB.Form frm090619 
   BorderStyle     =   1  '單線固定
   Caption         =   "獎金明細表"
   ClientHeight    =   1620
   ClientLeft      =   3300
   ClientTop       =   2112
   ClientWidth     =   3912
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3912
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1110
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1296
      Width           =   285
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2316
      TabIndex        =   4
      Top             =   24
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3108
      TabIndex        =   5
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   0
      Top             =   552
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2055
      MaxLength       =   1
      TabIndex        =   1
      Top             =   552
      Width           =   405
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1110
      MaxLength       =   1
      TabIndex        =   2
      Top             =   960
      Width           =   315
   End
   Begin VB.Label Label2 
      Caption         =   "顯示方式："
      Height          =   180
      Index           =   2
      Left            =   36
      TabIndex        =   12
      Top             =   1332
      Width           =   912
   End
   Begin VB.Label Label3 
      Caption         =   "(1.明細表 2.回執)"
      Height          =   180
      Left            =   1476
      TabIndex        =   11
      Top             =   1344
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "年"
      Height          =   180
      Index           =   2
      Left            =   1788
      TabIndex        =   10
      Top             =   612
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "季"
      Height          =   180
      Index           =   4
      Left            =   2496
      TabIndex        =   9
      Top             =   612
      Width           =   252
   End
   Begin VB.Label Label2 
      Caption         =   "年季："
      Height          =   180
      Index           =   0
      Left            =   36
      TabIndex        =   8
      Top             =   612
      Width           =   576
   End
   Begin VB.Label Label2 
      Caption         =   "對象："
      Height          =   180
      Index           =   3
      Left            =   36
      TabIndex        =   7
      Top             =   996
      Width           =   612
   End
   Begin VB.Label Label4 
      Caption         =   "(1.承辦人 2.繪圖人員 3.全部 )"
      Height          =   180
      Left            =   1500
      TabIndex        =   6
      Top             =   996
      Width           =   2376
   End
End
Attribute VB_Name = "frm090619"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Dim i As Integer, j As Integer, k As Integer, s As Integer, TextOk As Boolean, SeekAction As Integer, SeekRec As Variant
Dim StrSQL6 As String, strTemp1 As Variant, SeekTemp As String, DELMenu() As String, DELTemp() As String, SeekBmk1 As Variant, SeekBmk2 As Variant, SeekBmk3 As Variant
Dim strTemp(0 To 3) As String, PLeft(0 To 3) As Integer, Page As Integer, iPrint As Integer


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Or Len(txt1(1)) = 0 Then
        s = MsgBox("年季不可空白!!", , "USER 輸入錯誤")
        If Len(txt1(1)) = 0 Then txt1(1).SetFocus
        If Len(txt1(0)) = 0 Then txt1(0).SetFocus
        Exit Sub
     Else
        If Len(txt1(2)) = 0 Then
            s = MsgBox("對象不可空白!!", , "USER 輸入錯誤")
            txt1(2).SetFocus
            Exit Sub
        Else
            If Len(txt1(3)) = 0 Then
                s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                txt1(3).SetFocus
                Exit Sub
            Else
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/17 清除查詢印表記錄檔欄位
                Process
                Me.Enabled = True
                Screen.MousePointer = vbDefault
            End If
        End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
cnnConnection.Execute "DELETE FROM R090619 WHERE ID='" & strUserNum & "' "
StrSQL6 = ""
If Val(txt1(2)) = 1 Then
   pub_QL05 = pub_QL05 & ";" & Label2(3) & "1.承辦人" 'Add By Sindy 2010/12/17
   StrSQL6 = StrSQL6 + " AND s1.ST03<>'P13' "
Else
   If Val(txt1(2)) = 2 Then
      pub_QL05 = pub_QL05 & ";" & Label2(3) & "2.繪圖人員" 'Add By Sindy 2010/12/17
      StrSQL6 = StrSQL6 + " AND s1.ST03='P13' "
   End If
End If
pub_QL05 = pub_QL05 & ";" & Label2(0) & txt1(0) & Label1(2) & txt1(1) & Label1(4) 'Add By Sindy 2010/12/17
'92.04.03 nick add left join
'strSQL = "SELECT s1.ST02,SB04,SB05,SB04+SB05 FROM STAFFBONUS,STAFF s1,staff s2 WHERE SB01=s1.ST01(+) AND SB02=" & Val(txt1(0)) + 1911 & " AND SB03=" & Val(txt1(1)) & " and substr(s1.st03,1,2)=substr(s2.st03,1,2) and s2.st01='" & strUserNum & "' " & StrSQL6 & " ORDER BY 1 "
'Modified by Morgan 2018/5/23
'strSql = "SELECT s1.ST02,SB04,SB05,SB04+SB05 FROM STAFFBONUS,STAFF s1,staff s2 WHERE SB01=s1.ST01(+) AND SB02=" & Val(txt1(0)) + 1911 & " AND SB03=" & Val(txt1(1)) & " and substr(s2.st03,1,2)=substr(s1.st03,1,2)(+) and s2.st01='" & strUserNum & "' " & StrSQL6 & " ORDER BY 1 "
strSql = "SELECT s1.ST02,SB04,SB05,SB04+SB05 FROM STAFFBONUS,STAFF s1,staff s2 WHERE SB01=s1.ST01(+) AND SB02=" & Val(txt1(0)) + 1911 & " AND SB03=" & Val(txt1(1)) & " and substr(s2.st03,1,2)=substr(s1.st03(+),1,2) and s2.st01='" & strUserNum & "' " & StrSQL6 & " ORDER BY 1 "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        k = 0
        'FRM100.Show
        'FRM100.Tag = Trim(str(.RecordCount)) & "=0"
        DoEvents
        Do While .EOF = False
            For i = 0 To 3
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            cnnConnection.Execute "INSERT INTO R090619 VALUES('" & strTemp(0) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & ",'" & strUserNum & "') "
            .MoveNext
            k = k + 1
            'FRM100.Tag = Trim(str(.RecordCount)) & "=" & Trim(str(K))
            'FRM100.StrMenu
            DoEvents
        Loop
    End If
End With
CheckOC
'UNLOAD FRM100
If Val(txt1(3)) = 1 Then
   pub_QL05 = pub_QL05 & ";" & Label2(2) & "1.明細表" 'Add By Sindy 2010/12/17
   PrintData1
Else
   pub_QL05 = pub_QL05 & ";" & Label2(2) & "2.回執" 'Add By Sindy 2010/12/17
   PrintData2
End If
End Sub

Sub PrintData1()        '明細
strSql = "SELECT * FROM R090619 WHERE ID='" & strUserNum & "' "
CheckOC
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        .MoveFirst
        PrintTitle1
        Do While .EOF = False
            For i = 0 To 3
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            PrintDatil1
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle1
            End If
            .MoveNext
        Loop
    Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/17
      ShowNoData
      Exit Sub
    End If
End With
CheckOC
ShowLine
PrintEnd1
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd1()
'列印結尾
strSql = "SELECT '合  計',SUM(R112002),SUM(R112003),SUM(DECODE(R112002,0,0,R112002)) + SUm(DECODE(R112003,0,0,R112003)) FROM R090619 WHERE ID='" & strUserNum & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        For i = 0 To 3
            strTemp(i) = CheckStr(.Fields(i))
        Next i
        PrintDatil1
        If iPrint >= 14000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle1
        End If
    End If
End With
CheckOC
End Sub

Sub PrintDatil1() '列印資料

Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
For i = 1 To 3
    Printer.CurrentX = PLeft(i) + 1000 - Printer.TextWidth(Format(strTemp(i), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "####0")
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft1() '定陣列

Erase PLeft
PLeft(0) = 0
PLeft(1) = 1500
PLeft(2) = 3000
PLeft(3) = 4500
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(11000, iPrint + 150)
iPrint = iPrint + 300
End Sub

Sub PrintTitle1() '列印抬頭

GetPleft1
iPrint = 0
Printer.Orientation = 1
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 4000
Printer.CurrentY = iPrint
Printer.Print "獎金明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserNum
Printer.CurrentX = 8800
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print txt1(0) & " 年度  第 " & txt1(1) & " 季 "
Printer.CurrentX = 8800
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(11000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 15000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "姓名"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "獎金"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "額外獎勵"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "小計"
iPrint = iPrint + 300
If iPrint >= 15000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(11000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 15000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
    Exit Sub
End If
End Sub

Sub PrintTitle2() '列印抬頭

Printer.Orientation = 1
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
For i = 0 To 5
    Printer.Line (500, 1000 + i)-(10000, 1000 + i)
    Printer.Line (500 + i, 1000)-(500 + i, 10000)
    Printer.Line (10000 + i, 1000)-(10000 + i, 10000)
    Printer.Line (500, 10000 + i)-(10000, 10000 + i)
Next i
Printer.CurrentX = 3500
Printer.CurrentY = 1500
Printer.Print "回執"
Printer.Font.Size = 18
Printer.CurrentX = 600
Printer.CurrentY = 3000
Printer.Print "姓名：" & strTemp(0)
Printer.CurrentX = 1500
Printer.CurrentY = 6000
Printer.Print "茲收到 " & txt1(0) & " 年 " & txt1(1) & " 季 獎金 " & Format(strTemp(3), "###,###,###,###,##0") & " 元 "
Printer.CurrentX = 6000
Printer.CurrentY = 8000
Printer.Print "簽名："
End Sub

Sub PrintData2()        '回執
strSql = "SELECT * FROM R090619 WHERE ID='" & strUserNum & "' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 3
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            PrintTitle2
            .MoveNext
            If .EOF = False Then
                Printer.NewPage
            End If
        Loop
    Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/17
      ShowNoData
      Exit Sub
    End If
End With
CheckOC
Printer.EndDoc
ShowPrintOk
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(2) = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090619 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(Index)) <> 0 Then
      If IsNumeric(txt1(0)) = False Then
          s = MsgBox("年請輸入數字!!", , "USER 輸入錯誤")
          txt1(0).SetFocus
          txt1(0).SelStart = 0
          txt1(0).SelLength = Len(txt1(0))
          Exit Sub
      End If
     End If
Case 1
     If Len(txt1(Index)) <> 0 Then
         If IsNumeric(txt1(1)) = False Or Val(txt1(1)) < 1 Or Val(txt1(1)) > 4 Then
            s = MsgBox("季請輸入 1-4 數字!!", , "USER 輸入錯誤")
            txt1(1).SetFocus
            txt1(1).SelStart = 0
            txt1(1).SelLength = Len(txt1(1))
            Exit Sub
         End If
      End If
Case 2
      If Len(txt1(Index)) <> 0 Then
         If IsNumeric(txt1(2)) = False Or Val(txt1(2)) < 1 Or Val(txt1(2)) > 3 Then
            s = MsgBox("對象請輸入 1-3 數字!!", , "USER 輸入錯誤")
            txt1(2).SetFocus
            txt1(2).SelStart = 0
            txt1(2).SelLength = Len(txt1(2))
            Exit Sub
         End If
      End If
Case 3
      If Len(txt1(Index)) <> 0 Then
         If IsNumeric(txt1(3)) = False Or Val(txt1(3)) < 1 Or Val(txt1(3)) > 2 Then
            s = MsgBox("顯示方式請輸入 1-2 數字!!", , "USER 輸入錯誤")
            txt1(3).SetFocus
            txt1(3).SelStart = 0
            txt1(3).SelLength = Len(txt1(3))
            Exit Sub
         End If
      End If
Case Else
End Select
End Sub
