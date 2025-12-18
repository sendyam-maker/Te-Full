VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090705 
   BorderStyle     =   1  '單線固定
   Caption         =   "未齊備、未完稿、未發文查詢"
   ClientHeight    =   4260
   ClientLeft      =   2580
   ClientTop       =   1290
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4455
   Begin VB.Frame Frame1 
      Height          =   1785
      Left            =   90
      TabIndex        =   25
      Top             =   2400
      Width           =   4305
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   11
         Left            =   2430
         TabIndex        =   14
         Top             =   1440
         Width           =   525
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   10
         Left            =   2790
         TabIndex        =   12
         Top             =   1140
         Width           =   525
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   9
         Left            =   2430
         TabIndex        =   10
         Top             =   810
         Width           =   525
      End
      Begin VB.CheckBox Check1 
         Caption         =   "已墨圖完成未發文之時限："
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1470
         Width           =   2520
      End
      Begin VB.CheckBox Check1 
         Caption         =   "已草圖完成未墨圖齊備之時限："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1155
         Width           =   2850
      End
      Begin VB.CheckBox Check1 
         Caption         =   "已收文未草圖齊備之時限："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   855
         Width           =   2490
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   8
         Left            =   1050
         MaxLength       =   1
         TabIndex        =   8
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "天"
         Height          =   180
         Index           =   2
         Left            =   3000
         TabIndex        =   31
         Top             =   1470
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "天"
         Height          =   180
         Index           =   1
         Left            =   3390
         TabIndex        =   30
         Top             =   1155
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "天"
         Height          =   180
         Index           =   0
         Left            =   3000
         TabIndex        =   29
         Top             =   855
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "作業天數統計範圍："
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   28
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "(1.螢幕 2.報表)"
         Height          =   180
         Index           =   21
         Left            =   1365
         TabIndex        =   27
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "顯示方式："
         Height          =   180
         Index           =   22
         Left            =   120
         TabIndex        =   26
         Top             =   225
         Width           =   1035
      End
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   7
      Left            =   1032
      MaxLength       =   6
      TabIndex        =   7
      Top             =   2064
      Width           =   900
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   6
      Left            =   1056
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1765
      Width           =   900
   End
   Begin VB.TextBox Txt1 
      Height          =   270
      Index           =   5
      Left            =   1056
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1440
      Width           =   315
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   4
      Left            =   1500
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1128
      Width           =   270
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   3
      Left            =   1056
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1128
      Width           =   270
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   2
      Left            =   1848
      MaxLength       =   4
      TabIndex        =   2
      Top             =   792
      Width           =   675
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   1
      Left            =   1056
      MaxLength       =   4
      TabIndex        =   1
      Top             =   804
      Width           =   615
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   0
      Left            =   1056
      TabIndex        =   0
      Top             =   492
      Width           =   1650
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3084
      TabIndex        =   16
      Top             =   48
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2304
      TabIndex        =   15
      Top             =   48
      Width           =   756
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   1980
      TabIndex        =   33
      Top             =   2069
      Width           =   1920
      VariousPropertyBits=   27
      Size            =   "3387;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   1980
      TabIndex        =   32
      Top             =   1770
      Width           =   1920
      VariousPropertyBits=   27
      Size            =   "3387;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line3 
      X1              =   1185
      X2              =   1635
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      X1              =   1380
      X2              =   2265
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      Height          =   180
      Index           =   24
      Left            =   1800
      TabIndex        =   24
      Top             =   1188
      Width           =   2412
   End
   Begin VB.Label Label1 
      Caption         =   "(1.繪圖人員 2.承辦人)"
      Height          =   180
      Index           =   14
      Left            =   1416
      TabIndex        =   23
      Top             =   1512
      Width           =   1908
   End
   Begin VB.Label Label1 
      Caption         =   "繪圖人員："
      Height          =   180
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   1804
      Width           =   1248
   End
   Begin VB.Label Label1 
      Caption         =   "承辦人："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   2114
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   874
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   564
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   1184
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "查詢對象："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   1494
      Width           =   1092
   End
End
Attribute VB_Name = "frm090705"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/07 改成Form2.0 ; lbl1(index) ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 17) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 18) As String, k As Integer
Dim PLeft(0 To 17) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As String, SeekTemp2 As String
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(Txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         Txt1(0).SetFocus
         Exit Sub
     Else
         If Len(Txt1(5)) = 0 Then
             s = MsgBox("查詢對象不可空白!!", , "USER 輸入錯誤")
             Txt1(5).SetFocus
             Exit Sub
         Else
             If Len(Txt1(8)) = 0 Then
                 s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                 Txt1(8).SetFocus
                 Exit Sub
             Else
                 If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 Then
                     s = MsgBox("作業天數統計範圍至少核取一項!!", , "USER 輸入錯誤")
                     Check1(0).SetFocus
                     Exit Sub
                 Else
                     If Check1(0).Value = 1 Then
                         If Len(Txt1(9)) = 0 Then
                            s = MsgBox(Check1(0).Caption & " 不可空白!!", , "USER 輸入錯誤")
                            Txt1(9).SetFocus
                            Exit Sub
                         End If
                     End If
                     If Check1(1).Value = 1 Then
                         If Len(Txt1(10)) = 0 Then
                            s = MsgBox(Check1(1).Caption & " 不可空白!!", , "USER 輸入錯誤")
                            Txt1(10).SetFocus
                            Exit Sub
                         End If
                     End If
                     If Check1(2).Value = 1 Then
                         If Len(Txt1(11)) = 0 Then
                            s = MsgBox(Check1(2).Caption & " 不可空白!!", , "USER 輸入錯誤")
                            Txt1(11).SetFocus
                            Exit Sub
                         End If
                     End If
                     Screen.MousePointer = vbHourglass
                     Me.Enabled = False
                     ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/15 清除查詢印表記錄檔欄位
                     If Val(Txt1(8)) = 1 Then
                        pub_QL05 = pub_QL05 & ";" & Label1(22) & "1.螢幕" 'Add By Sindy 2010/12/15
                     Else
                        pub_QL05 = pub_QL05 & ";" & Label1(22) & "2.報表" 'Add By Sindy 2010/12/15
                     End If
                     'For i = 0 To 11
                     '   If (StrTemp99(i) <> txt1(i)) And (i <> 8) Then
                            Process
                     '       For j = 0 To 11
                     '           StrTemp99(j) = txt1(j)
                     '       Next j
                     '       Exit For
                     '   End If
                     'Next i
                     Process1
                     Me.Enabled = True
                     Screen.MousePointer = vbDefault
                 End If
             End If
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process1()
If Val(Txt1(8)) = 1 Then
    Me.Hide
    frm090705_1.Show
Else
    PrintData
End If
End Sub

Sub PrintDatil_1() '列印資料

'Printer.CurrentX = PLeft(1) + 500 - Printer.TextWidth(Format(StrTemp(1), "####0"))
'Printer.CurrentY = iPrint
'Printer.Print Format(StrTemp(1), "####0")
For i = 0 To 7
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub PrintDatil_2() '列印資料

'Printer.CurrentX = PLeft(1) + 500 - Printer.TextWidth(Format(StrTemp(1), "####0"))
'Printer.CurrentY = iPrint
'Printer.Print Format(StrTemp(1), "####0")
For i = 0 To 7
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub PrintDatil_3() '列印資料

'Printer.CurrentX = PLeft(1) + 500 - Printer.TextWidth(Format(StrTemp(1), "####0"))
'Printer.CurrentY = iPrint
'Printer.Print Format(StrTemp(1), "####0")
For i = 0 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub PrintTitle() '列印抬頭
iPrint = 0
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5000
Printer.CurrentY = iPrint
Printer.Print "未齊備/未完稿/未發文明細表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 7000
Printer.CurrentY = iPrint
Printer.Print "年月：" & Format(Format(GetTodayDate, "####/##/##"), "ee/MM")
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
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
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(15000, iPrint + 150)
iPrint = iPrint + 300
End Sub

Sub GetPleft_1()
Erase PLeft
'定陣列
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 2200
PLeft(3) = 4500
PLeft(4) = 8600
PLeft(5) = 9800
PLeft(6) = 11000
PLeft(7) = 12300
PLeft(8) = 14000
End Sub

Sub GetPleft_2()
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 2200
PLeft(3) = 4500
PLeft(4) = 8600
PLeft(5) = 9800
PLeft(6) = 11000
PLeft(7) = 12300
PLeft(8) = 14000
End Sub

Sub GetPleft_3()
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 2200
PLeft(3) = 4500
PLeft(4) = 8600
PLeft(5) = 9800
PLeft(6) = 11000
PLeft(7) = 12300
PLeft(8) = 14000
End Sub

Sub PrintTitle_1()
GetPleft_1 '列印抬頭
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print SeekTemp2
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_1
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "所別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
If Txt1(5) = "1" Then
    Printer.Print "繪圖人員"
Else
    Printer.Print "承辦人"
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "收文天數"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
If Val(Txt1(5)) = 1 Then
    Printer.Print "承辦人"
Else
    Printer.Print "繪圖人員"
End If
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_1
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_1
    Exit Sub
End If
End Sub

Sub PrintTitle_2() '列印抬頭
GetPleft_2
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print SeekTemp2
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_1
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "所別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
If Txt1(5) = "1" Then
    Printer.Print "繪圖人員"
Else
    Printer.Print "承辦人"
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "草圖完稿日"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "草圖完稿天數"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
If Val(Txt1(5)) = 1 Then
    Printer.Print "承辦人"
Else
    Printer.Print "繪圖人員"
End If
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_1
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_1
    Exit Sub
End If
End Sub

Sub PrintTitle_3() '列印抬頭
GetPleft_3
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print SeekTemp2
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_1
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "所別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
If Txt1(5) = "1" Then
    Printer.Print "繪圖人員"
Else
    Printer.Print "承辦人"
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "墨圖完稿日"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "墨圖完稿天數"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
If Val(Txt1(5)) = 1 Then
    Printer.Print "承辦人"
Else
    Printer.Print "繪圖人員"
End If
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_1
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_1
    Exit Sub
End If
End Sub

Sub PrintData()
PrintTitle
Page = 1
If Check1(0).Value = 1 Then
    strSql = "SELECT COUNT(*) FROM R090705 WHERE ID='" & strUserNum & "' AND R107014='1' "
    CheckOC2
    With adoRecordset1
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
            SeekTemp2 = "已收文未草圖齊備之時限 " & frm090705.Txt1(9).Text & " 天,共 " & str(Trim(CheckStr(.Fields(0)))) & " 件 "
        End If
    End With
    CheckOC2
    strSql = "SELECT r107001,R107002,R107003,R107004,R107005,R107006,R107007,R107008,R107013 FROM R090705 WHERE R107014='1' AND ID='" & strUserNum & "' ORDER BY 1,2,3 "
    CheckOC
    With adoRecordset
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
            .MoveFirst
            PrintTitle_1
            Do While .EOF = False
                For i = 0 To 8
                    strTemp(i) = CheckStr(.Fields(i))
                Next i
                strTemp(1) = StrToStr(strTemp(1), 8)
                strTemp(3) = StrToStr(strTemp(3), 18)
                strTemp(4) = StrToStr(strTemp(4), 8)
                strTemp(5) = StrToStr(strTemp(5), 8)
                strTemp(8) = StrToStr(strTemp(8), 8)
                PrintDatil_1
                If iPrint >= 9000 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitle
                    PrintTitle_1
                End If
                .MoveNext
            Loop
        Else
            'SeekTemp2 = "超過已收文未齊備之時限申請案  天, 非申請案  天, 申請案共  件, 非申請案共  件 "
        End If
    End With
    CheckOC2
End If
If Check1(1).Value = 1 Then
    iPrint = iPrint + 600
    strSql = "SELECT COUNT(*) FROM R090705 WHERE ID='" & strUserNum & "' AND R107014='2' "
    CheckOC2
    With adoRecordset1
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
            SeekTemp2 = "已草圖完成未墨圖齊備之時限 " & frm090705.Txt1(10).Text & " 天,共 " & str(Trim(CheckStr(.Fields(0)))) & " 件 "
        End If
    End With
    CheckOC2
    strSql = "SELECT r107001,R107002,R107003,R107004,R107005,R107006,R107009,R107010,R107013 FROM R090705 WHERE R107014='2' AND ID='" & strUserNum & "' ORDER BY 1,2,3 "
    CheckOC
    With adoRecordset
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
            .MoveFirst
            PrintTitle_2
            Do While .EOF = False
                For i = 0 To 8
                    strTemp(i) = CheckStr(.Fields(i))
                Next i
                strTemp(1) = StrToStr(strTemp(1), 8)
                strTemp(3) = StrToStr(strTemp(3), 18)
                strTemp(4) = StrToStr(strTemp(4), 8)
                strTemp(5) = StrToStr(strTemp(5), 8)
                strTemp(8) = StrToStr(strTemp(8), 8)
                PrintDatil_1
                If iPrint >= 9000 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitle
                    PrintTitle_2
                End If
                .MoveNext
            Loop
        Else
            'SeekTemp2 = "超過已收文未齊備之時限申請案  天, 非申請案  天, 申請案共  件, 非申請案共  件 "
        End If
    End With
    CheckOC2
End If
If Check1(2).Value = 1 Then
    iPrint = iPrint + 600
    strSql = "SELECT COUNT(*) FROM R090705 WHERE ID='" & strUserNum & "' AND R107014='3' "
    CheckOC2
    With adoRecordset1
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
            SeekTemp2 = "已墨圖完成未發文之時限 " & frm090705.Txt1(11).Text & " 天,共 " & str(Trim(CheckStr(.Fields(0)))) & " 件 "
        End If
    End With
    CheckOC2
    strSql = "SELECT r107001,R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107013 FROM R090705 WHERE R107014='3' AND ID='" & strUserNum & "' ORDER BY 1,2,3 "
    CheckOC
    With adoRecordset
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 And .RecordCount > 0 Then
            .MoveFirst
            PrintTitle_3
            Do While .EOF = False
                For i = 0 To 8
                    strTemp(i) = CheckStr(.Fields(i))
                Next i
                strTemp(1) = StrToStr(strTemp(1), 8)
                strTemp(3) = StrToStr(strTemp(3), 18)
                strTemp(4) = StrToStr(strTemp(4), 8)
                strTemp(5) = StrToStr(strTemp(5), 8)
                strTemp(8) = StrToStr(strTemp(8), 8)
                PrintDatil_1
                If iPrint >= 9000 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitle
                    PrintTitle_3
                End If
                .MoveNext
            Loop
        Else
            'SeekTemp2 = "超過已收文未齊備之時限申請案  天, 非申請案  天, 申請案共  件, 非申請案共  件 "
        End If
    End With
    CheckOC2
End If
CheckOC
Printer.EndDoc
End Sub

Sub Process()
cnnConnection.Execute "DELETE FROM R090705 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
If Len(Txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(Txt1(0), 1) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0) 'Add By Sindy 2010/12/15
End If
If Len(Txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & Txt1(1) & "' "
End If
If Len(Txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & Txt1(2) & "' "
End If
If Len(Txt1(1)) <> 0 Or Len(Txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & Txt1(1) & "-" & Txt1(2) 'Add By Sindy 2010/12/15
End If
StrSQL6 = ""
If Val(Txt1(5)) = 1 Then
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "1.繪圖人員" 'Add By Sindy 2010/12/15
    If Len(Txt1(3)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s1.st06>='" & Txt1(3) & "' "
    End If
    If Len(Txt1(4)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s1.st06<='" & Txt1(4) & "' "
    End If
Else
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "2.承辦人" 'Add By Sindy 2010/12/15
    If Len(Txt1(3)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s2.st06>='" & Txt1(3) & "' "
    End If
    If Len(Txt1(4)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s2.st06<='" & Txt1(4) & "' "
    End If
End If
If Len(Txt1(3)) <> 0 Or Len(Txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & Txt1(3) & "-" & Txt1(4) & Label1(24) 'Add By Sindy 2010/12/15
End If
If Len(Txt1(6)) <> 0 Then
    StrSQL6 = StrSQL6 + " and ep13='" & Txt1(6) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & Txt1(6) & lbl1(0) 'Add By Sindy 2010/12/15
Else
    StrSQL6 = StrSQL6 & " and ep13 is not null "
End If
If Len(Txt1(7)) <> 0 Then
    StrSQL6 = StrSQL6 + " and Ep05='" & Txt1(7) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(6) & Txt1(7) & lbl1(1) 'Add By Sindy 2010/12/15
End If
StrSQL6 = StrSQL6 + " and ep20 is null  "
If Val(Txt1(5)) = 2 Then
    '承辦人
    strSql = ""
    '已收文未草圖齊備
    If Check1(0).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(0).Caption & Txt1(9) & Label2(0) 'Add By Sindy 2010/12/15
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
        '92.04.03 nick add left join
        'strSQL = strSQL + " select DECODE(S1.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s1.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp05,0,0,0,0,0,s2.st02,'1' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  pa01=cp01 and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and Ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and cp05 is not null and (ep14 is null or ep14=0) " & strSQL1 & StrSQL6
        strSql = strSql + " select DECODE(S1.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s1.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp05,0,0,0,0,0,s2.st02,'1' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and Ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and cp05 is not null and (ep14 is null or ep14=0) " & strSQL1 & StrSQL6
    End If
    '已草圖完成未墨圖齊備
    If Check1(1).Value = 1 Then
        pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(1).Caption & Txt1(10) & Label2(1) 'Add By Sindy 2010/12/15
        If Len(strSql) <> 0 Then
            strSql = strSql + " union all "
        End If
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
        '92.04.03 nick add left join
        'strSQL = strSQL + " select DECODE(S1.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s1.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,Ep15,0,0,0,s2.st02,'2' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  pa01=cp01 and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep15 is not null and (ep17 is null or ep17=0) " & strSQL1 & StrSQL6
        strSql = strSql + " select DECODE(S1.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s1.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,Ep15,0,0,0,s2.st02,'2' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep15 is not null and (ep17 is null or ep17=0) " & strSQL1 & StrSQL6
    End If
    '已墨圖完成未發文
    If Check1(2).Value = 1 Then
        pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(2).Caption & Txt1(11) & Label2(2) 'Add By Sindy 2010/12/15
        If Len(strSql) <> 0 Then
            strSql = strSql + " union all "
        End If
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
        '92.04.03 nick add left join
        'strSQL = strSQL + " select DECODE(S1.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s1.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,0,0,Ep18,0,s2.st02,'3' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  pa01=cp01 and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep18 is not null and CP27 IS NULL  " & strSQL1 & StrSQL6
        strSql = strSql + " select DECODE(S1.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s1.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,0,0,Ep18,0,s2.st02,'3' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep18 is not null and CP27 IS NULL  " & strSQL1 & StrSQL6
    End If
    CheckOC
Else
    '繪圖人員
    strSql = ""
    '已收文未草圖齊備
    If Check1(0).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(0).Caption & Txt1(9) & Label2(0) 'Add By Sindy 2010/12/15
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
        '92.04.03 nick add left join
        'strSQL = strSQL + " select DECODE(S2.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s2.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp05,0,0,0,0,0,s1.st02,'1' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  pa01=cp01 and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and Ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and cp05 is not null and (ep14 is null or ep14=0) " & strSQL1 & StrSQL6
        strSql = strSql + " select DECODE(S2.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s2.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp05,0,0,0,0,0,s1.st02,'1' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and Ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and cp05 is not null and (ep14 is null or ep14=0) " & strSQL1 & StrSQL6
    End If
    '已草圖完成未墨圖齊備
    If Check1(1).Value = 1 Then
        pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(1).Caption & Txt1(10) & Label2(1) 'Add By Sindy 2010/12/15
        If Len(strSql) <> 0 Then
            strSql = strSql + " union all "
        End If
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
        '92.04.03 nick add left join
        'strSQL = strSQL + " select DECODE(S2.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s2.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,Ep15,0,0,0,s1.st02,'2' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  pa01=cp01 and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep15 is not null and (ep17 is null or ep17=0) " & strSQL1 & StrSQL6
        strSql = strSql + " select DECODE(S2.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s2.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,Ep15,0,0,0,s1.st02,'2' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep15 is not null and (ep17 is null or ep17=0) " & strSQL1 & StrSQL6
    End If
    '已墨圖完成未發文
    If Check1(2).Value = 1 Then
        pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(2).Caption & Txt1(11) & Label2(2) 'Add By Sindy 2010/12/15
        If Len(strSql) <> 0 Then
            strSql = strSql + " union all "
        End If
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
        '92.04.03 nick add left join
        'strSQL = strSQL + " select DECODE(S2.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s2.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,0,0,Ep18,0,s1.st02,'3' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  pa01=cp01 and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep18 is not null and CP27 IS NULL  " & strSQL1 & StrSQL6
        strSql = strSql + " select DECODE(S2.ST06,'1','北所','2','中所','3','南所','4','高所','其他'),s2.st02,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,0,0,Ep18,0,s1.st02,'3' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND  cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and ep13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep18 is not null and CP27 IS NULL  " & strSQL1 & StrSQL6
    End If
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/15
        .MoveFirst
        k = 0
        DoEvents
        Do While .EOF = False
            For i = 0 To 13
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            TestOk = True
            '發文日
            If Len(strTemp(6)) <> 0 And Val(strTemp(6)) <> 0 Then
                '取得實際工作天數
                strTemp(7) = str(GetWorkDay(GetTodayDate, strTemp(6)))
               'Modify By Cheng 2002/04/17
'                If Val(txt1(9)) < Val(strTemp(7)) Then
                If Val(Txt1(9)) > Val(strTemp(7)) Then
                    TestOk = False
                End If
                strTemp(6) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(6)))
            Else
                strTemp(6) = ""
            End If
            If Len(strTemp(8)) <> 0 And Val(strTemp(8)) <> 0 Then
                strTemp(9) = str(GetWorkDay(GetTodayDate, strTemp(8)))
               'Modify By Cheng 2002/04/17
'                If Val(txt1(10)) < Val(strTemp(9)) Then
                If Val(Txt1(10)) > Val(strTemp(9)) Then
                    TestOk = False
                End If
                strTemp(8) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(8)))
            Else
                strTemp(8) = ""
            End If
            If Len(strTemp(10)) <> 0 And Val(strTemp(10)) <> 0 Then
                strTemp(11) = str(GetWorkDay(GetTodayDate, strTemp(10)))
               'Modify By Cheng 2002/04/17
'                If Val(txt1(11)) < Val(strTemp(11)) Then
                If Val(Txt1(11)) > Val(strTemp(11)) Then
                    TestOk = False
                End If
                strTemp(10) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(10)))
            Else
                strTemp(10) = ""
            End If
            If TestOk = True Then
                strSql = "insert into r090705 values ('" & strTemp(0) & "','" & strTemp(1) & "','" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "'," & Val(strTemp(7)) & ",'" & strTemp(8) & "'," & Val(strTemp(9)) & ",'" & strTemp(10) & "'," & Val(strTemp(11)) & ",'" & strTemp(12) & "','" & strTemp(13) & "','" & strUserNum & "') "
                cnnConnection.Execute strSql
            End If
            .MoveNext
            k = k + 1
            'FRM100.Tag = Trim(str(.RecordCount)) & "=" & Trim(str(K))
            'FRM100.StrMenu
            DoEvents
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/15
    End If
End With
CheckOC
''UNLOAD FRM100
End Sub

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_Load()
m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
MoveFormToCenter Me
Txt1(0) = Systemkind_g_P
Txt1(8) = "1"
Select Case ProState
Case "1"  '個人
    Frame1.Left = 90
    Frame1.Top = 1100
    Me.Height = 3360
    Txt1(5) = "1"
    Txt1(6) = strUserNum
    Txt1(3).Enabled = False
    Txt1(4).Enabled = False
    Txt1(5).Enabled = False
    Txt1(6).Enabled = False
    Txt1(7).Enabled = False
Case "2"  '主管
    Frame1.Top = 2400
    Frame1.Left = 90
    Me.Height = 4635
Case Else
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090705 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
Txt1(Index).SelStart = 0
Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '系統類別
      'Add By Cheng 2002/01/07
      Me.Txt1(Index).Text = GetAllSysKind(Me.Txt1(Index))
     strTemp1 = Split(UCase(Systemkind_g_P), ",")
     strTemp2 = Split(UCase(Txt1(0)), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp2(i) = strTemp1(j) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限!! ", , "USER 權限問題")
            Txt1(0).SetFocus
            Txt1(0).SelStart = 0
            Txt1(0).SelLength = Len(Txt1(0))
            Exit Sub
        End If
     Next i
Case 2
     If RunNick(Txt1(Index - 1), Txt1(Index)) Then
        Txt1(Index - 1).SetFocus
        txt1_GotFocus (Index - 1)
        Exit Sub
     End If
Case 3
     Select Case Trim(Txt1(3))
     Case "1", "2", "", "3", "4", "5"
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          Txt1(3).SetFocus
          Txt1(3).SelStart = 0
          Txt1(3).SelLength = Len(Txt1(3))
          Exit Sub
     End Select
Case 4
     Select Case Trim(Txt1(4))
     Case "1", "2", "", "3", "4", "5"
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          Txt1(4).SetFocus
          Txt1(4).SelStart = 0
          Txt1(4).SelLength = Len(Txt1(4))
          Exit Sub
     End Select
     If RunNick(Txt1(Index - 1), Txt1(Index)) Then
        Txt1(Index - 1).SetFocus
        txt1_GotFocus (Index - 1)
        Exit Sub
     End If
Case 5
     Select Case Trim(Txt1(5))
     Case "1", "2", ""
     Case Else
          s = MsgBox("查詢對象只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(5).SetFocus
          Txt1(5).SelStart = 0
          Txt1(5).SelLength = Len(Txt1(5))
          Exit Sub
     End Select
Case 6
     lbl1(0).Caption = GetPrjSalesNM(Txt1(6))
     If Trim(Txt1(Index)) <> "" Then
        If Trim(lbl1(0)) = "" Then
            s = MsgBox("繪圖人員輸入錯誤！", , "錯誤！")
            Txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 7
     lbl1(1).Caption = GetPrjSalesNM(Txt1(7))
     If Trim(Txt1(Index)) <> "" Then
        If Trim(lbl1(1)) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            Txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 8
     Select Case Trim(Txt1(Index))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示方式只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(Index).SetFocus
          Txt1(Index).SelStart = 0
          Txt1(Index).SelLength = Len(Txt1(Index))
          Exit Sub
     End Select
Case 9, 10, 11
     For i = 1 To Len(Txt1(Index))
          strSql = Mid(Txt1(Index), i, 1)
          If InStr(1, "0123456789 ", strSql) = 0 Then
                s = MsgBox("作業天數只能輸入數字!!", , "USER 輸入錯誤")
                Txt1(Index).SetFocus
                Txt1(Index).SelStart = 0
                Txt1(Index).SelLength = Len(Txt1(Index))
                Exit Sub
          End If
     Next i
Case Else
End Select
End Sub

