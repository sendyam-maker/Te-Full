VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090701 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員達成情形"
   ClientHeight    =   3735
   ClientLeft      =   1425
   ClientTop       =   1410
   ClientWidth     =   4485
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "新細明體-ExtB"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4485
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   1092
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2775
      Width           =   315
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   3228
      TabIndex        =   12
      Top             =   20
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   2400
      TabIndex        =   11
      Top             =   0
      Width           =   800
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   1092
      MaxLength       =   1
      TabIndex        =   10
      Top             =   3225
      Width           =   330
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   1092
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2358
      Width           =   315
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   1092
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1605
      Width           =   885
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   2112
      MaxLength       =   5
      TabIndex        =   4
      Top             =   1200
      Width           =   900
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   3
      Left            =   1092
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1200
      Width           =   900
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1092
      MaxLength       =   4
      TabIndex        =   1
      Top             =   846
      Width           =   900
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1092
      TabIndex        =   0
      Top             =   468
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   2112
      MaxLength       =   4
      TabIndex        =   2
      Top             =   828
      Width           =   900
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   1548
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1980
      Width           =   375
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   1092
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1980
      Width           =   375
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   2010
      TabIndex        =   25
      Top             =   1620
      Width           =   2055
      Size            =   "3625;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "統計方式："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   105
      TabIndex        =   24
      Top             =   2835
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "(1.舊制  2.新制)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   12
      Left            =   1470
      TabIndex        =   23
      Top             =   2835
      Width           =   1635
   End
   Begin VB.Line Line2 
      X1              =   1656
      X2              =   2811
      Y1              =   1356
      Y2              =   1356
   End
   Begin VB.Label Label1 
      Caption         =   "(1.點數 % 2.件數 % 3.平均 % 4.繪圖人員)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   1440
      TabIndex        =   22
      Top             =   3225
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "(1.螢幕  2.報表)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   1404
      TabIndex        =   21
      Top             =   2412
      Width           =   1632
   End
   Begin VB.Label Label1 
      Caption         =   "列印順序："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   75
      TabIndex        =   20
      Top             =   3285
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "顯示方式："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   108
      TabIndex        =   19
      Top             =   2402
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "繪圖人員："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   108
      TabIndex        =   18
      Top             =   1638
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "發文年月："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   108
      TabIndex        =   17
      Top             =   1256
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   108
      TabIndex        =   16
      Top             =   874
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   15
      Top             =   492
      Visible         =   0   'False
      Width           =   1008
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   2565
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   1212
      X2              =   1812
      Y1              =   2076
      Y2              =   2076
   End
   Begin VB.Label Label1 
      Caption         =   "所別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   108
      TabIndex        =   14
      Top             =   2020
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   1980
      TabIndex        =   13
      Top             =   2016
      Width           =   2412
   End
End
Attribute VB_Name = "frm090701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; lbl1(index) ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, k As Integer
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 19) As String, strTemp3 As String, TestOk As Boolean, StrTemp7(0 To 19) As String, StrTemp99(0 To 7) As String
Dim PLeft(0 To 19) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, PLeft1(1 To 9) As Integer, SeekPrint As Integer, SeekPrintL As Integer
'Add By Cheng 2002/04/16
Dim strSQL71 As String
Dim strSQL73 As String
Dim strSQL74 As String
Dim strSQL72 As String
'Add By Cheng 2004/02/17
Dim strSQL8 As String, strSQL9 As String
'End
'Add By Cheng 2004/03/12
Dim m_Cnt As Integer '明細筆數
'End
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0 '確定
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         'Add By Cheng 2002/03/21
         If PUB_CheckKeyInYYMM(Me.txt1(3)) = -1 Then
            Me.txt1(3).SetFocus
            txt1_GotFocus 3
            Exit Sub
         End If
         If PUB_CheckKeyInYYMM(Me.txt1(4)) = -1 Then
            Me.txt1(4).SetFocus
            txt1_GotFocus 4
            Exit Sub
         End If
         'add by nickc 2005/03/04  加入判斷新舊制
         If Len(txt1(10)) = 0 Then
             s = MsgBox("統計方式不可空白!!", , "USER 輸入錯誤")
             If Len(txt1(10)) = 0 Then txt1(10).SetFocus
             Exit Sub
         End If
         If Len(txt1(3)) = 0 Or Len(txt1(4)) = 0 Then
             s = MsgBox("發文年月區間不可空白!!", , "USER 輸入錯誤")
             If Len(txt1(4)) = 0 Then txt1(4).SetFocus
             If Len(txt1(3)) = 0 Then txt1(3).SetFocus
             Exit Sub
         Else
             If Len(txt1(8)) = 0 Then
                 s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                 txt1(8).SetFocus
                 Exit Sub
             Else
                 If Len(txt1(9)) = 0 Then
                     s = MsgBox("列印順序不可空白!!", , "USER 輸入錯誤")
                     txt1(9).SetFocus
                     Exit Sub
                 Else
                     Screen.MousePointer = vbHourglass
                     Me.Enabled = False
                     ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/17 清除查詢印表記錄檔欄位
                     'If StrTemp99(0) <> txt1(0) Or StrTemp99(1) <> txt1(1) Or StrTemp99(2) <> txt1(2) Or StrTemp99(3) <> txt1(3) Or StrTemp99(4) <> txt1(4) Or StrTemp99(5) <> txt1(5) Or StrTemp99(6) <> txt1(6) Or StrTemp99(7) <> txt1(7) Then
                     'add by nick 2005/02/14 加入新制
                     If txt1(10) = "1" Then
                        pub_QL05 = pub_QL05 & ";" & Label1(11) & "1.舊制" 'Add By Sindy 2010/12/17
                        If Process = False Then
                            Me.Enabled = True
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                     Else
                        pub_QL05 = pub_QL05 & ";" & Label1(11) & "2.新制" 'Add By Sindy 2010/12/17
                        'edit by nickc 2005/04/11  先改即時運算
                        'If ProcessNew = False Then
                        If ProcessNew2 = False Then
                            Me.Enabled = True
                            Screen.MousePointer = vbDefault
                            Exit Sub
                        End If
                     End If
                     '   For i = 0 To 7
                     '       StrTemp99(i) = txt1(i)
                     '   Next i
                     'Else
                     'End If
                     If Val(txt1(8)) = 1 Then
                        pub_QL05 = pub_QL05 & ";" & Label1(5) & "1.螢幕" 'Add By Sindy 2010/12/17
                        Me.Hide
                        frm090701_1.Show
                     Else
                        pub_QL05 = pub_QL05 & ";" & Label1(5) & "2.報表" 'Add By Sindy 2010/12/17
                        PrintData
                     End If
                     Me.Enabled = True
                     Screen.MousePointer = vbDefault
                 End If
             End If
         End If
     End If
Case 1 '離開
     Unload Me
Case Else
End Select
End Sub

Sub PrintData()
'Modify By Cheng 2003/07/30
'strSQL = "select R102001,SUM(r102002),sum(r102003),sum(r102004),sum(r102005),sum(r102006),sum(r102007),sum(r102008),sum(r102009),sum(r102010),sum(r102011),sum(r102012),sum(r102013),sum(r102014),SUM(R102015),SUM(R102016),SUM(R102017),SUM(R102018) from r090701 WHERE ID='" & strUserNum & "' group by r102001"
'Modify By Cheng 2003/08/01
'strSQL = "Select R102001,SUM(r102002),sum(r102003),sum(r102004),sum(r102005),sum(r102006),sum(r102007),sum(r102008),sum(r102009),sum(r102010),sum(r102011),sum(r102012),sum(r102013),sum(r102014),SUM(R102015),SUM(R102016),SUM(R102017),SUM(R102018), ST06 from r090701, Staff WHERE R102001=ST01 And ID='" & strUserNum & "' Group By r102001, ST06 "
strSql = "Select R102001,SUM(r102002),sum(r102003),sum(r102004),sum(r102005),sum(r102006),sum(r102007),sum(r102008),sum(r102009),sum(r102010),sum(r102011),sum(r102012),sum(r102013),sum(r102014),SUM(R102015),SUM(R102016),SUM(R102017),SUM(R102018),SUM(R102019),SUM(R102020), ST06 from r090701, Staff WHERE R102001<>'99999' AND R102001=ST01 And ID='" & strUserNum & "' Group By r102001, ST06 "
strSql = strSql & " Having (nvl(Sum(R102002),0)+nvl(Sum(R102003),0)+nvl(Sum(R102004),0)+nvl(Sum(R102005),0)+nvl(Sum(R102006),0)+nvl(Sum(R102007),0)+nvl(Sum(R102008),0)+nvl(Sum(R102009),0)+nvl(Sum(R102010),0)+nvl(Sum(R102011),0)+nvl(Sum(R102012),0)+nvl(Sum(R102013),0)+nvl(Sum(R102014),0)+nvl(Sum(R102015),0)+nvl(Sum(R102016),0)+nvl(Sum(R102017),0)+nvl(Sum(R102018),0)+nvl(Sum(R102019),0)+nvl(Sum(R102020),0))  > 0 "
Select Case Val(frm090701.txt1(9))
'Modify By Cheng 2003/06/05
'Case 1
'     strSQL = strSQL + " ORDER BY r102001,sum(R102008) "
'Case 2
'     strSQL = strSQL + " ORDER BY R102001,sum(R102009) "
'Case 3
'     strSQL = strSQL + " ORDER BY R102001,sum(R102011) "
'Case 4
'     strSQL = strSQL + " ORDER BY R102001 "
Case 1 '點數
     pub_QL05 = pub_QL05 & ";" & Label1(6) & "1.點數 %" 'Add By Sindy 2010/12/17
     strSql = strSql + " ORDER BY " & SQLSum2("R102008") & " DESC "
Case 2 '件數
     pub_QL05 = pub_QL05 & ";" & Label1(6) & "2.件數 %" 'Add By Sindy 2010/12/17
     strSql = strSql + " ORDER BY " & SQLSum2("R102009") & " DESC "
Case 3 '平均
     pub_QL05 = pub_QL05 & ";" & Label1(6) & "3.平均 %" 'Add By Sindy 2010/12/17
     strSql = strSql + " ORDER BY " & SQLSum2("R102011") & " DESC "
Case 4 '繪圖人員
     pub_QL05 = pub_QL05 & ";" & Label1(6) & "4.繪圖人員" 'Add By Sindy 2010/12/17
     strSql = strSql + " ORDER BY ST06, R102001 "
Case Else
End Select
CheckOC
Page = 1: m_Cnt = 0
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/17
        m_Cnt = .RecordCount
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 19
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/17
    End If
End With
CheckOC
ShowLine
PrintEnd
ShowLine
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintTitle() '列印抬頭

GetPleft
iPrint = 0
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "繪圖人員達成情形統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
'Modify by Morgan 2011/2/16 修正百年問題
'Printer.Print "統計年月：" & Mid(txt1(3), 1, 2) & "/" & Mid(txt1(3), 3, 2) & "－" & Mid(txt1(4), 1, 2) & "/" & Mid(txt1(4), 3, 2)
Printer.Print "統計年月：" & (Val(txt1(3)) \ 100) & "/" & Right(txt1(3), 2) & "－" & (Val(txt1(4)) \ 100) & "/" & Right(txt1(4), 2)
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(16000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.Font.Underline = True
Printer.Font.Size = 9
Printer.CurrentX = (PLeft(2) + (Printer.TextWidth("件數") / 2) - (Printer.TextWidth("目標") / 2))
Printer.CurrentY = iPrint
Printer.Print "目標"
Printer.CurrentX = (PLeft(5) + (Printer.TextWidth("件數") / 2) - (Printer.TextWidth("目標達成") / 2))
Printer.CurrentY = iPrint
Printer.Print "目標達成"
Printer.CurrentX = ((((PLeft(9) + Printer.TextWidth("張數")) - PLeft(8)) / 2) + PLeft(8)) - (Printer.TextWidth("達成率 %") / 2)
Printer.CurrentY = iPrint
Printer.Print "達成率 %"
'edit by nickc 2005/04/13
If txt1(10) = "1" Then
      Printer.CurrentX = ((((PLeft(12) + Printer.TextWidth("件數")) - PLeft(11)) / 2) + PLeft(11)) - (Printer.TextWidth("其他新案") / 2)
      Printer.CurrentY = iPrint
      Printer.Print "其他新案"
      Printer.CurrentX = ((((PLeft(14) + Printer.TextWidth("件數")) - PLeft(13)) / 2) + PLeft(13)) - (Printer.TextWidth("其他舊案") / 2)
      Printer.CurrentY = iPrint
      Printer.Print "其他舊案"
Else
      Printer.CurrentX = ((((PLeft(12)) - PLeft(11)) / 2) + PLeft(11)) - (Printer.TextWidth("提供圖檔") / 2)
      Printer.CurrentY = iPrint
      Printer.Print "提供圖檔"
      Printer.CurrentX = ((((PLeft(13)) - PLeft(12)) / 2) + PLeft(12)) - (Printer.TextWidth("轉換案") / 2)
      Printer.CurrentY = iPrint
      Printer.Print "轉換案"
      Printer.CurrentX = ((((PLeft(15)) - PLeft(13)) / 2) + PLeft(13)) - (Printer.TextWidth("其他新舊案") / 2)
      Printer.CurrentY = iPrint
      Printer.Print "其他新舊案"
End If
'Add By Cheng 2003/08/01
Printer.CurrentX = ((((PLeft(17) + Printer.TextWidth("墨圖件數")) - PLeft(15)) / 2) + PLeft(15)) - (Printer.TextWidth("完　　成") / 2)
Printer.CurrentY = iPrint
Printer.Print "完　　成"
Printer.CurrentX = ((((PLeft(19) + Printer.TextWidth("墨圖件數")) - PLeft(18)) / 2) + PLeft(18)) - (Printer.TextWidth("逾　　時") / 2)
Printer.CurrentY = iPrint
Printer.Print "逾　　時"
Printer.Font.Underline = False
iPrint = iPrint + 300
Printer.CurrentX = ((((PLeft(12)) - PLeft(11)) / 2) + PLeft(11)) - (Printer.TextWidth("(0.6)") / 2)
Printer.CurrentY = iPrint
Printer.Print "(0.6)"
Printer.CurrentX = ((((PLeft(13)) - PLeft(12)) / 2) + PLeft(12)) - (Printer.TextWidth("(0.4)") / 2)
Printer.CurrentY = iPrint
Printer.Print "(0.4)"
iPrint = iPrint + 100
Printer.Line (PLeft(1), iPrint + 150)-(PLeft(4) - 50, iPrint + 150)
Printer.Line (PLeft(4), iPrint + 150)-(PLeft(7) - 50, iPrint + 150)
Printer.Line (PLeft(7), iPrint + 150)-(PLeft(11) - 50, iPrint + 150)
'edit by nickc 2005/04/18
'Printer.Line (PLeft(11), iPrint + 150)-(PLeft(13) - 50, iPrint + 150)
Printer.Line (PLeft(11), iPrint + 150)-(PLeft(12) - 50, iPrint + 150)
Printer.Line (PLeft(12), iPrint + 150)-(PLeft(13) - 50, iPrint + 150)
Printer.Line (PLeft(13), iPrint + 150)-(PLeft(15) - 50, iPrint + 150)
'Add By Cheng 2003/08/01
Printer.Line (PLeft(15), iPrint + 150)-(PLeft(18) - 50, iPrint + 150)
Printer.Line (PLeft(18), iPrint + 150)-(PLeft(19) + Printer.TextWidth("逾　　時"), iPrint + 150)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "繪圖人員"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "張數"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "張數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "張數"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "平均"
'edit by nickc 2004/05/13
If txt1(10) = "1" Then
      Printer.CurrentX = PLeft(11)
      Printer.CurrentY = iPrint
      Printer.Print "點數"
      Printer.CurrentX = PLeft(12)
      Printer.CurrentY = iPrint
      Printer.Print "件數"
      Printer.CurrentX = PLeft(13)
      Printer.CurrentY = iPrint
      Printer.Print "點數"
      Printer.CurrentX = PLeft(14)
      Printer.CurrentY = iPrint
      Printer.Print "件數"
Else
      Printer.CurrentX = ((((PLeft(12)) - PLeft(11)) / 2) + PLeft(11)) - (Printer.TextWidth("件數") / 2)
      Printer.CurrentY = iPrint
      Printer.Print "件數"
      Printer.CurrentX = ((((PLeft(13)) - PLeft(12)) / 2) + PLeft(12)) - (Printer.TextWidth("件數") / 2)
      Printer.CurrentY = iPrint
      Printer.Print "件數"
      Printer.CurrentX = ((((PLeft(15)) - PLeft(13)) / 2) + PLeft(13)) - (Printer.TextWidth("件數") / 2)
      Printer.CurrentY = iPrint
      Printer.Print "件數"
End If

Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "草圖件數"
'Printer.CurrentX = PLeft(16)
'Printer.CurrentY = iPrint
'Printer.Print "ACAD件數"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "墨圖件數"
'Add By Cheng 2003/08/01
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "草圖件數"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "墨圖件數"
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(16000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
End Sub

Sub PrintDatil() '列印資料

Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
'Modify By Cheng 2003/06/05
'取得員工姓名
'Printer.Print strTemp(0)
Printer.Print GetStaffName(strTemp(0), True)
Printer.CurrentX = PLeft(1) + 500 - Printer.TextWidth(Format(strTemp(1), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(1), "####0.00")
Printer.CurrentX = PLeft(2) + 500 - Printer.TextWidth(Format(strTemp(2), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(2), "####0.00")
Printer.CurrentX = PLeft(3) + 500 - Printer.TextWidth(Format(strTemp(3), "####0.0"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(3), "####0.0")
Printer.CurrentX = PLeft(4) + 500 - Printer.TextWidth(Format(strTemp(4), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(4), "####0.00")
Printer.CurrentX = PLeft(5) + 500 - Printer.TextWidth(Format(strTemp(5), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(5), "####0.00")
Printer.CurrentX = PLeft(6) + 500 - Printer.TextWidth(Format(strTemp(6), "####0.0"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(6), "####0.0")
Printer.CurrentX = PLeft(7) + 500 - Printer.TextWidth(Format(strTemp(7), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(7), "####0.00")
'Printer.CurrentX = PLeft(8) + 500 - Printer.TextWidth(Format(strTemp(8), "####0"))
Printer.CurrentX = PLeft(8) + 500 - Printer.TextWidth(Format(strTemp(8), "####0.00"))
Printer.CurrentY = iPrint
'Printer.Print Format(strTemp(8), "####0")
Printer.Print Format(strTemp(8), "####0.00")
'Printer.CurrentX = PLeft(9) + 500 - Printer.TextWidth(Format(strTemp(9), "####0"))
Printer.CurrentX = PLeft(9) + 500 - Printer.TextWidth(Format(strTemp(9), "####0.00"))
Printer.CurrentY = iPrint
'Printer.Print Format(strTemp(9), "####0")
Printer.Print Format(strTemp(9), "####0.00")
Printer.CurrentX = PLeft(10) + 500 - Printer.TextWidth(Format(strTemp(10), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(10), "####0.00")
Printer.CurrentX = PLeft(11) + 500 - Printer.TextWidth(Format(strTemp(11), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(11), "####0.00")
Printer.CurrentX = PLeft(12) + 500 - Printer.TextWidth(Format(strTemp(12), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(12), "####0.00")
Printer.CurrentX = PLeft(13) + 500 - Printer.TextWidth(Format(strTemp(13), "####0.00")) + IIf(txt1(10) = "2", 500, 0)
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(13), "####0.00")
'edit by nickc 2005/04/13
If txt1(10) = "1" Then
      Printer.CurrentX = PLeft(14) + 500 - Printer.TextWidth(Format(strTemp(14), "####0.00"))
      Printer.CurrentY = iPrint
      Printer.Print Format(strTemp(14), "####0.00")
End If
Printer.CurrentX = PLeft(15) + 500 - Printer.TextWidth(Format(strTemp(15), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(15), "####0.00")
'Printer.CurrentX = PLeft(16) + 500 - Printer.TextWidth(Format(strTemp(16), "####0"))
'Printer.CurrentY = iPrint
'Printer.Print Format(strTemp(16), "####0")
Printer.CurrentX = PLeft(17) + 500 - Printer.TextWidth(Format(strTemp(17), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(17), "####0.00")
'Add By Cheng 2003/08/01
Printer.CurrentX = PLeft(18) + 500 - Printer.TextWidth(Format(strTemp(18), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(18), "####0.00")
Printer.CurrentX = PLeft(19) + 500 - Printer.TextWidth(Format(strTemp(19), "####0.00"))
Printer.CurrentY = iPrint
Printer.Print Format(strTemp(19), "####0.00")
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft '定陣列

PLeft(0) = 0
For i = 1 To 14
    PLeft(i) = 1000 + (i - 1) * 800
Next i
PLeft(15) = PLeft(14) + 900
'PLeft(16) = PLeft(14) + 2000
PLeft(17) = PLeft(14) + 1800
PLeft(18) = PLeft(14) + 2700
PLeft(19) = PLeft(14) + 3600
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(16000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
End Sub

Sub PrintEnd()
'列印結尾
'Modify By Cheng 2003/08/01
'strSQL = "SELECT '合  計',SUM(R102002),SUM(R102003),SUM(R102004),SUM(R102005),SUM(R102006),SUM(R102007),SUM(R102008),SUM(R102009),SUM(R102010),SUM(R102011),SUM(R102012),SUM(R102013),sum(r102014),SUM(R102015),SUM(R102016),SUM(R102017),SUM(R102018) FROM r090701 WHERE ID='" & strUserNum & "' "
strSql = "SELECT '合  計',SUM(R102002),SUM(R102003),SUM(R102004),SUM(R102005),SUM(R102006),SUM(R102007),SUM(R102008),SUM(R102009),SUM(R102010),SUM(R102011),SUM(R102012),SUM(R102013),sum(r102014),SUM(R102015),SUM(R102016),SUM(R102017),SUM(R102018),SUM(R102019),SUM(R102020) FROM r090701 WHERE R102001<>'99999' AND ID='" & strUserNum & "' "
'strSQL = "select R102001,SUM(r102002),sum(r102003),sum(r102004),sum(r102005),sum(r102006),sum(r102007),sum(r102008),sum(r102009),sum(r102010),sum(r102011),sum(r102012),sum(r102013),sum(r102014),SUM(R102015),SUM(R102016),SUM(R102017),SUM(R102018) from r090701 WHERE ID='" & strUserNum & "' group by r102001"
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 19
                StrTemp7(i) = CheckStr(.Fields(i))
                If Val(StrTemp7(i)) = 0 And i > 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            StrTemp7(10) = str((Val(StrTemp7(8)) + Val(StrTemp7(7))) / 2)
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1) + 500 - Printer.TextWidth(Format(StrTemp7(1), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(1), "####0.00")
            Printer.CurrentX = PLeft(2) + 500 - Printer.TextWidth(Format(StrTemp7(2), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(2), "####0.00")
            Printer.CurrentX = PLeft(3) + 500 - Printer.TextWidth(Format(StrTemp7(3), "####0.0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(3), "####0.0")
            Printer.CurrentX = PLeft(4) + 500 - Printer.TextWidth(Format(StrTemp7(4), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(4), "####0.00")
            Printer.CurrentX = PLeft(5) + 500 - Printer.TextWidth(Format(StrTemp7(5), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(5), "####0.00")
            Printer.CurrentX = PLeft(6) + 500 - Printer.TextWidth(Format(StrTemp7(6), "####0.0"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(6), "####0.0")
            Printer.CurrentX = PLeft(7) + 500 - Printer.TextWidth(Format(Val(StrTemp7(7)) / m_Cnt, "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(Val(StrTemp7(7) / m_Cnt), "####0.00")
'            Printer.CurrentX = PLeft(8) + 500 - Printer.TextWidth(Format(StrTemp7(8), "####0"))
            Printer.CurrentX = PLeft(8) + 500 - Printer.TextWidth(Format(Val(StrTemp7(8)) / m_Cnt, "####0.00"))
            Printer.CurrentY = iPrint
'            Printer.Print Format(StrTemp7(8), "####0")
            Printer.Print Format(Val(StrTemp7(8)) / m_Cnt, "####0.00")
'            Printer.CurrentX = PLeft(9) + 500 - Printer.TextWidth(Format(StrTemp7(9), "####0"))
            Printer.CurrentX = PLeft(9) + 500 - Printer.TextWidth(Format(Val(StrTemp7(9)) / m_Cnt, "####0.00"))
            Printer.CurrentY = iPrint
'            Printer.Print Format(StrTemp7(9), "####0")
            Printer.Print Format(Val(StrTemp7(9)) / m_Cnt, "####0.00")
            Printer.CurrentX = PLeft(10) + 500 - Printer.TextWidth(Format(Val(StrTemp7(10)) / m_Cnt, "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(Val(StrTemp7(10)) / m_Cnt, "####0.00")
            Printer.CurrentX = PLeft(11) + 500 - Printer.TextWidth(Format(StrTemp7(11), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(11), "####0.00")
            Printer.CurrentX = PLeft(12) + 500 - Printer.TextWidth(Format(StrTemp7(12), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(12), "####0.00")
            Printer.CurrentX = PLeft(13) + 500 - Printer.TextWidth(Format(StrTemp7(13), "####0.00")) + IIf(txt1(10) = "2", 500, 0)
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(13), "####0.00")
            'edit by nickc 2005/04/13
            If txt1(10) = "1" Then
                  Printer.CurrentX = PLeft(14) + 500 - Printer.TextWidth(Format(StrTemp7(14), "####0.00"))
                  Printer.CurrentY = iPrint
                  Printer.Print Format(StrTemp7(14), "####0.00")
            End If
            Printer.CurrentX = PLeft(15) + 500 - Printer.TextWidth(Format(StrTemp7(15), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(15), "####0.00")
'            Printer.CurrentX = PLeft(16) + 500 - Printer.TextWidth(Format(StrTemp7(16), "####0"))
'            Printer.CurrentY = iPrint
'            Printer.Print Format(StrTemp7(16), "####0")
            Printer.CurrentX = PLeft(17) + 500 - Printer.TextWidth(Format(StrTemp7(17), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(17), "####0.00")
            'Add By Cheng 2003/08/01
            Printer.CurrentX = PLeft(18) + 500 - Printer.TextWidth(Format(StrTemp7(18), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(18), "####0.00")
            Printer.CurrentX = PLeft(19) + 500 - Printer.TextWidth(Format(StrTemp7(19), "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(StrTemp7(19), "####0.00")
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Function Process() As Boolean
On Error GoTo ErrorHandler
Process = False
cnnConnection.Execute "DELETE FROM r090701 WHERE ID='" & strUserNum & "' "
StrSQL7 = "": strSQL8 = "": strSQL9 = ""
'Move By Cheng 2003/06/05
strSQL71 = "": strSQL72 = "": strSQL73 = "": strSQL74 = ""
'系統類別
If Len(txt1(0)) <> 0 Then
   StrSQL7 = StrSQL7 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   'Add By Cheng 2003/06/05
   strSQL71 = strSQL71 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL73 = strSQL73 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL74 = strSQL74 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   'Add By Cheng 2004/02/17
   strSQL8 = strSQL8 + " And CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   'End
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/17
End If

StrSQL6 = ""
strSQL1 = ""
'申請國家(起)
If Len(txt1(1).Text) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & txt1(1) & "' "
    strSQL2 = strSQL2 + " AND TM10>='" & txt1(1) & "' "
    StrSQL3 = StrSQL3 + " AND LC15>='" & txt1(1) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09>='" & txt1(1) & "' "
End If
'申請國家(迄)
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & txt1(2) & "' "
    strSQL2 = strSQL2 + " AND TM10<='" & txt1(2) & "' "
    StrSQL3 = StrSQL3 + " AND LC15<='" & txt1(2) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09<='" & txt1(2) & "' "
End If
If Len(txt1(1).Text) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/17
End If
'Modify By Cheng 2004/02/17
'StrSQL7 = StrSQL7 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
'edit by nickc 2005/05/13
'StrSQL7 = StrSQL7 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " And CP57 Is Null "
StrSQL7 = StrSQL7 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " "
'End
'Add By Cheng 2003/06/05
'Modify By Cheng 2004/02/17
'strSQL71 = strSQL71 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31"))
'edit by nickc 2005/05/13
'strSQL71 = strSQL71 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " And CP57 Is Null "
strSQL71 = strSQL71 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & "  "
'End
'edit by nickc 2005/05/13
'strSQL8 = strSQL8 + " AND ((CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " And CP57 Is Null ) Or (CP57>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP57<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " And CP27 Is Null) Or (CP27 Is Null And CP57 Is Null) ) "
strSQL8 = strSQL8 + " AND ((CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " ) Or (CP57>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP57<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " And CP27 Is Null) Or (CP27 Is Null And CP57 Is Null) ) "
strSQL9 = strSQL9 + " AND (SH01>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND SH01<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " ) "
pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/17
'strSQL71 = ""
'strSQL72 = ""
'繪圖人員
If Len(txt1(5)) <> 0 Then
    StrSQL7 = StrSQL7 + " AND EP13='" & txt1(5) & "' "
    strSQL71 = strSQL71 + " AND EP13='" & txt1(5) & "' "
    strSQL72 = strSQL72 + " AND PE01='" & txt1(5) & "' "
    strSQL73 = strSQL73 + " AND EP13='" & txt1(5) & "' "
    strSQL74 = strSQL74 + " AND EP13='" & txt1(5) & "' "
    strSQL8 = strSQL8 + " AND EP13='" & txt1(5) & "' "
    strSQL9 = strSQL9 + " AND SH02='" & txt1(5) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & lbl1(0) 'Add By Sindy 2010/12/17
End If
'所別(起)
If Len(txt1(6)) <> 0 Then
    StrSQL7 = StrSQL7 + " AND ST06>='" & txt1(6) & "' "
    'Add By Cheng 2003/06/05
    strSQL71 = strSQL71 + " AND S2.ST06>='" & txt1(6) & "' "
    strSQL72 = strSQL72 + " AND ST06>='" & txt1(6) & "' "
    strSQL73 = strSQL73 + " AND S2.ST06>='" & txt1(6) & "' "
    strSQL74 = strSQL74 + " AND ST06>='" & txt1(6) & "' "
    strSQL8 = strSQL8 + " AND ST06>='" & txt1(6) & "' "
    strSQL9 = strSQL9 + " AND ST06>='" & txt1(6) & "' "
End If
'所別(迄)
If Len(txt1(7)) <> 0 Then
    StrSQL7 = StrSQL7 + " AND ST06<='" & txt1(7) & "' "
    'Add By Cheng 2003/06/05
    strSQL71 = strSQL71 + " AND S2.ST06<='" & txt1(7) & "' "
    strSQL72 = strSQL72 + " AND ST06<='" & txt1(7) & "' "
    strSQL73 = strSQL73 + " AND S2.ST06<='" & txt1(7) & "' "
    strSQL74 = strSQL74 + " AND ST06<='" & txt1(7) & "' "
    strSQL8 = strSQL8 + " AND ST06<='" & txt1(7) & "' "
    strSQL9 = strSQL9 + " AND ST06<='" & txt1(7) & "' "
End If
If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & "-" & txt1(7) & Label1(7) 'Add By Sindy 2010/12/17
End If
'Modify By Cheng 2003/07/31
'件數不計算墨圖不計件, 點數張數則都要
'StrSQL7 = StrSQL7 + " and EP20 IS NULL  "
''Add By Cheng 2003/06/05
'strSQL71 = strSQL71 + " and EP20 IS NULL  "
'Add By Cheng 2002/04/16
StrSQL7 = StrSQL7 + " and ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND ST04='1' ", " ")
strSQL9 = strSQL9 + " and ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND ST04='1' ", " ")

'add by nickc 205/04/11 離職不秀
    StrSQL7 = StrSQL7 + " AND ST04='1' "
    strSQL71 = strSQL71 + " AND S2.ST04='1' "
    strSQL72 = strSQL72 + " AND ST04='1' "
    strSQL73 = strSQL73 + " AND S2.ST04='1' "
    strSQL74 = strSQL74 + " AND ST04='1' "
    strSQL8 = strSQL8 + " AND ST04='1' "
    strSQL9 = strSQL9 + " AND ST04='1' "
    
CheckOC
'目標
'Modify By Cheng 2002/04/16
'strSQL = "INSERT INTO R090701 (r102001,r102002,r102003,r102004,id) select pe01,sum(pe11),sum(pe09),SUM(PE10),'" & strUserNum & "' from performance wheRE pe03>=" & Val(txt1(3)) + 191100 & " and pe03<=" & Val(txt1(4)) + 191100 & " and pe02='P' " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & " group by pe01,'" & strUserNum & "' "
strSql = "INSERT INTO R090701 (r102001,r102002,r102003,r102004,id) select pe01,sum(pe11),sum(pe09),SUM(PE10),'" & strUserNum & "' From Performance, STAFF Where PE01=ST01(+) AND PE03>=" & Val(txt1(3)) + 191100 & " and PE03<=" & Val(txt1(4)) + 191100 & " And PE02='P' " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & " AND ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND ST04='1' ", " ") & strSQL72 & " Group By PE01,'" & strUserNum & "' "
cnnConnection.Execute strSql
'目標達成
'Modify By Cheng 2002/04/16
'strSQL = " (SELECT EP13,SUM(CP18),COUNT(CP27),sum((ep16+ep19)/2),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") GROUP BY EP13,'" & strUserNum & "' )"
'Modify By Cheng 2003/07/31
'strSQL = " (SELECT EP13,SUM(CP18),COUNT(CP27),sum((ep16+ep19)/2),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' )"
'Modify By Cheng 2004/03/12
'繪圖在統計目標達成點數及件數時, 若墨圖不計件的不統計(而承辦人是不計件的件數不計但點數要計)
'strSQL = " (SELECT EP13, SUM(CP18), Sum(Decode(CP27, Null, 0, Decode(EP29, Null, 1, 0))), 0,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' )"
'strSQL = " (SELECT EP13, SUM(Decode(EP29, Null, CP18, 0)), Sum(Decode(EP29, Null, 1, 0)), 0,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' )"
strSql = " (SELECT EP13, SUM(CP18), Sum(1), 0,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09 AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") And EP29 Is Null " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' )"
'End
cnnConnection.Execute "insert into r090701 (r102001,r102005,r102006,r102007,id) " & strSql
'Add By Cheng 2004/03/25
'支援記錄計入目標達成件數
strSql = "SELECT SH02, 0, Sum(Decode(SH06, 'CFP', Nvl(SH05,0)/8, Nvl(SH05,0)/4)), 0,'" & strUserNum & "' FROM SupportHour,STAFF,PATENT WHERE SH02=ST01(+) And SH06=PA01(+) And SH07=PA02(+) And SH08=PA03(+) And SH09=PA04(+) And SH11='V' " & strSQL1 & strSQL9 & " GROUP BY SH02,'" & strUserNum & "' "
cnnConnection.Execute "insert into r090701 (r102001,r102005,r102006,r102007,id) " & strSql
'End
strSql = " (SELECT EP13, 0, 0, sum(nvl(ep16,0))/2,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL73 & _
                " AND EP15>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP15<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " AND ( EP18 Is Null Or (EP18<" & Val(ChangeTStringToWString(txt1(3) & "01")) & " Or EP18>" & Val(ChangeTStringToWString(txt1(4) & "31")) & ")) " & " GROUP BY EP13,'" & strUserNum & "' )"
cnnConnection.Execute "insert into r090701 (r102001,r102005,r102006,r102007,id) " & strSql
strSql = " (SELECT EP13, 0, 0, sum(nvl(ep19,0))/2,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL73 & _
                " AND EP18>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP18<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " AND ( EP15 Is Null Or (EP15<" & Val(ChangeTStringToWString(txt1(3) & "01")) & " Or EP15>" & Val(ChangeTStringToWString(txt1(4) & "31")) & ")) " & " GROUP BY EP13,'" & strUserNum & "' )"
cnnConnection.Execute "insert into r090701 (r102001,r102005,r102006,r102007,id) " & strSql
strSql = " (SELECT EP13, 0, 0, Sum((ep16+ep19)/2),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL73 & _
                " AND ((EP15>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP15<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & ") And (EP18>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " And EP18<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & ")) " & " GROUP BY EP13,'" & strUserNum & "' )"
cnnConnection.Execute "insert into r090701 (r102001,r102005,r102006,r102007,id) " & strSql
'達成率
strSql = "(SELECT R102001," & _
         "ROUND(DECODE(SUM(DECODE(R102002,0,0,NULL,0,R102002)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102002,0,0,NULL,0,R102002)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102007,0,0,NULL,0,R102007))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2)," & _
         "(ROUND(DECODE(SUM(DECODE(R102002,0,0,NULL,0,R102002)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102002,0,0,NULL,0,R102002)))*100),2) + ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2))/2, " & _
         "'" & strUserNum & "' " & _
         "FROM R090701 WHERE ID='" & strUserNum & "' GROUP BY R102001 )"
cnnConnection.Execute "INSERT INTO R090701 (R102001,R102008,R102009,R102010,R102011,ID) " & strSql
'Modify By Cheng 2003/07/31
'Begin
''其他新案
''Modify By Cheng 2002/04/16
''strSQL = " (SELECT EP13,SUM(CP18),COUNT(CP27),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=ST01(+) " & strSQL1 & StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") GROUP BY EP13,'" & strUserNum & "' )"
''Modify By Cheng 2003/07/31
''strSQL = " (SELECT EP13,SUM(CP18),COUNT(CP27),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' )"
'strSQL = " (SELECT EP13,SUM(CP18), Sum(Decode(CP27, Null, 0, Decode(EP29, Null, 1, 0))),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' )"
'cnnConnection.Execute "insert into r090701 (r102001,r102012,r102013,id) " & strSQL
'End
'                                1    2         3         4    5         6           7                       8                             9                               10                                           11                                                            12131415161718
                'Modify By Cheng 2003/06/05
                '目標資料前面已取得
'                strSQL = "SELECT EP13,PE11 AS A,PE09 AS B,pe10,SUM(CP18),COUNT(CP27),(sum(ep16)+sum(ep19))/2,((sum(cp18))/(pe11))*100 as g,((count(cp27))/(pe09))*100 as h,(((sum(ep16)+sum(ep19)) /2)/(sum(cp18)))*100,((((sum(cp18))/(pe11))*100) + (((count(cp27))/(pe09))*100))/2,0,0,0,0,0,0,0,'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,PERFORMANCE,STAFF,engineerprogress WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND EP13=PE01 AND CP01=PE02(+) AND SUBSTR(CP27,1,6)=PE03 and CP09=ep02(+)  AND CP01 NOT IN ('FCP','CFP') " & strSQL1 & StrSQL7 & " GROUP BY EP13,pe11,pe09,pe10,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,SUM(CP18),COUNT(CP27),0,0,0,0,0,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP01='P' AND (cp10 between '101' and '105') " & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,SUM(CP18),COUNT(CP27),0,0,0,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP01='P' AND (CP10 NOT BETWEEN '101' AND '105') " & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),0,0,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND instr(EP26,'ACAD')=0 AND EP15 IS NOT NULL " & strSQL1 & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),0,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND instr(EP26,'ACAD')>0 AND EP15 IS NOT NULL " & strSQL1 & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND EP18 IS NOT NULL " & strSQL1 & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'Modify By Cheng 2003/07/31
'                strSQL = "SELECT EP13, 0 AS A, 0 AS B, 0 As pe10, SUM(CP18),COUNT(CP27),(sum(ep16)+sum(ep19))/2,((sum(cp18))/(pe11))*100 as g,((count(cp27))/(pe09))*100 as h,(((sum(ep16)+sum(ep19)) /2)/(sum(cp18)))*100,((((sum(cp18))/(pe11))*100) + (((count(cp27))/(pe09))*100))/2,0,0,0,0,0,0,0,'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,PERFORMANCE,STAFF,engineerprogress WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND EP13=PE01 AND CP01=PE02(+) AND SUBSTR(CP27,1,6)=PE03 and CP09=ep02(+)  AND CP01 NOT IN ('FCP','CFP') " & strSQL1 & StrSQL7 & " GROUP BY EP13,pe11,pe09,pe10,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,SUM(CP18),COUNT(CP27),0,0,0,0,0,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP01='P' AND (cp10 between '101' and '105') " & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,SUM(CP18),COUNT(CP27),0,0,0,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP01='P' AND (CP10 NOT BETWEEN '101' AND '105') " & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),0,0,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND instr(EP26,'ACAD')=0 AND EP15 IS NOT NULL " & strSQL1 & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),0,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND instr(EP26,'ACAD')>0 AND EP15 IS NOT NULL " & strSQL1 & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND EP18 IS NOT NULL " & strSQL1 & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'Modify By Cheng 2003/07/31
'Begin
'                strSQL = "SELECT EP13, 0 AS A, 0 AS B, 0 As pe10, SUM(CP18), Sum(Decode(CP27, Null, 0, Decode(EP29, Null, 1, 0))),(sum(ep16)+sum(ep19))/2,((sum(cp18))/(pe11))*100 as g,((count(cp27))/(pe09))*100 as h,(((sum(ep16)+sum(ep19)) /2)/(sum(cp18)))*100,((((sum(cp18))/(pe11))*100) + (((count(cp27))/(pe09))*100))/2,0,0,0,0,0,0,0,'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,PERFORMANCE,STAFF,engineerprogress WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13=ST01(+) AND EP13=PE01 AND CP01=PE02(+) AND SUBSTR(CP27,1,6)=PE03 and CP09=ep02(+)  AND CP01<>'P' " & strSQL1 & StrSQL7 & " GROUP BY EP13,pe11,pe09,pe10,'" & strUserNum & "' "
'End
'其他新案
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(CP27, Null, 0, Decode(EP29, Null, 1, 0))),0,0,0,0,0,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP01<>'P' AND (CP10 Between '101' And '105') And EP15 Is Not Null " & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'Modify By Cheng 2004/02/17
'strSQL = "SELECT EP13,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(CP27, Null, 0, Decode(EP29, Null, 1, 0))),0,0,0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP01<>'P' AND (CP10 Between '101' And '105') And EP15 Is Not Null " & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'Modify By Cheng 2004/03/12
'不考慮發文日及取消收文日的問題
'strSQL = "SELECT EP13,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(EP20, Null, 1, 0)),0,0,0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP01 In ('CFP') And EP16>0 And substr(EP15,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP15,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL8 & " GROUP BY EP13,'" & strUserNum & "' "
strSql = "SELECT EP13,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(EP20, Null, 1, 0)),0,0,0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP') And EP16>0 And substr(EP15,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP15,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
'End
'End
'其他舊案
'Modify By Cheng 2004/02/17
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(CP27, Null, 0, Decode(EP29, Null, 1, 0))),0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP01<>'P' AND (CP10 Not Between '101' AND '105') And EP15 Is Null And EP18 Is Null " & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'Modify By Cheng 2004/03/12
'不考慮發文日及取消收文日的問題
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(EP29, Null, 1, 0)),0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP01 In ('CFP') And ((EP16 Is Null Or EP16<=0) And (EP19 Is Null or EP19<=0)) And substr(EP18,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP18,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL8 & " GROUP BY EP13,'" & strUserNum & "' "
'Modify by Morgan 2004/5/12
'不必判斷墨圖=0
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(EP29, Null, 1, 0)),0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP') And ((EP16 Is Null Or EP16<=0) And (EP19 Is Null or EP19<=0)) And substr(EP18,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP18,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(EP29, Null, 1, 0)),0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP') And (EP16 Is Null Or EP16<=0) And substr(EP18,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP18,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
'End
'草圖件數(不考慮發文日及取消收文日的問題)
'Modify By Cheng 2004/02/17
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND instr(EP26,'ACAD')=0 AND EP15 IS NOT NULL And EP20 Is Null " & strSQL1 & strSQL74 & _
'                " AND EP15>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP15<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP57 IS NULL And (EP16 Is Not Null And EP16 > 0 ) " & strSQL1 & strSQL74 & _
'                " AND EP15>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP15<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " GROUP BY EP13,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) And (EP16 Is Not Null And EP16 > 0 ) " & strSQL1 & strSQL74 & _
                " AND EP15>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP15<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " GROUP BY EP13,'" & strUserNum & "' "
'End
'Modify By Cheng 2003/07/31
'Begin
''ACDC件數
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),0,'" & strUserNum & "' FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND instr(EP26,'ACAD')>0 AND EP15 IS NOT NULL " & strSQL1 & StrSQL7 & " GROUP BY EP13,'" & strUserNum & "' "
'End
'墨圖件數(不考慮發文日及取消收文日的問題)
'Modify By Cheng 2004/02/17
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND EP18 IS NOT NULL And EP29 Is Null " & strSQL1 & strSQL74 & _
'                " AND EP18>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP18<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) AND CP57 Is NULL " & strSQL1 & strSQL74 & _
'                " AND EP18>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP18<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " GROUP BY EP13,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,COUNT(*),'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) " & strSQL1 & strSQL74 & _
                " AND EP18>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP18<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " GROUP BY EP13,'" & strUserNum & "' "
'End
cnnConnection.Execute " INSERT INTO r090701 " & strSql
'Add By Cheng 2004/03/25
'支援記錄算入草圖及墨圖件數
strSql = "SELECT SH02,0,0,0,0,0,0,0,0,0,0,0,0,0,0,Sum(Decode(SH06, 'CFP', Nvl(SH05,0)/8, Nvl(SH05,0)/4)),0,0,'" & strUserNum & "',0,0 FROM SupportHour,STAFF,PATENT WHERE SH02=ST01(+) And SH06=PA01(+) And SH07=PA02(+) And SH08=PA03(+) And SH09=PA04(+) And SH11='V' " & strSQL1 & strSQL9 & " GROUP BY SH02,'" & strUserNum & "' "
cnnConnection.Execute " INSERT INTO r090701 " & strSql
strSql = "SELECT SH02,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,Sum(Decode(SH06, 'CFP', Nvl(SH05,0)/8, Nvl(SH05,0)/4)),'" & strUserNum & "',0,0 FROM SupportHour,STAFF,PATENT WHERE SH02=ST01(+) And SH06=PA01(+) And SH07=PA02(+) And SH08=PA03(+) And SH09=PA04(+) And SH11='V' " & strSQL1 & strSQL9 & " GROUP BY SH02,'" & strUserNum & "' "
cnnConnection.Execute " INSERT INTO r090701 " & strSql
'End
'Add By Cheng 2003/07/31
'逾時草圖墨圖件數
CompOverTime
'Add By Cheng 2003/07/30
Process = True
Exit Function
ErrorHandler:
    MsgBox Err.Description
End Function

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_Load()
m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
MoveFormToCenter Me
txt1(0) = Systemkind_g
For i = 0 To 7
    StrTemp99(i) = ""
Next i
Me.txt1(3).Text = Left(strSrvDate(1), 4) - 1911 & Format(Mid(strSrvDate(1), 5, 2), "00")
Me.txt1(4).Text = Left(strSrvDate(1), 4) - 1911 & Format(Mid(strSrvDate(1), 5, 2), "00")
txt1(8) = "1"
txt1(9) = "4"
'add by nickc 2005/03/04 加入新舊制預設值
If strSrvDate(1) >= "20050315" Then
   txt1(10).Text = "2"
   Label1(1).Visible = False
   txt1(1).Visible = False
   txt1(2).Visible = False
   Line1.Visible = False
   txt1(0).Visible = False
   Label1(0).Visible = False
Else
   txt1(10).Text = "1"
End If
'Add By Cheng 2003/07/30
'若為個人
If ProState = "1" Then
    GetPersonalData
    '93.4.7 cancel by sonia
    'Me.txt1(3).Enabled = False
    'Me.txt1(4).Enabled = False
    '93.4.7 end
    Me.txt1(3).Text = Left(strSrvDate(1), 4) - 1911 & Format(Mid(strSrvDate(1), 5, 2), "00")
    Me.txt1(4).Text = Left(strSrvDate(1), 4) - 1911 & Format(Mid(strSrvDate(1), 5, 2), "00")
    Me.txt1(6).Enabled = False
    Me.txt1(7).Enabled = False
    Me.txt1(5).Enabled = False
End If
End Sub

Private Sub GetPersonalData()
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
   
   StrSQLa = "SELECT * FROM STAFF WHERE ST01='" & strUserNum & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      Me.txt1(6).Text = "" & rsA("ST06").Value
      Me.txt1(7).Text = "" & rsA("ST06").Value
      Me.txt1(5).Text = "" & rsA("ST01").Value
      Me.lbl1(0).Caption = "" & rsA("ST02").Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090701 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    'Add By Cheng 2003/06/05
    Select Case Index
    Case 9 '列印順序
        If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '系統類別
      'Add By Cheng 2002/01/07
      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
      
     strTemp1 = Split(UCase(Systemkind_g), ",")
     strTemp2 = Split(UCase(txt1(0)), ",")
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
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
     Next i
Case 5
     lbl1(0) = GetPrjSales(txt1(5))
Case 6
     Select Case Trim(txt1(6))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(6).SetFocus
          txt1(6).SelStart = 0
          txt1(6).SelLength = Len(txt1(6))
          Exit Sub
     End Select
Case 7
     Select Case Trim(txt1(7))
     Case "1", "2", "3", "4", "5", ""
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(7).SetFocus
          txt1(7).SelStart = 0
          txt1(7).SelLength = Len(txt1(7))
          Exit Sub
     End Select
Case 8
     Select Case Trim(txt1(8))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示方式只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(8).SetFocus
          txt1(8).SelStart = 0
          txt1(8).SelLength = Len(txt1(8))
          Exit Sub
     End Select
Case 9
        'Modify By Cheng 2003/06/05
        '不在此檢查
'     Select Case Trim(txt1(9))
'     Case "1", "2", ""
'     Case Else
'          s = MsgBox("列印順序只能輸入 1 或 2 !!", , "USER 輸入錯誤")
'          txt1(9).SetFocus
'          txt1(9).SelStart = 0
'          txt1(9).SelLength = Len(txt1(9))
'          Exit Sub
'     End Select
'add by nickc 2005/03/04
Case 10
   Select Case txt1(Index)
   Case "1"
      Label1(1).Visible = True
      txt1(1).Visible = True
      txt1(2).Visible = True
      Line1.Visible = True
'      txt1(0).Visible = True
'      Label1(0).Visible = True
   Case "2"
      Label1(1).Visible = False
      txt1(1).Visible = False
      txt1(2).Visible = False
      Line1.Visible = False
      txt1(0).Visible = False
      Label1(0).Visible = False
   Case Else
          MsgBox "請輸入 1 或 2 ！", , "選擇新舊制！"
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
   End Select
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 3, 4 '發文年月
   If PUB_CheckKeyInYYMM(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub

'Add By Cheng 2003/07/31
'計算逾時草圖墨圖件數
Private Sub CompOverTime()
Dim rsA As New ADODB.Recordset
Dim tmpcolor1 As Integer
Dim tmpcolor2 As Integer


    StrSQL6 = "": strSQL1 = ""
    strSQL1 = " AND CP05<=" & Val(Me.txt1(4).Text) + 191100 & "31 "
    StrSQL6 = " AND CP05<=" & Val(Me.txt1(4).Text) + 191100 & "31 "
    strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL1
    StrSQL6 = StrSQL6 & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & StrSQL6
    'Mark By Cheng 2004/03/12
    '不考慮發文日及取消收文日的問題
'    strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)>=" & (Val(txt1(3).Text) + 191100) & " And SUBSTR(CP27,1,6)<=" & (Val(txt1(4).Text) + 191100) & " AND CP57 IS NULL ) OR ( CP27 IS NULL AND SUBSTR(CP57,1,6)>=" & (Val(txt1(3).Text) + 191100) & " AND SUBSTR(CP57,1,6)<=" & (Val(txt1(4).Text) + 191100) & ")) and cp05>=19980101 "
'    StrSQL6 = StrSQL6 & " and CP27 IS NULL  and CP57 IS NULL and cp05>=19980101 "
    strSQL1 = strSQL1 & " And cp05>=19980101 "
    StrSQL6 = StrSQL6 & " And cp05>=19980101 "
    'End
    strSQL1 = strSQL1 & " AND ((EP14>=" & ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Val(Me.txt1(3).Text) + 191100 & "01"))) & " And EP14<=" & (Val(Me.txt1(4).Text) + 191100) & "31" & " ) Or (EP17>=" & ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Val(Me.txt1(3).Text) + 191100 & "01"))) & " And EP17<=" & (Val(Me.txt1(4).Text) + 191100) & "31" & ") ) "
    StrSQL6 = StrSQL6 & " AND ((EP14>=" & ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Val(Me.txt1(3).Text) + 191100 & "01"))) & " And EP14<=" & (Val(Me.txt1(4).Text) + 191100) & "31" & " ) Or (EP17>=" & ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Val(Me.txt1(3).Text) + 191100 & "01"))) & " And EP17<=" & (Val(Me.txt1(4).Text) + 191100) & "31" & ") ) "
'    strSQL1 = strSQL1 & " And ((EP15>=" & Val(Me.txt1(3).Text) + 191100 & "01" & " And EP15<=" & Val(Me.txt1(4).Text) + 191100 & "31" & " ) Or (EP18>=" & Val(Me.txt1(3).Text) + 191100 & "01" & " And EP18<=" & Val(Me.txt1(4).Text) + 191100 & "31" & " )) "
'    StrSQL6 = StrSQL6 & " And ((EP15>=" & Val(Me.txt1(3).Text) + 191100 & "01" & " And EP15<=" & Val(Me.txt1(4).Text) + 191100 & "31" & " ) Or (EP18>=" & Val(Me.txt1(3).Text) + 191100 & "01" & " And EP18<=" & Val(Me.txt1(4).Text) + 191100 & "31" & " )) "
    '申請國家
    If Me.txt1(1).Text <> "" Then
        strSQL1 = strSQL1 & " And PA09>='" & Me.txt1(1).Text & "' "
        StrSQL6 = StrSQL6 & " And PA09>='" & Me.txt1(1).Text & "' "
    End If
    If Me.txt1(2).Text <> "" Then
        strSQL1 = strSQL1 & " And PA09<='" & Me.txt1(2).Text & "' "
        StrSQL6 = StrSQL6 & " And PA09<='" & Me.txt1(2).Text & "' "
    End If
    '繪圖人員
    If Me.txt1(5).Text <> "" Then
        strSQL1 = strSQL1 & " And EP13='" & Me.txt1(5).Text & "' "
        StrSQL6 = StrSQL6 & " And EP13='" & Me.txt1(5).Text & "' "
    End If
    '所別
    If Me.txt1(6).Text <> "" Then
        strSQL1 = strSQL1 & " And S2.ST06>='" & Me.txt1(6).Text & "' "
        StrSQL6 = StrSQL6 & " And S2.ST06>='" & Me.txt1(6).Text & "' "
    End If
    If Me.txt1(7).Text <> "" Then
        strSQL1 = strSQL1 & " And S2.ST06<='" & Me.txt1(7).Text & "' "
        StrSQL6 = StrSQL6 & " And S2.ST06<='" & Me.txt1(7).Text & "' "
    End If
   'add by nickc 2005/03/01  加多國案且草圖不計件和墨圖不計件不秀
   strSQL1 = strSQL1 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
   StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/04/11     離職不秀
      strSQL1 = strSQL1 & " and s2.st04='1' "
      StrSQL6 = StrSQL6 & " and s2.st04='1' "

    'Modify By Cheng 2004/03/30
'    strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, EP13 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP13 Is Not Null " & StrSQL6
    
    'Modify by Morgan 2004/5/19
    '加專利種類
'    strSQL = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, EP13 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP13 Is Not Null " & StrSQL6
    strSql = "SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "),'' As 草期限, " & SQLDate("eP15") & ",0, EP29, " & SQLDate("eP17") & ", '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, EP13 , PA08 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
                " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP13 Is Not Null " & StrSQL6
                
  
    'End
'    strSQL = strSQL & " UNION all  SELECT SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)), Decode(pa09,'000',cpm03,cpm04),s1.st02, ROUND(cp18,2),DECODE(EP06,NULL,' ',0,' '," & SQLDate("EP06") & "), EP20, DECODE(EP14,NULL,'',0,''," & SQLDate("eP14") & "), '' As 草期限, " & SQLDate("eP15") & ",0, EP29, DECODE(EP17,NULL,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & "),0,DECODE(EP08,NULL,'',0,''," & SQLDate("EP08") & ")," & SQLDate("eP17") & "), '' As 墨期限, " & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,DECODE(PA09,'000',PTM03,PTM04)," & SQLDate("CP07") & "," & SQLDate("CP57") & ",ep21,ep22,ep23,ep24,ep25,ep16,ep19, CP10, EP13 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP " & _
'                 " WHERE EP02=CP09(+) AND cp01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp14=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP13 Is Not Null " & strSQL1
    rsA.CursorLocation = adUseClient
    rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            tmpcolor1 = 0: tmpcolor2 = 0
            '判斷案件性質
            'Modify by Morgan 2004/5/19
            '該依專利種類判斷
'            Select Case "" & rsA("CP10").Value
'            Case "103", "105" '設計申請
            Select Case ("" & rsA("PA08").Value)
               Case "3"
               
                If "" & rsA("EP20").Value = "" And "" & rsA.Fields(11).Value <> "" And "" & rsA.Fields(9).Value <> "" Then
                    tmpcolor1 = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(11).Value)), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(9).Value)))
                End If
                If "" & rsA("EP29").Value = "" And "" & rsA.Fields(16).Value <> "" And "" & rsA.Fields(14).Value <> "" Then
                    tmpcolor2 = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(16).Value)), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(14).Value)))
                End If
                If tmpcolor1 > 5 And "" & rsA("EP20").Value = "" Then
                    'Modify By Cheng 2004/02/17
                    '若草完日在發文年月區間
                    If Left(Val(DBDATE("" & rsA.Fields(11).Value)), 6) >= Val(Me.txt1(3).Text) + 191100 And Left(Val(DBDATE("" & rsA.Fields(11).Value)), 6) <= Val(Me.txt1(4).Text) + 191100 Then
                        strSql = "INSERT INTO R090701 (r102001, R102019, id) Values('" & rsA("EP13").Value & "', 1, '" & strUserNum & "' ) "
                        cnnConnection.Execute strSql
                    End If
                    'End
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                Else
'                    'Modify By Cheng 2004/03/15
'    '                '無發文日, 無取消收文日, 無草完日
'    '                If Len(Trim("" & rsA.Fields(19).Value)) = 0 And Len(Trim("" & rsA.Fields(25).Value)) = 0 And Len(Trim("" & rsA.Fields(11).Value)) = 0 Then
'                    '無草完日
'                    If Len(Trim("" & rsA.Fields(11).Value)) = 0 Then
'                    'End
'                        '若有草齊日
'                        If Len(Trim("" & rsA.Fields(9).Value)) <> 0 Then
'                            '若系統日超過草齊日5個工作天
'                            '若系統日>=迄年月的最後一天, 則用迄年月的最後一天和草齊日比, 否則用系統日和草齊日比
'    '                        If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(9).Value))) > 5 Then
'                            If GetWorkDay(IIf(Val(strSrvDate(1)) >= Val(Me.txt1(4).Text & "31") + 19110000, "" & Val(Me.txt1(4).Text & "31") + 19110000, strSrvDate(1)), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(9).Value))) > 5 Then
'                                strSQL = "INSERT INTO R090701 (r102001, R102019, id) Values('" & rsA("EP13").Value & "', 1, '" & strUserNum & "' ) "
'                                cnnConnection.Execute strSQL
'                            End If
'                        End If
'                    End If
                End If
                If tmpcolor2 > 3 And "" & rsA("EP29").Value = "" Then
                    'Modify By Cheng 2004/02/17
                    '若墨完日在發文年月區間
                    If Left(Val(DBDATE("" & rsA.Fields(16).Value)), 6) >= Val(Me.txt1(3).Text) + 191100 And Left(Val(DBDATE("" & rsA.Fields(16).Value)), 6) <= Val(Me.txt1(4).Text) + 191100 Then
                        strSql = "INSERT INTO R090701 (r102001, R102020, id) Values('" & rsA("EP13").Value & "', 1, '" & strUserNum & "' ) "
                        cnnConnection.Execute strSql
                    End If
                    'End
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                Else
'                    'Modify By Cheng 2004/03/15
'    '                '無發文日, 無取消收文日, 無墨完日
'    '                If Len(Trim("" & rsA.Fields(19).Value)) = 0 And Len(Trim("" & rsA.Fields(25).Value)) = 0 And Len(Trim("" & rsA.Fields(16).Value)) = 0 Then
'                    '無墨完日
'                    If Len(Trim("" & rsA.Fields(16).Value)) = 0 Then
'                    'End
'                        '若有墨齊日
'                        If Len(Trim("" & rsA.Fields(14).Value)) <> 0 Then
'                            '若系統日大於墨齊日3個工作天
'                            '若系統日>=迄年月的最後一天, 則用迄年月的最後一天和墨齊日比, 否則用系統日和墨齊日比
'    '                        If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(14).Value))) > 3 Then
'                            If GetWorkDay(IIf(Val(strSrvDate(1)) >= Val(Me.txt1(4).Text & "31") + 19110000, "" & Val(Me.txt1(4).Text & "31") + 19110000, strSrvDate(1)), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(14).Value))) > 3 Then
'                                strSQL = "INSERT INTO R090701 (r102001, R102020, id) Values('" & rsA("EP13").Value & "', 1, '" & strUserNum & "' ) "
'                                cnnConnection.Execute strSQL
'                            End If
'                        End If
'                    End If
                End If
            Case Else '非設計申請
                If "" & rsA("EP20").Value = "" And "" & rsA.Fields(11).Value <> "" And "" & rsA.Fields(9).Value <> "" Then
                    tmpcolor1 = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(11).Value)), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(9).Value)))
                End If
                If "" & rsA("EP29").Value = "" And "" & rsA.Fields(16).Value <> "" And "" & rsA.Fields(14).Value <> "" Then
                    tmpcolor2 = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(16).Value)), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(14).Value)))
                End If
                If tmpcolor1 > 4 And "" & rsA("EP20").Value = "" Then
                    'Modify By Cheng 2004/02/17
                    '若草完日在發文年月區間
                    If Left(Val(DBDATE("" & rsA.Fields(11).Value)), 6) >= Val(Me.txt1(3).Text) + 191100 And Left(Val(DBDATE("" & rsA.Fields(11).Value)), 6) <= Val(Me.txt1(4).Text) + 191100 Then
                        strSql = "INSERT INTO R090701 (r102001, R102019, id) Values('" & rsA("EP13").Value & "', 1, '" & strUserNum & "' ) "
                        cnnConnection.Execute strSql
                    End If
                    'End
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                Else
'                    'Modify By Cheng 2004/03/15
'    '                '無發文日, 無取消收文日, 無草完日
'    '                If Len(Trim("" & rsA.Fields(19).Value)) = 0 And Len(Trim("" & rsA.Fields(25).Value)) = 0 And Len(Trim("" & rsA.Fields(11).Value)) = 0 Then
'                    '無草完日
'                    If Len(Trim("" & rsA.Fields(11).Value)) = 0 Then
'                    'End
'                       '若有草齊日
'                       If Len(Trim("" & rsA.Fields(9).Value)) <> 0 Then
'                          '若系統日超過草齊日4個工作天
'                            '若系統日>=迄年月的最後一天, 則用迄年月的最後一天和草齊日比, 否則用系統日和草齊日比
'    '                      If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(9).Value))) > 4 Then
'                          If GetWorkDay(IIf(Val(strSrvDate(1)) >= Val(Me.txt1(4).Text & "31") + 19110000, "" & Val(Me.txt1(4).Text & "31") + 19110000, strSrvDate(1)), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(9).Value))) > 4 Then
'                            strSQL = "INSERT INTO R090701 (r102001, R102019, id) Values('" & rsA("EP13").Value & "', 1, '" & strUserNum & "' ) "
'                            cnnConnection.Execute strSQL
'                          End If
'                       End If
'                    End If
                End If
                If tmpcolor2 > 3 And "" & rsA("EP29").Value = "" Then
                    'Modify By Cheng 2004/02/17
                    '若墨完日在發文年月區間
                    If Left(Val(DBDATE("" & rsA.Fields(16).Value)), 6) >= Val(Me.txt1(3).Text) + 191100 And Left(Val(DBDATE("" & rsA.Fields(16).Value)), 6) <= Val(Me.txt1(4).Text) + 191100 Then
                        strSql = "INSERT INTO R090701 (r102001, R102020, id) Values('" & rsA("EP13").Value & "', 1, '" & strUserNum & "' ) "
                        cnnConnection.Execute strSql
                    End If
                    'End
'edit by nickc 2005/04/18 瓊玉說無完稿日不算逾時
'                Else
'                    'Modify By Cheng 2004/03/15
'    '                '無發文日, 無取消收文日, 無墨完日
'    '                If Len(Trim("" & rsA.Fields(19).Value)) = 0 And Len(Trim("" & rsA.Fields(25).Value)) = 0 And Len(Trim("" & rsA.Fields(16).Value)) = 0 Then
'                    '無墨完日
'                    If Len(Trim("" & rsA.Fields(16).Value)) = 0 Then
'                    'End
'                       '若有墨齊日
'                       If Len(Trim("" & rsA.Fields(14).Value)) <> 0 Then
'                          '若系統日大於墨齊日3個工作天
'                            '若系統日>=迄年月的最後一天, 則用迄年月的最後一天和草齊日比, 否則用系統日和草齊日比
'    '                      If GetWorkDay(strSrvDate(1), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(14).Value))) > 3 Then
'                          If GetWorkDay(IIf(Val(strSrvDate(1)) >= Val(Me.txt1(4).Text & "31") + 19110000, "" & Val(Me.txt1(4).Text & "31") + 19110000, strSrvDate(1)), ChangeTStringToWString(ChangeTDateStringToTString("" & rsA.Fields(14).Value))) > 3 Then
'                            strSQL = "INSERT INTO R090701 (r102001, R102020, id) Values('" & rsA("EP13").Value & "', 1, '" & strUserNum & "' ) "
'                            cnnConnection.Execute strSQL
'                          End If
'                       End If
'                    End If
                End If
            End Select
            rsA.MoveNext
        Wend
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing

End Sub

'add by nickc 2005/03/04 新制
Function ProcessNew() As Boolean

On Error GoTo ErrorHandler
ProcessNew = False
cnnConnection.Execute "DELETE FROM r090701 WHERE ID='" & strUserNum & "' "
StrSQL7 = "": strSQL8 = "": strSQL9 = ""

strSQL71 = "": strSQL72 = "": strSQL73 = "": strSQL74 = ""

StrSQL6 = ""
strSQL1 = ""
strSQL73 = strSQL73 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
strSQL1 = strSQL1 + " AND pe03>=" & Val(txt1(3)) + 191100 & " AND pe03<=" & Val(txt1(4)) + 191100 & " "
'繪圖人員
If Len(txt1(5)) <> 0 Then
    strSQL1 = strSQL1 + " AND pe01='" & txt1(5) & "' "
    strSQL73 = strSQL73 + " AND EP13='" & txt1(5) & "' "
End If
'所別(起)
If Len(txt1(6)) <> 0 Then
    strSQL1 = strSQL1 + " AND ST06>='" & txt1(6) & "' "
    strSQL73 = strSQL73 + " AND S2.ST06>='" & txt1(6) & "' "
End If
'所別(迄)
If Len(txt1(7)) <> 0 Then
    strSQL1 = strSQL1 + " AND ST06<='" & txt1(7) & "' "
    strSQL73 = strSQL73 + " AND S2.ST06<='" & txt1(7) & "' "
End If
'件數不計算墨圖不計件, 點數張數則都要
StrSQL7 = StrSQL7 + " and ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND ST04='1' ", " ")
strSQL9 = strSQL9 + " and ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND ST04='1' ", " ")
'add by nickc 205/04/11 離職不秀
    strSQL1 = strSQL1 + " AND ST04='1' "
    strSQL73 = strSQL73 + " AND S2.ST04='1' "
CheckOC
'改抓速度統計
strSql = "INSERT INTO R090701 (r102001,r102002,r102003,r102004,r102005,r102006,r102007,r102016,r102018,r102019,r102020,id) " & _
              " select pe01,sum(pe11)/2,sum(pe09)/2,SUM(PE10)/2,sum(nvl(ma40,0)),sum(nvl(ma37,0)),sum(nvl(ma47,0)),sum(decode(ma36,'1',nvl(ma43,0),0)),sum(decode(ma36,'2',nvl(ma43,0),0)),sum(decode(ma36,'1',nvl(ma51,0),0)),sum(decode(ma36,'2',nvl(ma51,0),0)),'" & strUserNum & "' From Performance, STAFF,monthassess " & _
              " Where PE01=ST01(+) and pe01=ma01(+) and pe03=ma02(+)  AND PE03>=" & Val(txt1(3)) + 191100 & " and PE03<=" & Val(txt1(4)) + 191100 & " And PE02='P' " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & " AND ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND ST04='1' ", " ") & strSQL72 & " Group By PE01,'" & strUserNum & "' "
cnnConnection.Execute strSql

'達成率
strSql = "(SELECT R102001," & _
         "ROUND(DECODE(SUM(DECODE(R102002,0,0,NULL,0,R102002)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102002,0,0,NULL,0,R102002)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102007,0,0,NULL,0,R102007))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2)," & _
         "(ROUND(DECODE(SUM(DECODE(R102002,0,0,NULL,0,R102002)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102002,0,0,NULL,0,R102002)))*100),2) + ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2))/2, " & _
         "'" & strUserNum & "' " & _
         "FROM R090701 WHERE ID='" & strUserNum & "' GROUP BY R102001 )"
cnnConnection.Execute "INSERT INTO R090701 (R102001,R102008,R102009,R102010,R102011,ID) " & strSql
'其他新案
strSql = "SELECT EP13,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(EP20, Null, 1, 0)),0,0,0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP') And EP16>0 And substr(EP15,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP15,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
'其他舊案
strSql = strSql + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(EP29, Null, 1, 0)),0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP') And (EP16 Is Null Or EP16<=0) And substr(EP18,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP18,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
cnnConnection.Execute " INSERT INTO r090701 " & strSql
'逾時草圖墨圖件數
'CompOverTime
ProcessNew = True
Exit Function
ErrorHandler:
    MsgBox Err.Description
    Resume Next
End Function

'即時運算
Function ProcessNew2() As Boolean
On Error GoTo ErrorHandler
ProcessNew2 = False
cnnConnection.Execute "DELETE FROM r090701 WHERE ID='" & strUserNum & "' "
StrSQL7 = "": strSQL8 = "": strSQL9 = ""
strSQL71 = "": strSQL72 = "": strSQL73 = "": strSQL74 = ""

'系統類別
'Removed by Morgan 2012/7/12 不必再限制系統(繪圖有可能辦FCP案但又無系統權限)
'If Len(txt1(0)) <> 0 Then
'   StrSQL7 = StrSQL7 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
'   strSQL71 = strSQL71 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
'   strSQL73 = strSQL73 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
'   strSQL74 = strSQL74 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
'    strSQL8 = strSQL8 + " And CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
'    pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/17
'End If

'add by nickc 2005/04/18
'StrSQL7 = StrSQL7 & " and cp107='Y' "
''strSQL71 = strSQL71 & " and cp107='Y' "
'strSQL73 = strSQL73 & " and cp107='Y' "
'strSQL74 = strSQL74 & " and cp107='Y' "
'strSQL8 = strSQL8 & " and cp107='Y' "

StrSQL6 = ""
strSQL1 = ""
'申請國家(起)
If Len(txt1(1).Text) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & txt1(1) & "' "
    strSQL2 = strSQL2 + " AND TM10>='" & txt1(1) & "' "
    StrSQL3 = StrSQL3 + " AND LC15>='" & txt1(1) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09>='" & txt1(1) & "' "
End If
'申請國家(迄)
If Len(txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & txt1(2) & "' "
    strSQL2 = strSQL2 + " AND TM10<='" & txt1(2) & "' "
    StrSQL3 = StrSQL3 + " AND LC15<='" & txt1(2) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09<='" & txt1(2) & "' "
End If
If Len(txt1(1).Text) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/17
End If
'edit by nickc 2005/05/13
'StrSQL7 = StrSQL7 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " And CP57 Is Null "
'strSQL71 = strSQL71 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " And CP57 Is Null "
'strSQL8 = strSQL8 + " AND ((CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " And CP57 Is Null ) Or (CP57>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP57<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " And CP27 Is Null) Or (CP27 Is Null And CP57 Is Null) ) "
StrSQL7 = StrSQL7 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " "
strSQL71 = strSQL71 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & "  "
strSQL8 = strSQL8 + " AND ((CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & "  ) Or (CP57>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP57<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " And CP27 Is Null) Or (CP27 Is Null And CP57 Is Null) ) "
strSQL9 = strSQL9 + " AND (SH01>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND SH01<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " ) "
pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/17

'繪圖人員
If Len(txt1(5)) <> 0 Then
    StrSQL7 = StrSQL7 + " AND EP13='" & txt1(5) & "' "
    strSQL71 = strSQL71 + " AND EP13='" & txt1(5) & "' "
    strSQL72 = strSQL72 + " AND PE01='" & txt1(5) & "' "
    strSQL73 = strSQL73 + " AND EP13='" & txt1(5) & "' "
    strSQL74 = strSQL74 + " AND EP13='" & txt1(5) & "' "
    strSQL8 = strSQL8 + " AND EP13='" & txt1(5) & "' "
    strSQL9 = strSQL9 + " AND SH02='" & txt1(5) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(5) & lbl1(0) 'Add By Sindy 2010/12/17
End If
'所別(起)
If Len(txt1(6)) <> 0 Then
    StrSQL7 = StrSQL7 + " AND ST06>='" & txt1(6) & "' "
    strSQL71 = strSQL71 + " AND S2.ST06>='" & txt1(6) & "' "
    strSQL72 = strSQL72 + " AND ST06>='" & txt1(6) & "' "
    strSQL73 = strSQL73 + " AND S2.ST06>='" & txt1(6) & "' "
    strSQL74 = strSQL74 + " AND ST06>='" & txt1(6) & "' "
    strSQL8 = strSQL8 + " AND ST06>='" & txt1(6) & "' "
    strSQL9 = strSQL9 + " AND ST06>='" & txt1(6) & "' "
End If
'所別(迄)
If Len(txt1(7)) <> 0 Then
    StrSQL7 = StrSQL7 + " AND ST06<='" & txt1(7) & "' "
    strSQL71 = strSQL71 + " AND S2.ST06<='" & txt1(7) & "' "
    strSQL72 = strSQL72 + " AND ST06<='" & txt1(7) & "' "
    strSQL73 = strSQL73 + " AND S2.ST06<='" & txt1(7) & "' "
    strSQL74 = strSQL74 + " AND ST06<='" & txt1(7) & "' "
    strSQL8 = strSQL8 + " AND ST06<='" & txt1(7) & "' "
    strSQL9 = strSQL9 + " AND ST06<='" & txt1(7) & "' "
End If
If Len(txt1(6)) <> 0 Or Len(txt1(7)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(6) & "-" & txt1(7) & Label1(7) 'Add By Sindy 2010/12/17
End If

'件數不計算墨圖不計件, 點數張數則都要
StrSQL7 = StrSQL7 + " and ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND ST04='1' ", " ")
strSQL9 = strSQL9 + " and ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND ST04='1' ", " ")

'Removed by Morgan 2013/8/5 離職也要--瓊玉
''add by nickc 205/04/11 離職不秀
'    StrSQL7 = StrSQL7 + " AND ST04='1' "
'    strSQL71 = strSQL71 + " AND S2.ST04='1' "
'    strSQL72 = strSQL72 + " AND ST04='1' "
'    strSQL73 = strSQL73 + " AND S2.ST04='1' "
'    strSQL74 = strSQL74 + " AND ST04='1' "
'    strSQL8 = strSQL8 + " AND ST04='1' "
'    strSQL9 = strSQL9 + " AND ST04='1' "

'add by nickc 2007/02/01
strSQL71 = strSQL71 & " and cp107='Y' "

CheckOC
'目標
strSql = "INSERT INTO R090701 (r102001,r102002,r102003,r102004,id) select pe01,sum(pe11),sum(pe09),SUM(PE10),'" & strUserNum & "' From Performance, STAFF Where PE01=ST01(+) AND PE03>=" & Val(txt1(3)) + 191100 & " and PE03<=" & Val(txt1(4)) + 191100 & " And PE02='P' " & IIf(Len(txt1(5)) <> 0, " AND PE01='" & txt1(5) & "' ", "") & " AND ST05 IN ('79','81','82','AC') " & IIf(Len(Me.txt1(5).Text) > 0, " AND ST04='1' ", " ") & strSQL72 & " Group By PE01,'" & strUserNum & "' "
cnnConnection.Execute strSql
'目標達成
'繪圖在統計目標達成點數及件數時, 若墨圖不計件的不統計(而承辦人是不計件的件數不計但點數要計)
'edit by nickc 2006/02/08 點數要扣銷案點
'strSQL = " SELECT EP13, SUM(CP18), Sum(cp103 * cp104), 0,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.Txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") And EP29 Is Null " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' "
'edit by nickc 2006/06/05 瓊玉說的，因為應該要控制與舊制相同，所以回歸舊點，墨圖計件才計，補收款或拆收據等其他相關案件性質或其他導致原申請案不計件但要畫圖，其他程序不畫圖但要計件的話，由工程師將該程序上繪圖人員，就會計件
'strSQL = " SELECT EP13, SUM(cp18-nvl(a1u07/1000,0)), sum(decode(ep29,null,cp103 * cp104,0)), 0,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & ") group by a1u03) WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) and ep02=a1u03(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' "

'edit by nickc 2006/07/03 瓊玉跟大愛討論說有繪圖人員的其他(910)就算草墨都不算也要計點，不計件
'strSQL = " SELECT EP13, SUM(cp18-nvl(a1u07/1000,0)), sum(decode(ep29,null,cp103 * cp104,0)), 0,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & ") group by a1u03) WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) And EP29 Is Null and ep02=a1u03(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' "
strSql = " SELECT EP13, SUM(cp18-nvl(a1u07/1000,0)), sum(decode(ep29,null,cp103 * cp104,0)), 0,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where CP27>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & ") group by a1u03) WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) And (EP29 Is Null or ep20||ep29||cp10='NN910') and ep02=a1u03(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' "

'cnnConnection.Execute "insert into r090701 (r102001,r102005,r102006,r102007,id) " & strSQL
'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數(原來沒有和工程師同步,本次一併更正)
'strSql = strSql & "union all SELECT SH02, 0, Sum(Decode(SH06, 'CFP', Nvl(SH05,0)/4, Nvl(SH05,0)/4)), 0,'" & strUserNum & "' FROM SupportHour,STAFF,PATENT WHERE SH02=ST01(+) And SH06=PA01(+) And SH07=PA02(+) And SH08=PA03(+) And SH09=PA04(+) And SH11='V' " & strSQL1 & strSQL9 & " GROUP BY SH02,'" & strUserNum & "' "
strSql = strSql & "union all SELECT SH02, 0, Sum(" & Sh2EPtCode & "), 0,'" & strUserNum & "' FROM SupportHour,STAFF,PATENT WHERE SH02=ST01(+) And SH06=PA01(+) And SH07=PA02(+) And SH08=PA03(+) And SH09=PA04(+) And SH11='V' " & strSQL1 & strSQL9 & " GROUP BY SH02,'" & strUserNum & "' "
'end 2014/4/1

'cnnConnection.Execute "insert into r090701 (r102001,r102005,r102006,r102007,id) " & strSQL
'edit by nickc 2006/02/07 應該要算達成，而不是承辦
'strSQL = strSQL & "union all SELECT EP13, 0, 0, sum(nvl(ep16,0))/2,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.Txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL73 & _
                " AND EP15>=" & Val(ChangeTStringToWString(Txt1(3) & "01")) & " AND EP15<=" & Val(ChangeTStringToWString(Txt1(4) & "31")) & " GROUP BY EP13,'" & strUserNum & "' "
strSql = strSql & "union all SELECT EP13, 0, 0, sum(nvl(ep16,0))/2,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' "
'cnnConnection.Execute "insert into r090701 (r102001,r102005,r102006,r102007,id) " & strSQL
'edit by nickc 2006/02/07 應該要算達成，而不是承辦
'strSQL = strSQL & "union all SELECT EP13, 0, 0, sum(nvl(ep19,0))/2,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.Txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL73 & _
                " AND EP18>=" & Val(ChangeTStringToWString(Txt1(3) & "01")) & " AND EP18<=" & Val(ChangeTStringToWString(Txt1(4) & "31")) & " GROUP BY EP13,'" & strUserNum & "' "
strSql = strSql & "union all SELECT EP13, 0, 0, sum(nvl(ep19,0))/2,'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL71 & " GROUP BY EP13,'" & strUserNum & "' "
'cnnConnection.Execute "insert into r090701 (r102001,r102005,r102006,r102007,id) " & strSQL
'edit by nickc 2006/02/08 可不用
'strSQL = strSQL & "union all SELECT EP13, 0, 0, Sum((ep16+ep19)/2),'" & strUserNum & "' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,STAFF S1,STAFF S2 WHERE EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP14=S1.ST01(+) AND EP13=S2.ST01(+) " & strSQL1 & StrSQL6 & " AND S2.ST05 IN ('79','81','82','AC')  " & IIf(Len(Me.Txt1(5).Text) > 0, " AND S2.ST04='1' ", " ") & " AND CP01 IN (" & SQLGrpStr("", 1) & ") " & strSQL73 & _
                " AND ((EP15>=" & Val(ChangeTStringToWString(Txt1(3) & "01")) & " AND EP15<=" & Val(ChangeTStringToWString(Txt1(4) & "31")) & ") And (EP18>=" & Val(ChangeTStringToWString(Txt1(3) & "01")) & " And EP18<=" & Val(ChangeTStringToWString(Txt1(4) & "31")) & ")) " & " GROUP BY EP13,'" & strUserNum & "' "
cnnConnection.Execute "insert into r090701 (r102001,r102005,r102006,r102007,id) (" & strSql & ") "
'達成率
strSql = "(SELECT R102001," & _
         "ROUND(DECODE(SUM(DECODE(R102002,0,0,NULL,0,R102002)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102002,0,0,NULL,0,R102002)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2)," & _
         "ROUND(DECODE(SUM(DECODE(R102004,0,0,NULL,0,R102004)),0,0,(SUM(DECODE(R102007,0,0,NULL,0,R102007))/SUM(DECODE(R102004,0,0,NULL,0,R102004)))*100),2)," & _
         "(ROUND(DECODE(SUM(DECODE(R102002,0,0,NULL,0,R102002)),0,0,(SUM(DECODE(R102005,0,0,NULL,0,R102005))/SUM(DECODE(R102002,0,0,NULL,0,R102002)))*100),2) + ROUND(DECODE(SUM(DECODE(R102003,0,0,NULL,0,R102003)),0,0,(SUM(DECODE(R102006,0,0,NULL,0,R102006))/SUM(DECODE(R102003,0,0,NULL,0,R102003)))*100),2))/2, " & _
         "'" & strUserNum & "' " & _
         "FROM R090701 WHERE ID='" & strUserNum & "' GROUP BY R102001 )"
cnnConnection.Execute "INSERT INTO R090701 (R102001,R102008,R102009,R102010,R102011,ID) " & strSql
'edit by nickc 2005/04/13 畫面不秀了
'strSQL = "SELECT EP13,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(EP20, Null, 1, 0)),0,0,0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP') And EP16>0 And substr(EP15,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP15,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
'strSQL = strSQL + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,SUM(CP18), Sum(Decode(EP29, Null, 1, 0)),0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP') And (EP16 Is Null Or EP16<=0) And substr(EP18,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP18,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
'add by nickc 2005/04/13 提供圖檔
                                  strSql = "SELECT EP13,0,0,0,0,0,0,0,0,0,0,count(*), 0,0,0,0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and cp106='Y' AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP','P') And EP16>0 And substr(EP15,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP15,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
'add by nickc 2005/04/13 轉換案
strSql = strSql + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,count(*),0,0,0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and ep20='N' and ep29 is null and cp103=0.4 and cp100=0  AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP','P')  And substr(EP18,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP18,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
'add by nickc 2005/04/18 在加入其它新舊案
    strSql = strSql & " union all SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,Sum(Decode(EP20, Null, 1, 0)),0,0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP') And EP16>0 And substr(EP15,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP15,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
strSql = strSql + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,Sum(Decode(EP29, Null, 1, 0)),0,0,0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S2,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=S2.ST01(+) AND CP01 In ('CFP') And (EP16 Is Null Or EP16<=0) And substr(EP18,1,6)>=" & Val(Me.txt1(3).Text) + 191100 & " And substr(EP18,1,6)<=" & Val(Me.txt1(4).Text) + 191100 & strSQL73 & " GROUP BY EP13,'" & strUserNum & "' "
'Modify by Morgan 2010/6/7 不必再判斷張數(取消 And (EP16 Is Not Null And EP16 > 0 ) 條件) --瓊玉
strSql = strSql + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,Sum(cp100 * cp101),0,0,'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) " & strSQL1 & strSQL74 & _
                " AND EP15>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP15<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " GROUP BY EP13,'" & strUserNum & "' "
'墨圖件數(不考慮發文日及取消收文日的問題)
strSql = strSql + " UNION all  SELECT EP13,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,Sum(cp103 * cp104),'" & strUserNum & "',0,0 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF,PATENT WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP02=CP09(+) AND  EP13=ST01(+) " & strSQL1 & strSQL74 & _
                " AND EP18>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND EP18<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " GROUP BY EP13,'" & strUserNum & "' "
cnnConnection.Execute " INSERT INTO r090701 " & strSql
'支援記錄算入草圖及墨圖件數
'Modified by Morgan 2013/11/1 比照速度考核按比例分配
'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數(原來沒有和工程師同步,本次一併更正)
'strSql = "SELECT SH02,0,0,0,0,0,0,0,0,0,0,0,0,0,0,Sum(Decode(SH06, 'CFP', Nvl(SH05,0)/4, Nvl(SH05,0)/4))*0.65*2,0,0,'" & strUserNum & "',0,0 FROM SupportHour,STAFF,PATENT WHERE SH02=ST01(+) And SH06=PA01(+) And SH07=PA02(+) And SH08=PA03(+) And SH09=PA04(+) And SH11='V' " & strSQL1 & strSQL9 & " GROUP BY SH02,'" & strUserNum & "' "
strSql = "SELECT SH02,0,0,0,0,0,0,0,0,0,0,0,0,0,0,Sum(" & Sh2EPtCode & ")*0.65*2,0,0,'" & strUserNum & "',0,0 FROM SupportHour,STAFF,PATENT WHERE SH02=ST01(+) And SH06=PA01(+) And SH07=PA02(+) And SH08=PA03(+) And SH09=PA04(+) And SH11='V' " & strSQL1 & strSQL9 & " GROUP BY SH02,'" & strUserNum & "' "
'end 2014/3/20
cnnConnection.Execute " INSERT INTO r090701 " & strSql
'Modified by Morgan 2013/11/1 比照速度考核按比例分配
'Modified by Morgan 2014/3/20 --2014/4/1起支援改每小時折計0.2基數(原來沒有和工程師同步,本次一併更正)
'strSql = "SELECT SH02,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,Sum(Decode(SH06, 'CFP', Nvl(SH05,0)/4, Nvl(SH05,0)/4))*0.35*2,'" & strUserNum & "',0,0 FROM SupportHour,STAFF,PATENT WHERE SH02=ST01(+) And SH06=PA01(+) And SH07=PA02(+) And SH08=PA03(+) And SH09=PA04(+) And SH11='V' " & strSQL1 & strSQL9 & " GROUP BY SH02,'" & strUserNum & "' "
strSql = "SELECT SH02,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,Sum(" & Sh2EPtCode & ")*0.35*2,'" & strUserNum & "',0,0 FROM SupportHour,STAFF,PATENT WHERE SH02=ST01(+) And SH06=PA01(+) And SH07=PA02(+) And SH08=PA03(+) And SH09=PA04(+) And SH11='V' " & strSQL1 & strSQL9 & " GROUP BY SH02,'" & strUserNum & "' "
'end 2014/3/20
cnnConnection.Execute " INSERT INTO r090701 " & strSql
'逾時草圖墨圖件數
CompOverTime

ProcessNew2 = True
Exit Function
ErrorHandler:
    MsgBox Err.Description
    Resume Next
End Function
