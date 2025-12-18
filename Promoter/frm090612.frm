VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090612 
   BorderStyle     =   1  '單線固定
   Caption         =   "未齊備,未完稿,未發文查詢"
   ClientHeight    =   5592
   ClientLeft      =   876
   ClientTop       =   1296
   ClientWidth     =   4044
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5592
   ScaleWidth      =   4044
   Begin VB.Frame Frame3 
      Height          =   465
      Left            =   45
      TabIndex        =   42
      Top             =   3210
      Width           =   3945
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   10
         Left            =   1095
         MaxLength       =   1
         TabIndex        =   43
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "顯示方式："
         Height          =   180
         Index           =   22
         Left            =   120
         TabIndex        =   45
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "(1.螢幕 2.報表)"
         Height          =   180
         Index           =   21
         Left            =   1425
         TabIndex        =   44
         Top             =   180
         Width           =   1305
      End
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   16
      Left            =   1872
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1120
      Width           =   675
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   15
      Left            =   1068
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1120
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   1788
      Left            =   45
      TabIndex        =   23
      Top             =   1440
      Width           =   3930
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   9
         Left            =   990
         MaxLength       =   6
         TabIndex        =   30
         Top             =   1452
         Width           =   900
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   8
         Left            =   1545
         MaxLength       =   3
         TabIndex        =   29
         Top             =   1140
         Width           =   525
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   7
         Left            =   990
         MaxLength       =   3
         TabIndex        =   28
         Top             =   1140
         Width           =   480
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   6
         Left            =   990
         MaxLength       =   6
         TabIndex        =   27
         Top             =   816
         Width           =   900
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   5
         Left            =   990
         MaxLength       =   1
         TabIndex        =   26
         Top             =   492
         Width           =   315
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   4
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   25
         Top             =   165
         Width           =   270
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   3
         Left            =   990
         MaxLength       =   1
         TabIndex        =   24
         Top             =   150
         Width           =   270
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   47
         Top             =   1470
         Width           =   1500
         VariousPropertyBits=   27
         Size            =   "2646;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   46
         Top             =   816
         Width           =   1500
         VariousPropertyBits=   27
         Size            =   "2646;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line4 
         X1              =   1176
         X2              =   1896
         Y1              =   1296
         Y2              =   1296
      End
      Begin VB.Line Line3 
         X1              =   1125
         X2              =   1575
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label Label1 
         Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
         Height          =   180
         Index           =   24
         Left            =   1764
         TabIndex        =   37
         Top             =   216
         Width           =   2124
      End
      Begin VB.Label Label1 
         Caption         =   "(1.承辦人 2.智權人員)"
         Height          =   180
         Index           =   14
         Left            =   1356
         TabIndex        =   36
         Top             =   504
         Width           =   1668
      End
      Begin VB.Label Label1 
         Caption         =   "業務區："
         Height          =   180
         Index           =   9
         Left            =   108
         TabIndex        =   35
         Top             =   1188
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "承辦人："
         Height          =   180
         Index           =   8
         Left            =   108
         TabIndex        =   34
         Top             =   864
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "智權人員："
         Height          =   180
         Index           =   6
         Left            =   108
         TabIndex        =   33
         Top             =   1512
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "所別："
         Height          =   180
         Index           =   4
         Left            =   108
         TabIndex        =   32
         Top             =   252
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "查詢對象："
         Height          =   180
         Index           =   3
         Left            =   108
         TabIndex        =   31
         Top             =   564
         Width           =   1092
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   45
      TabIndex        =   9
      Top             =   3615
      Width           =   3930
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   17
         Left            =   2100
         TabIndex        =   39
         Top             =   1020
         Width           =   525
      End
      Begin VB.CheckBox Check1 
         Caption         =   "已完稿未會稿之時限："
         Height          =   180
         Index           =   4
         Left            =   150
         TabIndex        =   40
         Top             =   1050
         Width           =   2115
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   14
         Left            =   2430
         TabIndex        =   10
         Top             =   1620
         Width           =   525
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   13
         Left            =   2430
         TabIndex        =   11
         Top             =   1320
         Width           =   525
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   12
         Left            =   2100
         TabIndex        =   12
         Top             =   735
         Width           =   525
      End
      Begin VB.TextBox Txt1 
         Height          =   264
         Index           =   11
         Left            =   2100
         TabIndex        =   13
         Top             =   456
         Width           =   525
      End
      Begin VB.CheckBox Check1 
         Caption         =   "已收文未齊備之時限："
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   504
         Width           =   2115
      End
      Begin VB.CheckBox Check1 
         Caption         =   "已齊備未完稿之時限："
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   16
         Top             =   765
         Width           =   2115
      End
      Begin VB.CheckBox Check1 
         Caption         =   "已會稿未會稿完成之時限："
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   15
         Top             =   1350
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "已會稿完成未發文之時限："
         Height          =   180
         Index           =   3
         Left            =   150
         TabIndex        =   14
         Top             =   1665
         Width           =   2475
      End
      Begin VB.Label Label2 
         Caption         =   "天"
         Height          =   180
         Index           =   4
         Left            =   2715
         TabIndex        =   41
         Top             =   1050
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "承辦天數統計範圍："
         Height          =   180
         Index           =   7
         Left            =   150
         TabIndex        =   22
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "天"
         Height          =   180
         Index           =   0
         Left            =   2748
         TabIndex        =   21
         Top             =   504
         Width           =   372
      End
      Begin VB.Label Label2 
         Caption         =   "天"
         Height          =   180
         Index           =   1
         Left            =   2715
         TabIndex        =   20
         Top             =   765
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "天"
         Height          =   180
         Index           =   2
         Left            =   2985
         TabIndex        =   19
         Top             =   1350
         Width           =   225
      End
      Begin VB.Label Label2 
         Caption         =   "天"
         Height          =   180
         Index           =   3
         Left            =   3030
         TabIndex        =   18
         Top             =   1665
         Width           =   225
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2400
      TabIndex        =   5
      Top             =   50
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3156
      TabIndex        =   6
      Top             =   50
      Width           =   756
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   0
      Left            =   1068
      TabIndex        =   0
      Top             =   504
      Width           =   1650
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   1
      Left            =   1068
      MaxLength       =   4
      TabIndex        =   1
      Top             =   816
      Width           =   615
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   2
      Left            =   1872
      MaxLength       =   4
      TabIndex        =   2
      Top             =   804
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   38
      Top             =   1162
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   1440
      X2              =   2160
      Y1              =   1240
      Y2              =   1240
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   108
      TabIndex        =   8
      Top             =   576
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   108
      TabIndex        =   7
      Top             =   864
      Width           =   1092
   End
   Begin VB.Line Line1 
      X1              =   1410
      X2              =   2295
      Y1              =   930
      Y2              =   930
   End
End
Attribute VB_Name = "frm090612"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ; lbl1(index); Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
'Modified by Lydia 2017/05/05 strTemp(0 To 17) => strTemp(0 To 18)
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 18) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 18) As String, StrSQL3 As String, k As Integer
Dim PLeft(0 To 17) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As String, SeekTemp2 As String
Dim bol911001checkRange As Boolean
'Add By Cheng 2003/07/07
Dim m_blnPrintPart1 As Boolean
Dim m_blnPrintPart2 As Boolean
Dim m_blnPrintPart3 As Boolean
Dim m_blnPrintPart4 As Boolean
Dim m_blnPrintPart5 As Boolean 'Added by Morgan 2021/8/4
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限


Private Sub Check1_Click(Index As Integer)
If Check1(Index).Value = 1 Then
    Select Case Index
    Case 0
        If Mid(Txt1(0).Text, 1, 3) = "CFP" Then
            Txt1(11).Text = "40"
        Else
            Txt1(11).Text = "3"
        End If
    Case 1
        'Modified by Morgan 2024/3/19 '多系統時預設7天--柏翰
        'If Mid(Txt1(0).Text, 1, 3) = "CFP" Then
        If Txt1(0).Text = "CFP" Then
            Txt1(12).Text = "14"
        Else
            Txt1(12).Text = "7"
        End If
    Case 2
        If Mid(Txt1(0).Text, 1, 3) = "CFP" Then
            Txt1(13).Text = "30"
        Else
            Txt1(13).Text = "30"
        End If
    Case 3
        If Mid(Txt1(0).Text, 1, 3) = "CFP" Then
            Txt1(14).Text = "4"
        Else
            Txt1(14).Text = "4"
        End If
    'Added by Morgan 2021/8/4 已完稿未會稿
    Case 4
        'Modified by Morgan 2024/3/19 '多系統時預設4天--柏翰
        'If Mid(Txt1(0).Text, 1, 3) = "CFP" Then
        If Txt1(0).Text = "CFP" Then
            Txt1(17).Text = "6"
        Else
            Txt1(17).Text = "4"
        End If
    Case Else
    End Select
'Add By Cheng 2003/05/06
Else
    Select Case Index
    Case 0
        Txt1(11).Text = ""
    Case 1
        Txt1(12).Text = ""
    Case 2
        Txt1(13).Text = ""
    Case 3
        Txt1(14).Text = ""
    'Added by Morgan 2021/8/4
    Case 4
        Txt1(17).Text = ""
    Case Else
    End Select
End If
End Sub

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
             If Len(Txt1(10)) = 0 Then
                 s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                 Txt1(10).SetFocus
                 Exit Sub
             Else
                 'Modified by Morgan 2021/8/4 +Check1(4)
                 If Check1(0).Value = 0 And Check1(1).Value = 0 And Check1(2).Value = 0 And Check1(3).Value = 0 And Check1(4).Value = 0 Then
                     s = MsgBox("承辦天數統計範圍至少核取一項!!", , "USER 輸入錯誤")
                     Check1(0).SetFocus
                     Exit Sub
                 Else
                     If Check1(0).Value = 1 Then
                         If Len(Txt1(11)) = 0 Then
                            s = MsgBox(Check1(0).Caption & " 不可空白!!", , "USER 輸入錯誤")
                            Txt1(11).SetFocus
                            Exit Sub
                         End If
                     End If
                     If Check1(1).Value = 1 Then
                         If Len(Txt1(12)) = 0 Then
                            s = MsgBox(Check1(1).Caption & " 不可空白!!", , "USER 輸入錯誤")
                            Txt1(12).SetFocus
                            Exit Sub
                         End If
                     End If
                     If Check1(2).Value = 1 Then
                         If Len(Txt1(13)) = 0 Then
                            s = MsgBox(Check1(2).Caption & " 不可空白!!", , "USER 輸入錯誤")
                            Txt1(13).SetFocus
                            Exit Sub
                         End If
                     End If
                     If Check1(3).Value = 1 Then
                         If Len(Txt1(14)) = 0 Then
                            s = MsgBox(Check1(3).Caption & " 不可空白!!", , "USER 輸入錯誤")
                            Txt1(14).SetFocus
                            Exit Sub
                         End If
                     End If
                                          
                     'Added by Morgan 2021/8/4
                     If Check1(4).Value = 1 Then
                         If Len(Txt1(17)) = 0 Then
                            s = MsgBox(Check1(4).Caption & " 不可空白!!", , "USER 輸入錯誤")
                            Txt1(17).SetFocus
                            Exit Sub
                         End If
                     End If
                     
                     ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/14 清除查詢印表記錄檔欄位
                     If Val(Txt1(10)) = 1 Then
                        pub_QL05 = pub_QL05 & ";" & Label1(22) & "1.螢幕" 'Add By Sindy 2010/12/14
                     Else
                        pub_QL05 = pub_QL05 & ";" & Label1(22) & "2.報表" 'Add By Sindy 2010/12/14
                     End If
                     Screen.MousePointer = vbHourglass
                     Me.Enabled = False
                     
'                     For i = 0 To 14
'                        If (StrTemp99(i) <> Txt1(i)) And (i <> 10) Then
                            Process
'                            For j = 0 To 14
'                                StrTemp99(j) = Txt1(j)
'                            Next j
'                            Exit For
'                        End If
'                     Next i

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
If Val(Txt1(10)) = 1 Then
    Me.Hide
    frm090612_1.Show
   'Added by Morgan 2024/3/15
   If Left(Pub_StrUserSt03, 2) = "P1" And (Check1(1).Value Or Check1(4).Value) Then
      frm090612_1.cmdOK(6).Visible = True
   End If
   'end 2024/3/15
Else
    PrintData
End If
End Sub

Sub StrMenu1()
i = 0
j = 0
If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='1' AND (R107001='' OR R107001 IS NULL) AND R107018='#' "
Else
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='1' AND R107001='" & strTemp3 & "' AND R107018='#' "
End If
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        i = Val(CheckStr(.Fields(0)))
    End If
End With
CheckOC2
If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='1' AND (R107001='' OR R107001 IS NULL) "
Else
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='1' AND R107001='" & strTemp3 & "' "
End If
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        j = Val(CheckStr(.Fields(0)))
    End If
End With
CheckOC2

'Frame1(0).Caption = "超過已收文未齊備之時限申請案  天, 非申請案  天, 申請案共  件, 非申請案共  件 "
If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107007,R107008,R107016 FROM R090612 WHERE R107017='1' AND ID='" & strUserNum & "' AND (R107001 IS NULL OR R107001='') ORDER BY 1"
Else
    strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107007,R107008,R107016 FROM R090612 WHERE R107017='1' AND ID='" & strUserNum & "' AND R107001='" & strTemp3 & "' ORDER BY 1"
End If
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        '93.6.15 CANCEL BY SONIA
        'If Page <> 1 Then
        '    Page = Page + 1
        '    Printer.NewPage
        'End If
        '93.6.15 END
        PrintTitle
        m_blnPrintPart1 = True
        .MoveFirst
        SeekTemp2 = "超過已收文未齊備之時限申請案 " & frm090612.Txt1(11).Text & " 天, 非申請案 " & frm090612.Txt1(11).Text & " 天, 申請案共 " & str(i) & " 件, 非申請案共 " & str(j - i) & " 件 "
        PrintTitle_1
        Do While .EOF = False
            For i = 0 To 7
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = StrToStr(strTemp(2), 18)
            strTemp(3) = StrToStr(strTemp(3), 6)
            strTemp(4) = StrToStr(strTemp(4), 6)
            strTemp(7) = StrToStr(strTemp(7), 4)
            PrintDatil_1
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle_1
            End If
            .MoveNext
        Loop
        ShowLine
        iPrint = iPrint + 300
    Else
        m_blnPrintPart1 = False
        SeekTemp2 = "超過已收文未齊備之時限申請案  天, 非申請案  天, 申請案共  件, 非申請案共  件 "
    End If
End With
CheckOC2
End Sub

Sub PrintDatil_1() '列印資料
For i = 0 To 7
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub StrMenu2()
'第2個
j = 0
If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='2' AND (R107001='' OR R107001 IS NULL) "
Else
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='2' AND R107001='" & strTemp3 & "' "
End If
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        j = Val(CheckStr(.Fields(0)))
    End If
End With
CheckOC2

If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107009,R107010,R107016 FROM R090612 WHERE R107017='2' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY 1"
Else
    strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107009,R107010,R107016 FROM R090612 WHERE R107017='2' AND ID='" & strUserNum & "' AND R107001='" & strTemp3 & "' ORDER BY 1"
End If
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        If m_blnPrintPart1 = False Then
            '93.6.15 CANCEL BY SONIA
            'If Page <> 1 Then
            '    Page = Page + 1
            '    Printer.NewPage
            'End If
            '93.6.15 END
            PrintTitle
        End If
        m_blnPrintPart2 = True
        .MoveFirst
        SeekTemp2 = "超過已齊備未完稿之時限 " & frm090612.Txt1(12).Text & " 天, 共 " & str(j) & " 件 "
        PrintTitle_2
        Do While .EOF = False
            For i = 0 To 7
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = StrToStr(strTemp(2), 18)
            strTemp(3) = StrToStr(strTemp(3), 6)
            strTemp(4) = StrToStr(strTemp(4), 6)
            strTemp(7) = StrToStr(strTemp(7), 4)
            PrintDatil_2
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle_2
            End If
            .MoveNext
        Loop
        ShowLine
        iPrint = iPrint + 300
    Else
        m_blnPrintPart2 = False
        SeekTemp2 = "超過已齊備未完稿之時限  天, 共  件 "
    End If
End With
CheckOC2

End Sub

Sub PrintDatil_2() '列印資料
For i = 0 To 7
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub StrMenu3()
'第3個
j = 0
If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='3' AND (R107001='' OR R107001 IS NULL) "
Else
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='3' AND R107001='" & strTemp3 & "' "
End If
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        j = Val(CheckStr(.Fields(0)))
    End If
End With
CheckOC2

If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107013,R107016 FROM R090612 WHERE R107017='3' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY 1"
Else
    strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107013,R107016 FROM R090612 WHERE R107017='3' AND ID='" & strUserNum & "' AND R107001='" & strTemp3 & "' ORDER BY 1"
End If
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        'Modified by Morgan 2021/8/4
        'If m_blnPrintPart1 = False And m_blnPrintPart2 = False Then
        If m_blnPrintPart1 = False And m_blnPrintPart2 = False And m_blnPrintPart5 = False Then
        'end 2021/8/4
            '93.6.15 CANCEL BY SONIA
            'If Page <> 1 Then
            '    Page = Page + 1
            '    Printer.NewPage
            'End If
            '93.6.15 END
            PrintTitle
        End If
        m_blnPrintPart3 = True
        .MoveFirst
        SeekTemp2 = "超過已會稿未會稿完成之時限 " & frm090612.Txt1(13).Text & " 天, 共 " & str(j) & " 件 "
        PrintTitle_3
        Do While .EOF = False
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = StrToStr(strTemp(2), 14)
            strTemp(3) = StrToStr(strTemp(3), 4)
            strTemp(4) = StrToStr(strTemp(4), 6)
            strTemp(8) = StrToStr(strTemp(8), 4)
            PrintDatil_3
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle_3
            End If
            .MoveNext
        Loop
        ShowLine
        iPrint = iPrint + 300
    Else
        m_blnPrintPart3 = False
        SeekTemp2 = "超過已會稿未會稿完成之時限  天, 共  件 "
    End If
End With
CheckOC2

End Sub

Sub PrintDatil_3() '列印資料
For i = 0 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub


Sub StrMenu4()
'第4個
j = 0
If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='4' AND (R107001='' OR R107001 IS NULL) "
Else
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='4' AND R107001='" & strTemp3 & "' "
End If
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        j = Val(CheckStr(.Fields(0)))
    End If
End With
CheckOC2

If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107014,R107015,R107016 FROM R090612 WHERE R107017='4' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY 1"
Else
    strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107011,R107012,R107014,R107015,R107016 FROM R090612 WHERE R107017='4' AND ID='" & strUserNum & "' AND R107001='" & strTemp3 & "' ORDER BY 1"
End If
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        If m_blnPrintPart1 = False And m_blnPrintPart2 = False And m_blnPrintPart5 = False And m_blnPrintPart3 = False Then
            '93.6.15 CANCEL BY SONIA
            'If Page <> 1 Then
            '    Page = Page + 1
            '    Printer.NewPage
            'End If
            '93.6.15 END
            PrintTitle
        End If
        m_blnPrintPart4 = True
        .MoveFirst
        SeekTemp2 = "超過已會稿完成未發文之時限 " & frm090612.Txt1(14).Text & " 天, 共 " & str(j) & " 件 "
        PrintTitle_4
        Do While .EOF = False
            For i = 0 To 9
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = StrToStr(strTemp(2), 10)
            strTemp(3) = StrToStr(strTemp(3), 4)
            strTemp(4) = StrToStr(strTemp(4), 6)
            strTemp(9) = StrToStr(strTemp(9), 4)
            PrintDatil_4
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle_4
            End If
            .MoveNext
        Loop
        ShowLine
        iPrint = iPrint + 300
    Else
        m_blnPrintPart4 = False
        SeekTemp2 = "超過已會稿完成未發文之時限  天, 共  件 "
    End If
End With
CheckOC2
End Sub

Sub PrintDatil_4() '列印資料

For i = 0 To 9
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub PrintTitle() '列印抬頭

iPrint = 0
'93.6.15 CANCEL BY SONIA
'Printer.Orientation = 2
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
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "年月：" & Format(Format(GetTodayDate, "####/##/##"), "ee/MM")
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
If Val(Txt1(5)) = 1 Then
    Printer.Print "承辦人：" & strTemp3
Else
    Printer.Print "智權人員：" & strTemp3
End If
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
PLeft(2) = 3500
PLeft(3) = 8000
PLeft(4) = 9500
PLeft(5) = 11000
PLeft(6) = 12500
PLeft(7) = 14000
End Sub

Sub GetPleft_2()
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 3500
PLeft(3) = 8000
PLeft(4) = 9500
PLeft(5) = 11000
PLeft(6) = 12500
PLeft(7) = 14000
End Sub

Sub GetPleft_3()
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 3500
PLeft(3) = 7000
PLeft(4) = 8000
PLeft(5) = 9500
PLeft(6) = 11000
PLeft(7) = 12500
PLeft(8) = 14000
End Sub

Sub GetPleft_4()
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1000
PLeft(2) = 3500
PLeft(3) = 6000
PLeft(4) = 7400
PLeft(5) = 8400
PLeft(6) = 9800
PLeft(7) = 11200
PLeft(8) = 12700
PLeft(9) = 14000
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
    Exit Sub 'Added by Morgan 2021/8/4
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "目次"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "收文天數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
If Val(Txt1(5)) = 1 Then
    Printer.Print "智權人員"
Else
    Printer.Print "承辦人"
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
    PrintTitle_2
    Exit Sub 'Added by Morgan 2021/8/4
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "目次"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "齊備天數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
If Val(Txt1(5)) = 1 Then
    Printer.Print "智權人員"
Else
    Printer.Print "承辦人"
End If
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_2
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_2
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
    PrintTitle_3
    Exit Sub 'Added by Morgan 2021/8/4
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "目次"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "會稿日"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "會稿天數"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
If Val(Txt1(5)) = 1 Then
    Printer.Print "智權人員"
Else
    Printer.Print "承辦人"
End If
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_3
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_3
    Exit Sub
End If
End Sub
Sub PrintTitle_4() '列印抬頭

GetPleft_4
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print SeekTemp2
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_4
    Exit Sub 'Added by Morgan 2021/8/4
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "目次"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "會稿日"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "會稿完成日"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "會完天數"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
If Val(Txt1(5)) = 1 Then
    Printer.Print "智權人員"
Else
    Printer.Print "承辦人"
End If
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_4
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_4
    Exit Sub
End If
End Sub

'Added by Morgan 2021/8/4 已完稿未會稿
Sub PrintTitle_5() '列印抬頭

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
    PrintTitle_5
    Exit Sub 'Added by Morgan 2021/8/4
End If
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "目次"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "完稿天數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
If Val(Txt1(5)) = 1 Then
    Printer.Print "智權人員"
Else
    Printer.Print "承辦人"
End If
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_5
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    PrintTitle_5
    Exit Sub
End If
End Sub

Sub StrMenu5()
j = 0
If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='5' AND (R107001='' OR R107001 IS NULL) "
Else
    strSql = "SELECT COUNT(*) FROM R090612 WHERE ID='" & strUserNum & "' AND R107017='5' AND R107001='" & strTemp3 & "' "
End If
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        j = Val(CheckStr(.Fields(0)))
    End If
End With
CheckOC2

If Len(Trim(strTemp3)) = 0 Then
    strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107011,R107013,R107016 FROM R090612 WHERE R107017='5' AND ID='" & strUserNum & "' AND (R107001='' OR R107001 IS NULL) ORDER BY 1"
Else
    strSql = "SELECT R107002,R107003,R107004,R107005,R107006,R107011,R107013,R107016 FROM R090612 WHERE R107017='5' AND ID='" & strUserNum & "' AND R107001='" & strTemp3 & "' ORDER BY 1"
End If
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        If m_blnPrintPart1 = False And m_blnPrintPart2 = False Then
            PrintTitle
        End If
        m_blnPrintPart5 = True
        .MoveFirst
        SeekTemp2 = "超過已完稿未會稿之時限 " & frm090612.Txt1(17).Text & " 天, 共 " & str(j) & " 件 "
        PrintTitle_5
        Do While .EOF = False
            For i = 0 To 7
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = StrToStr(strTemp(2), 18)
            strTemp(3) = StrToStr(strTemp(3), 6)
            strTemp(4) = StrToStr(strTemp(4), 6)
            strTemp(7) = StrToStr(strTemp(7), 4)
            PrintDatil_2
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
                PrintTitle_5
            End If
            .MoveNext
        Loop
        ShowLine
        iPrint = iPrint + 300
    Else
        m_blnPrintPart5 = False
        SeekTemp2 = "超過已完稿未會稿之時限  天, 共  件 "
    End If
End With
CheckOC2

End Sub
'end 2021/8/4

Sub PrintData()
strSql = "select distinct r107001 from r090612 where id='" & strUserNum & "' "
CheckOC
Page = 1
strTemp3 = " "
'93.6.15 ADD BY SONIA
Printer.Orientation = 2
'93.6.15 END
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            'Add By Cheng 2003/07/07
            '預設無資料列印
            m_blnPrintPart1 = False: m_blnPrintPart2 = False: m_blnPrintPart3 = False: m_blnPrintPart4 = False: m_blnPrintPart5 = False
            strTemp3 = CheckStr(.Fields(0))
'            If Page <> 1 Then
'                Page = Page + 1
'                Printer.NewPage
'            End If
'            PrintTitle
            If frm090612.Check1(0).Value = 1 Then
                StrMenu1
'                ShowLine
'                iPrint = iPrint + 300
            End If
            If frm090612.Check1(1).Value = 1 Then
                StrMenu2
'                ShowLine
'                iPrint = iPrint + 300
            End If
'            Page = Page + 1
'            Printer.NewPage
'            PrintTitle

            'Added by Morgan 2021/8/4 已完稿未會稿
            If frm090612.Check1(4).Value = 1 Then
                StrMenu5
            End If
            'end 2021/8/4
            
            If frm090612.Check1(2).Value = 1 Then
                StrMenu3
'                ShowLine
'                iPrint = iPrint + 300
            End If
            If frm090612.Check1(3).Value = 1 Then
                StrMenu4
'                ShowLine
'                iPrint = iPrint + 300
            End If
            '93.6.15 ADD BY SONIA
            
            Page = Page + 1
            Printer.NewPage
            '93.6.15 END
            .MoveNext
        Loop
        '93.6.15 ADD BY SONIA
        ShowPrintOk
        '93.6.15 END
    Else
      ShowNoData
      Exit Sub
    End If
End With
CheckOC
Printer.EndDoc
End Sub
'Modified by Lydia 2017/05/05 備份
'Sub Process()
Private Sub Process_old()
cnnConnection.Execute "DELETE FROM R090612 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
StrSQL3 = ""
StrSQL4 = ""
strSQL5 = ""
If Len(Txt1(0)) <> 0 Then
    strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(Txt1(0), 1) & ") "
    strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(Txt1(0), 2) & ") "
    StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(Txt1(0), 3) & ") "
    StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(Txt1(0), 4) & ") "
    strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(Txt1(0), 5) & ") "
    pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0) 'Add By Sindy 2010/12/14
End If
StrSQL6 = ""
If Len(Txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & Txt1(1) & "' "
    strSQL2 = strSQL2 + " AND TM10>='" & Txt1(1) & "' "
    StrSQL3 = StrSQL3 + " AND LC15>='" & Txt1(1) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09>='" & Txt1(1) & "' "
End If
If Len(Txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & Txt1(2) & "' "
    strSQL2 = strSQL2 + " AND TM10<='" & Txt1(2) & "' "
    StrSQL3 = StrSQL3 + " AND LC15<='" & Txt1(2) & "' "
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09<='" & Txt1(2) & "' "
End If
If Len(Txt1(1)) <> 0 Or Len(Txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & Txt1(1) & "-" & Txt1(2) 'Add By Sindy 2010/12/14
End If
If Val(Txt1(5)) = 1 Then
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "1.承辦人" 'Add By Sindy 2010/12/14
    If Len(Txt1(3)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s1.st06>='" & Txt1(3) & "' "
    End If
    If Len(Txt1(4)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s1.st06<='" & Txt1(4) & "' "
    End If
Else
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "2.智權人員" 'Add By Sindy 2010/12/14
    If Len(Txt1(3)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s2.st06>='" & Txt1(3) & "' "
    End If
    If Len(Txt1(4)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s2.st06<='" & Txt1(4) & "' "
    End If
End If
If Len(Txt1(3)) <> 0 Or Len(Txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & Txt1(3) & "-" & Txt1(4) & Label1(24) 'Add By Sindy 2010/12/14
End If
If Len(Txt1(6)) <> 0 Then
   StrSQL6 = StrSQL6 + " and ep05='" & Txt1(6) & "' "
   pub_QL05 = pub_QL05 & ";" & Label1(8) & Txt1(6) & lbl1(0) 'Add By Sindy 2010/12/14
End If
If Len(Txt1(7)) <> 0 Then
   StrSQL6 = StrSQL6 + " and cp12>='" & Txt1(7) & "' "
End If
If Len(Txt1(8)) <> 0 Then
   StrSQL6 = StrSQL6 + " and cp12<='" & Txt1(8) & "' "
End If
If Len(Txt1(7)) <> 0 Or Len(Txt1(8)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(9) & Txt1(7) & "-" & Txt1(8) 'Add By Sindy 2010/12/14
End If
If Len(Txt1(9)) <> 0 Then
   StrSQL6 = StrSQL6 + " and cp13='" & Txt1(9) & "' "
   pub_QL05 = pub_QL05 & ";" & Label1(6) & Txt1(9) & lbl1(1) 'Add By Sindy 2010/12/14
End If

'Modify By Cheng 2003/05/06
'取消是否算案件數的限制
'StrSQL6 = StrSQL6 + " and CP26 IS NULL  "
'modify by sonia 2017/3/8 取消cp05條件並後面語法皆改用cp158及cp159判斷
'StrSQL6 = StrSQL6 & " and cp05>=19980101 and cp27 is null and cp57 is null "
StrSQL6 = StrSQL6 & " and cp158=0 and cp159=0 "

If Val(Txt1(5)) = 1 Then
    '承辦人
    strSql = ""
    '已收文未齊備
    If Check1(0).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(0).Caption & Txt1(11) & Label2(0) 'Add By Sindy 2010/12/14
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
                   strSql = strSql + " select s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp05,0,0,0,0,0,0,0,0,s2.st02,'1',ep02,cp10,cp01 from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where ep02=cp09(+) and CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and cp05 is not null and EP06 IS NULL  " & strSQL1 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),cp05,0,0,0,0,0,0,0,0,s2.st02,'1',ep02,cp10,cp01 from engineerprogress,caseprogress,staff s1,staff s2,trademark,casepropertymap,patenttrademarkmap where ep02=cp09(+) and CP01=TM01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) and cp05 is not null and EP06 IS NULL  " & strSQL2 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),cp05,0,0,0,0,0,0,0,0,s2.st02,'1',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,lawcase,casepropertymap where ep02=cp09(+) and CP01=LC01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp05 is not null and EP06 IS NULL   " & StrSQL3 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,cp05,0,0,0,0,0,0,0,0,s2.st02,'1',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,hirecase,casepropertymap where ep02=cp09(+) and CP01=HC01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp05 is not null and EP06 IS NULL  " & StrSQL4 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),cp05,0,0,0,0,0,0,0,0,s2.st02,'1',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,servicepractice,casepropertymap where ep02=cp09(+) and CP01=SP01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp05 is not null and EP06 IS NULL  " & strSQL5 & StrSQL6
    End If
    '已齊備未完稿
    If Check1(1).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(1).Caption & Txt1(12) & Label2(1) 'Add By Sindy 2010/12/14
        If Len(strSql) <> 0 Then
            strSql = strSql + " union "
        End If
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
                   strSql = strSql + " select s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,ep06,0,0,0,0,0,0,s2.st02,'2',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) AND CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep06 is not null and EP09 IS NULL " & strSQL1 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),0,0,ep06,0,0,0,0,0,0,s2.st02,'2',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,trademark,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=TM01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) and ep06 is not null and EP09 IS NULL " & strSQL2 & StrSQL6
        strSql = strSql + " union      select s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),0,0,ep06,0,0,0,0,0,0,s2.st02,'2',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,lawcase,casepropertymap where EP02=CP09(+) and CP01=LC01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep06 is not null and EP09 IS NULL  " & StrSQL3 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,0,0,ep06,0,0,0,0,0,0,s2.st02,'2',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,hirecase,casepropertymap where EP02=CP09(+) and CP01=HC01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep06 is not null and EP09 IS NULL " & StrSQL4 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),0,0,ep06,0,0,0,0,0,0,s2.st02,'2',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,servicepractice,casepropertymap where EP02=CP09(+) and CP01=SP01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep06 is not null and EP09 IS NULL " & strSQL5 & StrSQL6
    End If
    '已會稿未會稿完成
    If Check1(2).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(2).Caption & Txt1(13) & Label2(2) 'Add By Sindy 2010/12/14
        If Len(strSql) <> 0 Then
            strSql = strSql + " union "
        End If
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
                   strSql = strSql + " select s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,0,0,s2.st02,'3',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep07 is not null and EP08 IS NULL  " & strSQL1 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,0,0,s2.st02,'3',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,trademark,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=TM01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) and ep07 is not null and EP08 IS NULL  " & strSQL2 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,0,0,s2.st02,'3',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,lawcase,casepropertymap where EP02=CP09(+) and CP01=LC01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep07 is not null and EP08 IS NULL   " & StrSQL3 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,0,0,0,0,ep09,ep07,0,0,0,s2.st02,'3',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,hirecase,casepropertymap where EP02=CP09(+) and CP01=HC01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep07 is not null and EP08 IS NULL  " & StrSQL4 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,0,0,s2.st02,'3',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,servicepractice,casepropertymap where EP02=CP09(+) and CP01=SP01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep07 is not null and EP08 IS NULL  " & strSQL5 & StrSQL6
    End If
    '已會稿完成未發文
    If Check1(3).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(3).Caption & Txt1(14) & Label2(3) 'Add By Sindy 2010/12/14
        If Len(strSql) <> 0 Then
            strSql = strSql + " union "
        End If
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
                   strSql = strSql + " select s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,ep08,0,s2.st02,'4',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep08 is not null and CP158=0 " & strSQL1 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,ep08,0,s2.st02,'4',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,trademark,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=TM01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) and ep08 is not null and CP158=0 " & strSQL2 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,ep08,0,s2.st02,'4',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,lawcase,casepropertymap where EP02=CP09(+) and CP01=LC01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep08 is not null and CP158=0 " & StrSQL3 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,0,0,0,0,ep09,ep07,0,ep08,0,s2.st02,'4',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,hirecase,casepropertymap where EP02=CP09(+) and CP01=HC01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep08 is not null and CP158=0 " & StrSQL4 & StrSQL6
        strSql = strSql + " UNION all  SELECT s1.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,ep08,0,s2.st02,'4',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,servicepractice,casepropertymap where EP02=CP09(+) and CP01=SP01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep08 is not null and CP158=0 " & strSQL5 & StrSQL6
    End If
    CheckOC
Else
    '智權人員
    strSql = ""
    '已收文未齊備
    If Check1(0).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(0).Caption & Txt1(11) & Label2(0) 'Add By Sindy 2010/12/14
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
                   strSql = strSql + " select s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),cp05,0,0,0,0,0,0,0,0,s1.st02,'1',ep02,cp10,cp01 from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and cp05 is not null and EP06 IS NULL  " & strSQL1 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),cp05,0,0,0,0,0,0,0,0,s1.st02,'1',ep02,cp10,cp01 from engineerprogress,caseprogress,staff s1,staff s2,trademark,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=TM01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) and cp05 is not null and EP06 IS NULL  " & strSQL2 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),cp05,0,0,0,0,0,0,0,0,s1.st02,'1',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,lawcase,casepropertymap where EP02=CP09(+) and CP01=LC01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp05 is not null and EP06 IS NULL   " & StrSQL3 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,cp05,0,0,0,0,0,0,0,0,s1.st02,'1',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,hirecase,casepropertymap where EP02=CP09(+) and CP01=HC01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp05 is not null and EP06 IS NULL  " & StrSQL4 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),cp05,0,0,0,0,0,0,0,0,s1.st02,'1',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,servicepractice,casepropertymap where EP02=CP09(+) and CP01=SP01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp05 is not null and EP06 IS NULL  " & strSQL5 & StrSQL6
    End If
    '已齊備未完稿
    If Check1(1).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(1).Caption & Txt1(12) & Label2(1) 'Add By Sindy 2010/12/14
        If Len(strSql) <> 0 Then
            strSql = strSql + " union "
        End If
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
                   strSql = strSql + " select s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,ep06,0,0,0,0,0,0,s1.st02,'2',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep06 is not null and EP09 IS NULL " & strSQL1 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),0,0,ep06,0,0,0,0,0,0,s1.st02,'2',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,trademark,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=TM01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) and ep06 is not null and EP09 IS NULL " & strSQL2 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),0,0,ep06,0,0,0,0,0,0,s1.st02,'2',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,lawcase,casepropertymap where EP02=CP09(+) and CP01=LC01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep06 is not null and EP09 IS NULL  " & StrSQL3 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,0,0,ep06,0,0,0,0,0,0,s1.st02,'2',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,hirecase,casepropertymap where EP02=CP09(+) and CP01=HC01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep06 is not null and EP09 IS NULL " & StrSQL4 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),0,0,ep06,0,0,0,0,0,0,s1.st02,'2',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,servicepractice,casepropertymap where EP02=CP09(+) and CP01=SP01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep06 is not null and EP09 IS NULL " & strSQL5 & StrSQL6
    End If
    '已會稿未會稿完成
    If Check1(2).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(2).Caption & Txt1(13) & Label2(2) 'Add By Sindy 2010/12/14
        If Len(strSql) <> 0 Then
            strSql = strSql + " union "
        End If
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
                   strSql = strSql + " select s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,0,0,s1.st02,'3',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep07 is not null and EP08 IS NULL  " & strSQL1 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,0,0,s1.st02,'3',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,trademark,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=TM01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) and ep07 is not null and EP08 IS NULL  " & strSQL2 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,0,0,s1.st02,'3',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,lawcase,casepropertymap where EP02=CP09(+) and CP01=LC01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep07 is not null and EP08 IS NULL   " & StrSQL3 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,0,0,0,0,ep09,ep07,0,0,0,s1.st02,'3',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,hirecase,casepropertymap where EP02=CP09(+) and CP01=HC01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep07 is not null and EP08 IS NULL  " & StrSQL4 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,0,0,s1.st02,'3',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,servicepractice,casepropertymap where EP02=CP09(+) and CP01=SP01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep07 is not null and EP08 IS NULL  " & strSQL5 & StrSQL6
    End If
    '已會稿完成未發文
    If Check1(3).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(3).Caption & Txt1(14) & Label2(3) 'Add By Sindy 2010/12/14
        If Len(strSql) <> 0 Then
            strSql = strSql + " union "
        End If
      'Modify By Cheng 2002/04/26
      '若已閉卷, 則在本所案號後加"*"號
                   strSql = strSql + " select s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),DECODE(PA09,'000',PTM03,PTM04),decode(pa09,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,ep08,0,s1.st02,'4',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,patent,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  and pa08=ptm02(+) and ep08 is not null and CP158=0 " & strSQL1 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊',''),nvl(tm05,nvl(tm06,tm07)),decode(tm10,'000',ptm03,ptm04),decode(tm10,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,ep08,0,s1.st02,'4',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,trademark,casepropertymap,patenttrademarkmap where EP02=CP09(+) and CP01=TM01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+)  and tm08=ptm02(+) and ep08 is not null and CP158=0 " & strSQL2 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊',''),nvl(lc05,nvl(lc06,lc07)),'',decode(lc15,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,ep08,0,s1.st02,'4',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,lawcase,casepropertymap where EP02=CP09(+) and CP01=LC01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep08 is not null and CP158=0 " & StrSQL3 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊',''),hc06,'',cpm03,0,0,0,0,ep09,ep07,0,ep08,0,s1.st02,'4',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,hirecase,casepropertymap where EP02=CP09(+) and CP01=HC01(+) and cp02=hc02(+) and cp03=hc03(+) and cp04=hc04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep08 is not null and CP158=0 " & StrSQL4 & StrSQL6
        strSql = strSql + " UNION all  SELECT s2.st02,ep01,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊',''),nvl(sp05,nvl(sp06,sp07)),'',decode(sp09,'000',cpm03,cpm04),0,0,0,0,ep09,ep07,0,ep08,0,s1.st02,'4',ep02,'','' from engineerprogress,caseprogress,staff s1,staff s2,servicepractice,casepropertymap where EP02=CP09(+) and CP01=SP01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and ep05=s1.st01(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and ep08 is not null and CP158=0 " & strSQL5 & StrSQL6
    End If
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/14
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 16
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            TestOk = True
            If Len(strTemp(6)) <> 0 And Val(strTemp(6)) <> 0 Then
                strTemp(7) = str(GetWorkDay(GetTodayDate, strTemp(6)))
                If Val(Txt1(11)) > Val(strTemp(7)) Then
                    TestOk = False
                End If
                strTemp(6) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(6)))
            Else
                strTemp(6) = ""
            End If
            If Len(strTemp(8)) <> 0 And Val(strTemp(8)) <> 0 Then
                strTemp(9) = str(GetWorkDay(GetTodayDate, strTemp(8)))
                If Val(Txt1(12)) > Val(strTemp(9)) Then
                    TestOk = False
                End If
                strTemp(8) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(8)))
            Else
                strTemp(8) = ""
            End If
            If Len(strTemp(10)) <> 0 And Val(strTemp(10)) <> 0 Then
                strTemp(10) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(10)))
            Else
                strTemp(10) = ""
            End If
            If Len(strTemp(11)) <> 0 And Val(strTemp(11)) <> 0 Then
                strTemp(12) = str(GetWorkDay(GetTodayDate, strTemp(11)))
                If Val(Txt1(13)) > Val(strTemp(12)) Then
                    TestOk = False
                End If
                strTemp(11) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(11)))
            Else
                strTemp(11) = ""
            End If
            If Len(strTemp(13)) <> 0 And Val(strTemp(13)) <> 0 Then
                strTemp(14) = str(GetWorkDay(GetTodayDate, strTemp(13)))
                If Val(Txt1(14)) > Val(strTemp(14)) Then
                    TestOk = False
                End If
                strTemp(13) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(13)))
            Else
                strTemp(13) = ""
            End If
            
            strTemp(18) = "" & .Fields("ep02") 'Added by Lydia 2017/05/05 收文號
            If TestOk = True Then
                strTemp(17) = ""
                If Val(strTemp(16)) = 1 Then
                    Select Case CheckStr(.Fields(19))
                    Case "P", "FCP", "CFP"
                         Select Case Val(CheckStr(.Fields(18)))
                         Case 101, 102, 103, 104, 105
                              strTemp(17) = "#"
                         Case Else
                              strTemp(17) = ""
                         End Select
                    Case "T", "FCT", "CFT", "TF"
                         Select Case Val(CheckStr(.Fields(18)))
                         Case "101"
                              strTemp(17) = "#"
                         Case Else
                              strTemp(17) = ""
                         End Select
                    Case Else
                    End Select
                End If
                'Modified by Lydia 2017/05/05 +收文號
                'strSql = "insert into r090612 values ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & ChgSQL(strTemp(3)) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "'," & Val(strTemp(7)) & ",'" & strTemp(8) & "'," & Val(strTemp(9)) & ",'" & strTemp(10) & "','" & strTemp(11) & "'," & Val(strTemp(12)) & ",'" & strTemp(13) & "'," & Val(strTemp(14)) & ",'" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "','" & strUserNum & "') "
                strSql = "insert into r090612 values ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & ChgSQL(strTemp(3)) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "'," & Val(strTemp(7)) & ",'" & strTemp(8) & "'," & Val(strTemp(9)) & ",'" & strTemp(10) & "','" & strTemp(11) & "'," & Val(strTemp(12)) & ",'" & strTemp(13) & "'," & Val(strTemp(14)) & ",'" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "','" & strUserNum & "','" & strTemp(18) & "') "
                cnnConnection.Execute strSql
            End If
            .MoveNext
            DoEvents
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/14
    End If
End With
CheckOC
End Sub

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_Load()
m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
Txt1(0) = Systemkind_g
Select Case ProState
Case "1"     '個人
    Frame1.Left = 72
    'Modified by Lydia 2017/05/05
    'Frame1.Top = 1152
    Frame1.Top = 1440
    Txt1(5).Text = "1"
    Txt1(6).Text = strUserNum
    Frame2.Enabled = False
    Frame2.Visible = False
    'Me.Height = 3708
    'Me.Width = 4068
    Txt1(0).TabIndex = 0
    Txt1(1).TabIndex = 1
    Txt1(2).TabIndex = 2
    'Added by Lydia 2017/05/05
    Txt1(15).TabIndex = 3
    Txt1(16).TabIndex = 4
    'end 2017/05/05
    'Modified by Lydia 2017/05/05 Index + 2 => ex. Check1(0).TabIndex = 3 ->Check1(0).TabIndex = 5
    Check1(0).TabIndex = 5
    Txt1(11).TabIndex = 6
    Check1(1).TabIndex = 7
    Txt1(12).TabIndex = 8
    Check1(2).TabIndex = 9
    Txt1(13).TabIndex = 10
    Check1(3).TabIndex = 11
    Txt1(14).TabIndex = 12
    Txt1(10).TabIndex = 13
    cmdOK(0).TabIndex = 14
    cmdOK(1).TabIndex = 15
    'end 2017/05/05
    
    Frame3.Top = Frame1.Top + Frame1.Height + 100 'Added by Morgan 2021/8/4
    
Case "2"     '主管
    'Modified by Lydia 2017/05/05
    'Frame1.Top = 3312
    'Modified by Morgan 2021/8/4
    'Frame1.Top = 3740
    Frame1.Top = 3615
    'end 2021/8/4
    Frame1.Left = 72
    Frame2.Enabled = True
    'Me.Height = 5496
    'Me.Width = 1068
    Txt1(0).TabIndex = 0
    Txt1(1).TabIndex = 1
    Txt1(2).TabIndex = 2
    'Added by Lydia 2017/05/05
    Txt1(15).TabIndex = 3
    Txt1(16).TabIndex = 4
    'end 2017/05/05
    'Modified by Lydia 2017/05/05 Index + 2 => ex. Txt1(3).TabIndex = 3 -> Txt1(3).TabIndex = 5
    Txt1(3).TabIndex = 5
    Txt1(4).TabIndex = 6
    Txt1(5).TabIndex = 7
    Txt1(6).TabIndex = 8
    Txt1(7).TabIndex = 9
    Txt1(8).TabIndex = 10
    Txt1(9).TabIndex = 11
    Txt1(10).TabIndex = 12
    Check1(0).TabIndex = 13
    Txt1(11).TabIndex = 14
    Check1(1).TabIndex = 15
    Txt1(12).TabIndex = 16
    'Added by Morgan 2021/8/5
    Check1(1).TabIndex = 17
    Txt1(12).TabIndex = 18
    'end 2021/8/5
    Check1(2).TabIndex = 19
    Txt1(13).TabIndex = 20
    Check1(3).TabIndex = 21
    Txt1(14).TabIndex = 22
    cmdOK(0).TabIndex = 23
    cmdOK(1).TabIndex = 24
    'end 2017/05/05
Case Else
End Select
MoveFormToCenter Me
Txt1(10) = "1"
bol911001checkRange = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090612 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   Txt1(Index).SelStart = 0
   Txt1(Index).SelLength = Len(Txt1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
       cmdOK(0).SetFocus
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
     strTemp1 = Split(UCase(Systemkind_g), ",")
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
'Modified by Lydia 2017/05/05 +16(案件性質)
Case 2, 8, 16
     If RunNick(Txt1(Index - 1), Txt1(Index)) Then
        Txt1(Index - 1).SetFocus
        txt1_GotFocus (Index - 1)
        Exit Sub
     End If
Case 3
     bol911001checkRange = True
     Select Case Trim(Txt1(3))
     Case "1", "2", "", "3", "4", "5"
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          Txt1(3).SetFocus
          Txt1(3).SelStart = 0
          Txt1(3).SelLength = Len(Txt1(3))
          bol911001checkRange = False
          Exit Sub
     End Select
Case 4
     If bol911001checkRange = True Then
        Select Case Trim(Txt1(4))
        Case "1", "2", "", "3", "4", "5"
        Case Else
             s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
             Txt1(4).SetFocus
             Txt1(4).SelStart = 0
             Txt1(4).SelLength = Len(Txt1(4))
             Exit Sub
        End Select
        If RunNick(Txt1(3), Txt1(4)) Then
           Txt1(3).SetFocus
           txt1_GotFocus (3)
           Exit Sub
        End If
     End If
     bol911001checkRange = True
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
        If Trim(lbl1(0).Caption) = "" Then
            s = MsgBox("承辦人輸入錯誤！", , "錯誤！")
            Txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 9
     lbl1(1).Caption = GetPrjSalesNM(Txt1(9))
     If Trim(Txt1(Index)) <> "" Then
        If Trim(lbl1(1).Caption) = "" Then
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
            Txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 10
     Select Case Trim(Txt1(10))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示方式只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          Txt1(10).SetFocus
          Txt1(10).SelStart = 0
          Txt1(10).SelLength = Len(Txt1(10))
          Exit Sub
     End Select
Case 11, 12, 13, 14
     For i = 1 To Len(Txt1(Index))
          strSql = Mid(Txt1(Index), i, 1)
          If InStr(1, "0123456789 ", strSql) = 0 Then
                s = MsgBox("承辦天數只能輸入數字!!", , "USER 輸入錯誤")
                Txt1(Index).SetFocus
                Txt1(Index).SelStart = 0
                Txt1(Index).SelLength = Len(Txt1(Index))
                Exit Sub
          End If
     Next i
Case Else
End Select
End Sub

'Added by Lydia 2017/05/05 改成以案件進度為基礎
Private Sub Process()
cnnConnection.Execute "DELETE FROM R090612 WHERE ID='" & strUserNum & "' "
strSQL1 = " AND CP158=0 AND CP159=0" '已收文未發文
strSQL2 = strSQL1
StrSQL3 = strSQL1
StrSQL4 = strSQL1
strSQL5 = strSQL1
StrSQL6 = ""

'-------------系統別
If Len(Txt1(0)) <> 0 Then
    strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(Txt1(0), 1) & ")"
    strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(Txt1(0), 2) & ")"
    StrSQL3 = StrSQL3 & " AND CP01 IN (" & SQLGrpStr(Txt1(0), 3) & ")"
    StrSQL4 = StrSQL4 & " AND CP01 IN (" & SQLGrpStr(Txt1(0), 4) & ")"
    strSQL5 = strSQL5 & " AND CP01 IN (" & SQLGrpStr(Txt1(0), 5) & ")"
    pub_QL05 = pub_QL05 & ";" & Label1(0) & Txt1(0)
End If

'-------------國別
If Len(Txt1(1)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09>='" & Trim(Txt1(1)) & "'"
    strSQL2 = strSQL2 + " AND TM10>='" & Trim(Txt1(1)) & "'"
    StrSQL3 = StrSQL3 + " AND LC15>='" & Trim(Txt1(1)) & "'"
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09>='" & Trim(Txt1(1)) & "'"
End If
If Len(Txt1(2)) <> 0 Then
    strSQL1 = strSQL1 + " AND PA09<='" & Trim(Txt1(2)) & "'"
    strSQL2 = strSQL2 + " AND TM10<='" & Trim(Txt1(2)) & "'"
    StrSQL3 = StrSQL3 + " AND LC15<='" & Trim(Txt1(2)) & "'"
    StrSQL4 = StrSQL4
    strSQL5 = strSQL5 + " AND SP09<='" & Trim(Txt1(2)) & "'"
End If
If Len(Txt1(1)) <> 0 Or Len(Txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & Txt1(1) & "-" & Txt1(2)
End If

'-------------所別
If Val(Txt1(5)) = 1 Then
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "1.承辦人"
    If Len(Txt1(3)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s1.st06>='" & Trim(Txt1(3)) & "'"
    End If
    If Len(Txt1(4)) <> 0 Then
        StrSQL6 = StrSQL6 + " and s1.st06<='" & Trim(Txt1(4)) & "'"
    End If
Else
    pub_QL05 = pub_QL05 & ";" & Label1(3) & "2.智權人員"
    If Len(Txt1(3)) <> 0 Then
        strSQL1 = strSQL1 + " and s2.st06>='" & Trim(Txt1(3)) & "'"
        strSQL2 = strSQL2 + " and s2.st06>='" & Trim(Txt1(3)) & "'"
        StrSQL3 = StrSQL3 + " and s2.st06>='" & Trim(Txt1(3)) & "'"
        StrSQL4 = StrSQL4 + " and s2.st06>='" & Trim(Txt1(3)) & "'"
        strSQL5 = strSQL5 + " and s2.st06>='" & Trim(Txt1(3)) & "'"
    End If
    If Len(Txt1(4)) <> 0 Then
        strSQL1 = strSQL1 + " and s2.st06<='" & Trim(Txt1(4)) & "'"
        strSQL2 = strSQL2 + " and s2.st06<='" & Trim(Txt1(4)) & "'"
        StrSQL3 = StrSQL3 + " and s2.st06<='" & Trim(Txt1(4)) & "'"
        StrSQL4 = StrSQL4 + " and s2.st06<='" & Trim(Txt1(4)) & "'"
        strSQL5 = strSQL5 + " and s2.st06<='" & Trim(Txt1(4)) & "'"
    End If
End If
If Len(Txt1(3)) <> 0 Or Len(Txt1(4)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & Txt1(3) & "-" & Txt1(4) & Label1(24)
End If

'-------------承辦人
If Len(Txt1(6)) <> 0 Then
   StrSQL6 = StrSQL6 + " and ep05='" & Trim(Txt1(6)) & "'"
   pub_QL05 = pub_QL05 & ";" & Label1(8) & Txt1(6) & lbl1(0)
End If

'-------------業務區
If Len(Txt1(7)) <> 0 Then
   strSQL1 = strSQL1 + " and cp12>='" & Trim(Txt1(7)) & "'"
   strSQL2 = strSQL2 + " and cp12>='" & Trim(Txt1(7)) & "'"
   StrSQL3 = StrSQL3 + " and cp12>='" & Trim(Txt1(7)) & "'"
   StrSQL4 = StrSQL4 + " and cp12>='" & Trim(Txt1(7)) & "'"
   strSQL5 = strSQL5 + " and cp12>='" & Trim(Txt1(7)) & "'"
End If
If Len(Txt1(8)) <> 0 Then
   strSQL1 = strSQL1 + " and cp12<='" & Trim(Txt1(8)) & "'"
   strSQL2 = strSQL2 + " and cp12<='" & Trim(Txt1(8)) & "'"
   StrSQL3 = StrSQL3 + " and cp12<='" & Trim(Txt1(8)) & "'"
   StrSQL4 = StrSQL4 + " and cp12<='" & Trim(Txt1(8)) & "'"
   strSQL5 = strSQL5 + " and cp12<='" & Trim(Txt1(8)) & "'"
End If
If Len(Txt1(7)) <> 0 Or Len(Txt1(8)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(9) & Txt1(7) & "-" & Txt1(8)
End If

'-------------智權人員
If Len(Txt1(9)) <> 0 Then
   strSQL1 = strSQL1 + " and cp13='" & Trim(Txt1(9)) & "'"
   strSQL2 = strSQL2 + " and cp13='" & Trim(Txt1(9)) & "'"
   StrSQL3 = StrSQL3 + " and cp13='" & Trim(Txt1(9)) & "'"
   StrSQL4 = StrSQL4 + " and cp13='" & Trim(Txt1(9)) & "'"
   strSQL5 = strSQL5 + " and cp13='" & Trim(Txt1(9)) & "'"
   pub_QL05 = pub_QL05 & ";" & Label1(6) & Txt1(9) & lbl1(1)
End If

'-------------案件性質
If Len(Trim(Txt1(15))) <> 0 Then
   strSQL1 = strSQL1 + " and cp10>='" & Trim(Txt1(15)) & "'"
   strSQL2 = strSQL2 + " and cp10>='" & Trim(Txt1(15)) & "'"
   StrSQL3 = StrSQL3 + " and cp10>='" & Trim(Txt1(15)) & "'"
   StrSQL4 = StrSQL4 + " and cp10>='" & Trim(Txt1(15)) & "'"
   strSQL5 = strSQL5 + " and cp10>='" & Trim(Txt1(15)) & "'"
End If
If Len(Trim(Txt1(16))) <> 0 Then
   strSQL1 = strSQL1 + " and cp10<='" & Trim(Txt1(16)) & "'"
   strSQL2 = strSQL2 + " and cp10<='" & Trim(Txt1(16)) & "'"
   StrSQL3 = StrSQL3 + " and cp10<='" & Trim(Txt1(16)) & "'"
   StrSQL4 = StrSQL4 + " and cp10<='" & Trim(Txt1(16)) & "'"
   strSQL5 = strSQL5 + " and cp10<='" & Trim(Txt1(16)) & "'"
End If
If Len(Trim(Txt1(15))) <> 0 Or Len(Trim(Txt1(16))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & Txt1(15) & "-" & Txt1(16)
End If

'----------抓案件進度,基本檔
strSQL1 = " select cp01,cp02,cp03,cp04,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(PA57,'Y','＊','') caseno,nvl(pa05,nvl(pa06,pa07)) casename,DECODE(PA09,'000',PTM03,PTM04) ptm0304,decode(pa09,'000',cpm03,cpm04) cpm0304,cp05,s2.st02 智權人員,'1',cp10 from caseprogress,staff s2,patent,casepropertymap,patenttrademarkmap where CP01=PA01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) and pa08=ptm02(+) and cp05 is not null" & strSQL1
strSQL2 = " select cp01,cp02,cp03,cp04,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(TM29,'Y','＊','') caseno,nvl(TM05,nvl(TM06,TM07)) casename,DECODE(TM10,'000',PTM03,PTM04) ptm0304,decode(TM10,'000',cpm03,cpm04) cpm0304,cp05,s2.st02 智權人員,'1',cp10 from caseprogress,staff s2,trademark,casepropertymap,patenttrademarkmap where CP01=TM01(+) and cp02=TM02(+) and cp03=TM03(+) and cp04=TM04(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+) and tm09=ptm02(+) and cp05 is not null" & strSQL2
StrSQL3 = " select cp01,cp02,cp03,cp04,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(LC08,'Y','＊','') caseno,nvl(LC05,nvl(LC06,LC07)) casename,'' ptm0304,decode(LC15,'000',cpm03,cpm04) cpm0304,cp05,s2.st02 智權人員,'1',cp10 from caseprogress,staff s2,lawcase,casepropertymap where CP01=LC01(+) and cp02=LC02(+) and cp03=LC03(+) and cp04=LC04(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp05 is not null" & StrSQL3
StrSQL4 = " select cp01,cp02,cp03,cp04,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(HC09,'Y','＊','') caseno,nvl(HC06,'') casename,'' ptm0304,NVL(CPM03,'') cpm0304,cp05,s2.st02 智權人員,'1',cp10 from caseprogress,staff s2,hirecase,casepropertymap where CP01=HC01(+) and cp02=HC02(+) and cp03=HC03(+) and cp04=HC04(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp05 is not null" & StrSQL4
strSQL5 = " select cp01,cp02,cp03,cp04,cp09,cp01||'-'||cp02||'-'||cp03||'-'||cp04||DECODE(SP15,'Y','＊','') caseno,nvl(sp05,nvl(sp06,sp07)) casename,'' ptm0304,decode(sp09,'000',cpm03,cpm04) cpm0304,cp05,s2.st02 智權人員,'1',cp10 from caseprogress,staff s2,servicepractice,casepropertymap where CP01=SP01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) and cp13=s2.st01(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp05 is not null" & strSQL5

If Val(Txt1(5)) = 1 Then
   '指定對象-承辦人
   strExc(1) = " s1.st02 "
   '非指定的另一方
   strExc(2) = " 智權人員 "
Else
   '指定對象-智權人員
   strExc(1) = " 智權人員 "
   strExc(2) = " s1.st02 "
End If

    strSql = ""
    '已收文未齊備
    If Check1(0).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(0).Caption & Txt1(11) & Label2(0)
      '若已閉卷, 則在本所案號後加"*"號
               strSql = strSql + " select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,cp05,0,0,0,0,0,0,0,0," & strExc(2) & " sname2,'1',ep02 收文號,cp10,cp01 from engineerprogress,staff s1,(" & strSQL1 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and EP06 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,cp05,0,0,0,0,0,0,0,0," & strExc(2) & " sname2,'1',ep02 收文號,cp10,cp01 from engineerprogress,staff s1,(" & strSQL2 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and EP06 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,cp05,0,0,0,0,0,0,0,0," & strExc(2) & " sname2,'1',ep02 收文號,'','' from engineerprogress,staff s1,(" & StrSQL3 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and EP06 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,cp05,0,0,0,0,0,0,0,0," & strExc(2) & " sname2,'1',ep02 收文號,'','' from engineerprogress,staff s1,(" & StrSQL4 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and EP06 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,cp05,0,0,0,0,0,0,0,0," & strExc(2) & " sname2,'1',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL5 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and EP06 IS NULL " & StrSQL6
    End If
    '已齊備未完稿
    If Check1(1).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(1).Caption & Txt1(12) & Label2(1)
        If Len(strSql) <> 0 Then
            strSql = strSql + " union "
        End If

      '若已閉卷, 則在本所案號後加"*"號
               strSql = strSql + " select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,ep06,0,0,0,0,0,0," & strExc(2) & " sname2,'2',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL1 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep06 is not null and EP09 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,ep06,0,0,0,0,0,0," & strExc(2) & " sname2,'2',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL2 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep06 is not null and EP09 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,ep06,0,0,0,0,0,0," & strExc(2) & " sname2,'2',ep02 收文號,'','' from engineerprogress,staff s1,(" & StrSQL3 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep06 is not null and EP09 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,ep06,0,0,0,0,0,0," & strExc(2) & " sname2,'2',ep02 收文號,'','' from engineerprogress,staff s1,(" & StrSQL4 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep06 is not null and EP09 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,ep06,0,0,0,0,0,0," & strExc(2) & " sname2,'2',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL5 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep06 is not null and EP09 IS NULL " & StrSQL6
    End If
    '已會稿未會稿完成
    If Check1(2).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(2).Caption & Txt1(13) & Label2(2)
        If Len(strSql) <> 0 Then
            strSql = strSql + " union "
        End If
      '若已閉卷, 則在本所案號後加"*"號
               strSql = strSql + " select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,0,0,ep09,ep07,0,0,0," & strExc(2) & " sname2,'3',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL1 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep07 is not null and EP08 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,0,0,ep09,ep07,0,0,0," & strExc(2) & " sname2,'3',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL2 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep07 is not null and EP08 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,0,0,ep09,ep07,0,0,0," & strExc(2) & " sname2,'3',ep02 收文號,'','' from engineerprogress,staff s1,(" & StrSQL3 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep07 is not null and EP08 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,0,0,ep09,ep07,0,0,0," & strExc(2) & " sname2,'3',ep02 收文號,'','' from engineerprogress,staff s1,(" & StrSQL4 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep07 is not null and EP08 IS NULL " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,0,0,ep09,ep07,0,0,0," & strExc(2) & " sname2,'3',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL5 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep07 is not null and EP08 IS NULL " & StrSQL6
    End If
    '已會稿完成未發文
    If Check1(3).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(3).Caption & Txt1(14) & Label2(3)
        If Len(strSql) <> 0 Then
            strSql = strSql + " union "
        End If
      '若已閉卷, 則在本所案號後加"*"號
               strSql = strSql + " select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,0,0,ep09,ep07,0,ep08,0," & strExc(2) & " sname2,'4',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL1 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep08 is not null " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,0,0,ep09,ep07,0,ep08,0," & strExc(2) & " sname2,'4',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL2 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep08 is not null " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,0,0,ep09,ep07,0,ep08,0," & strExc(2) & " sname2,'4',ep02 收文號,'','' from engineerprogress,staff s1,(" & StrSQL3 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep08 is not null " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,0,0,ep09,ep07,0,ep08,0," & strExc(2) & " sname2,'4',ep02 收文號,'','' from engineerprogress,staff s1,(" & StrSQL4 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep08 is not null " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,0,0,ep09,ep07,0,ep08,0," & strExc(2) & " sname2,'4',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL5 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep08 is not null " & StrSQL6
    End If

    'Added by Morgan 2021/8/4 --王副總
    '已完稿未會稿
    If Check1(4).Value = 1 Then
      pub_QL05 = pub_QL05 & ";" & Label1(7) & Check1(4).Caption & Txt1(17) & Label2(4)
        If Len(strSql) <> 0 Then
            strSql = strSql + " union "
        End If
      '若已閉卷, 則在本所案號後加"*"號
      'Modified by Morgan 2024/3/18 +ep06(Word輸出計算總天數要用)
               strSql = strSql + " select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,ep06,0,ep09,0,0,0,0," & strExc(2) & " sname2,'5',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL1 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep09>0 and ep34='Y' and ep07 is null " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,ep06,0,ep09,0,0,0,0," & strExc(2) & " sname2,'5',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL2 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep09>0 and ep34='Y' and ep07 is null " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,ep06,0,ep09,0,0,0,0," & strExc(2) & " sname2,'5',ep02 收文號,'','' from engineerprogress,staff s1,(" & StrSQL3 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep09>0 and ep34='Y' and ep07 is null " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,ep06,0,ep09,0,0,0,0," & strExc(2) & " sname2,'5',ep02 收文號,'','' from engineerprogress,staff s1,(" & StrSQL4 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep09>0 and ep34='Y' and ep07 is null " & StrSQL6
      strSql = strSql + "UNION ALL select " & strExc(1) & " sname1,ep01,caseno,casename,ptm0304,cpm0304,0,0,ep06,0,ep09,0,0,0,0," & strExc(2) & " sname2,'5',ep02 收文號,'','' from engineerprogress,staff s1,(" & strSQL5 & ") x1 where ep02=cp09 and ep05=s1.st01(+) and ep09>0 and ep34='Y' and ep07 is null " & StrSQL6
    End If
    
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount)
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 16
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            TestOk = True
            If Len(strTemp(6)) <> 0 And Val(strTemp(6)) <> 0 Then
                strTemp(7) = str(GetWorkDay(GetTodayDate, strTemp(6)))
                If Val(Txt1(11)) > Val(strTemp(7)) Then
                    TestOk = False
                End If
                strTemp(6) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(6)))
            Else
                strTemp(6) = ""
            End If
            If Len(strTemp(8)) <> 0 And Val(strTemp(8)) <> 0 Then
                strTemp(9) = str(GetWorkDay(GetTodayDate, strTemp(8)))
                If Val(Txt1(12)) > Val(strTemp(9)) Then
                    TestOk = False
                End If
                strTemp(8) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(8)))
            Else
                strTemp(8) = ""
            End If
            If Len(strTemp(10)) <> 0 And Val(strTemp(10)) <> 0 Then
               'Added by Morgan 2021/8/4
               If Val(strTemp(16)) = 5 Then
                  strTemp(9) = str(GetWorkDay(GetTodayDate, DBDATE(strTemp(8)))) 'Added by Morgan 2024/3/18 +總承辦天數
                  strTemp(12) = str(GetWorkDay(GetTodayDate, strTemp(10)))
                  If Val(Txt1(17)) > Val(strTemp(12)) Then
                     TestOk = False
                  End If
               End If
               'end 2021/8/4
                strTemp(10) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(10)))
            Else
                strTemp(10) = ""
            End If
            If Len(strTemp(11)) <> 0 And Val(strTemp(11)) <> 0 Then
                strTemp(12) = str(GetWorkDay(GetTodayDate, strTemp(11)))
                If Val(Txt1(13)) > Val(strTemp(12)) Then
                    TestOk = False
                End If
                strTemp(11) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(11)))
            Else
                strTemp(11) = ""
            End If
            If Len(strTemp(13)) <> 0 And Val(strTemp(13)) <> 0 Then
                strTemp(14) = str(GetWorkDay(GetTodayDate, strTemp(13)))
                If Val(Txt1(14)) > Val(strTemp(14)) Then
                    TestOk = False
                End If
                strTemp(13) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(13)))
            Else
                strTemp(13) = ""
            End If
            
            strTemp(18) = "" & .Fields("收文號")
            If TestOk = True Then
                strTemp(17) = ""
                If Val(strTemp(16)) = 1 Then
                    Select Case CheckStr(.Fields(19))
                    Case "P", "FCP", "CFP"
                         Select Case Val(CheckStr(.Fields(18)))
                         Case 101, 102, 103, 104, 105
                              strTemp(17) = "#"
                         Case Else
                              strTemp(17) = ""
                         End Select
                    Case "T", "FCT", "CFT", "TF"
                         Select Case Val(CheckStr(.Fields(18)))
                         Case "101"
                              strTemp(17) = "#"
                         Case Else
                              strTemp(17) = ""
                         End Select
                    Case Else
                    End Select
                End If
                strSql = "insert into r090612 values ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & ChgSQL(strTemp(3)) & "','" & strTemp(4) & "','" & strTemp(5) & "','" & strTemp(6) & "'," & Val(strTemp(7)) & ",'" & strTemp(8) & "'," & Val(strTemp(9)) & ",'" & strTemp(10) & "','" & strTemp(11) & "'," & Val(strTemp(12)) & ",'" & strTemp(13) & "'," & Val(strTemp(14)) & ",'" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "','" & strUserNum & "','" & strTemp(18) & "') "
                cnnConnection.Execute strSql, intI
            End If
            .MoveNext
            DoEvents
        Loop
    Else
        InsertQueryLog (0)
    End If
End With
CheckOC
End Sub

