VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090212_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "公報簡訊查詢列印"
   ClientHeight    =   3270
   ClientLeft      =   360
   ClientTop       =   2970
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   9315
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1176
      TabIndex        =   0
      Top             =   456
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   1176
      TabIndex        =   1
      Top             =   802
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   1176
      TabIndex        =   10
      Top             =   1494
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   330
      Left            =   1176
      TabIndex        =   12
      Top             =   1840
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   330
      Left            =   1176
      TabIndex        =   13
      Top             =   2190
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   0
      Left            =   1176
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1163
      Width           =   720
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   108
      TabIndex        =   28
      Top             =   2550
      Width           =   3825
      Begin VB.ComboBox Combo2 
         Height          =   276
         Left            =   756
         Style           =   2  '單純下拉式
         TabIndex        =   14
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   2
         Left            =   105
         TabIndex        =   29
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   3
      Left            =   7425
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1163
      Width           =   945
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   2
      Left            =   5190
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1163
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   1
      Left            =   2988
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1163
      Width           =   1005
   End
   Begin VB.TextBox Text5 
      Height          =   330
      Left            =   2520
      TabIndex        =   11
      Top             =   1494
      Width           =   1215
   End
   Begin VB.CommandButton Command 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   8208
      TabIndex        =   16
      Top             =   12
      Width           =   1092
   End
   Begin VB.CommandButton Command 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7416
      TabIndex        =   15
      Top             =   12
      Width           =   756
   End
   Begin MSForms.TextBox txt1 
      Height          =   330
      Index           =   3
      Left            =   8448
      TabIndex        =   9
      Top             =   1148
      Width           =   850
      VariousPropertyBits=   -1466941413
      ScrollBars      =   3
      Size            =   "1499;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   330
      Index           =   2
      Left            =   6264
      TabIndex        =   7
      Top             =   1148
      Width           =   850
      VariousPropertyBits=   -1466941413
      ScrollBars      =   3
      Size            =   "1499;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   330
      Index           =   1
      Left            =   4044
      TabIndex        =   5
      Top             =   1148
      Width           =   850
      VariousPropertyBits=   -1466941413
      ScrollBars      =   3
      Size            =   "1499;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txt1 
      Height          =   330
      Index           =   0
      Left            =   1860
      TabIndex        =   3
      Top             =   1148
      Width           =   850
      VariousPropertyBits=   -1466941413
      ScrollBars      =   3
      Size            =   "1499;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   2790
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label11 
      Caption         =   "((("
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   816
      TabIndex        =   27
      Top             =   1187
      Width           =   300
   End
   Begin VB.Label Label10 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   7284
      TabIndex        =   26
      Top             =   1199
      Width           =   132
   End
   Begin VB.Label Label9 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5088
      TabIndex        =   25
      Top             =   1193
      Width           =   120
   End
   Begin VB.Label Label8 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2880
      TabIndex        =   24
      Top             =   1169
      Width           =   180
   End
   Begin VB.Label Label7 
      Height          =   300
      Left            =   2415
      TabIndex        =   23
      Top             =   471
      Width           =   4965
   End
   Begin VB.Label Label6 
      Caption         =   "報表別：                (1.查詢用2.存檔用)"
      Height          =   180
      Left            =   165
      TabIndex        =   22
      Top             =   2250
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "顯示方式：            (1.螢幕 2.報表)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   135
      TabIndex        =   21
      Top             =   1890
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "公告日期："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   210
      TabIndex        =   20
      Top             =   1569
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "關鍵字："
      Height          =   180
      Left            =   60
      TabIndex        =   19
      Top             =   1223
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "國際分類：                                  ( 可模糊比對查詢 )"
      Height          =   180
      Index           =   0
      Left            =   144
      TabIndex        =   18
      Top             =   877
      Width           =   4900
   End
   Begin VB.Label Label1 
      Caption         =   "索引："
      Height          =   180
      Left            =   504
      TabIndex        =   17
      Top             =   531
      Width           =   612
   End
End
Attribute VB_Name = "frm090212_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; txt1(index) ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset, p As New ADODB.Recordset, strSQL1 As String, strSql As String, StrTmpNick As String, StrTmpNick1 As String, StrTmpNick2 As String
Public SQLSTRING As String
Dim SHELL As String, a(6) As Integer, B(6) As Integer, SSTRING As String, s As Integer
Dim PLeft(0 To 5) As Integer, iPrint As Integer, Page As Integer, i As Integer, j As Integer, SeekPrint As Integer, SeekPrintL As Integer, SeekTempPrint As String
 
Private Sub Command_Click(Index As Integer)
Select Case Index
       Case 0
         If Len(Trim(Text1.Text)) = 0 And Len(Trim(Text2.Text)) = 0 And Len(Trim(Text4.Text)) = 0 And Len(Trim(Text5.Text)) = 0 And Len(Trim(txt1(0))) = 0 And Len(Trim(txt1(1))) = 0 And Len(Trim(txt1(2))) = 0 And Len(Trim(txt1(3))) = 0 Then
            s = MsgBox("條件不可空白!!  最少輸入一項!!")
            Text1.SetFocus
            Text1_GotFocus
            Exit Sub
         End If
         If Len(Trim(Text6.Text)) = 0 Then
            s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
            Text6.SetFocus
            Text6_GotFocus
            Exit Sub
         Else
            If Text6 = "2" Then
                If Len(Trim(Text7.Text)) = 0 Then
                    s = MsgBox("報表別不可空白!!", , "USER 輸入錯誤")
                    Text7.SetFocus
                    Text7_GotFocus
                    Exit Sub
                End If
            End If
         End If
         j = Combo2.ListIndex
         Set Printer = Printers(j)
         
         'Printer.PaperSize = 39
         'Printer.Orientation = 2
         DoEvents
         
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         Process
         Me.Enabled = True
         Screen.MousePointer = vbDefault
       Case 1
         Unload Me
       Case Else
End Select
End Sub

Sub Process()
ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/15 清除查詢印表記錄檔欄位
If Len(Trim(txt1(0))) <> 0 Or Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & Label11 & Combo1(0).Text & " " & txt1(0) & Label8 & Combo1(1).Text & " " & txt1(1) & Label9 & Combo1(2).Text & " " & txt1(2) & Label10 & Combo1(3).Text & " " & txt1(3) 'Add By Sindy 2010/12/15
End If
strSQL1 = ""
If Len(Trim(txt1(0))) <> 0 Then
   If Combo1(0).Text = "" Or Combo1(0).Text = "AND" Then
      strSQL1 = strSQL1 & " (instr(BB08,'" & txt1(0) & "')>0 ) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(0).Text & " (instR(BB08,'" & txt1(0) & "')>0 ) "
   End If
End If
If Len(Trim(txt1(1))) <> 0 Then
   If Combo1(1).Text = "" Then
      strSQL1 = strSQL1 & " AND (instr(BB08,'" & txt1(1) & "')>0 ) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(1).Text & " (instr(BB08,'" & txt1(1) & "')>0 ) "
   End If
   strSQL1 = " (" & strSQL1 & ") "
End If
If Len(Trim(txt1(2))) <> 0 Then
   If Combo1(2).Text = "" Then
      strSQL1 = strSQL1 & " AND (instr(BB08,'" & txt1(2) & "')>0 ) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(2).Text & " (instr(BB08,'" & txt1(2) & "')>0) "
   End If
   strSQL1 = " (" & strSQL1 & ") "
End If
If Len(Trim(txt1(3))) <> 0 Then
   If Combo1(3).Text = "" Then
      strSQL1 = strSQL1 & " AND (instr(BB08,'" & txt1(3) & "')>0) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(3).Text & " (instr(BB08,'" & txt1(3) & "')>0) "
   End If
   strSQL1 = " (" & strSQL1 & ") "
End If
If Len(strSQL1) <> 0 Then
   strSQL1 = " AND " & strSQL1
End If
If Len(Trim(Text1)) <> 0 Then
    strSQL1 = strSQL1 + " AND BB06='" & Text1.Text & "' "
    pub_QL05 = pub_QL05 & ";" & Label1 & Text1 & Label7 'Add By Sindy 2010/12/15
End If
If Len(Trim(Text2)) <> 0 Then
   'Modify By Cheng 2002/03/04
'    strSQL1 = strSQL1 + " AND (BB03='" & Text2 & "' OR BB04='" & Text2 & "' OR BB05='" & Text2 & "') "
    strSQL1 = strSQL1 + " AND (BB03 Like '%" & Text2 & "%' OR BB04 Like '%" & Text2 & "%' OR BB05 Like '%" & Text2 & "%') "
    pub_QL05 = pub_QL05 & ";" & Left(Label2(0), 5) & Text2 & " (可模糊比對查詢)" 'Add By Sindy 2010/12/15
End If
If Len(Trim(Text4)) <> 0 Then
    strSQL1 = strSQL1 + " AND BB07>=" & Val(ChangeTStringToWString(Text4)) & " "
End If
If Len(Trim(Text5)) <> 0 Then
    strSQL1 = strSQL1 + " AND BB07<=" & Val(ChangeTStringToWString(Text5)) & " "
End If
If Len(Trim(Text4)) <> 0 Or Len(Trim(Text5)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label4 & Text4 & "-" & Text5 'Add By Sindy 2010/12/15
End If
If Text6 = "1" Then '查詢
   pub_QL05 = pub_QL05 & ";" & Left(Label5, 5) & "1.螢幕" 'Add By Sindy 2010/12/15
Else
   pub_QL05 = pub_QL05 & ";" & Left(Label5, 5) & "2.報表" 'Add By Sindy 2010/12/15
End If
If Text7 = "1" Then '查詢用
   pub_QL05 = pub_QL05 & ";" & Left(Label6, 4) & "1.查詢用" 'Add By Sindy 2010/12/15
Else
   pub_QL05 = pub_QL05 & ";" & Left(Label6, 4) & "2.存檔用" 'Add By Sindy 2010/12/15
End If
'If Len(Trim(Text3)) <> 0 Then
'    StrSQL1 = StrSQL1 + Text3
'End If
'Modify By Cheng 2003/01/06
'依公告日期及頁數由小至大排序
'strSQL = "SELECT BB01 as 頁數,BB02 as 公告號數," & SQLDate("BB07") & " as 公告日期,BB08 as 內容摘要       ,BB03||','||BB04||','||BB05 as 國際分類,BBI02 as 索引 FROM BULLETINBRIEF,BULLETINBRIEFINDEX WHERE BB06=BBI01(+) " & strSQL1 & " order by 2 desc,1 "
'Modify By Cheng 2003/01/13
'strSQL = "SELECT BB01 as 頁數,BB02 as 公告號數," & SQLDate("BB07") & " as 公告日期,BB08 as 內容摘要       ,BB03||','||BB04||','||BB05 as 國際分類,BBI02 as 索引 FROM BULLETINBRIEF,BULLETINBRIEFINDEX WHERE BB06=BBI01(+) " & strSQL1 & " order by BB07, BB01 "
strSql = "SELECT BB01 as 頁數,BB02 as 公告號數," & SQLDate("BB07") & " as 公告日期,BB08 as 內容摘要       ,BB03||','||BB04||','||BB05 as 國際分類,BBI02 as 索引 FROM BULLETINBRIEF,BULLETINBRIEFINDEX WHERE BB06=BBI01(+) " & strSQL1 & " order by To_Number(BB07), To_Number(BB01) "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/15
        If Text6 = "1" Then '查詢
            frm090212_2.Show
            Screen.MousePointer = vbDefault
            Set frm090212_2.Adodc1.Recordset = adoRecordset
            'Add By Cheng 2002/03/04
            frm090212_2.lbl.Caption = "公報簡訊：　" & frm090212_2.Adodc1.Recordset.RecordCount & " 筆"
        Else '列印
            If Text7 = "1" Then '查詢用
                Printer.Orientation = 2
                DoEvents
                PrintData1
                Set Printer = Printers(SeekPrint)
                Printer.Orientation = SeekPrintL
            Else '存檔用
                Printer.Orientation = 1
                DoEvents
                PrintData2
                Set Printer = Printers(SeekPrint)
                Printer.Orientation = SeekPrintL
            End If
        End If
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/15
        ShowNoData
        Exit Sub
    End If
End With
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    pemain.CursorLocation = adUseClient
    Combo1(0).AddItem "", 0
    Combo1(0).AddItem "AND", 1
    Combo1(0).AddItem "NOT", 2
    Combo1(1).AddItem "", 0
    Combo1(1).AddItem "AND", 1
    Combo1(1).AddItem "OR", 2
    Combo1(1).AddItem "AND NOT", 3
    Combo1(1).AddItem "OR NOT", 4
    Combo1(2).AddItem "", 0
    Combo1(2).AddItem "AND", 1
    Combo1(2).AddItem "OR", 2
    Combo1(2).AddItem "AND NOT", 3
    Combo1(2).AddItem "OR NOT", 4
    Combo1(3).AddItem "", 0
    Combo1(3).AddItem "AND", 1
    Combo1(3).AddItem "OR", 2
    Combo1(3).AddItem "AND NOT", 3
    Combo1(3).AddItem "OR NOT", 4
    For i = 0 To 3
      Combo1(i).Text = Combo1(i).List(1)
    Next i
    strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
       Set Printer = Printers(i)
       Combo2.AddItem Printer.DeviceName, j
       j = j + 1
       If Printer.DeviceName = strSql Then
           SeekPrint = i
       End If
   Next i
   Combo2.Text = Combo2.List(SeekPrint)
   Text6 = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
    Set frm090212_1 = Nothing
End Sub

Sub PrintData1()
Page = 1
With adoRecordset
    .MoveFirst
    PrintTitle1
    iPrint = iPrint + 300
    Do While .EOF = False
        If iPrint >= 10000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle1
            'Add By Cheng 2003/01/13
            '多加一列
            iPrint = iPrint + 300
        End If
        'Modify By Cheng 2003/06/06
'        Printer.CurrentX = PLeft(0) + 500 - (Printer.TextWidth(CheckStr(.Fields(0))))
        Printer.CurrentX = PLeft(0) + 625 - (Printer.TextWidth(CheckStr(.Fields(0))))
        Printer.CurrentY = iPrint
        Printer.Print CheckStr(.Fields(0))
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iPrint
        Printer.Print CheckStr(.Fields(1))
        Printer.CurrentX = PLeft(3)
        Printer.CurrentY = iPrint
        Printer.Print CheckStr(.Fields(2))
        'Add By Cheng 2003/03/20
        '國際分類
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iPrint
        Printer.Print Left(CheckStr(.Fields(4)), 12)
        '索引
        Printer.CurrentX = PLeft(5)
        Printer.CurrentY = iPrint
        Printer.Print Left(CheckStr(.Fields(5)), 8)
        If LenB(StrConv(CheckStr(.Fields(3)), vbFromUnicode)) > 70 Then
            StrTmpNick = CheckStr(.Fields(3))
            StrTmpNick2 = StrTmpNick
            Do While Len(Trim(StrTmpNick)) <> 0
                StrTmpNick1 = StrToStr(StrTmpNick, 35)
               Printer.CurrentX = PLeft(2)
               Printer.CurrentY = iPrint
               Printer.Print StrTmpNick1
               iPrint = iPrint + 300
               StrTmpNick = Replace(StrTmpNick, StrTmpNick1, "")
               If StrTmpNick = StrTmpNick2 Then
                  StrTmpNick = Replace(StrTmpNick, Left(StrTmpNick1, Len(StrTmpNick1) - 1), "")
                  StrTmpNick2 = StrTmpNick
               Else
                  StrTmpNick2 = StrTmpNick
               End If
            Loop
        Else
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(3))
        End If
        iPrint = iPrint + 300
        .MoveNext
    Loop
End With
Printer.Line (0, iPrint + 100)-(18000, iPrint + 100)
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintData2()
Page = 1
With adoRecordset
    .MoveFirst
    PrintTitle2
    iPrint = iPrint + 300
    Do While .EOF = False
        If iPrint >= 15000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle2
            'Add By Cheng 2003/01/13
            '多加一列
            iPrint = iPrint + 300
        End If
        'Modify By Cheng 2003/06/06
'        Printer.CurrentX = PLeft(0) + 500 - (Printer.TextWidth(CheckStr(.Fields(0))))
        Printer.CurrentX = PLeft(0) + 625 - (Printer.TextWidth(CheckStr(.Fields(0))))
        Printer.CurrentY = iPrint
        Printer.Print CheckStr(.Fields(0))
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iPrint
        Printer.Print CheckStr(.Fields(1))
        Printer.CurrentX = PLeft(2)
        Printer.CurrentY = iPrint
        Printer.Print CheckStr(.Fields(5))
        iPrint = iPrint + 300
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        Printer.Print CheckStr(.Fields(4))
        'iPrint = iPrint - 300
        'StrToStr = StrConv(MidB(StrConv(Strindex, vbFromUnicode), 1, StrIndex2 * 2), vbUnicode)
        If LenB(StrConv(CheckStr(.Fields(3)), vbFromUnicode)) > 36 Then
            StrTmpNick = CheckStr(.Fields(3))
            StrTmpNick2 = StrTmpNick
            Do While Len(Trim(StrTmpNick)) <> 0
                StrTmpNick1 = StrToStr(StrTmpNick, 18)
               StrTmpNick = Replace(StrTmpNick, StrTmpNick1, "")
               If StrTmpNick = StrTmpNick2 Then
                  StrTmpNick = Replace(StrTmpNick, Left(StrTmpNick1, Len(StrTmpNick1) - 1), "")
                  StrTmpNick2 = StrTmpNick
               Else
                  StrTmpNick2 = StrTmpNick
               End If
            Loop
        Else
            Printer.CurrentX = PLeft(3)
            Printer.CurrentY = iPrint
            Printer.Print CheckStr(.Fields(3))
        End If
        iPrint = iPrint + 300
        
        .MoveNext
    Loop
End With
Printer.Line (0, iPrint + 100)-(11000, iPrint + 100)
Printer.EndDoc
ShowPrintOk
End Sub

Private Sub Text1_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text1.IMEMode = 2
   CloseIme
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
If p.State = adStateOpen Then p.Close
strExc(0) = "SELECT BBI02 FROM BULLETINBRIEFINDEX WHERE BBI01='" & Text1.Text & "'"
p.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
If p.BOF And p.EOF Then Label7.Caption = "": Exit Sub
If IsNull(p.Fields(0).Value) Then
    Label7.Caption = ""
Else
    Label7.Caption = p.Fields(0).Value
End If
End Sub

Private Sub Text2_GotFocus()
'edit by nickc 2007/07/11 切換輸入法改用API
'Text2.IMEMode = 2
CloseIme
Text2.SelStart = 0
    Text2.SelLength = Len(Text2)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
'edit by nickc 2007/07/11 切換輸入法改用API
'Text4.IMEMode = 2
CloseIme
Text4.SelStart = 0
    Text4.SelLength = Len(Text4)
End Sub

Private Sub Text4_LostFocus()
If Text4.Text <> "" Then
If CheckIsTaiwanDate(Text4.Text) = False Then
    Text4.SetFocus
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4)
End If
End If
End Sub

Private Sub Text5_GotFocus()
'edit by nickc 2007/07/11 切換輸入法改用API
'Text5.IMEMode = 2
CloseIme
Text5.SelStart = 0
    Text5.SelLength = Len(Text5)
End Sub

Private Sub Text5_LostFocus()
If Text5.Text <> "" Then
If CheckIsTaiwanDate(Text5.Text) = False Then
    Text5.SetFocus
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5)
    Exit Sub
End If
If Text4.Text > Text5.Text Then MsgBox "輸入日期範圍錯誤", vbInformation: Text4.SetFocus: Exit Sub
End If
End Sub

Private Sub Text6_GotFocus()
'edit by nickc 2007/07/11 切換輸入法改用API
'Text6.IMEMode = 2
CloseIme
Text6.SelStart = 0
    Text6.SelLength = Len(Text6)
End Sub

Private Sub Text6_LostFocus()
    If Text6.Text <> "1" And Text6.Text <> "2" And Trim(Text6.Text) <> "" Then
        MsgBox "顯示方式輸入錯誤", vbInformation
        Text6.SetFocus
        Text6_GotFocus
    End If
End Sub

Private Sub Text7_GotFocus()
'edit by nickc 2007/07/11 切換輸入法改用API
'Text7.IMEMode = 2
CloseIme
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7)
End Sub

Private Sub Text7_LostFocus()

    If Text7.Text <> "1" And Text7.Text <> "2" And Trim(Text7.Text) <> "" Then
        MsgBox "報表別輸入錯誤", vbInformation
        Text7.SetFocus
        Text7_GotFocus
    End If
End Sub

Sub GetPleft1()
PLeft(0) = 0
PLeft(1) = 1250
PLeft(2) = 2750
PLeft(3) = 11500
'Add By Cheng 2003/03/20
PLeft(4) = 12750 '國際分類
PLeft(5) = 14250 '索引
End Sub

Private Sub PrintTitle1()
GetPleft1
'Printer.Orientation = 1
iPrint = 500
Printer.Font.Size = 24
Printer.Font.Name = "細明體"
Printer.Font.Bold = True
Printer.CurrentX = GetPrPosX(16000, "公報簡訊查詢表")
Printer.CurrentY = iPrint
Printer.Print "公報簡訊查詢表"
Printer.Font.Bold = False
Printer.Font.Size = 12
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint + 300
Printer.Print "列印人： " & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint + 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(GetTaiwanTodayDate)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint + 600
Printer.Print "頁　　次：" & Page
Printer.Line (0, iPrint + 900)-(18000, iPrint + 900)
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint + 1200
Printer.Print "公告頁數"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint + 1200
Printer.Print "公告號數"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint + 1200
Printer.Print "內容"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint + 1200
Printer.Print "公告日期"
'Add By Cheng 2003/03/20
'標題欄--國際分類
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint + 1200
Printer.Print "國際分類"
'標題欄--索引
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint + 1200
Printer.Print "索引"
Printer.Line (0, iPrint + 1500)-(18000, iPrint + 1500)
iPrint = 1800
End Sub

Sub GetPleft2()
PLeft(0) = 0
PLeft(1) = 1100
PLeft(2) = 2600
PLeft(3) = 6500
End Sub

Private Sub PrintTitle2()
GetPleft2
'Printer.Orientation = 1
iPrint = 500
Printer.Font.Size = 24
Printer.Font.Name = "細明體"
Printer.Font.Bold = True
Printer.CurrentX = GetPrPosX(10000, "公報簡訊表")
Printer.CurrentY = iPrint
Printer.Print "公報簡訊表"
Printer.Font.Bold = False
Printer.Font.Size = 12
Printer.CurrentX = 0
Printer.CurrentY = iPrint + 300
Printer.Print "列印人： " & strUserName
Printer.CurrentX = 0
Printer.CurrentY = iPrint + 600
Printer.Print "公告日期：" & Format(ChangeTStringToTDateString(Text4.Text) & " ", "@@@@@@@@@") & " - " & ChangeTStringToTDateString(Text5.Text)
Printer.CurrentX = 9000
Printer.CurrentY = iPrint + 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(GetTaiwanTodayDate)
Printer.CurrentX = 9000
Printer.CurrentY = iPrint + 600
Printer.Print "頁次： " & Page
Printer.Line (0, iPrint + 900)-(11000, iPrint + 900)
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint + 1200
Printer.Print "公告頁數"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint + 1200
Printer.Print "公告號數"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint + 1200
Printer.Print "索引"
Printer.Line (0, iPrint + 1500)-(PLeft(3) - 200, iPrint + 1500)
Printer.CurrentX = PLeft(0) '+ GetPrPosX(Pleft(3), "國 際 分 類")
Printer.CurrentY = iPrint + 1800
Printer.Print "國 際 分 類"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint + 1800
Printer.Print "內容"
Printer.Line (0, iPrint + 2100)-(11000, iPrint + 2100)
iPrint = 2400
End Sub

Private Sub txt1_GotFocus(Index As Integer)
'edit by nickc 2007/07/11 切換輸入法改用API
'txt1(Index).IMEMode = 1
OpenIme
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
'edit by nickc 2007/07/11 切換輸入法改用API
CloseIme
End Sub
