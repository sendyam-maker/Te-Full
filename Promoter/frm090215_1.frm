VERSION 5.00
Begin VB.Form frm090215_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "期刊資料查詢列印"
   ClientHeight    =   3435
   ClientLeft      =   1830
   ClientTop       =   2220
   ClientWidth     =   8115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   8115
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   1056
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1206
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   1056
      MaxLength       =   8
      TabIndex        =   10
      Top             =   1578
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   330
      Left            =   1056
      MaxLength       =   2
      TabIndex        =   12
      Top             =   1950
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   330
      Left            =   1056
      MaxLength       =   1
      TabIndex        =   13
      Top             =   2325
      Width           =   375
   End
   Begin VB.TextBox txt1 
      Height          =   330
      Index           =   0
      Left            =   1788
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   1
      Top             =   492
      Width           =   672
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Index           =   0
      Left            =   1050
      TabIndex        =   8
      Text            =   "Combo3"
      Top             =   864
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   90
      TabIndex        =   29
      Top             =   2760
      Width           =   3825
      Begin VB.ComboBox Combo2 
         Height          =   276
         Left            =   765
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
         TabIndex        =   30
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   330
      Index           =   1
      Left            =   3708
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   3
      Top             =   492
      Width           =   672
   End
   Begin VB.TextBox txt1 
      Height          =   330
      Index           =   2
      Left            =   5508
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   5
      Top             =   492
      Width           =   672
   End
   Begin VB.TextBox txt1 
      Height          =   330
      Index           =   3
      Left            =   7344
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   7
      Top             =   492
      Width           =   672
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   0
      Left            =   1056
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   507
      Width           =   720
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   1
      Left            =   2664
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   507
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   2
      Left            =   4500
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   507
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   3
      Left            =   6348
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   507
      Width           =   945
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   6828
      TabIndex        =   16
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6048
      TabIndex        =   15
      Top             =   10
      Width           =   756
   End
   Begin VB.TextBox Text5 
      Height          =   330
      Left            =   2856
      MaxLength       =   8
      TabIndex        =   11
      Top             =   1578
      Width           =   1215
   End
   Begin VB.Label LBL1 
      Height          =   300
      Left            =   1440
      TabIndex        =   28
      Top             =   1965
      Width           =   6570
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
      Left            =   2520
      TabIndex        =   27
      Top             =   513
      Width           =   180
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
      Left            =   4404
      TabIndex        =   26
      Top             =   492
      Width           =   120
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
      Left            =   6204
      TabIndex        =   25
      Top             =   492
      Width           =   132
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
      Left            =   732
      TabIndex        =   24
      Top             =   531
      Width           =   300
   End
   Begin VB.Label Label7 
      Caption         =   "(請輸入西元年)(年月日)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   4200
      TabIndex        =   23
      Top             =   1647
      Width           =   2676
   End
   Begin VB.Label Label6 
      Caption         =   "顯示方式：            (1.螢幕2.報表列印)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   75
      TabIndex        =   22
      Top             =   2400
      Width           =   3285
   End
   Begin VB.Label Label5 
      Caption         =   "索引："
      Height          =   180
      Left            =   144
      TabIndex        =   21
      Top             =   2025
      Width           =   612
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "出版日期：　　　　　　　　－　　　　　　　　　"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   108
      TabIndex        =   20
      Top             =   1653
      Width           =   3732
   End
   Begin VB.Label Label3 
      Caption         =   "作者："
      Height          =   180
      Left            =   96
      TabIndex        =   19
      Top             =   1281
      Width           =   612
   End
   Begin VB.Label Label2 
      Caption         =   "資料出處："
      Height          =   180
      Index           =   0
      Left            =   96
      TabIndex        =   18
      Top             =   924
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "關鍵字："
      Height          =   180
      Left            =   60
      TabIndex        =   17
      Top             =   567
      Width           =   852
   End
End
Attribute VB_Name = "frm090215_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; txt1(index) ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim pemain As New ADODB.Recordset
Public SQLSTRING As String
Dim SHELL As String, a(6) As Integer, B(6) As Integer, SSTRING As String, StrSQL6 As String, strSql As String, StrTmpNick As String
Dim ADDR(6) As Integer, iPrint As Integer, Page As Integer, i As Integer, j As Integer, strSQL1 As String, SeekPrint As Integer, SeekPrintL As Integer, SeekTempPrint As String
'Add By Cheng 2002/03/04
Dim rsP As New ADODB.Recordset
 
Private Sub cmdOK_Click(Index As Integer)
Select Case Index
       Case 0
         If Len(Trim(txt1(0))) = 0 And Len(Trim(txt1(1))) = 0 And Len(Trim(txt1(2))) = 0 And Len(Trim(txt1(3))) = 0 And Len(Trim(Combo3(0))) = 0 And Len(Trim(Text3)) = 0 And Len(Trim(Text4)) = 0 And Len(Trim(Text5)) = 0 And Len(Trim(Text6)) = 0 Then
            MsgBox ("最少輸入一項條件!!")
            txt1(0).SetFocus
            txt1_GotFocus (0)
            Exit Sub
         End If
         If Len(Trim(Text7)) = 0 Then
         End If
         j = Combo2.ListIndex
         Set Printer = Printers(j)
         Screen.MousePointer = vbHourglass
         Me.Enabled = False
         Process
         Me.Enabled = True
         Screen.MousePointer = vbDefault
       Case 1
         Unload Me
End Select
End Sub

Sub Process()
ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/15 清除查詢印表記錄檔欄位
If Len(Trim(txt1(0))) <> 0 Or Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1 & Label11 & Combo1(0).Text & " " & txt1(0) & Label8 & Combo1(1).Text & " " & txt1(1) & Label9 & Combo1(2).Text & " " & txt1(2) & Label10 & Combo1(3).Text & " " & txt1(3) 'Add By Sindy 2010/12/15
End If
strSQL1 = ""
If Len(Trim(txt1(0))) <> 0 Then
   If Combo1(0).Text = "" Or Combo1(0).Text = "AND" Then
      strSQL1 = strSQL1 & " (instr(PE01||PE06,'" & txt1(0) & "')>0) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(0).Text & " (instr(PE01||PE06,'" & txt1(0) & "')>0) "
   End If
End If
If Len(Trim(txt1(1))) <> 0 Then
   If Combo1(1).Text = "" Then
      strSQL1 = strSQL1 & " AND (instr(PE01||PE06,'" & txt1(1) & "')>0) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(1).Text & " (instr(PE01||PE06,'" & txt1(1) & "')>0) "
   End If
   strSQL1 = " (" & strSQL1 & ") "
End If
If Len(Trim(txt1(2))) <> 0 Then
   If Combo1(2).Text = "" Then
      strSQL1 = strSQL1 & " AND (instr(PE01||PE06,'" & txt1(2) & "')>0) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(2).Text & " (instr(PE01||PE06,'" & txt1(2) & "')>0) "
   End If
   strSQL1 = " (" & strSQL1 & ") "
End If
If Len(Trim(txt1(3))) <> 0 Then
   If Combo1(3).Text = "" Then
      strSQL1 = strSQL1 & " AND (instr(PE01||PE06,'" & txt1(3) & "')>0) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(3).Text & " (instr(PE01||PE06,'" & txt1(3) & "')>0) "
   End If
   strSQL1 = " (" & strSQL1 & ") "
End If

If Len(strSQL1) <> 0 Then
   strSQL1 = " AND " & strSQL1
End If
StrSQL6 = ""
If Len(Trim(Combo3(0))) <> 0 Then
   StrSQL6 = StrSQL6 & " and pe03='" & Combo3(0).Text & "' "
   pub_QL05 = pub_QL05 & ";" & Label2(0) & Combo3(0).Text 'Add By Sindy 2010/12/15
End If
If Len(Trim(Text3)) <> 0 Then
   StrSQL6 = StrSQL6 + " and pe02='" & Text3.Text & "' "
   pub_QL05 = pub_QL05 & ";" & Label3 & Text3 'Add By Sindy 2010/12/15
End If
If Len(Trim(Text4)) <> 0 Then
   StrSQL6 = StrSQL6 & " and pe05>=" & Val(Text4.Text) & " "
End If
If Len(Trim(Text5)) <> 0 Then
   StrSQL6 = StrSQL6 & " AND PE05<=" & Val(Text5.Text) & " "
End If
If Len(Trim(Text4)) <> 0 Or Len(Trim(Text5)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Left(Label4, 5) & Text4 & "-" & Text5 'Add By Sindy 2010/12/15
End If
If Len(Trim(Text6)) <> 0 Then
   StrSQL6 = StrSQL6 & " AND PE07='" & Text6.Text & "' "
   pub_QL05 = pub_QL05 & ";" & Label5 & Text6 & lbl1 'Add By Sindy 2010/12/15
End If
If Text7 = "1" Then
   pub_QL05 = pub_QL05 & ";" & Left(Label6, 5) & "1.螢幕" 'Add By Sindy 2010/12/15
Else
   pub_QL05 = pub_QL05 & ";" & Left(Label6, 5) & "2.報表列印" 'Add By Sindy 2010/12/15
End If

'strSQL = "SELECT PE03,decode(pe03,'中國專利報',PE05,DECODE(PE05,'','',(SUBSTR(PE05,1,4)-1911)||(SUBSTR(PE05,5,2))||(SUBSTR(PE05,7,2)))),PE04,NVL(PI02,PI01),PE01,PE02 FROM PERIODICAL,PERIODICALINDEX WHERE PE07=PI01(+) " & StrSQL1 & StrSQL6 & " ORDER BY 1,2,3 "
'Modify By Cheng 2003/11/27
'strSQL = "SELECT PE01,decode(pe03,'中國專利報'," & SQLDate("PE05", False) & ",DECODE(PE05,'','',(SUBSTR(PE05,1,4)-1911)||'/'||(SUBSTR(PE05,5,2))||'/'||(SUBSTR(PE05,7,2)))),NVL(PI02,PI01),PE03,PE02,PE06,PE04 FROM PERIODICAL,PERIODICALINDEX WHERE PE07=PI01(+) " & strSQL1 & StrSQL6 & " ORDER BY 1,2,3 "
'Modify by Morgan 2010/8/16 百年蟲 (SUBSTR(PE05,1,4)-1911)||'/'||(SUBSTR(PE05,5,2))||'/'||(SUBSTR(PE05,7,2)))-->substrb(' '||sqldatet(PE05),-9)
strSql = "SELECT PE03, decode(pe03,'中國專利報',DECODE(PE05,'','',SUBSTR(PE05,1,4)||'/'||SUBSTR(PE05,5,2)||'/'||SUBSTR(PE05,7,2)),DECODE(PE05,'','',substrb(' '||sqldatet(PE05),-9))), PE04, NVL(PI02,PI01), PE01, PE02, PE06 FROM PERIODICAL,PERIODICALINDEX " & _
                " WHERE PE07=PI01(+) " & strSQL1 & StrSQL6 & " ORDER BY 5, 2, 4 "
'End
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/15
      .MoveFirst
      If Text7 = "1" Then
         Set frm090215_2.Adodc1.Recordset = adoRecordset
         Set frm090215_2.grd1.Recordset = frm090215_2.Adodc1.Recordset
         frm090215_2.SetGrd
         Me.Hide
         frm090215_2.Show
         frm090215_2.Lbl.Caption = "期刊資料： " & frm090215_2.Adodc1.Recordset.RecordCount & " 筆"
      Else
         PRINT_REPORT
         Set Printer = Printers(SeekPrint)
         Printer.Orientation = SeekPrintL
      End If
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/15
      ShowNoData
      Exit Sub
   End If
End With
'CheckOC
End Sub

Private Sub Combo3_Validate(Index As Integer, Cancel As Boolean)
'edit by nickc 2007/07/11 切換輸入法改用API
CloseIme
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'pemain.CursorLocation = adUseClient
lbl1.Caption = ""
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
   Text7 = "1"

   'Add By Cheng 2002/03/04
   Me.Combo3(0).Clear
   If rsP.State <> adStateClosed Then rsP.Close
   rsP.CursorLocation = adUseClient
   rsP.Open "Select Distinct PE03 From Periodical Order By PE03 ", cnnConnection, adOpenStatic, adLockReadOnly
   If rsP.RecordCount > 0 Then
      rsP.MoveFirst
      While Not rsP.EOF
         Me.Combo3(0).AddItem "" & rsP.Fields(0).Value
         rsP.MoveNext
      Wend
   End If
   If rsP.State <> adStateClosed Then rsP.Close
   Set rsP = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Printer = Printers(SeekPrint)
Printer.Orientation = SeekPrintL
Set frm090215_1 = Nothing
End Sub

Private Sub Combo3_GotFocus(Index As Integer)
Combo3(0).SelStart = 0
Combo3(0).SelLength = Len(Combo3(0))
'edit by nickc 2007/07/11 切換輸入法改用API
'Combo3(0).IMEMode = 1
OpenIme

End Sub

Private Sub Text3_GotFocus()
Text3.SelStart = 0
Text3.SelLength = Len(Text3)
'edit by nickc 2007/07/11 切換輸入法改用API
'Text3.IMEMode = 1
OpenIme

End Sub

Private Sub Text3_Validate(Cancel As Boolean)
'edit by nickc 2007/07/11 切換輸入法改用API
CloseIme
End Sub

Private Sub Text4_GotFocus()
Text4.SelStart = 0
Text4.SelLength = Len(Text4)
'edit by nickc 2007/07/11 切換輸入法改用API
'Text4.IMEMode = 2
CloseIme

End Sub

Private Sub Text4_LostFocus()
If Text4.Text <> "" Then
    If CheckIsDate(Text4.Text) = False Then
    Text4.SetFocus
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4)
    End If
End If
End Sub



Private Sub Text5_GotFocus()
Text5.SelStart = 0
Text5.SelLength = Len(Text5)
'edit by nickc 2007/07/11 切換輸入法改用API
'Text5.IMEMode = 2
CloseIme

End Sub

Private Sub Text5_LostFocus()
If Text5.Text <> "" Then
    If CheckIsDate(Text5.Text) = False Then
        Text5.SetFocus
        Text5.SelStart = 0
        Text5.SelLength = Len(Text5)
        Exit Sub
    End If
    If Val(Text4.Text) > Val(Text5.Text) Then
    MsgBox "輸入日期範圍錯誤", vbInformation: Text4.SetFocus
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4)
    End If
End If
End Sub


Private Sub Text6_GotFocus()
Text6.SelStart = 0
Text6.SelLength = Len(Text6)
'edit by nickc 2007/07/11 切換輸入法改用API
'Text6.IMEMode = 2
CloseIme
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_LostFocus()
lbl1.Caption = ""
If Text6.Text <> "" Then
   strSql = "SELECT PI02 FROM PERIODICALINDEX WHERE PI01='" & Text6.Text & "' "
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         lbl1.Caption = CheckStr(.Fields(0))
      Else
         MsgBox ("無此索引代號!!")
         Text6.SetFocus
         Text6_GotFocus
         lbl1.Caption = ""
         Exit Sub
      End If
   End With
End If
CheckOC
End Sub

Private Sub Text7_GotFocus()
Text7.SelStart = 0
Text7.SelLength = Len(Text7)
'edit by nickc 2007/07/11 切換輸入法改用API
'Text7.IMEMode = 2
CloseIme
End Sub

Private Sub Text7_LostFocus()
Select Case Text7
Case "1", "2", ""
Case Else
      MsgBox ("只能輸入 1 或 2 !!")
      Text7.SetFocus
      Text7.SelLength = Len(Text7)
      Exit Sub
End Select
End Sub

'Private Sub GET_SQLSTRING()
'A(1) = InStr(SHELL, "+")
'A(2) = InStr(SHELL, "|")
'A(3) = InStr(SHELL, "AND")
'A(4) = InStr(SHELL, "OR")
'A(5) = InStr(SHELL, "and")
'A(6) = InStr(SHELL, "or")
'If A(1) = 0 And A(2) = 0 And A(3) = 0 And A(4) = 0 And A(5) = 0 And A(6) = 0 Then GoTo FIR
'Do While Not (InStr(SHELL, "+") = 0 And InStr(SHELL, "|") = 0 And InStr(SHELL, "AND") = 0 And InStr(SHELL, "OR") = 0 And InStr(SHELL, "and") = 0 And InStr(SHELL, "or") = 0)
'For i = 1 To 6
'   B(i) = A(i)
'Next i
'For i = 1 To 5
'    If B(i) > B(i + 1) Then
'       j = B(i)
'       B(i) = B(i + 1)
'       B(i + 1) = j
'    End If
'Next i
'For i = 1 To 6
'   If B(i) = 0 Then
'      j = j
'   Else
'      j = B(i): Exit For
'   End If
'Next i
'   SSTRING = Left(SHELL, j - 2)
'If j = A(1) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) AND "
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) AND "
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 3) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 3) & "')>0) AND "
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (instr(PE01,'" & SSTRING & "')>0 OR instr(PE06,'" & SSTRING & "')>0) AND "
'   End If
'   SHELL = Mid(SHELL, j + 2)

'ElseIf j = A(2) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) OR "
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) OR "
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 3) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 3) & "')>0) OR "
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (instr(PE01,'" & SSTRING & "')>0 OR instr(PE06,'" & SSTRING & "')>0) OR "
'   End If
'   SHELL = Mid(SHELL, j + 2)
'ElseIf j = A(3) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) AND "
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) AND "
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 3) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 3) & "')>0) AND "
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (instr(PE01,'" & SSTRING & "')>0 OR instr(PE06,'" & SSTRING & "')>0) AND "
'   End If
'   SHELL = Mid(SHELL, j + 4)
'ElseIf j = A(4) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) OR "
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) OR "
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 3) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 3) & "')>0) OR "
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (instr(PE01,'" & SSTRING & "')>0 OR instr(PE06,'" & SSTRING & "')>0) OR "
'   End If
'   SHELL = Mid(SHELL, j + 3)
'ElseIf j = A(5) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) AND "
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) AND "
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 3) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 3) & "')>0) AND "
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (instr(PE01,'" & SSTRING & "')>0 OR instr(PE06,'" & SSTRING & "')>0) AND "
'   End If
'   SHELL = Mid(SHELL, j + 4)

'ElseIf j = A(6) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) OR "
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 5) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 5) & "')>0) OR "
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SSTRING, 3) & "')>0 OR instr(PE06,'" & Mid(SSTRING, 3) & "')>0) OR "
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (instr(PE01,'" & SSTRING & "')>0 OR instr(PE06,'" & SSTRING & "')>0) OR "
'   End If
'   SHELL = Mid(SHELL, j + 3)
'End If
'A(1) = InStr(SHELL, "+")
'A(2) = InStr(SHELL, "|")
'A(3) = InStr(SHELL, "AND")
'A(4) = InStr(SHELL, "OR")
'A(5) = InStr(SHELL, "and")
'A(6) = InStr(SHELL, "or")
'Loop
'If Mid(SHELL, 1, 3) = "NOT" Or Mid(SHELL, 1, 3) = "not" Or Mid(SHELL, 1, 1) = "-" Then
'   If Mid(SHELL, 1, 3) = "NOT" Then
'      SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SHELL, 5) & "')>0 OR instr(PE06,'" & Mid(SHELL, 5) & "')>0) "
'   ElseIf Mid(SHELL, 1, 3) = "not" Then
'      SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SHELL, 5) & "')>0 OR instr(PE06,'" & Mid(SHELL, 5) & "')>0) "
'   ElseIf Mid(SHELL, 1, 1) = "-" Then
'      SQLSTRING = SQLSTRING + "NOT" + " (instr(PE01,'" & Mid(SHELL, 3) & "')>0 OR instr(PE06,'" & Mid(SHELL, 3) & "')>0) "
'   End If
'Else
'   SQLSTRING = SQLSTRING + " (instr(PE01,'" & SHELL & "')>0 OR instr(PE06,'" & SHELL & "')>0) "
'End If
'Exit Sub
'FIR:
'If Mid(SHELL, 1, 3) = "NOT" Or Mid(SHELL, 1, 3) = "not" Or Mid(SHELL, 1, 1) = "-" Then
'   If Mid(SHELL, 1, 3) = "NOT" Then
'      SQLSTRING = SQLSTRING + "AND NOT" + " (instr(PE01,'" & Mid(SHELL, 5) & "')>0 OR instr(PE06,'" & Mid(SHELL, 5) & "')>0) "
'   ElseIf Mid(SHELL, 1, 3) = "not" Then
'      SQLSTRING = SQLSTRING + "AND NOT" + " (instr(PE01,'" & Mid(SHELL, 5) & "')>0 OR instr(PE06,'" & Mid(SHELL, 5) & "')>0) "
'   ElseIf Mid(SHELL, 1, 1) = "-" Then
'      SQLSTRING = SQLSTRING + "AND NOT" + " (instr(PE01,'" & Mid(SHELL, 3) & "')>0 OR instr(PE06,'" & Mid(SHELL, 3) & "')>0) "
'   End If
'Else
'   SQLSTRING = SQLSTRING + " AND (instr(PE01,'" & SHELL & "')>0 OR instr(PE06,'" & SHELL &"')>0) "
'End If
'End Sub'

Private Sub GET_TITLE()
GET_ADDRESS
iPrint = 600
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.CurrentX = 6500: Printer.CurrentY = iPrint
Printer.Print "期刊資料表"
Printer.Font.Size = 12
Printer.CurrentX = ADDR(1): Printer.CurrentY = iPrint + 300
Printer.Print "列印人： " & strUserName
Printer.CurrentX = ADDR(6): Printer.CurrentY = iPrint + 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(GetTaiwanTodayDate)
Printer.CurrentX = ADDR(6): Printer.CurrentY = iPrint + 600
Printer.Print "頁  次：" & Page
Printer.Line (ADDR(1), iPrint + 900)-(15500, iPrint + 900)
Printer.CurrentX = ADDR(1): Printer.CurrentY = iPrint + 1200
Printer.Print "資料出處"
Printer.CurrentX = ADDR(2): Printer.CurrentY = iPrint + 1200
Printer.Print "出版日期"
Printer.CurrentX = ADDR(3): Printer.CurrentY = iPrint + 1200
Printer.Print "版,頁"
Printer.CurrentX = ADDR(4): Printer.CurrentY = iPrint + 1200
Printer.Print "索引"
Printer.CurrentX = ADDR(5): Printer.CurrentY = iPrint + 1200
Printer.Print "標題"
Printer.CurrentX = ADDR(6): Printer.CurrentY = iPrint + 1200
Printer.Print "作者"
Printer.Line (ADDR(1), iPrint + 1500)-(15500, iPrint + 1500)
End Sub

Private Sub PRINT_REPORT()
Printer.Orientation = 2
DoEvents
Page = 1
GET_TITLE

iPrint = 2400
Printer.Font.Size = 12
'If pemain.State = adStateOpen Then pemain.Close
'pemain.Open SQLSTRING, cnnConnection, adOpenStatic, adLockReadOnly
'If pemain.BOF And pemain.EOF Then Exit Sub
'If Not pemain.BOF Then pemain.MoveFirst
i = 0
adoRecordset.MoveFirst
Do While Not adoRecordset.EOF
Printer.CurrentX = ADDR(1): Printer.CurrentY = iPrint
Printer.Print StrToStr(CheckStr(adoRecordset.Fields(0)), 5)
Printer.CurrentX = ADDR(2): Printer.CurrentY = iPrint
'Select Case Len(CheckStr(adoRecordset.Fields(1)))
'Case 0
'      StrTmpNick = ""
'Case 6
'      StrTmpNick = ChangeTStringToTDateString(CheckStr(adoRecordset.Fields(1)))
'Case 8
'      StrTmpNick = ChangeWStringToWDateString(CheckStr(adoRecordset.Fields(1)))
'Case Else
'End Select
Printer.Print CheckStr(adoRecordset.Fields(1))
Printer.CurrentX = ADDR(3) + Printer.TextWidth("版,頁") - Printer.TextWidth(CheckStr(adoRecordset.Fields(2))): Printer.CurrentY = iPrint
Printer.Print CheckStr(adoRecordset.Fields(2))
Printer.CurrentX = ADDR(4): Printer.CurrentY = iPrint
Printer.Print StrToStr(CheckStr(adoRecordset.Fields(3)), 4)
Printer.CurrentX = ADDR(5): Printer.CurrentY = iPrint
Printer.Print StrToStr(CheckStr(adoRecordset.Fields(4)), 30)
Printer.CurrentX = ADDR(6): Printer.CurrentY = iPrint
Printer.Print StrToStr(CheckStr(adoRecordset.Fields(5)), 7)
i = i + 1
iPrint = iPrint + 300
adoRecordset.MoveNext
If adoRecordset.EOF Then Exit Do
    If i Mod 25 = 0 Then
        Printer.NewPage
        Page = Page + 1
        GET_TITLE
        iPrint = 2400
        i = 0
    End If
Loop
Printer.EndDoc
ShowPrintOk
End Sub

Private Sub GET_ADDRESS()
ADDR(1) = 0
ADDR(2) = 1700
ADDR(3) = 3300
ADDR(4) = 4100
ADDR(5) = 5300
ADDR(6) = 13000
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
