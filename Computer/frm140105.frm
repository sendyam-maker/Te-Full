VERSION 5.00
Begin VB.Form frm140105 
   BorderStyle     =   1  '單線固定
   Caption         =   "分所內商延展、第二期註冊費未續辦清單"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   3870
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   2
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1110
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   1
      Left            =   2580
      MaxLength       =   8
      TabIndex        =   1
      Top             =   750
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   315
      Index           =   0
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   0
      Top             =   750
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   2865
      TabIndex        =   4
      Top             =   30
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1860
      TabIndex        =   3
      Top             =   30
      Width           =   975
   End
   Begin VB.Label lblST 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   2490
      TabIndex        =   10
      Top             =   1170
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   180
      TabIndex        =   9
      Top             =   1170
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2250
      X2              =   2940
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Left            =   180
      TabIndex        =   8
      Top             =   810
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所別："
      Height          =   180
      Left            =   180
      TabIndex        =   7
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(1:北;2:中;3:南;4:高)"
      Height          =   180
      Left            =   1620
      TabIndex        =   6
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label lblst06 
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   450
      Width           =   315
   End
End
Attribute VB_Name = "frm140105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Dim strSql As String
Dim strTemp(0 To 7) As String, i As Integer, j As Integer, s As Integer, PLeft(0 To 7) As Integer, iPrint As Integer, Page As Integer


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
     If Trim(txt1(0)) = "" Then
        MsgBox "法定期限區間起不可空白！", vbExclamation
        txt1(0).SetFocus
        Exit Sub
     End If
     If Trim(txt1(1)) = "" Then
        MsgBox "法定期限區間迄不可空白！", vbExclamation
        txt1(1).SetFocus
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     Me.Enabled = False
     StrMenu
     Me.Enabled = True
     Screen.MousePointer = vbDefault
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    lblst06 = pub_strUserOffice
    txt1(0) = ChangeWDateStringToTString(DateAdd("m", -7, ChangeWStringToWDateString(Mid(strSrvDate(1), 1, 6) & "01")))
    txt1(1) = ChangeWDateStringToTString(DateAdd("d", -1, DateAdd("m", -6, ChangeWStringToWDateString(Mid(strSrvDate(1), 1, 6) & "01"))))
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm140105 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         If KeyAscii < 48 And KeyAscii > 57 And KeyAscii <> 13 And KeyAscii <> 8 Then
            KeyAscii = 0
         End If
      'Add By Sindy 2010/11/26
      Case 2
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim s
Cancel = True
Select Case Index
    Case 0, 1
        If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
           Me.txt1(Index).SetFocus
           txt1_GotFocus Index
           Exit Sub
        End If
        If Index = 1 Then
           If RunNick2(txt1(Index - 1), txt1(Index)) Then
               txt1(Index - 1).SetFocus
               txt1_GotFocus (Index - 1)
               Exit Sub
           End If
         End If
    Case 2
        lblST.Caption = GetPrjSalesNM(txt1(Index))
        If Trim(txt1(Index)) <> "" Then
             If Trim(lblST.Caption) = "" Then
                 s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
                 txt1(Index).SetFocus
                 txt1_GotFocus (Index)
                 Exit Sub
             End If
        End If
    Case Else
End Select
Cancel = False
End Sub

Sub StrMenu()
Page = 1
'2007/8/17 modify by sonia 加判斷專用期止日小於系統日
'strSQL = "select tm34,np02||'-'||np03||'-'||np04||'-'||np05,NVL(DECODE(tm10,'000',CPM03,CPM04),to_char(np07)),tm05,NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06)),NVL(S1.ST02,np10)," & SQLDate("NP08") & "," & SQLDate("NP09") & " from nextprogress,trademark,casepropertymap,staff S1,staff S2,customer " & _
'          " Where np02=cpm01(+) and to_char(np07)=cpm02(+) and np02=tm01 and np03=tm02 and np04=tm03 and np05=tm04 and (np06 ='N' or np06 is null) and np07 in (102,716) and np02='T' and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and np09>=" & ChangeTStringToWString(txt1(0)) & " and np09<=" & ChangeTStringToWString(txt1(1)) & " and cu13=s2.st01(+) and s2.st06='" & lblst06 & "' and np10=s1.st01(+) " & IIf(Trim(txt1(2)) = "", "", " and np10='" & txt1(2) & "' ") & " order by 1,2"
'2011/12/14 modify by sonia 加銷卷日期tm57或tm73
strSql = "select tm34,decode(decode('" & lblst06 & "','1',tm57,tm73),null,'　','△')||np02||'-'||np03||'-'||np04||'-'||np05,NVL(DECODE(tm10,'000',CPM03,CPM04),to_char(np07)),tm05,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),NVL(S1.ST02,np10)," & SQLDate("NP08") & "," & SQLDate("NP09") & " from nextprogress,trademark,casepropertymap,staff S1,staff S2,customer " & _
         " Where np02=cpm01(+) and to_char(np07)=cpm02(+) and np02=tm01 and np03=tm02 and np04=tm03 and np05=tm04 and tm22<=" & GetTodayDate & " and (np06 ='N' or np06 is null) and np07 in (102,716) and np02='T' and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) and np09>=" & ChangeTStringToWString(txt1(0)) & " and np09<=" & ChangeTStringToWString(txt1(1)) & " and cu13=s2.st01(+) and s2.st06='" & lblst06 & "' and np10=s1.st01(+) " & IIf(Trim(txt1(2)) = "", "", " and np10='" & txt1(2) & "' ") & " order by tm34,np02||'-'||np03||'-'||np04||'-'||np05"
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 Then
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 7
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = StrToStr(strTemp(2), 6)
            strTemp(3) = StrToStr(strTemp(3), 14)
            strTemp(4) = StrToStr(strTemp(4), 14)
            PrintDatil
            If iPrint > Printer.ScaleHeight - 1000 Then
                Page = Page + 1
                Printer.CurrentX = PLeft(0)
                Printer.CurrentY = iPrint
                Printer.Print String(204, "-")
                Printer.NewPage
                PrintTitle
            End If
           .MoveNext
        Loop
    Else
       ShowNoData
       Exit Sub
    End If
End With
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print String(204, "-")
Printer.EndDoc
ShowPrintOk
CheckOC
End Sub

Private Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth("內商延展、第二期註冊費未續辦清單") / 2)
Printer.CurrentY = iPrint
Printer.Print "內商延展、第二期註冊費未續辦清單"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth("所別：" & lblst06 & "(1:北 2:中 3:南 4:高)") / 2)
Printer.CurrentY = iPrint
Printer.Print "所別：" & lblst06 & "(1:北 2:中 3:南 4:高)"
iPrint = iPrint + 300
Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth("法定期限：" & Format(ChangeTStringToTDateString(txt1(0)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(1))) / 2)
Printer.CurrentY = iPrint
Printer.Print "法定期限：" & Format(ChangeTStringToTDateString(txt1(0)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(1))
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")) - 500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
'2011/12/14 add by sonia
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "注意：本所案號前有△者表示此案已銷卷"
'2011/12/14 end
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")) - 500
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print String(204, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "分所案號"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "申請人"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "本所期限"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "法定期限"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print String(204, "-")
iPrint = iPrint + 300
End Sub

Private Sub PrintDatil()
For i = 0 To 7
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Private Sub GetPleft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1500
PLeft(2) = 3400
PLeft(3) = 4500 + 500
PLeft(4) = 7500 + 500 + 500
PLeft(5) = 10500 + 500 + 500 + 500
PLeft(6) = 12100 + 500 + 500 + 500
PLeft(7) = 13600 + 500 + 500 + 500
End Sub
