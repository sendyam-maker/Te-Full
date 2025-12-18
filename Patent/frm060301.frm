VERSION 5.00
Begin VB.Form frm060301 
   BorderStyle     =   1  '單線固定
   Caption         =   "請求公告表/公告函"
   ClientHeight    =   1488
   ClientLeft      =   3936
   ClientTop       =   3180
   ClientWidth     =   3804
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1488
   ScaleWidth      =   3804
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1128
      MaxLength       =   7
      TabIndex        =   0
      Top             =   468
      Width           =   1035
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1128
      MaxLength       =   1
      TabIndex        =   2
      Top             =   804
      Width           =   420
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   3
      Left            =   1128
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "FCP"
      Top             =   1152
      Width           =   555
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2796
      TabIndex        =   8
      Top             =   12
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2004
      TabIndex        =   7
      Top             =   12
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   3384
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1152
      Width           =   345
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2976
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1152
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   1812
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1152
      Width           =   1035
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2232
      MaxLength       =   7
      TabIndex        =   1
      Top             =   468
      Width           =   1035
   End
   Begin VB.Line Line3 
      X1              =   1128
      X2              =   3543
      Y1              =   1296
      Y2              =   1296
   End
   Begin VB.Line Line2 
      X1              =   1692
      X2              =   2817
      Y1              =   576
      Y2              =   576
   End
   Begin VB.Line Line1 
      X1              =   1620
      X2              =   1830
      Y1              =   648
      Y2              =   648
   End
   Begin VB.Label Label1 
      Caption         =   "(1.管制表 2.申請書)"
      Height          =   180
      Index           =   3
      Left            =   1620
      TabIndex        =   12
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1185
      Width           =   930
   End
   Begin VB.Label Label1 
      Caption         =   "報表性質："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   870
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "核准日："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   510
      Width           =   720
   End
End
Attribute VB_Name = "frm060301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 8) As String
Dim PLeft(0 To 8) As Integer, strTemp1 As Variant, strTemp2 As Variant, Bol1 As Boolean

Private Sub cmdok_Click(Index As Integer)
 Dim rsTmp As New ADODB.Recordset
 Dim strTmp As String
Select Case Index
Case 0
     Printer.Orientation = 2
     DoEvents
     ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/8 清除查詢印表記錄檔欄位
     If Len(txt1(1)) = 0 Then
         'Add By Cheng 2002/03/20
         If PUB_CheckKeyInDate(Me.txt1(0)) = -1 Then
            Me.txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
         
         If Len(txt1(4)) = 0 Then
             s = MsgBox("本所案號不可空白!!", , "USER 輸入錯誤")
             txt1(4).SetFocus
             txt1_GotFocus (4)
             Exit Sub
         Else
             If Len(txt1(2)) = 0 Then
                s = MsgBox("報表性質不可空白!!", , "USER 輸入錯誤")
                txt1(2).SetFocus
                Exit Sub
             Else
                 Select Case Val(txt1(2))
                 Case 2
                     pub_QL05 = pub_QL05 & ";" & Label1(1) & "2.申請書" 'Add By Sindy 2010/12/8
                     strTmp = txt1(3) & txt1(4)
                     pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/8
                     If txt1(5) = "" Then
                        strTmp = strTmp & "0"
                     Else
                        strTmp = strTmp & txt1(5)
                        pub_QL05 = pub_QL05 & "-" & txt1(5) 'Add By Sindy 2010/12/8
                     End If
                     If txt1(6) = "" Then
                        strTmp = strTmp & "00"
                     Else
                        strTmp = strTmp & txt1(6)
                        pub_QL05 = pub_QL05 & "-" & txt1(6) 'Add By Sindy 2010/12/8
                     End If
                     Screen.MousePointer = vbHourglass
                     NowPrint strTmp & "&418", "03", "00", False, strUserNum, 0
                     Screen.MousePointer = vbDefault
                     InsertQueryLog ("") 'Add By Sindy 2010/12/8
                     MsgBox "列印完成 !", vbInformation
                 Case 1
                      pub_QL05 = pub_QL05 & ";" & Label1(1) & "1.管制表" 'Add By Sindy 2010/12/8
                      Bol1 = False
                      Screen.MousePointer = vbHourglass
                      Me.Enabled = False
                      Process
                      Me.Enabled = True
                      Screen.MousePointer = vbDefault
                 Case Else
                 End Select
             End If
         End If
     Else
         'Add By Cheng 2002/03/20
         If PUB_CheckKeyInDate(Me.txt1(0)) = -1 Then
            Me.txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Sub
         End If
         If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
            Me.txt1(1).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
         'Modify by Morgan 2010/8/13 百年蟲
         'If txt1(0) > txt1(1) Then
         If Val(txt1(0)) > Val(txt1(1)) Then
            MsgBox "核准日範圍不正確，請重新輸入 !", vbCritical
            txt1(0).SetFocus
            Exit Sub
         End If
         If Len(txt1(2)) = 0 Then
            s = MsgBox("報表性質不可空白!!", , "USER 輸入錯誤")
            txt1(2).SetFocus
            Exit Sub
         Else
            Select Case Val(txt1(2))
            Case 2
               pub_QL05 = pub_QL05 & ";" & Label1(1) & "2.申請書" 'Add By Sindy 2010/12/8
               '910624 Sieg
               If Len(Trim(txt1(0))) <> 0 Then
                  strSQL1 = " AND PA20>=" & Val(ChangeTStringToWString(txt1(0))) & " "
               End If
               If Len(Trim(txt1(1))) <> 0 Then
                  strSQL1 = strSQL1 & " AND PA20<=" & Val(ChangeTStringToWString(txt1(1))) & " "
               End If
               If Len(Trim(txt1(0))) <> 0 Or Len(Trim(txt1(1))) <> 0 Then
                  pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/12/8
               End If
               'Modify by Morgan 2004/1/2
               'strSQL = "SELECT PA01||PA02||PA03||PA04 FROM PATENT WHERE PA01='FCP' AND PA14 IS NULL AND PA57 IS NULL " & strSQL1
               strSql = "SELECT PA01||PA02||PA03||PA04 FROM PATENT WHERE PA01='FCP' AND PA14 IS NULL " & strSQL1
               intI = 1
               Set rsTmp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
               If intI = 1 Then
                  Screen.MousePointer = vbHourglass
                  InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/12/8
                  Do While Not rsTmp.EOF
                     NowPrint rsTmp.Fields(0) & "&418", "03", "00", False, strUserNum, 0
                     rsTmp.MoveNext
                  Loop
                  Screen.MousePointer = vbDefault
                  ShowPrintOk
               Else
                  InsertQueryLog (0) 'Add By Sindy 2010/12/8
                  ShowNoData
               End If
            Case 1
                pub_QL05 = pub_QL05 & ";" & Label1(1) & "1.管制表" 'Add By Sindy 2010/12/8
                Bol1 = True
                Screen.MousePointer = vbHourglass
                Me.Enabled = False
                Process
                Me.Enabled = True
                Screen.MousePointer = vbDefault
            Case Else
            End Select
         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R060301 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
If Bol1 = True Then
   If Len(Trim(txt1(0))) <> 0 Then
      strSQL1 = strSQL1 & " AND PA20>=" & Val(ChangeTStringToWString(txt1(0))) & " "
   End If
   If Len(Trim(txt1(1))) <> 0 Then
      strSQL1 = strSQL1 & " AND PA20<=" & Val(ChangeTStringToWString(txt1(1))) & " "
   End If
   If Len(Trim(txt1(0))) <> 0 Or Len(Trim(txt1(1))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/12/8
   End If
Else
    strSQL1 = " AND PA01='FCP' "
    If Len(txt1(4)) <> 0 Then
        strSQL1 = strSQL1 & " AND PA02='" & txt1(4) & "' "
    End If
    pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/8
    If Len(txt1(5)) = 0 Then
        strSQL1 = strSQL1 + " AND PA03='0' "
    Else
        strSQL1 = strSQL1 + " AND PA03='" & txt1(5) & "' "
        pub_QL05 = pub_QL05 & "-" & txt1(5) 'Add By Sindy 2010/12/8
    End If
    If Len(txt1(6)) = 0 Then
        strSQL1 = strSQL1 + " AND PA04='00' "
    Else
        strSQL1 = strSQL1 + " AND PA04='" & txt1(6) & "' "
        pub_QL05 = pub_QL05 & "-" & txt1(6) 'Add By Sindy 2010/12/8
    End If
End If
strSQL1 = strSQL1 + " AND '1'=PA16(+) AND 'Y'=cp31(+) "
'Modify by Morgan 2004/1/2
'strSQL = "SELECT " & SQLDate("CP27") & "," & SQLDate("PA20") & ",PA11,PA01||'-'||PA02||'-'||PA03||'-'||PA04,NVL(PA05,NVL(PA06,PA07)),NVL(DECODE(PA09,'000',PTM03,PTM04),PA08),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),NVL(S1.ST02,CP14),NVL(S2.ST02,CP13),'" & strUserNum & "'  FROM PATENT,CASEPROGRESS,CASEPROPERTYMAP,PATENTTRADEMARKMAP,STAFF S1,STAFF S2 WHERE PA01='FCP' AND PA01=cp01(+) AND PA02=cp02(+) AND PA03=cp03(+) AND PA04=cp04(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP14=S1.ST01(+) AND CP13=S2.ST01(+) AND PA14 IS NULL AND PA57 IS NULL " & strSQL1
strSql = "SELECT " & SQLDate("CP27") & "," & SQLDate("PA20") & ",PA11,PA01||'-'||PA02||'-'||PA03||'-'||PA04,NVL(PA05,NVL(PA06,PA07)),NVL(DECODE(PA09,'000',PTM03,PTM04),PA08),NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),PA57,PA89,'" & strUserNum & "' FROM PATENT,CASEPROGRESS,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE PA01='FCP' AND PA01=cp01(+) AND PA02=cp02(+) AND PA03=cp03(+) AND PA04=cp04(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND PA14 IS NULL " & strSQL1
cnnConnection.Execute "INSERT INTO R060301 " & strSql
strSql = "SELECT * FROM R060301 WHERE ID='" & strUserNum & "' "
CheckOC

adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/8
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/8
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
PrintData
ShowPrintOk
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
strSql = "SELECT * FROM R060301 WHERE ID='" & strUserNum & "' "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(4) = StrToStr(strTemp(4), 11)
            strTemp(5) = StrToStr(strTemp(5), 4)
            strTemp(6) = StrToStr(strTemp(6), 4)
            strTemp(7) = Space(4) & StrToStr(strTemp(7), 4)
            strTemp(8) = Space(7) & StrToStr(strTemp(8), 4)
            PrintDatil
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End With
End If
Printer.EndDoc
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
Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "請求公告函") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "請求公告函"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 7500 - (Printer.TextWidth("准駁日期：" & Format(ChangeTStringToTDateString(txt1(0)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(1))) / 2)
Printer.CurrentY = iPrint
Printer.Print "准駁日期：" & Format(ChangeTStringToTDateString(txt1(0)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(1))
iPrint = iPrint + 300
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
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(300, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "核准日"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "申請案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "專利種類"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
'Modify by Morgan 2004/1/2
'Printer.Print "承辦人"
Printer.Print "是否閉券"

Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
'Modify by Morgan 2004/1/2
'Printer.Print "智權人員"
'Modify by Amy 2025/08/06 原:不續辦但准通知
Printer.Print "後續准駁簡單報告"

iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print String(300, "-")
iPrint = iPrint + 300
End Sub

Private Sub PrintDatil()
For i = 0 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Private Sub GetPleft()
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1300
PLeft(2) = 2500
PLeft(3) = 5000
PLeft(4) = 7000
PLeft(5) = 10000
PLeft(6) = 11500
PLeft(7) = 12900
PLeft(8) = 14200

End Sub

Private Sub Form_Load()
MoveFormToCenter Me
If InStr(1, UCase(GetSystemKindByNick), "FCP") = 0 Then
    s = MsgBox(strUserName & "  沒有使用 FCP 的權限!!", , "越權")
    Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm060301 = Nothing
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
Case 1
   If txt1(0) <> "" Then
      If Val(txt1(0)) > Val(txt1(1)) Then
         MsgBox "核准日範圍不正確，請重新輸入 !", vbCritical
         txt1(0).SetFocus
      End If
   End If
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0, 1
         If Not ChkDate(txt1(Index).Text) Then
            Cancel = True
            TextInverse txt1(Index)
         End If
      Case 2
           Select Case Val(txt1(2))
           Case 1, 2
           Case Else
                s = MsgBox("報表性質只能 1 或 2 !!", , "USER 輸入錯誤")
                Cancel = True
           End Select
   End Select
   If Cancel Then TextInverse txt1(Index)
End Sub
