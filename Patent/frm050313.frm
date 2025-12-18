VERSION 5.00
Begin VB.Form frm050313 
   BorderStyle     =   1  '單線固定
   Caption         =   "取消收文明細表"
   ClientHeight    =   1995
   ClientLeft      =   3960
   ClientTop       =   1935
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4200
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1320
      Width           =   420
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2790
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   3150
      TabIndex        =   5
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2355
      TabIndex        =   4
      Top             =   48
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2670
      MaxLength       =   7
      TabIndex        =   2
      Top             =   960
      Width           =   960
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1335
      MaxLength       =   7
      TabIndex        =   1
      Top             =   960
      Width           =   960
   End
   Begin VB.Label lblReason 
      Height          =   510
      Left            =   1830
      TabIndex        =   10
      Top             =   1350
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "取消原因："
      Height          =   180
      Left            =   420
      TabIndex        =   9
      Top             =   1350
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   420
      TabIndex        =   8
      Top             =   630
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "至"
      Height          =   180
      Left            =   2385
      TabIndex        =   7
      Top             =   1035
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "取消收文日期："
      Height          =   180
      Left            =   75
      TabIndex        =   6
      Top             =   990
      Width           =   1260
   End
End
Attribute VB_Name = "frm050313"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, i As Integer, j As Integer, s As Integer, k As Integer
Dim strTemp(0 To 10) As String, PLeft(0 To 10) As Integer, Page As Integer, iPrint As Integer, StrTemp5(0 To 2) As String
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean
'Add By Cheng 2002/09/19
Dim strTemp1 As Variant, strTemp2 As Variant
 
Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
   'Add By Cheng 2002/09/16
   blnClkSure = False
   'Add By Cheng 2002/09/19
   If Len(txt1(2)) = 0 Then
      s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
      txt1(2).SetFocus
      txt1(2).SelStart = 0
      txt1(2).SelLength = Len(txt1(2))
      Exit Sub
   Else
      strTemp1 = Split(UCase(GetSystemKindByNick), ",")
      strTemp2 = Split(UCase(txt1(2)), ",")
      For i = 0 To UBound(strTemp2)
         s = 0
         For j = 0 To UBound(strTemp1)
             If strTemp1(j) = strTemp2(i) Then
                 s = 1
                 Exit For
             End If
         Next j
         If s = 0 Then
             s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
             txt1(2).SetFocus
             txt1(2).SelStart = 0
             txt1(2).SelLength = Len(txt1(2))
             Exit Sub
         End If
      Next i
   End If
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
      'Add By Cheng 2002/09/16
      If Me.txt1(0).Text <> "" And Me.txt1(1).Text <> "" Then
         If Val(Me.txt1(0).Text) > Val(Me.txt1(1).Text) Then
            MsgBox "取消收文日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            blnClkSure = True
            Me.txt1(0).SetFocus
            txt1_GotFocus 1
            Exit Sub
         End If
      End If
     
     If Len(Trim(txt1(1))) = 0 Then
        s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         txt1_GotFocus (0)
        Exit Sub
     Else
         'Add By Cheng 2002/09/19
         Me.lblReason.Caption = bbGetReasonOfRelief(Me.txt1(3).Text)
         If Me.txt1(3).Text <> "" And Me.lblReason.Caption = "" Then
            MsgBox "取消原因代號輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txt1(3).SetFocus
            txt1_GotFocus 3
            Exit Sub
         End If
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
        Me.Enabled = False
        Process
        Me.Enabled = True
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Private Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050313 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
'strTemp1 = Split(GetSystemKindByNick, ",")
'Modify By Cheng 2002/09/19
'strsql1 = strsql1 & " AND CP01 IN (" & SQLGrpStr(GetSystemKindByNick, 1) & ") "
'strsql2 = strsql2 & " AND CP01 IN (" & SQLGrpStr(GetSystemKindByNick, 2) & ") "
strSQL1 = strSQL1 & " AND CP01 IN (" & SQLGrpStr(Me.txt1(2).Text, 1) & ") "
strSQL2 = strSQL2 & " AND CP01 IN (" & SQLGrpStr(Me.txt1(2).Text, 2) & ") "
pub_QL05 = pub_QL05 & ";" & Label2 & txt1(2) 'Add By Sindy 2010/12/7

If Len(Trim(txt1(0))) <> 0 Then
   strSQL1 = strSQL1 + " AND CP57>=" & Val(ChangeTStringToWString(txt1(0))) & " "
   strSQL2 = strSQL2 + " AND CP57>=" & Val(ChangeTStringToWString(txt1(0))) & " "
End If
If Len(Trim(txt1(1))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP57<=" & Val(ChangeTStringToWString(txt1(1))) & " "
   strSQL2 = strSQL2 & " AND CP57<=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(0))) <> 0 Or Len(Trim(txt1(1))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/12/7
End If
'Add By Cheng 2002/09/19
If Me.txt1(3).Text <> "" Then
   strSQL1 = strSQL1 & " AND ROR01='" & Me.txt1(3).Text & "' "
   strSQL2 = strSQL2 & " AND ROR01='" & Me.txt1(3).Text & "' "
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & lblReason 'Add By Sindy 2010/12/7
End If

strSql = ""
strSql = "SELECT " & SQLDate("CP57") & ",ROR02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PTM03,DECODE(PA09,'000',CPM03,CPM04),NVL(S1.ST02,CP13),NVL(S2.ST02,CP14),CP16,CP18,'" & strUserNum & "' FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP,STAFF S1,STAFF S2,REASONOFRELIEF WHERE CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP58=ROR01(+) " & strSQL1
strSql = strSql + " union all select " & SQLDate("CP57") & ",ROR02," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),'',DECODE(SP09,'000',CPM03,CPM04),NVL(S1.ST02,CP13),NVL(S2.ST02,CP14),CP16,CP18,'" & strUserNum & "' FROM CASEPROGRESS,SERVICEPRACTICE,CASEPROPERTYMAP,STAFF S1,STAFF S2,REASONOFRELIEF WHERE CP13=S1.ST01(+) AND CP14=S2.ST01(+) AND cp01=cpm01(+) AND cp10=cpm02(+) AND cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND CP58=ROR01(+) " & strSQL2
cnnConnection.Execute "INSERT INTO R050313 " & strSql
CheckOC
k = 0
strSql = "SELECT * FROM R050313 WHERE ID='" & strUserNum & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/7
Else
   InsertQueryLog (0)  'Add By Sindy 2010/12/7
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
PrintData

Screen.MousePointer = vbDefault
End Sub

Private Sub PrintData()
strSql = "SELECT * FROM R050313 WHERE ID='" & strUserNum & "' ORDER BY R012001,R012004 "
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 10
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(1) = StrToStr(strTemp(1), 2)
            strTemp(4) = StrToStr(strTemp(4), 15)
            strTemp(5) = StrToStr(strTemp(5), 4)
            strTemp(6) = StrToStr(strTemp(6), 4)
            strTemp(7) = StrToStr(strTemp(7), 3)
            strTemp(8) = StrToStr(strTemp(8), 3)
            'StrConv(MidB(StrConv(StrTemp(0), vbFromUnicode), 1, 8), vbUnicode)
            If iPrint > 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            PrintDatil
            .MoveNext
        Loop
    End With
End If
CheckOC
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint > 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
strSql = "SELECT COUNT(*),SUM(R012010),SUM(R012011) FROM R050313 WHERE ID='" & strUserNum & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        For i = 0 To 2
            StrTemp5(i) = CheckStr(adoRecordset.Fields(i))
        Next i
        adoRecordset.MoveNext
    Loop
Else
    StrTemp5(0) = "0"
    StrTemp5(1) = "0"
    StrTemp5(2) = "0"
End If
CheckOC
PrintEnd

Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd()
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(200, "-")
'iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "合  計：" & StrTemp5(0) & " 筆"
Printer.CurrentX = 10000
Printer.CurrentY = iPrint
Printer.Print "合  計："
'Modify By Cheng 2002/09/19
'Printer.CurrentX = 14400 - Printer.TextWidth(StrTemp5(1))
Printer.CurrentX = 14400 - Printer.TextWidth(StrTemp5(1)) + 500
Printer.CurrentY = iPrint
Printer.Print StrTemp5(1)
'Modify By Cheng 2002/09/19
'Printer.CurrentX = 15100 - Printer.TextWidth(StrTemp5(2))
Printer.CurrentX = 15100 - Printer.TextWidth(StrTemp5(2)) + 500
Printer.CurrentY = iPrint
Printer.Print StrTemp5(2)
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1800
PLeft(2) = 2500
PLeft(3) = 3500 + 250
PLeft(4) = 5500 + 250
PLeft(5) = 9800 + 250
PLeft(6) = 11000 + 250
PLeft(7) = 12100 + 250
PLeft(8) = 13000 + 250
PLeft(9) = 13900 + 250 + 250
PLeft(10) = 14700 + 250 + 250
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
Printer.Print "取消收文明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "取消收文日期：" & Format(ChangeTStringToTDateString(txt1(0)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(1))
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人　：" & strUserName
'Add By Cheng 2002/09/19
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "取消原因：" & Me.lblReason.Caption

Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
'Add By Cheng 2002/09/19
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "系統類別：" & Me.txt1(2).Text

Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "取消收文日"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "原因"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文日"
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
Printer.Print "智權人員"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "費用"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "點數"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub PrintDatil()
For i = 0 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
'Modify By Cheng 2002/09/19
'Printer.CurrentX = 14400 - Printer.TextWidth(strTemp(9))
Printer.CurrentX = 14400 - Printer.TextWidth(strTemp(9)) + 500
Printer.CurrentY = iPrint
Printer.Print strTemp(9)
'Modify By Cheng 2002/09/19
'Printer.CurrentX = 15100 - Printer.TextWidth(strTemp(10))
Printer.CurrentX = 15100 - Printer.TextWidth(strTemp(10)) + 500
Printer.CurrentY = iPrint
Printer.Print strTemp(10)
iPrint = iPrint + 300
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'Add By Cheng 2002/09/19
txt1(2).Text = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050313 = Nothing
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
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If
'Add By Cheng 2002/09/19
Case 2
   strTemp1 = Split(UCase(GetSystemKindByNick), ",")
   strTemp2 = Split(UCase(txt1(2)), ",")
   For i = 0 To UBound(strTemp2)
      s = 0
      For j = 0 To UBound(strTemp1)
          If strTemp1(j) = strTemp2(i) Then
              s = 1
              Exit For
          End If
      Next j
      If s = 0 Then
          s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
          txt1(2).SetFocus
          txt1(2).SelStart = 0
          txt1(2).SelLength = Len(txt1(2))
          Exit Sub
      End If
   Next i
'Add By Cheng 2002/09/19
Case 3 '取消原因
      Me.lblReason.Caption = bbGetReasonOfRelief(Me.txt1(3).Text)
      If Me.txt1(3).Text <> "" And Me.lblReason.Caption = "" Then
         MsgBox "取消原因代號輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.txt1(3).SetFocus
         txt1_GotFocus 3
         Exit Sub
      End If
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 0, 1 '取消收文日期起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub

'Add By Cheng 2002/09/19
'取得取消原因
Private Function bbGetReasonOfRelief(strROR01 As String) As String
   Dim rsA As New ADODB.Recordset
   Dim StrSQLa As String
   
   StrSQLa = "Select * From ReasonOfRelief Where ROR01='" & strROR01 & "'"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      bbGetReasonOfRelief = "" & rsA.Fields(1).Value
   Else
      bbGetReasonOfRelief = ""
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function
