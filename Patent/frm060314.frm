VERSION 5.00
Begin VB.Form frm060314 
   BorderStyle     =   1  '單線固定
   Caption         =   "FCP新案收發文明細表"
   ClientHeight    =   1665
   ClientLeft      =   1650
   ClientTop       =   5370
   ClientWidth     =   3375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3375
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2364
      TabIndex        =   5
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1572
      TabIndex        =   4
      Top             =   24
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   1008
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1296
      Width           =   192
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1920
      MaxLength       =   7
      TabIndex        =   2
      Top             =   876
      Width           =   840
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1008
      MaxLength       =   7
      TabIndex        =   1
      Top             =   876
      Width           =   840
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1008
      MaxLength       =   1
      TabIndex        =   0
      Top             =   456
      Width           =   210
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   1905
      Y1              =   972
      Y2              =   972
   End
   Begin VB.Label Label1 
      Caption         =   "(1.收/發文日 2.本所案號)"
      Height          =   180
      Index           =   4
      Left            =   1236
      TabIndex        =   10
      Top             =   1356
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "(1.收文 2.發文)"
      Height          =   180
      Index           =   3
      Left            =   1296
      TabIndex        =   9
      Top             =   516
      Width           =   1536
   End
   Begin VB.Label Label1 
      Caption         =   "排序條件："
      Height          =   180
      Index           =   2
      Left            =   72
      TabIndex        =   8
      Top             =   1356
      Width           =   912
   End
   Begin VB.Label Label1 
      Caption         =   "日期："
      Height          =   180
      Index           =   1
      Left            =   72
      TabIndex        =   7
      Top             =   960
      Width           =   576
   End
   Begin VB.Label Label1 
      Caption         =   "列印別："
      Height          =   180
      Index           =   0
      Left            =   72
      TabIndex        =   6
      Top             =   540
      Width           =   756
   End
End
Attribute VB_Name = "frm060314"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String
Dim iPrint As Integer, Page As Integer, strTemp(0 To 7) As String
Dim PLeft(0 To 7) As Integer, strTemp1 As Variant, strTemp2 As Variant
'Add By Cheng 2002/09/17
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
   'Add By Cheng 2002/09/17
   blnClkSure = False
   
      Printer.Orientation = 2
      DoEvents
     If Len(txt1(0)) = 0 Then
        s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        Exit Sub
     Else
         If Len(txt1(2)) = 0 Then
             s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
             txt1(1).SetFocus
             txt1_GotFocus (1)
             Exit Sub
         Else
            'Add By Cheng 2002/03/20
            If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            'Add By Cheng 2002/09/17
            If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
               If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
                  MsgBox "日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
            
             If Len(txt1(3)) = 0 Then
                 s = MsgBox("排序條件不可空白!!", , "USER 輸入錯誤")
                 txt1(3).SetFocus
                 Exit Sub
             Else
                 Screen.MousePointer = vbHourglass
                 Me.Enabled = False
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
Screen.MousePointer = vbHourglass
ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
cnnConnection.Execute "DELETE FROM R060314 WHERE ID='" & strUserNum & "' "
Select Case Val(txt1(0))
Case 1
    pub_QL05 = pub_QL05 & ";" & Label1(0) & "1.收文" 'Add By Sindy 2010/12/13
    pub_QL05 = pub_QL05 & ";收文" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/13
    'Modify By Cheng 2002/10/31
    '國籍為申請人的國籍
'     '91.10.28 MODIFY BY SONIA
'     'strSQL = "SELECT " & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(PA09,'000',NA03,NA04),NVL(ST02,CP13),NVL(NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06))),NVL(PA05,NVL(PA06,PA07))," & SQLDate("PA10") & ",NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,FAGENT,CUSTOMER,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=ST01(+) AND FA10=NA01(+) AND " & SQLNewFag("PA26", "CU") & " AND " & SQLNewFag("PA75", "FA") & " AND CP01='FCP' AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " AND CP31='Y' AND CP57 IS NULL "
'     strSQL = "SELECT " & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(PA09,'000',NA03,NA04),NVL(ST02,CP13),NVL(NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06))),NVL(PA05,NVL(PA06,PA07))," & SQLDate("PA10") & ",NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,FAGENT,CUSTOMER,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=ST01(+) AND FA10=NA01(+) AND " & SQLNewFag("PA26", "CU") & " AND " & SQLNewFag("PA75", "FA") & " AND CP01='FCP' AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " AND CP10>='101' AND CP10<='105' AND CP09<'B' AND CP57 IS NULL "
'     '91.10.28 END
    strSql = "SELECT " & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(PA09,'000',NA03,NA04),NVL(ST02,CP13),NVL(NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90))),NVL(PA05,NVL(PA06,PA07))," & SQLDate("PA10") & ",NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,FAGENT,CUSTOMER,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=ST01(+) AND CU10=NA01(+) AND " & SQLNewFag("PA26", "CU") & " AND " & SQLNewFag("PA75", "FA") & " AND CP01='FCP' AND CP05>=" & Val(ChangeTStringToWString(txt1(1))) & " AND CP05<=" & Val(ChangeTStringToWString(txt1(2))) & " AND CP10>='101' AND CP10<='105' AND CP09<'B' AND CP57 IS NULL "
Case 2
    pub_QL05 = pub_QL05 & ";" & Label1(0) & "2.發文" 'Add By Sindy 2010/12/13
    pub_QL05 = pub_QL05 & ";發文" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/13
    'Modify By Cheng 2002/10/31
    '國籍為申請人的國籍
'     '91.10.28 MODIFY BY SONIA
'     'strSQL = "SELECT " & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(PA09,'000',NA03,NA04),NVL(ST02,CP13),NVL(NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06))),NVL(PA05,NVL(PA06,PA07))," & SQLDate("PA10") & ",NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,FAGENT,CUSTOMER,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=ST01(+) AND FA10=NA01(+) AND " & SQLNewFag("PA26", "CU") & " AND " & SQLNewFag("PA75", "FA") & " AND CP01='FCP' AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " AND CP31='Y' AND CP57 IS NULL "
'     strSQL = "SELECT " & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(PA09,'000',NA03,NA04),NVL(ST02,CP13),NVL(NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),NVL(CU04,NVL(cu05||' '||cu88||' '||cu89||' '||cu90,CU06))),NVL(PA05,NVL(PA06,PA07))," & SQLDate("PA10") & ",NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,FAGENT,CUSTOMER,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=ST01(+) AND FA10=NA01(+) AND " & SQLNewFag("PA26", "CU") & " AND " & SQLNewFag("PA75", "FA") & " AND CP01='FCP' AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " AND CP10>='101' AND CP10<='105' AND CP09<'B' AND CP57 IS NULL "
'     '91.10.28 END
    strSql = "SELECT " & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,DECODE(PA09,'000',NA03,NA04),NVL(ST02,CP13),NVL(NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90))),NVL(PA05,NVL(PA06,PA07))," & SQLDate("PA10") & ",NVL(DECODE(PA09,'000',CPM03,CPM04),CP10),'" & strUserNum & "'  FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,FAGENT,CUSTOMER,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND CP13=ST01(+) AND CU10=NA01(+) AND " & SQLNewFag("PA26", "CU") & " AND " & SQLNewFag("PA75", "FA") & " AND CP01='FCP' AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " AND CP10>='101' AND CP10<='105' AND CP09<'B' AND CP57 IS NULL "
Case Else
End Select
If txt1(3) = "1" Then
   pub_QL05 = pub_QL05 & ";" & Label1(2) & "1.收/發文日" 'Add By Sindy 2010/12/13
Else
   pub_QL05 = pub_QL05 & ";" & Label1(2) & "2.本所案號" 'Add By Sindy 2010/12/13
End If
cnnConnection.Execute "INSERT INTO R060314 " & strSql
strSql = "SELECT * FROM R060314 WHERE ID='" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/13
Else
   InsertQueryLog (0) 'Add By Sindy 2010/12/13
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
PrintData
Screen.MousePointer = vbDefault
End Sub

Sub PrintData()
If txt1(3) = "1" Then
    strSql = "SELECT * FROM R060314 WHERE ID='" & strUserNum & "' ORDER BY R047001 "
Else
    strSql = "SELECT * FROM R060314 WHERE ID='" & strUserNum & "' ORDER BY R047002 "
End If
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 7
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(2) = StrToStr(strTemp(2), 4)
            strTemp(3) = StrToStr(strTemp(3), 4)
            strTemp(4) = StrToStr(strTemp(4), 15)
            strTemp(5) = StrToStr(strTemp(5), 10)
            strTemp(7) = StrToStr(strTemp(7), 6)
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
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
If txt1(0) = "1" Then
    Printer.Print "收文筆數：" & CheckStr(adoRecordset.RecordCount)
Else
    Printer.Print "發文筆數：" & CheckStr(adoRecordset.RecordCount)
End If
Printer.EndDoc
ShowPrintOk
CheckOC
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.FontName = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5800
Printer.CurrentY = iPrint
Printer.Print "FCP 新案收/發文明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
If txt1(0) = "1" Then
    Printer.Print "收文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
Else
    Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
End If
iPrint = iPrint + 300
Printer.CurrentX = 500
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
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
If txt1(0) = "1" Then
    Printer.Print "收文日"
Else
    Printer.Print "發文日"
End If
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "國    籍"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "代理人/申請人名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "專利名稱"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "申請日"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
If iPrint >= 10000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
End Sub


Sub PrintDatil()
For i = 0 To 7
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1600
PLeft(2) = 3500
PLeft(3) = 5000
PLeft(4) = 6500
PLeft(5) = 10100
PLeft(6) = 13000
PLeft(7) = 14500
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm060314 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/17
   Select Case Index
   Case 0
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 3
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     If txt1(0) = "" Then Exit Sub 'Added by Morgan 2011/12/1 沒輸不必檢查
     Select Case Val(txt1(0))
     Case 1, 2
     Case Else
          s = MsgBox("列印別只能 1 或 2 !!", , "USER 輸入錯誤")
          txt1(0).SetFocus
          txt1(0).SelStart = 0
          txt1(0).SelLength = Len(txt1(0))
          Exit Sub
     End Select
Case 2
   'Modify By Cheng 2002/09/17
   If blnClkSure = False Then
      If RunNick(txt1(1), txt1(2)) Then
         txt1(1).SetFocus
         txt1_GotFocus (1)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If
Case 3
   'Modify By Cheng 2002/09/26
   If Me.txt1(3).Text <> "" Then
     Select Case Val(txt1(3))
     Case 1, 2
     Case Else
          s = MsgBox("排序條件只能 1 或 2 !!", , "USER 輸入錯誤")
          txt1(3).SetFocus
          txt1(3).SelStart = 0
          txt1(3).SelLength = Len(txt1(3))
          Exit Sub
     End Select
   End If
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '日期起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
'Add By Cheng 2002/09/27
Case 3
   Select Case Val(txt1(3))
   Case 1, 2
   Case Else
      s = MsgBox("排序條件只能 1 或 2 !!", , "USER 輸入錯誤")
      Cancel = True
      txt1(3).SetFocus
      txt1(3).SelStart = 0
      txt1(3).SelLength = Len(txt1(3))
      Exit Sub
   End Select
End Select
End Sub
