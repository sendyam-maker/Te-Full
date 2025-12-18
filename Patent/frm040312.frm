VERSION 5.00
Begin VB.Form frm040312 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸發明案參考資料表"
   ClientHeight    =   1455
   ClientLeft      =   3630
   ClientTop       =   3690
   ClientWidth     =   3180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3180
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1824
      MaxLength       =   1
      TabIndex        =   0
      Top             =   828
      Width           =   315
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2316
      TabIndex        =   3
      Top             =   20
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1488
      TabIndex        =   2
      Top             =   20
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1428
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1152
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "是否列印上次資料："
      Height          =   180
      Index           =   5
      Left            =   144
      TabIndex        =   10
      Top             =   864
      Width           =   1632
   End
   Begin VB.Label Label1 
      Caption         =   "(Y/N)"
      Height          =   180
      Index           =   4
      Left            =   2196
      TabIndex        =   9
      Top             =   864
      Width           =   456
   End
   Begin VB.Label Label1 
      Caption         =   "(Y/N)"
      Height          =   180
      Index           =   3
      Left            =   1824
      TabIndex        =   8
      Top             =   1176
      Width           =   456
   End
   Begin VB.Label lbl1 
      Height          =   180
      Left            =   564
      TabIndex        =   7
      Top             =   492
      Width           =   768
   End
   Begin VB.Label Label1 
      Caption         =   "是否確定列印："
      Height          =   180
      Index           =   2
      Left            =   144
      TabIndex        =   6
      Top             =   1212
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "筆資料"
      Height          =   180
      Index           =   1
      Left            =   1380
      TabIndex        =   5
      Top             =   504
      Width           =   624
   End
   Begin VB.Label Label1 
      Caption         =   "現有"
      Height          =   180
      Index           =   0
      Left            =   132
      TabIndex        =   4
      Top             =   492
      Width           =   480
   End
End
Attribute VB_Name = "frm040312"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, strSQL2 As String, i As Integer, j As Integer, s As Integer
Dim strTemp(0 To 8) As String, strTemp1 As Variant, strTemp2 As Variant, StrTemp8(0 To 1) As String, k As Integer
Dim PLeft(0 To 8) As Integer, iPrint As Integer, Page As Integer, strTemp3(0 To 4) As String, StrTemp5(0 To 4) As String, StrTemp6(0 To 4) As String, F4312BOL As Boolean

Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
   If txt1(1).Text = "N" Then
      Unload Me
   Else
     If Len(Trim(txt1(0))) <> 0 And Len(Trim(txt1(1))) <> 0 Then
        F4312BOL = True
        Screen.MousePointer = vbHourglass
        If UCase(Trim(txt1(0))) = "N" Then
            s = MsgBox("確定重新產生資料??  ", vbYesNo, "舊資料將被刪除!!")
            If s = 6 Then
                Process
                cnnConnection.Execute "DELETE FROM PERMITRECORD "
            Else
                If s = 7 Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
        If UCase(Trim(txt1(1))) = "Y" Then
            If F4312BOL = True Then
               PrintData
               Form_Load
            End If
        End If
        Screen.MousePointer = vbDefault
     Else
        s = MsgBox("皆不可空白!!", , "USER 輸入錯誤")
        Exit Sub
     End If
   End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R040312 WHERE ID='" & strUserNum & "' "
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        DoEvents
        k = 0
        Do While .EOF = False
            For i = 0 To 4
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            For i = 5 To 7
                strTemp(i + 1) = CheckStr(.Fields(i))
            Next i
            strSql = "SELECT CP57 FROM CASEPROGRESS WHERE CP31='Y' AND CP01='" & SystemNumber(strTemp(0), 1) & "' AND CP02='" & SystemNumber(strTemp(0), 2) & "' AND CP03='" & SystemNumber(strTemp(0), 3) & "' AND CP04='" & SystemNumber(strTemp(0), 4) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                strTemp(5) = CheckStr(adoRecordset1.Fields(0))
            Else
                strTemp(5) = ""
            End If
            CheckOC2
            strTemp(2) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(2)))
            strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
            strTemp(5) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(5)))
            strTemp(8) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(8)))
            Select Case strTemp(4)
            Case "1"
                 strTemp(4) = "准"
            Case "2"
                 strTemp(4) = "駁"
            Case Else
                 strTemp(4) = ""
            End Select
            strSql = "INSERT INTO R040312 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "','" & ChgSQL(strTemp(3)) & "','" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & ChgSQL(strTemp(8)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            k = k + 1
            DoEvents
            .MoveNext
        Loop
    End With
Else
   ShowNoData
   F4312BOL = False
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Screen.MousePointer = vbDefault
End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "大陸發明案參考資料表"
iPrint = iPrint + 300
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
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
Printer.Print "CF案號"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "申請日"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "准駁日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "結果"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "取消收文日"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "大陸案號"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "申請日"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 2500
PLeft(2) = 5000
'Modify By Cheng 2002/06/20
'PLeft(3) = 6000
'PLeft(4) = 7000
'PLeft(5) = 7700
'PLeft(6) = 9000
'PLeft(7) = 11000
'PLeft(8) = 14000
PLeft(3) = 6000 + 250
PLeft(4) = 7000 + 250
PLeft(5) = 7700 + 250
PLeft(6) = 9000 + 250
PLeft(7) = 11000 + 250
PLeft(8) = 14000 + 250
End Sub

Sub PrintDatil()
For i = 0 To 8
    Printer.CurrentX = PLeft(i)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(i)
Next i
iPrint = iPrint + 300
End Sub


Sub PrintData()
Screen.MousePointer = vbHourglass
strSql = "SELECT * FROM R040312 WHERE ID='" & strUserNum & "' ORDER BY R028001,R028004,R028007 "
CheckOC
strTemp3(0) = " "
strTemp3(1) = " "
strTemp3(2) = " "
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
            If strTemp3(0) <> strTemp(0) Then
                strTemp3(0) = strTemp(0)
                strTemp3(1) = strTemp(1)
                strTemp3(2) = strTemp(2)
            Else
                strTemp(0) = ""
                If strTemp3(1) <> strTemp(1) Then
                    strTemp3(1) = strTemp(1)
                    strTemp3(2) = strTemp(2)
                Else
                    strTemp(1) = ""
                    If strTemp3(2) <> strTemp(2) Then
                        strTemp3(2) = strTemp(2)
                    Else
                        strTemp(2) = ""
                    End If
                End If
            End If
            strTemp(1) = StrToStr(strTemp(1), 10)
            strTemp(7) = StrToStr(strTemp(7), 10)
            If iPrint >= 10000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            PrintDatil
            .MoveNext
        Loop
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print "共 " & Trim(str(.RecordCount)) & " 筆"
        Printer.EndDoc
    End With
    ShowPrintOk
    Screen.MousePointer = vbDefault
Else
   ShowNoData
   Screen.MousePointer = vbDefault
End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'STRSQL="SELECT PR01||'-'||PR02||'-'||PR03||'-'||PR04,NVL(P1.PA05,NVL(P1.PA06,P1.PA07)),P1.PA10,P1.PA20,P1.PA16,CP57,CM01||'-'||CM02||'-'||CM03||'-'||CM04,NVL(P2.PA05,NVL(P2.PA06,P2.PA07)),P2.PA10,CP31 FROM PERMITRECORD,CASEMAP,PATENT P1,PATENT P2,CASEPROGRESS WHERE CM05=PR01 AND CM06=PR02 AND CM07=PR03 AND CM08=PR04 AND CM10='2' AND PR01=P1.PA01(+) AND PR02=P1.PA02(+) AND PR03=P1.PA03(+) AND PR04=P1.PA04(+) AND CM01=P2.PA01(+) AND CM02=P2.PA02(+) AND CM03=P2.PA03(+) AND CM04=P2.PA04(+) AND PR01=CP01(+) AND PR02=CP02(+) AND PR03=CP03(+) AND PR04=CP04(+) AND CP31='Y' "
strSql = "SELECT PR01||'-'||PR02||'-'||PR03||'-'||PR04,NVL(P1.PA05,NVL(P1.PA06,P1.PA07)),P1.PA10,P1.PA20,P1.PA16,CM01||'-'||CM02||'-'||CM03||'-'||CM04,NVL(P2.PA05,NVL(P2.PA06,P2.PA07)),P2.PA10 FROM PERMITRECORD,CASEMAP,PATENT P1,PATENT P2 WHERE PR01=CM05(+) AND PR02=CM06(+) AND PR03=CM07(+) AND PR04=CM08(+) AND CM10='2' AND PR01=P1.PA01(+) AND PR02=P1.PA02(+) AND PR03=P1.PA03(+) AND PR04=P1.PA04(+) AND CM01=P2.PA01(+) AND CM02=P2.PA02(+) AND CM03=P2.PA03(+) AND CM04=P2.PA04(+) AND PR01 IN (" & SQLGrpStr("", 1) & ") "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
lbl1.Caption = Trim(str(adoRecordset.RecordCount))
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm040312 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case UCase(txt1(Index))
Case "Y", "N", ""
Case Else
     Select Case Index
     Case 0
          s = MsgBox("是否列印上次資料只能輸入 Y 或 N !!", , "USER 輸入錯誤")
     Case 1
          s = MsgBox("是否確定列印只能輸入 Y 或 N !!", , "USER 輸入錯誤")
     Case Else
     End Select
     txt1(Index).SetFocus
     txt1(Index).SelStart = 0
     txt1(Index).SelLength = Len(txt1(Index))
End Select
End Sub

