VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090106_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "委查人委查明細表"
   ClientHeight    =   4584
   ClientLeft      =   876
   ClientTop       =   1008
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4584
   ScaleWidth      =   6600
   Begin VB.Frame Frame1 
      Caption         =   "列印順序："
      ForeColor       =   &H000000FF&
      Height          =   1850
      Left            =   744
      TabIndex        =   10
      Top             =   1680
      Width           =   4950
      Begin VB.OptionButton OptBtn 
         Caption         =   "依收件日順序"
         Height          =   315
         Index           =   3
         Left            =   210
         TabIndex        =   6
         Top             =   1380
         Width           =   2175
      End
      Begin VB.OptionButton OptBtn 
         Caption         =   "依委查日順序"
         Height          =   315
         Index           =   2
         Left            =   210
         TabIndex        =   5
         Top             =   1010
         Width           =   2175
      End
      Begin VB.OptionButton OptBtn 
         Caption         =   "依委查人順序"
         Height          =   315
         Index           =   1
         Left            =   210
         TabIndex        =   4
         Top             =   640
         Width           =   2175
      End
      Begin VB.OptionButton OptBtn 
         Caption         =   "依委查單號順序"
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   270
         Width           =   2175
      End
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   3
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3672
      Width           =   255
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   1
      Left            =   1692
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1200
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   2
      Left            =   2856
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1224
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   0
      Left            =   1692
      MaxLength       =   6
      TabIndex        =   0
      Top             =   672
      Width           =   825
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4668
      TabIndex        =   8
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   5496
      TabIndex        =   9
      Top             =   120
      Width           =   800
   End
   Begin MSForms.Label LblTmq02NM 
      Height          =   255
      Left            =   2610
      TabIndex        =   17
      Top             =   690
      Width           =   1875
      VariousPropertyBits=   27
      Size            =   "3307;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      Caption         =   "（請輸入民國年）"
      Height          =   336
      Left            =   3852
      TabIndex        =   16
      Top             =   1224
      Width           =   1692
   End
   Begin VB.Label Label6 
      Caption         =   "顯示方式："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   768
      TabIndex        =   15
      Top             =   3684
      Width           =   900
   End
   Begin VB.Label Label5 
      Caption         =   "（1：螢幕 2：報表）"
      Height          =   336
      Left            =   2064
      TabIndex        =   14
      Top             =   3684
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "委查人："
      Height          =   300
      Left            =   768
      TabIndex        =   13
      Top             =   696
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "委查日期："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   768
      TabIndex        =   12
      Top             =   1224
      Width           =   912
   End
   Begin VB.Label Label3 
      Caption         =   "－"
      Height          =   288
      Left            =   2604
      TabIndex        =   11
      Top             =   1260
      Width           =   276
   End
End
Attribute VB_Name = "frm090106_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/11 改成Form2.0 ; LblTmq02NM ; Printer列印未改--'Memo by Lydia 2024/11/15 Printer逐字檢查Unicode文字改以圖片方式列印
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer
Dim iPrint As Integer, Page As Integer, PLeft(0 To 10) As Integer
Dim strSql As String, strTemp(0 To 10) As String, OrderCondition As String
Dim Rs As New ADODB.Recordset
Dim BlnCheck As Boolean
Dim Xo As Long, Yo As Long 'Added by Lydia 2024/11/15
Dim bolIsTMA As Boolean 'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)

Private Sub cmdOK_Click()
   Txtdata_LostFocus (0)
   If BlnCheck Then Exit Sub
   Txtdata_LostFocus (2)
   If BlnCheck Then Exit Sub
   'Add By Cheng 2002/03/21
   If PUB_CheckKeyInDate(Me.Txtdata(1)) = -1 Then
      Me.Txtdata(1).SetFocus
      Txtdata_GotFocus 1
      Exit Sub
   End If
   If PUB_CheckKeyInDate(Me.Txtdata(2)) = -1 Then
      Me.Txtdata(2).SetFocus
      Txtdata_GotFocus 2
      Exit Sub
   End If

   OptBtn_Check
   If BlnCheck Then Exit Sub
   Txtdata_LostFocus (3)
   If BlnCheck Then Exit Sub
   Me.Enabled = False
   Me.Hide
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/14 清除查詢印表記錄檔欄位
   If Txtdata(3) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label6 & "1：螢幕" 'Add By Sindy 2010/12/14
      Query_Sub
   Else
      pub_QL05 = pub_QL05 & ";" & Label6 & "2：報表" 'Add By Sindy 2010/12/14
      Print_Sub
   End If
      If bolToEndByNick = True Then
         cmdExit_Click
         Exit Sub
      End If
   Me.Enabled = True
   Me.Show
End Sub

Private Sub cmdExit_Click()
   Unload Me
   Set frm090106_2 = Nothing
End Sub

Private Sub Form_Load()
   Me.Height = 4965
   Me.Width = 7665
   MoveFormToCenter Me
   BlnCheck = False: bolToEndByNick = False
   Txtdata(3) = "1"
   Me.OptBtn(0).Value = True
End Sub

Private Sub OptBtn_Check()
Dim BoxOptBtn As OptionButton
   BlnCheck = True
   For Each BoxOptBtn In frm090106_1.OptBtn
      If BoxOptBtn.Value Then
         BlnCheck = False
      End If
   Next
   If BlnCheck Then
      s = MsgBox("請選擇列印順序", , "使用者輸入錯誤")
   End If
   End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm090106_1 = Nothing
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse Txtdata(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Txtdata_LostFocus(Index As Integer)
Dim strTemp As String
Dim strTemp1 As String
   BlnCheck = False
   Select Case Index
   Case 0
      If Txtdata(0) <> Empty Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(txtData(0).Text, strTemp, strTemp1) Then
         'Modified by Lydia 2018/07/02
         'If ClsPDGetStaff(Txtdata(0).Text, strTemp, strTemp1) Then
         strTemp = GetStaffName(Txtdata(0).Text, True)
         If strTemp <> "" Then
         'end 2018/07/02
            LblTmq02NM.Caption = strTemp
         Else
            LblTmq02NM.Caption = ""
            Txtdata(0).SetFocus
            TextInverse Txtdata(0)
            BlnCheck = True
         Exit Sub
         End If
      Else
         LblTmq02NM.Caption = ""
      End If
   Case 2
      If Txtdata(1) = Empty And Txtdata(2) = Empty Then
         s = MsgBox("請輸入委查日期條件", , "使用者輸入錯誤")
         Txtdata(1).SetFocus
         BlnCheck = True
         Exit Sub
      End If
      'Modify by Morgan 2010/8/16 百年蟲
      'If Txtdata(1) > Txtdata(2) Then
      If Val(Txtdata(1)) > Val(Txtdata(2)) Then
         s = MsgBox("委查日期範圍錯誤", , "使用者輸入錯誤")
         Txtdata(1).SetFocus
         TextInverse Txtdata(1)
         BlnCheck = True
         Exit Sub
      End If
   Case 3
      If Txtdata(3) <> "1" And Txtdata(3) <> "2" Then
         s = MsgBox("顯示方式只可輸入'1'或'2'", , "使用者輸入錯誤")
         Txtdata(3).SetFocus
         TextInverse Txtdata(3)
         BlnCheck = True
         Exit Sub
      End If
   Case Else
   End Select
End Sub

Private Sub Query_Sub()
   frm090106_2.Show
   frm090106_2.Hide
   frm090106_2.MousePointer = vbHourglass
   frm090106_2.GridData
   frm090106_2.MousePointer = vbDefault
   If frm090106_2.Enabled = True Then
      frm090106_2.Show
   Else
      s = MsgBox("資料庫中沒有符合的資料!!", , "請檢查條件")
   End If
   Do
      DoEvents
      If bolToEndByNick = True Then
         cmdExit_Click
         Exit Sub
      End If
   Loop Until Not frm090106_2.Visible
   Unload frm090106_2
End Sub

Private Sub Print_Sub()
Dim OrderStr As String, SubSQL As String, strCondition As String
   Printer.Orientation = 2
   Screen.MousePointer = vbHourglass
   
   'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)
   If DBDATE(Txtdata(1)) >= 查名單網中系統啟用日 Or DBDATE(Txtdata(2)) >= 查名單網中系統啟用日 Then
      bolIsTMA = True
   Else
      bolIsTMA = False
   End If
   'end 2024//11/15
   
   If Me.OptBtn(0).Value Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         OrderStr = "委查單號"
      Else
      'end 2024/11/15
         OrderStr = "TMQ01"
      End If
      OrderCondition = "委查單號"
      pub_QL05 = pub_QL05 & ";" & Frame1.Caption & OptBtn(0).Caption  'Add By Sindy 2010/12/14
   Else
      If Me.OptBtn(1).Value Then
         'Added by Lydia 2024/11/15 查名單(網中)
         If bolIsTMA = True Then
            OrderStr = "委查人, 委查單號"
         Else
         'end 2024/11/15
            OrderStr = "TMQ02, TMQ01"
         End If
         OrderCondition = "委查人"
         pub_QL05 = pub_QL05 & ";" & Frame1.Caption & OptBtn(1).Caption  'Add By Sindy 2010/12/14
      Else
         If Me.OptBtn(2).Value Then
            'Added by Lydia 2024/11/15 查名單(網中)
            If bolIsTMA = True Then
               OrderStr = "委查日期, 委查單號"
            Else
            'end 2024/11/15
               OrderStr = "TMQ04, TMQ01"
            End If
            OrderCondition = "委查日"
            pub_QL05 = pub_QL05 & ";" & Frame1.Caption & OptBtn(2).Caption  'Add By Sindy 2010/12/14
         Else
            If Me.OptBtn(3).Value Then
               'Added by Lydia 2024/11/15 查名單(網中)
               If bolIsTMA = True Then
                  OrderStr = "收件日期, 委查單號"
               Else
               'end 2024/11/15
                  OrderStr = "TMQ05, TMQ01"
               End If
               OrderCondition = "收件日"
               pub_QL05 = pub_QL05 & ";" & Frame1.Caption & OptBtn(3).Caption  'Add By Sindy 2010/12/14
            End If
         End If
      End If
   End If
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   SubSQL = "": strCondition = ""
   If Txtdata(0) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         strCondition = strCondition + " AND TMA08 = '" & Txtdata(0) & "'"
      Else
      'end 2024/11/15
         strCondition = strCondition + " AND TMQ02 = '" & Txtdata(0) & "'"
      End If
      pub_QL05 = pub_QL05 & ";" & Label1 & Txtdata(0) & LblTmq02NM 'Add By Sindy 2010/12/14
   End If
   If Txtdata(1) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TO_CHAR(TMA04,'YYYYMMDD')>=" & Val(ChangeTStringToWString(Txtdata(1))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ04>=" & Val(ChangeTStringToWString(Txtdata(1))) & ""
      End If
   End If
   If Txtdata(2) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TO_CHAR(TMA04,'YYYYMMDD')<=" & Val(ChangeTStringToWString(Txtdata(2))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ04<=" & Val(ChangeTStringToWString(Txtdata(2))) & ""
      End If
   End If
   If Txtdata(1) <> Empty Or Txtdata(2) <> Empty Then
      pub_QL05 = pub_QL05 & ";" & Label2 & Txtdata(1) & "-" & Txtdata(2) 'Add By Sindy 2010/12/14
   End If
   
   If Len(SubSQL) <> 0 Then
      SubSQL = " WHERE " & Mid(SubSQL, 5)
   End If
   'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入＞＞TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
   If bolIsTMA = True Then
      strSql = " SELECT TMA01 AS 委查單號, " & PUB_GetTMAforClass & " AS 類別組群, NVL(S1.ST02, TMA08) 委查人, TO_CHAR(TMA04,'YYYYMMDD') AS 委查日期, TMA09 AS 收件日期, TMA36 AS 中文, TMA37 AS 英文, TMA38 AS 圖形, NVL(S2.ST02, TMA10) AS 查名人, TMA14 AS 查覆日期, NVL(TMA11,TMA12) AS 期限日期" & _
               " FROM TMQAPPFORM, STAFF S1, STAFF S2 " & SubSQL & strCondition & " AND TMA08 = S1.ST01(+) AND TMA10 = S2.ST01(+) AND TO_CHAR(TMA04,'YYYYMMDD')>='20240601' ORDER BY " & OrderStr
   Else
   'end 2024/11/15
      strSql = "SELECT TMQ01, TMQ03, NVL(S1.ST02, TMQ02), TMQ04, TMQ05, TMQ07, TMQ08, TMQ09, NVL(S2.ST02, TMQ10), TMQ11, TMQ06 FROM TRADEMARKQUERY, STAFF S1, STAFF S2 " & SubSQL & strCondition & " AND TMQ02 = S1.ST01(+) AND TMQ10 = S2.ST01(+) ORDER BY " & OrderStr
   End If
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Rs.RecordCount > 0 Then
      InsertQueryLog (Rs.RecordCount) 'Add By Sindy 2010/12/14
      With Rs
         .MoveFirst
         Page = 1
         PrintTitle
         j = 0
'         'FRM100.Show
         DoEvents
'         'FRM100.Tag = Trim(str(.RecordCount)) & "=0"
         Do While .EOF = False
            For i = 0 To 10
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Modified by Lydia 2024/11/15 23>>42
            strTemp(1) = StrConv(MidB(StrConv(strTemp(1), vbFromUnicode), 1, 42), vbUnicode)
            strTemp(2) = StrConv(MidB(StrConv(strTemp(2), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(3) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(3)))
            strTemp(4) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(4)))
            strTemp(5) = Format(strTemp(5), "#,###")
            strTemp(6) = Format(strTemp(6), "#,###")
            strTemp(7) = Format(strTemp(7), "#,###")
            strTemp(8) = StrConv(MidB(StrConv(strTemp(8), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(9) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(9)))
            strTemp(10) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(10)))
            PrintDetail
            If iPrint > 10000 Then
               Page = Page + 1
               Printer.CurrentX = 500
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               Printer.NewPage
               PrintTitle
            End If
            j = j + 1
'            'FRM100.Tag = Trim(str(.RecordCount)) & "=" & Trim(str(j))
'            'FRM100.StrMenu
            DoEvents
            .MoveNext
         Loop
         Printer.CurrentX = 500
         Printer.CurrentY = iPrint
         Printer.Print String(200, "-")
         Printer.EndDoc
         s = MsgBox("列印完成!!", , "列印成功")
'         Unload frm100
      End With
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/14
      s = MsgBox("資料庫中沒有符合的資料!!", , "沒有資料")
   End If
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub PrintTitle() '列印抬頭

   GetPleft
   iPrint = 500
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6500
   Printer.CurrentY = iPrint
   Printer.Print "委查人委查明細表"
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   'Modified by Lydia 2024/11/15
   'Printer.CurrentX = 500
   'Printer.CurrentY = iPrint
   'Printer.Print "列印人　：" & strUserName
   Xo = 500
   Yo = iPrint
   PUB_PrintUnicodeText "列印人　：" & strUserName, Xo, Yo, 0
   'end 2024/11/15
   Printer.CurrentX = 6700
   Printer.CurrentY = iPrint
   Printer.Print "委查日期：" & ChangeTStringToTDateString(Txtdata(1))
   Printer.CurrentX = 8700
   Printer.CurrentY = iPrint
   Printer.Print "－"
   Printer.CurrentX = 8950
   Printer.CurrentY = iPrint
   Printer.Print ChangeTStringToTDateString(Txtdata(2))
   Printer.CurrentX = 13300
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print "列印順序：" & OrderCondition
   Printer.CurrentX = 13300
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   Printer.Font.Underline = True
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "委查單號"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   'Added by Lydia 2024/11/15 查名單(網中)
   If bolIsTMA = True Then
      Printer.Print "類別組群"
   Else
   'end 2024/11/15
      Printer.Print "組  　　　群"
   End If
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "委查人"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "委查日期"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "收件日期"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "中文"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "英文"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "圖形"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "查名人"
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "查覆日期"
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print "期限日期"
   Printer.Font.Underline = False
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
End Sub

Private Sub PrintDetail()

   For i = 0 To 4
      'Added by Lydia 2024/11/15 逐字檢查Unicode文字改以圖片方式列印
      If i = 2 Then '委查人
         Xo = PLeft(i)
         Yo = iPrint
         PUB_PrintUnicodeText strTemp(i), Xo, Yo, 0
      Else
      'end 2024/11/15
         Printer.CurrentX = PLeft(i)
         Printer.CurrentY = iPrint
         Printer.Print strTemp(i)
      End If
   Next i
   For i = 5 To 7
      Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i))
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   Next i
   For i = 8 To 10
      'Added by Lydia 2024/11/15 逐字檢查Unicode文字改以圖片方式列印
      If i = 8 Then '查名人
         Xo = PLeft(i)
         Yo = iPrint
         PUB_PrintUnicodeText strTemp(i), Xo, Yo, 0
      Else
      'end 2024/11/15
         Printer.CurrentX = PLeft(i)
         Printer.CurrentY = iPrint
         Printer.Print strTemp(i)
      End If
   Next i
   iPrint = iPrint + 300
End Sub

Private Sub GetPleft()
   Erase PLeft
   PLeft(0) = 500
   PLeft(1) = 1700
   PLeft(2) = 6300
   PLeft(3) = 7500
   PLeft(4) = 8600
   PLeft(5) = 9700
   PLeft(6) = 10500
   PLeft(7) = 11300
   PLeft(8) = 12100
   PLeft(9) = 13300
   PLeft(10) = 14400
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '期限日期起, 迄
   If PUB_CheckKeyInDate(Me.Txtdata(Index)) = -1 Then
      Cancel = True
      Me.Txtdata(Index).SetFocus
      Txtdata_GotFocus Index
   End If
End Select
End Sub
