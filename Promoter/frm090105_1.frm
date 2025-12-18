VERSION 5.00
Begin VB.Form frm090105_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "期限過期明細表"
   ClientHeight    =   4224
   ClientLeft      =   1536
   ClientTop       =   1332
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4224
   ScaleWidth      =   6360
   Begin VB.Frame Frame1 
      Caption         =   "列印順序："
      ForeColor       =   &H000000FF&
      Height          =   1850
      Left            =   672
      TabIndex        =   9
      Top             =   1392
      Width           =   4950
      Begin VB.OptionButton OptBtn 
         Caption         =   "依收件日順序"
         Height          =   315
         Index           =   3
         Left            =   210
         TabIndex        =   5
         Top             =   1380
         Width           =   2175
      End
      Begin VB.OptionButton OptBtn 
         Caption         =   "依委查日順序"
         Height          =   315
         Index           =   2
         Left            =   210
         TabIndex        =   4
         Top             =   1010
         Width           =   2175
      End
      Begin VB.OptionButton OptBtn 
         Caption         =   "依查名人順序"
         Height          =   315
         Index           =   1
         Left            =   216
         TabIndex        =   3
         Top             =   624
         Width           =   2175
      End
      Begin VB.OptionButton OptBtn 
         Caption         =   "依委查單號順序"
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   270
         Width           =   2175
      End
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   3
      Left            =   1584
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3348
      Width           =   255
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   1
      Left            =   1620
      MaxLength       =   7
      TabIndex        =   0
      Top             =   888
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   2
      Left            =   2784
      MaxLength       =   7
      TabIndex        =   1
      Top             =   888
      Width           =   825
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4452
      TabIndex        =   7
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   5280
      TabIndex        =   8
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label7 
      Caption         =   "（請輸入民國年）"
      Height          =   336
      Left            =   3780
      TabIndex        =   14
      Top             =   936
      Width           =   1692
   End
   Begin VB.Label Label6 
      Caption         =   "顯示方式："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   696
      TabIndex        =   13
      Top             =   3396
      Width           =   900
   End
   Begin VB.Label Label5 
      Caption         =   "（1：螢幕 2：報表）"
      Height          =   336
      Left            =   1992
      TabIndex        =   12
      Top             =   3396
      Width           =   1692
   End
   Begin VB.Label Label2 
      Caption         =   "期限日期："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   696
      TabIndex        =   11
      Top             =   936
      Width           =   912
   End
   Begin VB.Label Label3 
      Caption         =   "－"
      Height          =   288
      Left            =   2532
      TabIndex        =   10
      Top             =   972
      Width           =   276
   End
End
Attribute VB_Name = "frm090105_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/11 Form2.0已檢查 (無需修改的物件); Printer列印未改--'Memo by Lydia 2024/11/15 Printer逐字檢查Unicode文字改以圖片方式列印
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
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
   If Txtdata(3) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label6 & "1：螢幕" 'Add By Sindy 2010/12/13
      Query_Sub
   Else
      pub_QL05 = pub_QL05 & ";" & Label6 & "2：報表" 'Add By Sindy 2010/12/13
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
   Set frm090105_2 = Nothing
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
   For Each BoxOptBtn In frm090105_1.OptBtn
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
   Set frm090105_1 = Nothing
End Sub

Private Sub Txtdata_GotFocus(Index As Integer)
   TextInverse Txtdata(Index)
End Sub

Private Sub txtdata_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdOK_Click
   Else
      KeyAscii = UpperCase(KeyAscii)
   End If
End Sub

Private Sub Txtdata_LostFocus(Index As Integer)
Dim strTemp As String
Dim strTemp1 As String
   BlnCheck = False
   Select Case Index
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
   frm090105_2.Show
   frm090105_2.Hide
   frm090105_2.MousePointer = vbHourglass
   frm090105_2.GridData
   frm090105_2.MousePointer = vbDefault
   If frm090105_2.Enabled = True Then
      frm090105_2.Show
   Else
      s = MsgBox("資料庫中沒有符合的資料!!", , "請檢查條件")
   End If
   Do
      DoEvents
      If bolToEndByNick = True Then
         cmdExit_Click
         Exit Sub
      End If
   Loop Until Not frm090105_2.Visible
   Unload frm090105_2
End Sub

Private Sub Print_Sub()
Dim OrderStr As String, SubSQL As String
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
      pub_QL05 = pub_QL05 & ";" & Frame1.Caption & OptBtn(0).Caption 'Add By Sindy 2010/12/13
   Else
      If Me.OptBtn(1).Value Then
         'Added by Lydia 2024/11/15 查名單(網中)
         If bolIsTMA = True Then
            OrderStr = "委查人, 委查單號"
         Else
         'end 2024/11/15
            OrderStr = "TMQ10, TMQ01"
         End If
         OrderCondition = "查名人"
         pub_QL05 = pub_QL05 & ";" & Frame1.Caption & OptBtn(1).Caption 'Add By Sindy 2010/12/13
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
            pub_QL05 = pub_QL05 & ";" & Frame1.Caption & OptBtn(2).Caption 'Add By Sindy 2010/12/13
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
               pub_QL05 = pub_QL05 & ";" & Frame1.Caption & OptBtn(3).Caption 'Add By Sindy 2010/12/13
            End If
         End If
      End If
   End If
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   SubSQL = ""
   If Txtdata(1) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND NVL(TMA11,TMA12)>=" & Val(ChangeTStringToWString(Txtdata(1))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ06>=" & Val(ChangeTStringToWString(Txtdata(1))) & ""
      End If
   End If
   If Txtdata(2) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND NVL(TMA11,TMA12)<=" & Val(ChangeTStringToWString(Txtdata(2))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ06<=" & Val(ChangeTStringToWString(Txtdata(2))) & ""
      End If
   End If
   If Txtdata(1) <> Empty Or Txtdata(2) <> Empty Then
      pub_QL05 = pub_QL05 & ";" & Label2 & Txtdata(1) & "-" & Txtdata(2) 'Add By Sindy 2010/12/13
   End If
   
   If Len(SubSQL) <> 0 Then
      SubSQL = " WHERE " & Mid(SubSQL, 5)
   End If
   'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入＞＞TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
   If bolIsTMA = True Then
      strSql = " SELECT NVL(TMA11,TMA12) AS 期限日期, " & PUB_GetTMAforClass & " AS 類別組群, TMA01 AS 委查單號, NVL(S1.ST02, TMA08) 委查人, TO_CHAR(TMA04,'YYYYMMDD') AS 委查日期, TMA09 AS 收件日期, " & _
               " TMA36 AS 中文, TMA37 AS 英文, TMA38 AS 圖形, NVL(S2.ST02, TMA10) AS 查名人, TMA14 AS 查覆日期 " & _
               " FROM TMQAPPFORM, STAFF S1, STAFF S2 " & SubSQL & " AND ((TMA14 IS NULL AND NVL(TMA11,TMA12) < " & strSrvDate(1) & ") OR (TMA14 IS NOT NULL AND TMA14 > NVL(TMA11,TMA12))) AND TMA08 = S1.ST01(+) AND TMA10 = S2.ST01(+) AND TO_CHAR(TMA04,'YYYYMMDD')>='20240601' ORDER BY " & OrderStr
   Else
   'end 2024/11/15
      strSql = "SELECT TMQ06, TMQ03, TMQ01, NVL(S1.ST02, TMQ02), TMQ04, TMQ05, TMQ07, TMQ08, TMQ09, NVL(S2.ST02, TMQ10), TMQ11 FROM TRADEMARKQUERY, STAFF S1, STAFF S2 " & SubSQL
      strSql = strSql + " AND ((TMQ11 IS NULL AND TMQ06 < " & strSrvDate(1) & ") OR (TMQ11 IS NOT NULL AND TMQ11 > TMQ06)) AND TMQ02 = S1.ST01(+) AND TMQ10 = S2.ST01(+) ORDER BY " & OrderStr
   End If
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Rs.RecordCount > 0 Then
      InsertQueryLog (Rs.RecordCount) 'Add By Sindy 2010/12/13
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
            strTemp(0) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(0)))
            'Modified by Lydia 2024/11/15 23>>42
            strTemp(1) = StrConv(MidB(StrConv(strTemp(1), vbFromUnicode), 1, 42), vbUnicode)
            strTemp(3) = StrConv(MidB(StrConv(strTemp(3), vbFromUnicode), 1, 8), vbUnicode)
            strTemp(4) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(4)))
            strTemp(5) = ChangeTStringToTDateString(ChangeWStringToTString(strTemp(5)))
            strTemp(6) = Format(strTemp(6), "#,###")
            strTemp(7) = Format(strTemp(7), "#,###")
            strTemp(8) = Format(strTemp(8), "#,###")
            strTemp(9) = StrConv(MidB(StrConv(strTemp(9), vbFromUnicode), 1, 8), vbUnicode)
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
            DoEvents
            .MoveNext
         Loop
         Printer.CurrentX = 500
         Printer.CurrentY = iPrint
         Printer.Print String(200, "-")
         Printer.EndDoc
         s = MsgBox("列印完成!!", , "列印成功")
      End With
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/13
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
   Printer.CurrentX = 6700
   Printer.CurrentY = iPrint
   Printer.Print "期限過期明細表"
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
   Printer.Print "期限日期：" & ChangeTStringToTDateString(Txtdata(1))
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
   Printer.Print "期限日期"
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
   Printer.Print "委查單號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "委查人"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "委查日期"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "收件日期"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "中文"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "英文"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "圖形"
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "查名人"
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print "查覆日期"
   Printer.Font.Underline = False
   iPrint = iPrint + 300
   Printer.CurrentX = 500
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
End Sub

Private Sub PrintDetail()

   For i = 0 To 5
      'Added by Lydia 2024/11/15 逐字檢查Unicode文字改以圖片方式列印
      If i = 3 Then '委查人
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
   For i = 6 To 8
   
      Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(strTemp(i))
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   Next i
   For i = 9 To 10
      'Added by Lydia 2024/11/15 逐字檢查Unicode文字改以圖片方式列印
      If i = 9 Then '查名人
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
   PLeft(1) = 1600
   PLeft(2) = 6200
   PLeft(3) = 7400
   PLeft(4) = 8600
   PLeft(5) = 9700
   PLeft(6) = 10800
   PLeft(7) = 11600
   PLeft(8) = 12400
   PLeft(9) = 13200
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
