VERSION 5.00
Begin VB.Form frm090109_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "商品組群委查統計表"
   ClientHeight    =   2928
   ClientLeft      =   576
   ClientTop       =   1020
   ClientWidth     =   7380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2928
   ScaleWidth      =   7380
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   4
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "1"
      Top             =   1770
      Width           =   255
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   3
      Left            =   1260
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2295
      Width           =   255
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   1
      Left            =   1224
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1260
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   2
      Left            =   2424
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1272
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   0
      Left            =   1224
      MaxLength       =   39
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "確定(&U)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5460
      TabIndex        =   5
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   6288
      TabIndex        =   6
      Top             =   120
      Width           =   800
   End
   Begin VB.Label Label9 
      Caption         =   "（1：群組  2：中文筆數  3：英文筆數  4：圖形筆數 ）"
      Height          =   330
      Left            =   1635
      TabIndex        =   15
      Top             =   1785
      Width           =   5445
   End
   Begin VB.Label Label8 
      Caption         =   "排序方式："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   330
      TabIndex        =   14
      Top             =   1785
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "(以，區隔)"
      Height          =   255
      Left            =   6030
      TabIndex        =   13
      Top             =   735
      Width           =   900
   End
   Begin VB.Label Label7 
      Caption         =   "（請輸入民國年）"
      Height          =   330
      Left            =   3360
      TabIndex        =   12
      Top             =   1290
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "顯示方式："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   330
      TabIndex        =   11
      Top             =   2310
      Width           =   900
   End
   Begin VB.Label Label5 
      Caption         =   "（1：螢幕 2：報表）"
      Height          =   330
      Left            =   1635
      TabIndex        =   10
      Top             =   2310
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "商品組群："
      Height          =   300
      Left            =   336
      TabIndex        =   9
      Top             =   768
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "查覆日期："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   336
      TabIndex        =   8
      Top             =   1296
      Width           =   912
   End
   Begin VB.Label Label3 
      Caption         =   "－"
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   1380
      Width           =   270
   End
End
Attribute VB_Name = "frm090109_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/11 Form2.0已檢查 (無需修改的物件); Printer列印未改--'Memo by Lydia 2024/11/14 Printer逐字檢查Unicode文字改以圖片方式列印
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer
Dim iPrint As Integer, Page As Integer, PLeft(0 To 4) As Integer, SubTotal(1 To 4) As Variant
Dim strSql As String, strTemp(0 To 4) As String
Dim StrArray As Variant, ComArray As Variant
Dim Rs As New ADODB.Recordset
Dim BlnCheck As Boolean
Dim Xo As Long, Yo As Long 'Added by Lydia 2024/11/14
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
   'Add By Cheng 2002/03/26
   If Me.Txtdata(4).Text < "1" Or Me.Txtdata(4).Text > "4" Then
      MsgBox "排序方式輸入錯誤!!!", vbExclamation
      Me.Txtdata(4).SetFocus
      Txtdata_GotFocus 4
      Exit Sub
   End If
   
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
   Set frm090109_2 = Nothing
End Sub

Private Sub Form_Load()
   Me.Height = 3705
   Me.Width = 7470
   MoveFormToCenter Me
   BlnCheck = False: bolToEndByNick = False
   Txtdata(3) = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm090109_1 = Nothing
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
   Case 0
      If Check_Txtdata0 Then
         Txtdata(0).SetFocus
         BlnCheck = True
      End If
   Case 2
      If Txtdata(1) = Empty And Txtdata(2) = Empty Then
         s = MsgBox("請輸入查覆日期條件", , "使用者輸入錯誤")
         Txtdata(1).SetFocus
         BlnCheck = True
         Exit Sub
      End If
      'Modify by Morgan 2010/8/16 百年蟲
      'If Txtdata(1) > Txtdata(2) Then
      If Val(Txtdata(1)) > Val(Txtdata(2)) Then
         s = MsgBox("查覆日期範圍錯誤", , "使用者輸入錯誤")
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
'檢查委查組群是否有重覆
Private Function Check_Txtdata0() As Boolean
   StrArray = ""
   If Len(Txtdata(0)) <> 0 Then
      StrArray = Split(Txtdata(0), ",")
      For i = 0 To UBound(StrArray)
         For j = i + 1 To UBound(StrArray)
            If StrArray(i) = StrArray(j) Then
               Check_Txtdata0 = True
               MsgBox "委查組群重覆輸入，請查明再輸!", vbCritical
               Exit Function
            End If
         Next j
      Next i
   End If
End Function

Private Sub Query_Sub()
   frm090109_2.Show
   frm090109_2.Hide
   frm090109_2.MousePointer = vbHourglass
   frm090109_2.GridData
   frm090109_2.MousePointer = vbDefault
   If frm090109_2.Enabled = True Then
      frm090109_2.Show
   Else
      s = MsgBox("資料庫中沒有符合的資料!!", , "請檢查條件")
   End If
   Do
      DoEvents
      If bolToEndByNick = True Then
         cmdExit_Click
         Exit Sub
      End If
   Loop Until Not frm090109_2.Visible
   Unload frm090109_2
End Sub

Private Sub Print_Sub()
Dim SubSQL As String
   Printer.Orientation = 2
   Screen.MousePointer = vbHourglass
   
   'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)
   If DBDATE(Txtdata(1)) >= 查名單網中系統啟用日 Or DBDATE(Txtdata(2)) >= 查名單網中系統啟用日 Then
      bolIsTMA = True
   Else
      bolIsTMA = False
   End If
   'end 2024//11/15
   
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   cnnConnection.Execute "DELETE FROM R090109 WHERE ID='" & strUserNum & "' "
   SubSQL = ""
   For i = 1 To 4
      SubTotal(i) = 0
   Next i
   If Txtdata(1) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TMA14>=" & Val(ChangeTStringToWString(Txtdata(1))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " AND TMQ11>=" & Val(ChangeTStringToWString(Txtdata(1))) & ""
      End If
   End If
   If Txtdata(2) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         SubSQL = SubSQL + " AND TMA14<=" & Val(ChangeTStringToWString(Txtdata(2))) & ""
      Else
      'end 2024/11/15
         SubSQL = SubSQL + " TMQ11<=" & Val(ChangeTStringToWString(Txtdata(2))) & ""
      End If
   End If
   If Txtdata(1) <> Empty Or Txtdata(2) <> Empty Then
      pub_QL05 = pub_QL05 & ";" & Label2 & Txtdata(1) & "-" & Txtdata(2) 'Add By Sindy 2010/12/14
   End If
   If Txtdata(0) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1 & Txtdata(0) 'Add By Sindy 2010/12/14
   End If
   If Txtdata(4) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label8 & "1：群組" 'Add By Sindy 2010/12/14
   ElseIf Txtdata(4) = "2" Then
      pub_QL05 = pub_QL05 & ";" & Label8 & "2：中文筆數" 'Add By Sindy 2010/12/14
   ElseIf Txtdata(4) = "3" Then
      pub_QL05 = pub_QL05 & ";" & Label8 & "3：英文筆數" 'Add By Sindy 2010/12/14
   Else
      pub_QL05 = pub_QL05 & ";" & Label8 & "4：圖形筆數" 'Add By Sindy 2010/12/14
   End If
   
   If Len(SubSQL) <> 0 Then
      SubSQL = " WHERE " & Mid(SubSQL, 5)
   End If
   'Added by Lydia 2024/11/15 查名單(網中)：排除1120904-1120928期間資料匯入＞＞TO_CHAR(TMA04,'YYYYMMDD')>='20240601'
   If bolIsTMA = True Then
      strSql = "SELECT " & PUB_GetTMAforClass & " AS 類別組群, SUM(NVL(TMA36, 0)) AS 中文, SUM(NVL(TMA37, 0)) AS 英文, SUM(NVL(TMA38, 0)) AS 圖形 " & _
               "FROM TMQAPPFORM " & SubSQL & " AND TO_CHAR(TMA04,'YYYYMMDD')>='20240601' GROUP BY " & PUB_GetTMAforClass & " ORDER  BY 1"
   Else
   'end 2024/11/15
      strSql = "SELECT TMQ03, NVL(TMQ07, 0), NVL(TMQ08, 0), NVL(TMQ09, 0) FROM TRADEMARKQUERY " & SubSQL & " ORDER BY TMQ03"
   End If
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Rs.RecordCount > 0 Then
      With Rs
         .MoveFirst
         j = 0
         ''FRM100.Show
         DoEvents
'         'FRM100.Tag = Trim(str(.RecordCount)) & "=0"
         Do While .EOF = False
            For i = 0 To 3
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            Insert_Temp
            j = j + 1
'            'FRM100.Tag = Trim(str(.RecordCount)) & "=" & Trim(str(j))
 '           'FRM100.StrMenu
            DoEvents
            .MoveNext
         Loop
      End With
      PrintProcess
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/14
      s = MsgBox("資料庫中沒有符合的資料!!", , "沒有資料")
   End If
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   Screen.MousePointer = vbDefault
End Sub
'寫入暫存檔
Private Sub Insert_Temp()
Dim BlnIsNew As Boolean, tmpArray As Variant
   'Modify By Cheng 2002/03/26
'   ComArray = Split(strTemp(0), ",")
   ComArray = Split(strTemp(0), ".")
   For i = 0 To UBound(ComArray)
      BlnIsNew = False
      If Len(Txtdata(0)) <> 0 Then
         tmpArray = Filter(StrArray, ComArray(i))
         For j = 0 To UBound(tmpArray)
            If tmpArray(j) <> Empty Then
               BlnIsNew = True
               j = UBound(tmpArray) + 1
            End If
        Next j
      Else
         BlnIsNew = True
      End If
      If BlnIsNew Then
         strSql = "INSERT INTO R090109 VALUES('" & ComArray(i) & "'," & Val(strTemp(1)) & "," & Val(strTemp(2)) & "," & Val(strTemp(3)) & ",'" & strUserNum & "')"
         cnnConnection.Execute strSql
      End If
   Next i
End Sub

Private Sub PrintProcess()
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   'Modify By Cheng 2002/03/26
'   strSQL = "SELECT R001001, SUM(NVL(R001002, 0)), SUM(NVL(R001003, 0)), SUM(NVL(R001004, 0)), SUM(NVL(R001002, 0) + NVL(R001003, 0) + NVL(R001004, 0)) FROM R090109 WHERE ID='" & strUserNum & "' GROUP BY R001001 ORDER BY R001001"
   '以組群排序
   If Me.Txtdata(4).Text = "1" Then
      strSql = "SELECT R001001 AS 組群, SUM(NVL(R001002, 0)) AS 中文筆數, SUM(NVL(R001003, 0)) AS 英文筆數, SUM(NVL(R001004, 0)) AS 圖形筆數, SUM(NVL(R001002, 0) + NVL(R001003, 0) + NVL(R001004, 0)) FROM R090109 WHERE ID='" & strUserNum & "' GROUP BY R001001 ORDER BY 組群"
   '以中文筆數排序
   ElseIf Me.Txtdata(4).Text = "2" Then
      strSql = "SELECT R001001 AS 組群, SUM(NVL(R001002, 0)) AS 中文筆數, SUM(NVL(R001003, 0)) AS 英文筆數, SUM(NVL(R001004, 0)) AS 圖形筆數, SUM(NVL(R001002, 0) + NVL(R001003, 0) + NVL(R001004, 0)) FROM R090109 WHERE ID='" & strUserNum & "' GROUP BY R001001 ORDER BY 中文筆數 desc"
   '以英文筆數排序
   ElseIf Me.Txtdata(4).Text = "3" Then
      strSql = "SELECT R001001 AS 組群, SUM(NVL(R001002, 0)) AS 中文筆數, SUM(NVL(R001003, 0)) AS 英文筆數, SUM(NVL(R001004, 0)) AS 圖形筆數, SUM(NVL(R001002, 0) + NVL(R001003, 0) + NVL(R001004, 0)) FROM R090109 WHERE ID='" & strUserNum & "' GROUP BY R001001 ORDER BY 英文筆數 desc"
   '以圖形筆數排序
   Else
      strSql = "SELECT R001001 AS 組群, SUM(NVL(R001002, 0)) AS 中文筆數, SUM(NVL(R001003, 0)) AS 英文筆數, SUM(NVL(R001004, 0)) AS 圖形筆數, SUM(NVL(R001002, 0) + NVL(R001003, 0) + NVL(R001004, 0)) FROM R090109 WHERE ID='" & strUserNum & "' GROUP BY R001001 ORDER BY 圖形筆數 desc"
   End If
   
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Rs.RecordCount > 0 Then
      InsertQueryLog (Rs.RecordCount) 'Add By Sindy 2010/12/14
      cnnConnection.Execute "DELETE FROM R090109 WHERE ID='" & strUserNum & "' "
      With Rs
        .MoveFirst
         Page = 1
         PrintTitle
         j = 0
'         'FRM100.Show
         DoEvents
         ''FRM100.Tag = Trim(str(.RecordCount)) & "=0"
         Do While .EOF = False
            For i = 0 To 4
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            For i = 1 To 4
               SubTotal(i) = SubTotal(i) + CDec(.Fields(i))
            Next i
            PrintDetail
            If iPrint > 14000 Then
               Page = Page + 1
               Printer.CurrentX = 500
               Printer.CurrentY = iPrint
               Printer.Print String(200, "-")
               Printer.NewPage
               PrintTitle
            End If
            j = j + 1
'            'FRM100.Tag = Trim(str(.RecordCount)) & "=" & Trim(str(j))
 '           'FRM100.StrMenu
            DoEvents
            .MoveNext
         Loop
         PrintTotal
         Printer.EndDoc
         s = MsgBox("列印完成!!", , "列印成功")
         'Unload frm100
      End With
   Else
      InsertQueryLog (0) 'Add By Sindy 2010/12/14
      s = MsgBox("資料庫中沒有符合的資料!!", , "沒有資料")
   End If
End Sub

Private Sub PrintTitle() '列印抬頭
   Printer.Orientation = 1
   GetPleft
   iPrint = 500
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4000
   Printer.CurrentY = iPrint
   Printer.Print "商品組群委查統計表"
   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   'Modified by Lydia 2024/11/15
   'Printer.CurrentX = 300
   'Printer.CurrentY = iPrint
   'Printer.Print "列印人　：" & strUserName
   Xo = 300
   Yo = iPrint
   PUB_PrintUnicodeText "列印人　：" & strUserName, Xo, Yo, 0
   'end 2024/11/15
   Printer.CurrentX = 4400
   Printer.CurrentY = iPrint
   Printer.Print "查覆日期：" & ChangeTStringToTDateString(Txtdata(1))
   Printer.CurrentX = 6400
   Printer.CurrentY = iPrint
   Printer.Print "－"
   Printer.CurrentX = 6650
   Printer.CurrentY = iPrint
   Printer.Print ChangeTStringToTDateString(Txtdata(2))
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   'Add By Cheng 2002/03/26
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print "排序方式：" & IIf(Me.Txtdata(4).Text = "1", "商品組群", IIf(Me.Txtdata(4).Text = "2", "中文筆數", IIf(Me.Txtdata(4).Text = "3", "英文筆數", "圖形筆數")))
   
   Printer.CurrentX = 9200
   Printer.CurrentY = iPrint
   Printer.Print "頁　　次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(146, "-")
   iPrint = iPrint + 300
   Printer.Font.Underline = True
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   'Added by Lydia 2024/11/15 查名單(網中)
   If bolIsTMA = True Then
      Printer.Print "商品類別組群"
   Else
   'end 2024/11/15
      Printer.Print "商品組群"
   End If
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "中　文"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "英　文"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "圖　形"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "小　計"
   Printer.Font.Underline = False
   iPrint = iPrint + 300
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(146, "-")
   iPrint = iPrint + 300
End Sub

Private Sub PrintDetail()
   For i = 0 To 0
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(i)
   Next i
   For i = 1 To 4
      Printer.CurrentX = PLeft(i) + 700 - Printer.TextWidth(Format(strTemp(i), "###,###,###,###"))
      Printer.CurrentY = iPrint
      Printer.Print Format(strTemp(i), "###,###,###,###")
   Next i
   iPrint = iPrint + 300
End Sub

Private Sub PrintTotal()
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(146, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 900
   Printer.CurrentY = iPrint
   Printer.Print "總　計："
   For i = 1 To 4
      Printer.CurrentX = PLeft(i) + 700 - Printer.TextWidth(Format(SubTotal(i), "###,###,###,###"))
      Printer.CurrentY = iPrint
      Printer.Print Format(SubTotal(i), "###,###,###,###")
   Next i
   iPrint = iPrint + 300
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(85, "=")
End Sub

Private Sub GetPleft()
   Erase PLeft
   PLeft(0) = 900
   PLeft(1) = 3100
   PLeft(2) = 5300
   PLeft(3) = 7500
   PLeft(4) = 9700
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '查覆日期起, 迄
   If PUB_CheckKeyInDate(Me.Txtdata(Index)) = -1 Then
      Cancel = True
      Me.Txtdata(Index).SetFocus
      Txtdata_GotFocus Index
   End If
Case 4 '排序
   If Me.Txtdata(Index).Text < "1" Or Me.Txtdata(Index).Text > "4" Then
      MsgBox "排序方式輸入錯誤!!!", vbExclamation
      Cancel = True
      Me.Txtdata(Index).SetFocus
      Txtdata_GotFocus Index
   End If
End Select
End Sub
