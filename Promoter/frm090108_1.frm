VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090108_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "委查人委查統計表"
   ClientHeight    =   3672
   ClientLeft      =   1776
   ClientTop       =   1056
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3672
   ScaleWidth      =   6360
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   5
      Left            =   2490
      MaxLength       =   7
      TabIndex        =   1
      Top             =   792
      Width           =   500
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   4
      Left            =   1572
      MaxLength       =   7
      TabIndex        =   0
      Top             =   792
      Width           =   500
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   3
      Left            =   1572
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2376
      Width           =   255
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   1
      Left            =   1572
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1824
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   2
      Left            =   2736
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1824
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   0
      Left            =   1584
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1272
      Width           =   825
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4416
      TabIndex        =   6
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   5256
      TabIndex        =   7
      Top             =   120
      Width           =   800
   End
   Begin MSForms.Label LblTmq02NM 
      Height          =   255
      Left            =   2490
      TabIndex        =   16
      Top             =   1290
      Width           =   1875
      VariousPropertyBits=   27
      Size            =   "3307;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      Caption         =   "－"
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   810
      Width           =   270
   End
   Begin VB.Label Label4 
      Caption         =   "業務區："
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   648
      TabIndex        =   14
      Top             =   792
      Width           =   912
   End
   Begin VB.Label Label7 
      Caption         =   "（請輸入民國年）"
      Height          =   330
      Left            =   3660
      TabIndex        =   13
      Top             =   1830
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "顯示方式："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   648
      TabIndex        =   12
      Top             =   2364
      Width           =   900
   End
   Begin VB.Label Label5 
      Caption         =   "（1：螢幕 2：報表）"
      Height          =   330
      Left            =   1920
      TabIndex        =   11
      Top             =   2370
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "委查人："
      Height          =   300
      Left            =   648
      TabIndex        =   10
      Top             =   1296
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "委查日期："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   648
      TabIndex        =   9
      Top             =   1824
      Width           =   912
   End
   Begin VB.Label Label3 
      Caption         =   "－"
      Height          =   288
      Left            =   2484
      TabIndex        =   8
      Top             =   1860
      Width           =   276
   End
End
Attribute VB_Name = "frm090108_1"
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
Dim iPrint As Integer, Page As Integer, PLeft(0 To 5) As Integer
'Modify By Cheng 2002/08/07
'Dim SubTotal(2 To 5) As Integer, GradeTotal(2 To 5) As Integer
Dim SubTotal(2 To 5) As Integer, SubTotalA(2 To 5) As Integer, GradeTotal(2 To 5) As Integer
'Modify By Cheng 2002/08/07
'Dim strsql As String, strTemp(0 To 7) As String, TmpKey1 As String
Dim strSql As String, strTemp(0 To 7) As String, TmpKey1 As String, TmpKey1A As String
Dim Rs As New ADODB.Recordset
Dim BlnCheck As Boolean
Dim Xo As Long, Yo As Long 'Added by Lydia 2024/11/15
Dim bolIsTMA As Boolean 'Added by Lydia 2024/11/15 判斷日期條件，資料改抓查名單(網中)

Private Sub cmdOK_Click()
   Txtdata_LostFocus (5)
   If BlnCheck Then Exit Sub
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
   Set frm090108_2 = Nothing
End Sub

Private Sub Form_Load()
   Me.Height = 4065
   Me.Width = 7470
   MoveFormToCenter Me
   BlnCheck = False: bolToEndByNick = False
   Txtdata(3) = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm090108_1 = Nothing
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
   Case 5
      If Txtdata(4) > Txtdata(5) Then
         s = MsgBox("業務區範圍錯誤", , "使用者輸入錯誤")
         Txtdata(4).SetFocus
         TextInverse Txtdata(4)
         BlnCheck = True
         Exit Sub
      End If
   Case Else
   End Select
End Sub

Private Sub Query_Sub()
   frm090108_2.Show
   frm090108_2.Hide
   frm090108_2.MousePointer = vbHourglass
   frm090108_2.GridData
   frm090108_2.MousePointer = vbDefault
   If frm090108_2.Enabled = True Then
      frm090108_2.Show
   Else
      s = MsgBox("資料庫中沒有符合的資料!!", , "請檢查條件")
   End If
   Do
      DoEvents
      If bolToEndByNick = True Then
         cmdExit_Click
         Exit Sub
      End If
   Loop Until Not frm090108_2.Visible
   Unload frm090108_2
End Sub

Private Sub Print_Sub()
Dim SubSQL As String, strCondition As String
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
   SubSQL = "": strCondition = ""
   For i = 2 To 5
      SubTotal(i) = 0: GradeTotal(i) = 0
      'Add By Cheng 2002/08/07
      SubTotalA(i) = 0
   Next i
   If Txtdata(4) <> Empty Then
      strCondition = strCondition + " AND ST03>='" & Txtdata(4) & "'"
   End If
   If Txtdata(5) <> Empty Then
      strCondition = strCondition + " AND ST03<='" & Txtdata(5) & "'"
   End If
   If Txtdata(4) <> Empty Or Txtdata(5) <> Empty Then
      pub_QL05 = pub_QL05 & ";" & Label4 & Txtdata(4) & "-" & Txtdata(5) 'Add By Sindy 2010/12/14
   End If
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
      strSql = "SELECT DECODE(A0902, NULL, ' ', NVL(A0902, ST15)) AS 部門,NVL(ST02, TMA08) AS 委查人, SUM(NVL(TMA36, 0)) AS 中文, SUM(NVL(TMA37, 0)) AS 英文, SUM(NVL(TMA38, 0)) AS 圖形, SUM(NVL(TMA36, 0) + NVL(TMA37, 0) + NVL(TMA38, 0)) AS 小計, NVL(ST15, '   ') AS ST15, TMA08 " & _
               "FROM TMQAPPFORM, STAFF, ACC090 " & SubSQL & strCondition & " AND TMA08 = ST01(+) AND ST15 = A0901(+) AND TO_CHAR(TMA04,'YYYYMMDD')>='20240601' GROUP BY ST15, TMA08, DECODE(A0902, NULL, ' ', NVL(A0902, ST15)), NVL(ST02, TMA08)"
   Else
   'end 2024/11/15
      'Modify By Cheng 2002/08/07
   '   strsql = "SELECT DECODE(A0902, NULL, ' ', NVL(A0902, ST03)) AS 部門, NVL(ST02, TMQ02) AS 委查人, SUM(NVL(TMQ07, 0)) AS 中文, SUM(NVL(TMQ08, 0)) AS 英文, SUM(NVL(TMQ09, 0)) AS 圖形, SUM(NVL(TMQ07, 0) + NVL(TMQ08, 0) + NVL(TMQ09, 0)), NVL(ST03, ' '), TMQ02 FROM TRADEMARKQUERY, STAFF, ACC090 " & SubSQL & strCondition & " AND TMQ02 = ST01(+) AND ST03 = A0901(+) GROUP BY ST03, TMQ02, DECODE(A0902, NULL, ' ', NVL(A0902, ST03)), NVL(ST02, TMQ02)"
      strSql = "SELECT DECODE(A0902, NULL, ' ', NVL(A0902, ST15)) AS 部門, NVL(ST02, TMQ02) AS 委查人, SUM(NVL(TMQ07, 0)) AS 中文, SUM(NVL(TMQ08, 0)) AS 英文, SUM(NVL(TMQ09, 0)) AS 圖形, SUM(NVL(TMQ07, 0) + NVL(TMQ08, 0) + NVL(TMQ09, 0)), NVL(ST15, '   '), TMQ02 FROM TRADEMARKQUERY, STAFF, ACC090 " & SubSQL & strCondition & " AND TMQ02 = ST01(+) AND ST15 = A0901(+) GROUP BY ST15, TMQ02, DECODE(A0902, NULL, ' ', NVL(A0902, ST15)), NVL(ST02, TMQ02)"
   End If
   Rs.CursorLocation = adUseClient
   Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If Rs.RecordCount > 0 Then
      InsertQueryLog (Rs.RecordCount) 'Add By Sindy 2010/12/14
      TmpKey1 = ""
      'Add By Cheng 2002/08/07
      TmpKey1A = ""
      With Rs
         .MoveFirst
         Page = 1
         PrintTitle
         j = 0
'         'FRM100.Show
         DoEvents
 '        'FRM100.Tag = Trim(str(.RecordCount)) & "=0"
         Do While .EOF = False
            For i = 0 To 7
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(1) = StrConv(MidB(StrConv(strTemp(1), vbFromUnicode), 1, 10), vbUnicode)
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
   If Rs.State <> adStateClosed Then
      Rs.Close
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub PrintTitle() '列印抬頭
   Printer.Orientation = 1
   GetPleft
   iPrint = 500
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 4200
   Printer.CurrentY = iPrint
   Printer.Print "委查人委查統計表"
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
   Printer.Print "委查日期：" & ChangeTStringToTDateString(Txtdata(1))
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
   Printer.Print "部　門"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "委查人"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "中　文"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "英　文"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "圖　形"
   Printer.CurrentX = PLeft(5)
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
   If TmpKey1 <> Empty And TmpKey1 <> strTemp(6) Then
      PrintSubTotal
      'Add By Cheng 2002
      If TmpKey1A <> Left(strTemp(6), 2) Then
         PrintSubTotalA
      End If
   Else
      If TmpKey1 <> Empty Then strTemp(0) = ""
   End If
   For i = 0 To 1
      'Added by Lydia 2024/11/15 逐字檢查Unicode文字改以圖片方式列印
      If i = 1 Then
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
   For i = 2 To 5
      SubTotal(i) = SubTotal(i) + CDec(strTemp(i))
      'Add By Cheng 2002/08/07
      SubTotalA(i) = SubTotalA(i) + CDec(strTemp(i))
      GradeTotal(i) = GradeTotal(i) + CDec(strTemp(i))
      Printer.CurrentX = PLeft(i) + 700 - Printer.TextWidth(Format(strTemp(i), "###,###,###,###"))
      Printer.CurrentY = iPrint
      Printer.Print Format(strTemp(i), "###,###,###,###")
   Next i
   iPrint = iPrint + 300
   TmpKey1 = strTemp(6)
   'Add By Cheng 2002/08/07
   TmpKey1A = Left(strTemp(6), 2)
End Sub

Private Sub PrintSubTotal()
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(146, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 3000
   Printer.CurrentY = iPrint
   Printer.Print "小　計："
   For i = 2 To 5
      Printer.CurrentX = PLeft(i) + 700 - Printer.TextWidth(Format(SubTotal(i), "###,###,###,###"))
      Printer.CurrentY = iPrint
      Printer.Print Format(SubTotal(i), "###,###,###,###")
   Next i
   iPrint = iPrint + 300
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(146, " ")
   iPrint = iPrint + 300
   For i = 2 To 5
      SubTotal(i) = 0
   Next i
End Sub

'Add By Cheng 2002/08/07
Private Sub PrintSubTotalA()
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(146, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 3000
   Printer.CurrentY = iPrint
   Printer.Print "合　計："
   For i = 2 To 5
      Printer.CurrentX = PLeft(i) + 700 - Printer.TextWidth(Format(SubTotalA(i), "###,###,###,###"))
      Printer.CurrentY = iPrint
      Printer.Print Format(SubTotalA(i), "###,###,###,###")
   Next i
   iPrint = iPrint + 300
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(146, " ")
   iPrint = iPrint + 300
   For i = 2 To 5
      SubTotalA(i) = 0
   Next i
End Sub

Private Sub PrintTotal()
   PrintSubTotal
   'Add By Cheng 2002/08/07
   PrintSubTotalA
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(146, "-")
   iPrint = iPrint + 300
   Printer.CurrentX = 3000
   Printer.CurrentY = iPrint
   Printer.Print "總　計："
   For i = 2 To 5
      Printer.CurrentX = PLeft(i) + 700 - Printer.TextWidth(Format(GradeTotal(i), "###,###,###,###"))
      Printer.CurrentY = iPrint
      Printer.Print Format(GradeTotal(i), "###,###,###,###")
   Next i
   iPrint = iPrint + 300
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(85, "=")
End Sub

Private Sub GetPleft()
   Erase PLeft
   PLeft(0) = 300
   PLeft(1) = 2250
   PLeft(2) = 4250
   PLeft(3) = 6250
   PLeft(4) = 8250
   PLeft(5) = 10250
End Sub

Private Sub Txtdata_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '委查日期起, 迄
   If PUB_CheckKeyInDate(Me.Txtdata(Index)) = -1 Then
      Cancel = True
      Me.Txtdata(Index).SetFocus
      Txtdata_GotFocus Index
   End If
End Select
End Sub
