VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090107_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "查名人查覆統計表"
   ClientHeight    =   3324
   ClientLeft      =   996
   ClientTop       =   1836
   ClientWidth     =   6336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3324
   ScaleWidth      =   6336
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   3
      Left            =   1668
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1992
      Width           =   255
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   1
      Left            =   1668
      MaxLength       =   7
      TabIndex        =   1
      Top             =   1440
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   2
      Left            =   2832
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1440
      Width           =   825
   End
   Begin VB.TextBox Txtdata 
      Height          =   300
      Index           =   0
      Left            =   1668
      MaxLength       =   6
      TabIndex        =   0
      Top             =   912
      Width           =   825
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   4428
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   5256
      TabIndex        =   5
      Top             =   120
      Width           =   800
   End
   Begin MSForms.Label LblTmq10NM 
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   930
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
      Left            =   3840
      TabIndex        =   11
      Top             =   1422
      Width           =   1692
   End
   Begin VB.Label Label6 
      Caption         =   "顯示方式："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   744
      TabIndex        =   10
      Top             =   2004
      Width           =   900
   End
   Begin VB.Label Label5 
      Caption         =   "（1：螢幕 2：報表）"
      Height          =   336
      Left            =   2040
      TabIndex        =   9
      Top             =   2004
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "查名人："
      Height          =   300
      Left            =   744
      TabIndex        =   8
      Top             =   936
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "查覆日期："
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   744
      TabIndex        =   7
      Top             =   1464
      Width           =   912
   End
   Begin VB.Label Label3 
      Caption         =   "－"
      Height          =   288
      Left            =   2580
      TabIndex        =   6
      Top             =   1500
      Width           =   276
   End
End
Attribute VB_Name = "frm090107_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/11 改成Form2.0 ; LblTmq10NM ; Printer列印未改--'Memo by Lydia 2024/11/15 Printer逐字檢查Unicode文字改以圖片方式列印
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim s As Integer, i As Integer, j As Integer
Dim iPrint As Integer, Page As Integer, PLeft(0 To 4) As Integer, SubTotal(1 To 4) As Integer
Dim strSql As String, strTemp(0 To 4) As String
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
   Set frm090107_2 = Nothing
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
   Set frm090107_1 = Nothing
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
            LblTmq10NM.Caption = strTemp
         Else
            LblTmq10NM.Caption = ""
            Txtdata(0).SetFocus
            TextInverse Txtdata(0)
            BlnCheck = True
         Exit Sub
         End If
      Else
         LblTmq10NM.Caption = ""
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

Private Sub Query_Sub()
   frm090107_2.Show
   frm090107_2.Hide
   frm090107_2.MousePointer = vbHourglass
   frm090107_2.GridData
   frm090107_2.MousePointer = vbDefault
   If frm090107_2.Enabled = True Then
      frm090107_2.Show
   Else
      s = MsgBox("資料庫中沒有符合的資料!!", , "請檢查條件")
   End If
   Do
      DoEvents
      If bolToEndByNick = True Then
         cmdExit_Click
         Exit Sub
      End If
   Loop Until Not frm090107_2.Visible
   Unload frm090107_2
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
   For i = 1 To 4
      SubTotal(i) = 0
   Next i
   If Txtdata(0) <> Empty Then
      'Added by Lydia 2024/11/15 查名單(網中)
      If bolIsTMA = True Then
         strCondition = strCondition + " AND TMA10 = '" & Txtdata(0) & "'"
      Else
      'end 2024/11/15
         strCondition = strCondition + " AND TMQ10 = '" & Txtdata(0) & "'"
      End If
      pub_QL05 = pub_QL05 & ";" & Label1 & Txtdata(0) & LblTmq10NM 'Add By Sindy 2010/12/14
   End If
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
         SubSQL = SubSQL + " AND TMQ11<=" & Val(ChangeTStringToWString(Txtdata(2))) & ""
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
      strSql = "SELECT NVL(ST02, TMA10) AS 查名人, SUM(NVL(TMA36, 0)) AS 中文, SUM(NVL(TMA37, 0)) AS 英文, SUM(NVL(TMA38, 0)) AS 圖形, SUM(NVL(TMA36, 0) + NVL(TMA37, 0) + NVL(TMA38, 0)) AS 小計, TMA10 " & _
               "FROM TMQAPPFORM, STAFF " & SubSQL & strCondition & " AND TMA10 = ST01(+) AND TO_CHAR(TMA04,'YYYYMMDD')>='20240601' GROUP BY TMA10, NVL(ST02, TMA10)"
   Else
   'end 2024/11/15
      strSql = "SELECT NVL(ST02, TMQ10), SUM(NVL(TMQ07, 0)), SUM(NVL(TMQ08, 0)), SUM(NVL(TMQ09, 0)), SUM(NVL(TMQ07, 0) + NVL(TMQ08, 0) + NVL(TMQ09, 0)), TMQ10 FROM TRADEMARKQUERY, STAFF " & SubSQL & strCondition & " AND TMQ10 = ST01(+) GROUP BY TMQ10, NVL(ST02, TMQ10)"
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
         DoEvents
         Do While .EOF = False
            For i = 0 To 4
               strTemp(i) = CheckStr(.Fields(i))
            Next i
            For i = 1 To 4
               SubTotal(i) = SubTotal(i) + CDec(.Fields(i))
            Next i
            PrintDetail
            If iPrint > 15000 Then
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
         PrintTotal
         Printer.EndDoc
         s = MsgBox("列印完成!!", , "列印成功")
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
   Printer.Print "查名人查覆統計表"
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
   Printer.Print "查名人"
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
      'Modified by Lydia 2024/11/15 逐字檢查Unicode文字改以圖片方式列印
      'Printer.CurrentX = PLeft(i)
      'Printer.CurrentY = iPrint
      'Printer.Print strTemp(i)
      Xo = PLeft(i)
      Yo = iPrint
      PUB_PrintUnicodeText strTemp(i), Xo, Yo, 0
      'end 2024/11/15
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
End Select
End Sub
